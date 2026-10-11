#Requires -Version 5.1
<#
  屏幕划词翻译（英文 → 中文）
  ────────────────────────────────────────────────────────────
  常驻后台。按全局热键 → 鼠标框选屏幕上任意区域 →
  自带 OCR 识别文字 → 翻译成中文 → 在选区旁边弹卡片显示对照。

  为什么是 PowerShell 5.1：
    OCR 用的是 Windows 系统自带的 WinRT 引擎（Windows.Media.Ocr），
    零安装、离线、免费。但 WinRT 的类型投影只在【Windows PowerShell 5.1】
    里可用，PowerShell 7 / Core 会直接报 "Unable to find type"。
    所以本脚本必须用 powershell.exe 跑，不能用 pwsh。

  启动：双击同目录下的 "启动 屏幕翻译.lnk"
  退出：右键托盘图标 → 退出
#>
param(
    [switch]$KeepConsole,     # 保留控制台窗口，排查问题时用
    [switch]$SelfTest,        # 自检：只验证 OCR 和翻译接口是否可用，然后退出
    [switch]$Demo,            # 演示：直接弹一张样例卡片并截图，用于核对排版
    [switch]$OcrShot,         # 调试：只截指定区域跑一次 OCR，结果写进 .ocr.txt
    [int]$RectX = 0, [int]$RectY = 0, [int]$RectW = 800, [int]$RectH = 300
)

$ErrorActionPreference = 'Stop'

# ══ 配置区 ═══════════════════════════════════════════════════
$HotkeyMods   = 0x0002 -bor 0x0001 -bor 0x4000   # Ctrl + Alt + (不重复触发)
$HotkeyKey    = 0x5A                             # Z  →  即 Ctrl+Alt+Z
$HotkeyName   = 'Ctrl+Alt+Z'

# 下面这几个尺寸写的是「100% 缩放下的基准值」，实际生效值在 DPI 系数算出来后再乘
$BaseCardW        = 460      # 结果卡片宽度
$BaseCardMaxBody  = 360      # 卡片正文最大高度，超出则内部滚动
$BaseMinSelectPx  = 8        # 小于这个尺寸的选框视为误触，直接忽略
$TranslateTimeout = 6        # 翻译接口超时（秒）。正常翻译 0.1~0.8 秒；给 15 秒的话
                             # 断网时要依次干等三个接口超时，用户盯着「翻译中…」等一分多钟
$ChunkSize        = 450      # 单次请求的最大字符数。
                             # 这个值是被备用接口 MyMemory 的 500 字符硬上限卡出来的；
                             # 有道实测 2000 字符会报错、6400 字符直接 URI 过长，
                             # 取 450 三个接口都能吃下，也不容易触发限流。
# ═════════════════════════════════════════════════════════════

# ══ 1. DPI 感知（必须在加载 WinForms / 创建任何窗口之前） ═════
# 本机是 2560x1600 @150% 缩放。不设这一句的话，进程是 DPI-unaware，
# 系统会把截图缩到 66% 再给我们，OCR 质量明显下降。
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class StDpi {
    [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int v);
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int i);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern uint GetDpiForSystem();
}
'@
[void][StDpi]::SetProcessDpiAwareness(2)   # 2 = PerMonitorDpiAware

# 系统缩放系数。字体按「点」算会自动跟着 DPI 放大，
# 所以卡片的宽高、内边距、间距也必须一起乘这个系数，否则 150% 缩放下
# 字会比框大一圈（标签折行、按钮文字被截）；反过来度量文字的 Graphics
# 也得在同一 DPI 下量，不然算出的高度只有实际的三分之二。
$Scale = [double][StDpi]::GetDpiForSystem() / 96.0
if ($Scale -le 0 -or $Scale -gt 8) { $Scale = 1.0 }

$CardWidth   = [int][Math]::Round($BaseCardW * $Scale)
$CardMaxBody = [int][Math]::Round($BaseCardMaxBody * $Scale)
$MinSelectPx = [int][Math]::Round($BaseMinSelectPx * $Scale)

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
# 不加载这个程序集，下面就找不到 System.WindowsRuntimeSystemExtensions
Add-Type -AssemblyName System.Runtime.WindowsRuntime

if (-not $KeepConsole) {
    try {
        $h = [StDpi]::GetConsoleWindow()
        if ($h -ne [IntPtr]::Zero) { [void][StDpi]::ShowWindow($h, 0) }   # SW_HIDE
    }
    catch { }
}

# ══ 2. WinRT 类型 + 异步互操作 ═══════════════════════════════
# 首次引用必须带 ContentType=WindowsRuntime 全名，之后短名才可用
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]

# IAsyncOperation<T> → .NET Task 的泛型 AsTask（反射拿，没有更简单的写法）
$AsTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and
        $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    })[0]

function Await($op, $type) {
    $t = $AsTaskGeneric.MakeGenericMethod($type).Invoke($null, @($op))
    $t.Wait(-1) | Out-Null
    $t.Result
}

$script:OcrEngine = $null
function Get-OcrEngine {
    if (-not $script:OcrEngine) {
        $script:OcrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
        if (-not $script:OcrEngine) {
            $script:OcrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage(
                (New-Object Windows.Globalization.Language 'en-US'))
        }
    }
    return $script:OcrEngine
}

# ══ 3. 全局热键（C# 宿主一个 message-only 窗口收 WM_HOTKEY） ═══
Add-Type -ReferencedAssemblies 'System.Windows.Forms' -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class StHotkey : NativeWindow
{
    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public const int WM_HOTKEY = 0x0312;
    private const int ID = 0x5711;

    public event Action Fired;

    public StHotkey()
    {
        CreateParams cp = new CreateParams();
        cp.Caption = "StHotkeySink";
        cp.Parent = new IntPtr(-3);        // HWND_MESSAGE：message-only 窗口
        CreateHandle(cp);
    }

    public bool Hook(uint modifiers, uint vk) { return RegisterHotKey(Handle, ID, modifiers, vk); }
    public void Unhook() { try { UnregisterHotKey(Handle, ID); } catch { } }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == ID)
        {
            Action f = Fired;
            if (f != null) { f(); }
        }
        base.WndProc(ref m);
    }
}
'@

# ══ 4. 取色 / 字体 ═══════════════════════════════════════════
$C_Bg     = [System.Drawing.ColorTranslator]::FromHtml('#252526')
$C_Line   = [System.Drawing.ColorTranslator]::FromHtml('#3C3C3C')
$C_Title  = [System.Drawing.ColorTranslator]::FromHtml('#9A9A9A')
$C_Text   = [System.Drawing.ColorTranslator]::FromHtml('#E8E8E8')
$C_Dim    = [System.Drawing.ColorTranslator]::FromHtml('#B0B0B0')
$C_Accent = [System.Drawing.ColorTranslator]::FromHtml('#4EA1FF')
$C_Err    = [System.Drawing.ColorTranslator]::FromHtml('#F48771')
$C_Btn    = [System.Drawing.ColorTranslator]::FromHtml('#3A3D41')
$C_BtnHov = [System.Drawing.ColorTranslator]::FromHtml('#4A4E53')

function New-UiFont($size, $bold = $false) {
    $style = if ($bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    return New-Object System.Drawing.Font 'Microsoft YaHei UI', $size, $style
}

# 量一段文字在指定宽度下折行后的高度。
# 关键：这个 Graphics 的 DPI 必须跟真实屏幕一致，否则 150% 缩放下量出来的
# 高度只有实际的三分之二，标签框会矮一截、文字被切掉半行。
$script:MeasureBmp = New-Object System.Drawing.Bitmap 1, 1
$script:MeasureBmp.SetResolution([single](96.0 * $Scale), [single](96.0 * $Scale))
$script:MeasureG   = [System.Drawing.Graphics]::FromImage($script:MeasureBmp)
function Measure-TextHeight($text, $font, $width) {
    if ([string]::IsNullOrEmpty($text)) { return [int](18 * $Scale) }
    $sz = $script:MeasureG.MeasureString($text, $font, [int]$width)
    return [int][Math]::Ceiling($sz.Height)
}

# ══ 5. 截图 + OCR ════════════════════════════════════════════
function Capture-Region([int]$x, [int]$y, [int]$w, [int]$h) {
    $bmp = New-Object System.Drawing.Bitmap -ArgumentList $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size -ArgumentList $w, $h))
    $g.Dispose()
    return $bmp
}

# 重采样成指定倍数 + 指定 DPI 的新位图（OCR 前处理用）
function Convert-Bitmap($src, [double]$factor, [int]$dpi) {
    $nw = [int][Math]::Round($src.Width * $factor)
    $nh = [int][Math]::Round($src.Height * $factor)
    $dst = New-Object System.Drawing.Bitmap -ArgumentList $nw, $nh
    $dst.SetResolution([single]$dpi, [single]$dpi)
    $g = [System.Drawing.Graphics]::FromImage($dst)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.DrawImage($src, 0, 0, $nw, $nh)
    $g.Dispose()
    return $dst
}

# 识别。返回 @{ Result; Factor }。Factor 是喂给引擎前把图放大的倍数，
# 后面拿到词框坐标要除回去。
#
# 为什么要放大：WinRT 的 OCR 模型是按「大概 30px 高的正文」训练的，屏幕上的
# 英文（网页正文 12-16px，本机 150% 缩放后也就 18-24 物理像素）对它来说太小，
# 实测会整段丢字——同一张图，一行字只认出右半边、左半边凭空消失。放大 2 倍
# 之后这一行就完整认出来了。太大的区域（>100 万像素）不放大，避免一张全屏图
# 变成 400 万像素拖慢识别。
function Invoke-Ocr($bmp) {
    $factor = 1.0
    $src = $bmp
    if (($bmp.Width * $bmp.Height) -le 1000000) {
        $factor = 2.0
        $src = Convert-Bitmap $bmp 2.0 96
    }

    $ms = New-Object System.IO.MemoryStream
    $src.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)      # 全程内存，不落盘
    $ras = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($ms)
    $ras.Seek(0) | Out-Null

    $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($ras)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $sb = Await ($dec.GetSoftwareBitmapAsync([Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
            [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied)) ([Windows.Graphics.Imaging.SoftwareBitmap])

    $eng = Get-OcrEngine
    if (-not $eng) { throw '系统没有可用的 OCR 语言包（需要 en-US）' }
    $res = Await ($eng.RecognizeAsync($sb)) ([Windows.Media.Ocr.OcrResult])
    if ($factor -ne 1.0) { $src.Dispose() }
    return [pscustomobject]@{ Result = $res; Factor = $factor }
}

# 拼一行文字。不能直接信 $line.Text：引擎自己按字距分词的，遇到字距稍宽就
# 把一个词劈成两半，而 $line.Text 一律用空格把词接起来，于是 "proxy" 变成
# "p roxy"、"try" 变成 "t ry"，翻译出来就是"p - Roxy"这种鬼东西。
# 这里改成按词框之间的实际间距判断：间距够大才当空格。
function Join-OcrWords($words, [double]$factor) {
    $sb = New-Object System.Text.StringBuilder
    $prevRight = $null
    $lineH = 0.0
    foreach ($w in $words) {
        $b = $w.BoundingRect
        if ($b.Height -gt $lineH) { $lineH = $b.Height }
    }
    $spaceGap = [Math]::Max(3.0, $lineH * 0.28 / $factor)
    foreach ($w in $words) {
        $b = $w.BoundingRect
        if ($null -ne $prevRight) {
            if (($b.X - $prevRight) -gt $spaceGap) { [void]$sb.Append(' ') }
        }
        [void]$sb.Append($w.Text)
        $prevRight = $b.X + $b.Width
    }
    return $sb.ToString()
}

# 返回 [{Text, X, Y, W, H}]，坐标已换算成屏幕物理像素
function Get-OcrLines($ocr, [int]$originX, [int]$originY) {
    $out = New-Object System.Collections.ArrayList
    $f = $ocr.Factor
    # 注意：$result.Lines 是 __ComObject，直接索引会触发成员枚举把多行粘成一段，
    # 必须先 @() 包一层
    foreach ($line in @($ocr.Result.Lines)) {
        $ws = @($line.Words)
        if ($ws.Count -eq 0) { continue }
        $x1 = [double]::MaxValue; $y1 = [double]::MaxValue
        $x2 = [double]::MinValue; $y2 = [double]::MinValue
        foreach ($w in $ws) {
            $b = $w.BoundingRect          # OcrLine 没有 BoundingRect，只能取词框的并集
            if ($b.X -lt $x1) { $x1 = $b.X }
            if ($b.Y -lt $y1) { $y1 = $b.Y }
            if (($b.X + $b.Width) -gt $x2) { $x2 = $b.X + $b.Width }
            if (($b.Y + $b.Height) -gt $y2) { $y2 = $b.Y + $b.Height }
        }
        if ($x2 -lt $x1) { continue }
        $null = $out.Add([pscustomobject]@{
                Text = Join-OcrWords $ws $f
                X    = $originX + [int][Math]::Round($x1 / $f)
                Y    = $originY + [int][Math]::Round($y1 / $f)
                W    = [int][Math]::Round(($x2 - $x1) / $f)
                H    = [int][Math]::Round(($y2 - $y1) / $f)
            })
    }
    return $out
}

# ══ 6. 翻译（主：Google 免费接口；备：MyMemory） ══════════════
function Split-Chunks($text, [int]$size) {
    $chunks = New-Object System.Collections.ArrayList
    $lines = $text -split "`r?`n"
    $cur = ''
    foreach ($ln in $lines) {
        if ($cur.Length -gt 0 -and ($cur.Length + $ln.Length + 1) -gt $size) {
            $null = $chunks.Add($cur); $cur = ''
        }
        if ($ln.Length -gt $size) {
            # 单行就超长：按句子边界硬切
            $rest = $ln
            while ($rest.Length -gt $size) {
                $cut = $rest.LastIndexOf('. ', [Math]::Min($size, $rest.Length - 2))
                if ($cut -lt $size / 2) { $cut = $size - 1 }
                if ($cur.Length -gt 0) { $null = $chunks.Add($cur); $cur = '' }
                $null = $chunks.Add($rest.Substring(0, $cut + 1))
                $rest = $rest.Substring($cut + 1)
            }
            $cur = $rest
        }
        else {
            if ($cur.Length -gt 0) { $cur += "`n" }
            $cur += $ln
        }
    }
    if ($cur.Length -gt 0) { $null = $chunks.Add($cur) }
    return $chunks
}

# PS 5.1 的 Invoke-WebRequest 是照响应头里的 charset 解码的；这几个接口返回的是
# UTF-8 但头里要么不写 charset、要么写错，于是被当成 ISO-8859-1 解，中文全变成
# 「æ ¼à¡Ý」这种乱码。绕开它，直接拿原始字节自己用 UTF-8 解。
function Get-JsonFromUrl($url) {
    $r = Invoke-WebRequest -Uri $url -TimeoutSec $TranslateTimeout -UseBasicParsing
    $bytes = $r.RawContentStream.ToArray()
    return ([System.Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json)
}

# ── 主接口：有道（国内服务，稳定不限流，实测 300-400ms） ──
function Invoke-YoudaoTranslate($text) {
    $url = 'https://aidemo.youdao.com/trans?from=en&to=zh-CHS&q=' + [uri]::EscapeDataString($text)
    $j = Get-JsonFromUrl $url
    if ($j.errorCode -ne 0) { throw "有道返回错误码 $($j.errorCode)" }
    $t = ($j.translation -join "`n")
    if ([string]::IsNullOrWhiteSpace($t)) { throw '有道返回了空结果' }
    return $t
}

# ── 备接口：Google 免费 gtx（质量最好，但按 IP 限流，实测会 429） ──
function Invoke-GoogleTranslate($text) {
    $url = 'https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=zh-CN&dt=t&q=' +
    [uri]::EscapeDataString($text)
    $j = Get-JsonFromUrl $url
    $parts = @()
    foreach ($seg in @($j[0])) {
        if ($seg -and $seg[0]) { $parts += [string]$seg[0] }
    }
    if ($parts.Count -eq 0) { throw 'Google 接口返回了空结果' }
    return (-join $parts)
}

function Invoke-MyMemoryTranslate($text) {
    $url = 'https://api.mymemory.translated.net/get?q=' + [uri]::EscapeDataString($text) + '&langpair=en|zh-CN'
    $j = Get-JsonFromUrl $url
    $t = [string]$j.responseData.translatedText
    if ([string]::IsNullOrWhiteSpace($t)) { throw 'MyMemory 返回了空结果' }
    if ($t -match 'QUERY LENGTH LIMIT|INVALID') { throw "MyMemory 拒绝：$t" }
    return $t
}

# 三个免费接口按可靠性排序，前一个失败就自动退到下一个。
# 全部是公开的非官方接口，随时可能失效——所以留了三个，并且失败会明确告诉用户。
$script:Engines = @(
    @{ Name = '有道'; Fn = { param($c) Invoke-YoudaoTranslate $c } },
    @{ Name = 'Google'; Fn = { param($c) Invoke-GoogleTranslate $c } },
    @{ Name = 'MyMemory'; Fn = { param($c) Invoke-MyMemoryTranslate $c } }
)

$script:Cache = @{}      # 相同内容不重复请求，长按热键时很省事

function Invoke-Translate($text) {
    if ($script:Cache.ContainsKey($text)) {
        return [pscustomobject]@{ Text = $script:Cache[$text]; Engine = '缓存'; Failed = $false }
    }

    $chunks = Split-Chunks $text $ChunkSize
    $out = New-Object System.Collections.ArrayList
    $usedEngine = $null
    $lastErr = $null

    foreach ($c in $chunks) {
        $piece = $null
        foreach ($e in $script:Engines) {
            try {
                $piece = & $e.Fn $c
                if (-not $usedEngine) { $usedEngine = $e.Name }
                break
            }
            catch { $lastErr = $e.Name + ': ' + $_.Exception.Message }
        }
        if (-not $piece) { throw "所有翻译接口都失败了（最后一个：$lastErr）" }
        $null = $out.Add($piece)
    }

    $result = (-join $out)
    $script:Cache[$text] = $result
    return [pscustomobject]@{ Text = $result; Engine = $usedEngine; Failed = $false }
}

# ══ 7. 结果卡片 ══════════════════════════════════════════════
function New-CardButton($parent, $text, $x, $y, $w) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Location = New-Object System.Drawing.Point $x, $y
    $bh = [int](30 * $Scale)
    $b.Size = New-Object System.Drawing.Size -ArgumentList $w, $bh
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderColor = $C_Line
    $b.FlatAppearance.MouseOverBackColor = $C_BtnHov
    $b.BackColor = $C_Btn
    $b.ForeColor = $C_Text
    $b.Font = New-UiFont 8.5
    $parent.Controls.Add($b)
    return $b
}

function Show-Card($translation, $source, $sel, $note) {
    $padX   = [int](18 * $Scale)
    $tagRow = [int](17 * $Scale)     # “译文 / 原文”小标题占的高度
    $gap    = [int](13 * $Scale)     # 段间距
    $fTrans = New-UiFont 11.5
    $fSrc = New-UiFont 9
    $fTag = New-UiFont 8

    # 一律按「右侧已经让出一条滚动条」的宽度排版并丈量。要是按满宽量、之后再
    # 让滚动条挤窄，Label 会重新折行、实际比量出来的高，末尾被裁掉，
    # WinForms 还会顺手拉一条横向滚动条把正文推歪。
    $wFull = $CardWidth - $padX * 2
    $contentW = $wFull - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth - [int](6 * $Scale)

    # MeasureString 量出的行高比 Label 实际渲染的略矮（行距取整），留 15% 余量，
    # 否则长文本最后一行会被 Label 自己吃掉。
    $hTrans = [int][Math]::Ceiling((Measure-TextHeight $translation $fTrans $contentW) * 1.15)
    $hSrc = [int][Math]::Ceiling((Measure-TextHeight $source $fSrc $contentW) * 1.15)

    # 正文区纵向排布：译文标题 / 译文 / 分隔线 / 原文标题 / 原文
    $yTransTag = $gap
    $yTrans = $yTransTag + $tagRow
    $ySep = $yTrans + $hTrans + $gap
    $ySrcTag = $ySep + $gap
    $ySrc = $ySrcTag + $tagRow
    $bodyH = $ySrc + $hSrc + $gap

    $viewH = [Math]::Min($bodyH, $CardMaxBody)
    $btnH = [int](46 * $Scale)
    $cardH = $viewH + $btnH

    $card = New-Object System.Windows.Forms.Form
    $card.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $card.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $card.ShowInTaskbar = $false
    $card.TopMost = $true
    $card.BackColor = $C_Bg
    $card.ClientSize = New-Object System.Drawing.Size $CardWidth, $cardH
    $card.KeyPreview = $true
    $card.Add_Paint({
            param($s, $e)
            $pen = New-Object System.Drawing.Pen $C_Line, 1
            $e.Graphics.DrawRectangle($pen, 0, 0, $s.ClientSize.Width - 1, $s.ClientSize.Height - 1)
            $pen.Dispose()
        })

    # 正文放在可滚动面板里，长文本不会撑破卡片
    $body = New-Object System.Windows.Forms.Panel
    $body.Location = New-Object System.Drawing.Point 0, 0
    $body.Size = New-Object System.Drawing.Size $CardWidth, $viewH
    $body.BackColor = $C_Bg
    # 常开：不超高时不会冒出滚动条，超高时至少能滚，不会把正文闷在可视区外
    $body.AutoScroll = $true
    $card.Controls.Add($body)

    # 高度必须给准。给个虚高（比如 4000）会把可滚动区域撑爆，
    # 滚动条粗细和滑块位置全乱；给矮了则最后一行被切掉。
    function Add-CardLabel($text, $x, $y, $w, $h, $font, $color) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text
        $l.Location = New-Object System.Drawing.Point $x, $y
        $l.AutoSize = $false
        $l.Size = New-Object System.Drawing.Size $w, $h
        $l.BackColor = [System.Drawing.Color]::Transparent
        $l.ForeColor = $color
        $l.Font = $font
        $body.Controls.Add($l)
        return $l
    }

    $tagTrans = Add-CardLabel '译文' $padX $yTransTag $contentW $tagRow $fTag $C_Title
    $null = Add-CardLabel $translation $padX $yTrans $contentW ($hTrans + [int](4 * $Scale)) $fTrans $C_Text

    $sep = New-Object System.Windows.Forms.Panel
    $sep.Location = New-Object System.Drawing.Point $padX, $ySep
    $sep.Size = New-Object System.Drawing.Size $contentW, 1
    $sep.BackColor = $C_Line
    $body.Controls.Add($sep)

    $tagSrc = Add-CardLabel '原文' $padX $ySrcTag $contentW $tagRow $fTag $C_Title
    $null = Add-CardLabel $source $padX $ySrc $contentW ($hSrc + [int](4 * $Scale)) $fSrc $C_Dim

    $body.AutoScrollMinSize = New-Object System.Drawing.Size 0, $bodyH

    # 底部按钮条（跟正文之间拉一条分隔线）
    $barLine = New-Object System.Windows.Forms.Panel
    $barLine.Location = New-Object System.Drawing.Point 0, $viewH
    $barLine.Size = New-Object System.Drawing.Size $CardWidth, 1
    $barLine.BackColor = $C_Line
    $card.Controls.Add($barLine)

    $barY = $viewH + [int](9 * $Scale)
    $btnW = [int](88 * $Scale)
    $btnCopy = New-CardButton $card '复制译文' $padX $barY $btnW
    $btnClose = New-CardButton $card '关闭' ($CardWidth - $padX - $btnW) $barY $btnW

    $script:LastTranslation = $translation
    # 回调里一律用事件自带的 sender（$s）拿控件，不要引用 Show-Card 的局部变量：
    # 函数返回后那些局部变量就没了，回调跑起来时它们全是 $null。
    $btnCopy.Add_Click({
            param($s, $e)
            if ($script:LastTranslation) { [System.Windows.Forms.Clipboard]::SetText($script:LastTranslation) }
            $s.Text = '已复制'
        })
    $btnClose.Add_Click({ param($s, $e) $s.FindForm().Close() })
    $card.Add_KeyDown({
            param($s, $e)
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $s.Close() }
        })
    # 点卡片外面自动收起。刚弹出时焦点还没落稳，前 400ms 的 Deactivate 一律忽略，
    # 否则卡片会在显示出来的瞬间被自己关掉
    $card.Add_Deactivate({
            param($s, $e)
            if (((Get-Date) - $script:CardBorn).TotalMilliseconds -lt 400) { return }
            $s.Close()
        })

    # 位置：优先贴选区右边，右边放不下改左边，再不行贴左边缘；纵向夹进工作区
    $scr = [System.Windows.Forms.Screen]::FromRectangle($sel)
    $wa = $scr.WorkingArea
    $gap = 10
    $x = $sel.Right + $gap
    if (($x + $CardWidth) -gt ($wa.Right - 4)) { $x = $sel.Left - $CardWidth - $gap }
    if ($x -lt ($wa.Left + 4)) {
        $x = [Math]::Max(($wa.Left + 4), [Math]::Min($sel.Left, ($wa.Right - $CardWidth - 4)))
    }
    $y = $sel.Top
    if (($y + $cardH) -gt ($wa.Bottom - 4)) { $y = $wa.Bottom - $cardH - 4 }
    if ($y -lt ($wa.Top + 4)) { $y = $wa.Top + 4 }
    $card.Location = New-Object System.Drawing.Point $x, $y

    if ($note) { $tagTrans.Text = "译文    ($note)" }

    $script:CardBorn = Get-Date
    $card.Show()
    # -WindowStyle Hidden 启动时第一次 ShowWindow 会被强制隐藏，这里补一刀确保显形
    [void][StDpi]::ShowWindow($card.Handle, 5)
    $card.BringToFront()
    $card.Activate()
    [System.Windows.Forms.Application]::DoEvents()
    return $card
}

# ══ 8. 划词主流程 ════════════════════════════════════════════
function Start-Selection {
    if ($script:Busy) { return }
    $script:Busy = $true
    $sel = $null        # 出错时 catch 里要用，先占位
    $source = ''
    try {
        $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
        $ov = New-Object System.Windows.Forms.Form
        $ov.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $ov.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
        $ov.Bounds = $vs
        $ov.TopMost = $true
        $ov.ShowInTaskbar = $false
        $ov.BackColor = [System.Drawing.Color]::Black
        $ov.Opacity = 0.30
        $ov.Cursor = [System.Windows.Forms.Cursors]::Cross
        $ov.KeyPreview = $true
        try {
            $ov.GetType().GetProperty('DoubleBuffered',
                [System.Reflection.BindingFlags]'NonPublic,Instance').SetValue($ov, $true, $null)
        }
        catch { }

        $script:SelStart = $null
        $script:SelRect = [System.Drawing.Rectangle]::Empty

        $ov.Add_MouseDown({
                param($s, $e)
                if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
                    $script:Ov.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $script:Ov.Close(); return
                }
                if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    $script:SelStart = $e.Location
                    $script:SelRect = New-Object System.Drawing.Rectangle $e.X, $e.Y, 0, 0
                    $script:Ov.Invalidate()
                }
            })

        $ov.Add_MouseMove({
                param($s, $e)
                if (-not $script:SelStart) { return }
                $p = $script:SelStart
                $x = [Math]::Min($p.X, $e.X); $y = [Math]::Min($p.Y, $e.Y)
                $w = [Math]::Abs($e.X - $p.X); $h = [Math]::Abs($e.Y - $p.Y)
                $script:SelRect = New-Object System.Drawing.Rectangle $x, $y, $w, $h
                $script:Ov.Invalidate()
            })

        $ov.Add_MouseUp({
                param($s, $e)
                if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
                if ($script:SelRect.Width -lt $MinSelectPx -or $script:SelRect.Height -lt $MinSelectPx) {
                    $script:Ov.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                }
                else {
                    $script:Ov.DialogResult = [System.Windows.Forms.DialogResult]::OK
                }
                $script:Ov.Close()
            })

        $ov.Add_KeyDown({
                param($s, $e)
                if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                    $script:Ov.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $script:Ov.Close()
                }
            })

        $ov.Add_Paint({
                param($s, $e)
                $r = $script:SelRect
                if ($r.Width -le 0 -or $r.Height -le 0) { return }
                $pen = New-Object System.Drawing.Pen $C_Accent, 2
                $e.Graphics.DrawRectangle($pen, $r)
                $pen.Dispose()
            })

        $script:Ov = $ov
        $res = $ov.ShowDialog()
        $rect = $script:SelRect
        $ov.Dispose()
        $script:Ov = $null

        if ($res -ne [System.Windows.Forms.DialogResult]::OK) { return }
        if ($rect.Width -lt $MinSelectPx -or $rect.Height -lt $MinSelectPx) { return }

        # 遮罩是铺在 VirtualScreen 上的，换算回屏幕物理坐标
        $sel = New-Object System.Drawing.Rectangle ($vs.X + $rect.X), ($vs.Y + $rect.Y), $rect.Width, $rect.Height

        # 先截图再弹占位卡片。反过来的话，当选框贴着屏幕左边、卡片被挤到
        # 选框身上时，截出来的图里就带着卡片自己的「识别中…」，OCR 认出一堆
        # 不该有的字。截图只要几十毫秒，用户感觉不出这点延迟。
        $bmp = Capture-Region $sel.X $sel.Y $sel.Width $sel.Height
        $card = Show-Card '识别中…' '' $sel $null

        $ocr = Invoke-Ocr $bmp
        $bmp.Dispose()
        $lines = Get-OcrLines $ocr $sel.X $sel.Y

        if ($lines.Count -eq 0) {
            $card.Close()
            Show-Card '没有识别到文字' '（换一块区域，或把内容放大一点再试）' $sel $null | Out-Null
            return
        }

        $source = (($lines | ForEach-Object { $_.Text }) -join "`n")

        $card.Close()
        $card = Show-Card '翻译中…' $source $sel $null

        $t = Invoke-Translate $source
        $card.Close()
        Show-Card $t.Text $source $sel $t.Engine | Out-Null
    }
    catch {
        # 翻译失败也要把已经 OCR 出来的原文留着，否则断网时用户什么都拿不到
        try {
            $anchor = if ($null -ne $sel) { $sel } else { [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea }
            Show-Card ('翻译失败：' + $_.Exception.Message) $source $anchor $null | Out-Null
        }
        catch { }
    }
    finally {
        $script:Busy = $false
    }
}

# ══ 9. 托盘图标 ══════════════════════════════════════════════
function New-TrayIcon {
    $bmp = New-Object System.Drawing.Bitmap 32, 32
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)
    $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml('#3C8CE7'))
    $g.FillEllipse($brush, 1, 1, 30, 30)
    $font = New-Object System.Drawing.Font 'Microsoft YaHei UI', 16, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rect = New-Object System.Drawing.RectangleF 0, 0, 32, 32
    $g.DrawString('译', $font, [System.Drawing.Brushes]::White, $rect, $sf)
    $font.Dispose(); $brush.Dispose(); $sf.Dispose(); $g.Dispose()
    $hicon = $bmp.GetHicon()
    $bmp.Dispose()
    return [System.Drawing.Icon]::FromHandle($hicon)
}

# ══ 10. 自检模式 ═════════════════════════════════════════════
if ($SelfTest) {
    "=== DPI ==="
    "  物理分辨率: " + [StDpi]::GetSystemMetrics(0) + 'x' + [StDpi]::GetSystemMetrics(1)
    "=== OCR 语言包 ==="
    foreach ($l in [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages) { "  - " + $l.LanguageTag }
    "=== OCR 实测 ==="
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $bmp = Capture-Region 0 0 600 400
    $r = Invoke-Ocr $bmp
    $sw.Stop()
    $ls = Get-OcrLines $r 0 0
    "  耗时 $($sw.ElapsedMilliseconds)ms，识别到 $($ls.Count) 行"
    $bmp.Dispose()
    "=== 翻译实测 ==="
    foreach ($sample in @('Hello world, this is a test.', 'The operation completed successfully.')) {
        $sw.Restart()
        try {
            $t = Invoke-Translate $sample
            $sw.Stop()
            "  [$($sw.ElapsedMilliseconds)ms / $($t.Engine)] $($t.Text)"
        }
        catch { "  失败: " + $_.Exception.Message }
    }
    exit 0
}

# ══ 10a. OCR 调试（把某个区域截下来单独跑一次 OCR，结果落盘） ══
if ($OcrShot) {
    $bmp = Capture-Region $RectX $RectY $RectW $RectH
    $bmp.Save("$PSScriptRoot\.crop.png", [System.Drawing.Imaging.ImageFormat]::Png)
    $ocr = Invoke-Ocr $bmp
    $lines = Get-OcrLines $ocr $RectX $RectY
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("区域=$RectX,$RectY ${RectW}x${RectH}  放大=$($ocr.Factor)x  行数=$($lines.Count)")
    foreach ($l in $lines) {
        [void]$sb.AppendLine(("   [{0},{1} {2}x{3}] {4}" -f $l.X, $l.Y, $l.W, $l.H, $l.Text))
    }
    [System.IO.File]::WriteAllText("$PSScriptRoot\.ocr.txt", $sb.ToString(),
        (New-Object System.Text.UTF8Encoding $true))
    exit 0
}

# ══ 10b. 演示模式（核对卡片排版用） ══════════════════════════
if ($Demo) {
    $consume = New-Object System.Windows.Forms.Form
    $consume.Show(); $consume.Hide(); $consume.Dispose()

    $sel = New-Object System.Drawing.Rectangle 300, 300, 640, 220
    $src = "The operation completed successfully.`nPlease check your network connection and try again.`nThis dialog will close automatically after a few seconds."
    $dst = "操作已成功完成。`n请检查您的网络连接，然后重试。`n此对话框将在几秒后自动关闭。"
    $card = Show-Card $dst $src $sel '有道'

    Start-Sleep -Milliseconds 1500
    $cx = $card.Location.X - 20
    $cy = $card.Location.Y - 20
    $bmp = New-Object System.Drawing.Bitmap -ArgumentList ($card.Width + 40), ($card.Height + 40)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($cx, $cy, 0, 0, (New-Object System.Drawing.Size -ArgumentList ($card.Width + 40), ($card.Height + 40)))
    $g.Dispose()
    $bmp.Save('D:\kk\codex\screen-translator\.demo.png', [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $card.Close()
    "已保存 .demo.png  卡片位置=$($card.Location)  尺寸=$($card.Width)x$($card.Height)"
    exit 0
}

# ══ 11. 主流程 ═══════════════════════════════════════════════
$script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon

# -WindowStyle Hidden 启动时 STARTUPINFO 里是 SW_HIDE，Windows 会强制隐藏
# 本进程的【第一次】ShowWindow。用一个一次性窗口把它吃掉，后面才画得出来。
$consume = New-Object System.Windows.Forms.Form
$consume.Show()
$consume.Hide()
$consume.Dispose()

$script:Busy = $false
$script:LastTranslation = ''

$script:NotifyIcon.Icon = New-TrayIcon
$script:NotifyIcon.Text = '屏幕翻译  (' + $HotkeyName + ')'
$script:NotifyIcon.Visible = $true

$hk = New-Object StHotkey
$hk.add_Fired({ Start-Selection })
$hooked = $hk.Hook([uint32]$HotkeyMods, [uint32]$HotkeyKey)

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miGo     = $menu.Items.Add('立即划词翻译')
$miHelp   = $menu.Items.Add('使用说明')
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$miExit   = $menu.Items.Add('退出')
$miGo.Add_Click({ Start-Selection })
$miHelp.Add_Click({
        [void][System.Windows.Forms.MessageBox]::Show(
            "用法：" + "`n" +
            "  1. 按 " + $HotkeyName + "（或右键托盘图标选「立即划词翻译」）" + "`n" +
            "  2. 屏幕变暗后，用鼠标拖拽框住要翻译的英文" + "`n" +
            "  3. 松手，右侧会弹出译文卡片" + "`n" +
            "  4. 按 Esc 或点卡片外面收起；框选途中按 Esc / 右键可取消" + "`n`n" +
            "说明：" + "`n" +
            "  · 识别用 Windows 自带 OCR，离线完成，不联网" + "`n" +
            "  · 翻译会把识别到的文字发到免费的在线翻译接口" + "`n" +
            "  · 卡片上的「复制译文」按钮可把译文放进剪贴板",
            '屏幕翻译 - 使用说明',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    })
$miExit.Add_Click({
        $hk.Unhook()
        $script:NotifyIcon.Visible = $false
        $script:NotifyIcon.Dispose()
        [System.Windows.Forms.Application]::Exit()
    })
$script:NotifyIcon.ContextMenuStrip = $menu

$script:NotifyIcon.Add_MouseClick({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Start-Selection }
    })

if (-not $hooked) {
    $script:NotifyIcon.ShowBalloonTip(8000, '热键注册失败',
        "快捷键 $HotkeyName 已被别的程序占用。`n右键托盘图标选「立即划词翻译」仍可使用。",
        [System.Windows.Forms.ToolTipIcon]::Warning)
}

[System.Windows.Forms.Application]::Run()
