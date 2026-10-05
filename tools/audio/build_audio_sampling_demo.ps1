param(
    [string]$SourceWav = "",
    [string]$OutputPptx = "",
    [string]$PythonExe = ""
)

$ErrorActionPreference = "Stop"

function RgbInt([string]$Hex) {
    $h = $Hex.TrimStart("#")
    $r = [Convert]::ToInt32($h.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($h.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function AddText($Slide, [string]$Text, [double]$X, [double]$Y, [double]$W, [double]$H, [int]$Size, [string]$Color, [bool]$Bold = $false) {
    $box = $Slide.Shapes.AddTextbox(1, $X, $Y, $W, $H)
    $box.TextFrame2.TextRange.Text = $Text
    $box.TextFrame2.TextRange.Font.NameFarEast = "Microsoft YaHei UI"
    $box.TextFrame2.TextRange.Font.Name = "Microsoft YaHei UI"
    $box.TextFrame2.TextRange.Font.Size = $Size
    $box.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RgbInt $Color
    if ($Bold) { $box.TextFrame2.TextRange.Font.Bold = -1 }
    $box.TextFrame2.WordWrap = -1
    $box.TextFrame2.MarginLeft = 0
    $box.TextFrame2.MarginRight = 0
    $box.TextFrame2.MarginTop = 0
    $box.TextFrame2.MarginBottom = 0
    return $box
}

function AddCard($Slide, [double]$X, [double]$Y, [double]$W, [double]$H, [string]$Fill = "#FFFFFF", [string]$Line = "#D9DDD4") {
    $shape = $Slide.Shapes.AddShape(5, $X, $Y, $W, $H)
    $shape.Fill.ForeColor.RGB = RgbInt $Fill
    $shape.Line.ForeColor.RGB = RgbInt $Line
    $shape.Line.Weight = 1.25
    return $shape
}

function AddMedia($Slide, [string]$Path, [double]$X, [double]$Y, [double]$W = 42, [double]$H = 42) {
    $media = $Slide.Shapes.AddMediaObject2($Path, 0, -1, $X, $Y, $W, $H)
    try {
        $media.AnimationSettings.PlaySettings.PlayOnEntry = 0
        $media.AnimationSettings.PlaySettings.HideWhileNotPlaying = 0
        $media.MediaFormat.Volume = 0.86
    } catch {}
    return $media
}

function AddFooter($Slide, [string]$Text) {
    AddText $Slide $Text 44 505 860 22 9 "#737A84" $false | Out-Null
}

if ($PSScriptRoot) {
    $scriptDir = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $scriptDir = Join-Path (Get-Location) "tools\audio"
}
$rootDir = Resolve-Path (Join-Path $scriptDir "..\..")
$assetDir = Join-Path $rootDir "assets\audio"
$previewDir = Join-Path $rootDir "build\audio-preview"
New-Item -ItemType Directory -Force -Path $assetDir | Out-Null
New-Item -ItemType Directory -Force -Path $previewDir | Out-Null

if (-not $PythonExe) {
    $PythonExe = "python"
}

$assetScript = Join-Path $scriptDir "build_audio_assets.py"
$assetArgs = @($assetScript, "--out-dir", $assetDir)
if ($SourceWav) { $assetArgs += @("--source", $SourceWav) }
& $PythonExe @assetArgs | Write-Host

$manifestPath = Join-Path $assetDir "manifest.json"
$manifest = Get-Content -Path $manifestPath -Encoding UTF8 -Raw | ConvertFrom-Json

# Resolve portable manifest filenames before handing them to PowerPoint.
foreach ($clip in @($manifest.clips) + @($manifest.main_clips)) {
    if (-not [IO.Path]::IsPathRooted($clip.file)) { $clip.file = Join-Path $assetDir $clip.file }
}
foreach ($property in @("spectra_image", "waveform_image")) {
    if (-not [IO.Path]::IsPathRooted($manifest.$property)) { $manifest.$property = Join-Path $assetDir $manifest.$property }
}

if (-not $OutputPptx) {
    $OutputPptx = Join-Path $rootDir "materials\slides\采样定理_音频采样率播放Demo_重生成.pptx"
}

$workingPptx = Join-Path $scriptDir "_working_audio_sampling_demo.pptx"
if (Test-Path $workingPptx) {
    Remove-Item -Path $workingPptx -Force
}

$pp = $null
$pres = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.Visible = -1
    $pres = $pp.Presentations.Add()
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540

    $blank = 12
    $accentTeal = "#0B6E69"
    $accentAmber = "#BB6B00"
    $accentRed = "#D14256"
    $ink = "#20242A"
    $muted = "#59616C"
    $paper = "#F7F8F4"

    # Slide 1
    $s1 = $pres.Slides.Add(1, $blank)
    $s1.Background.Fill.ForeColor.RGB = RgbInt $paper
    AddText $s1 "音频采样率 Demo" 62 58 620 54 34 $ink $true | Out-Null
    AddText $s1 "为什么记录人耳可听范围，常用 44.1 kHz / 48 kHz？" 63 120 770 38 20 $muted $false | Out-Null
    $band = $s1.Shapes.AddShape(1, 62, 190, 835, 126)
    $band.Fill.ForeColor.RGB = RgbInt "#FFFFFF"
    $band.Line.ForeColor.RGB = RgbInt "#D9DDD4"
    AddText $s1 "人耳上限常取约 20 kHz" 96 220 260 28 18 $ink $true | Out-Null
    AddText $s1 "Nyquist 要求：采样率 fs > 2 × fmax，所以 fs 必须高于 40 kHz。" 96 258 670 34 18 $ink $false | Out-Null
    AddText $s1 "工程上还要给抗混叠滤波器留过渡带，因此 CD 采用 44.1 kHz，专业音频常用 48 kHz。" 96 293 720 30 14 $muted $false | Out-Null
    $line = $s1.Shapes.AddShape(1, 93, 370, 774, 18)
    $line.Fill.ForeColor.RGB = RgbInt "#DFE7E1"
    $line.Line.Visible = 0
    $mark20 = $s1.Shapes.AddShape(1, 93, 366, 690, 26)
    $mark20.Fill.ForeColor.RGB = RgbInt "#67B6AF"
    $mark20.Line.Visible = 0
    AddText $s1 "0 Hz" 90 398 70 18 10 $muted $false | Out-Null
    AddText $s1 "20 kHz" 756 398 70 18 10 $muted $false | Out-Null
    AddText $s1 "44.1 kHz 采样 → Nyquist 22.05 kHz" 645 338 220 21 11 $accentTeal $true | Out-Null
    AddFooter $s1 "本文件不含商业录音；若用于《Melody》，请替换为你拥有使用权的 WAV 片段。"

    # Slide 2
    $s2 = $pres.Slides.Add(2, $blank)
    $s2.Background.Fill.ForeColor.RGB = RgbInt "#FFFFFF"
    AddText $s2 "同一段声音，换采样率播放" 44 38 560 42 29 $ink $true | Out-Null
    AddText $s2 "放映时点击每张卡片里的扬声器图标，按 44.1 kHz → 4 kHz 的顺序比较。" 46 86 760 30 16 $muted $false | Out-Null
    $clips = @($manifest.main_clips)
    $positions = @(
        @(46, 142), @(342, 142), @(638, 142),
        @(194, 330), @(490, 330)
    )
    $colors = @($accentTeal, $accentAmber, "#7B4FB4", $accentRed, "#6A737D")
    for ($i = 0; $i -lt $clips.Count; $i++) {
        $clip = $clips[$i]
        $pos = $positions[$i]
        AddCard $s2 $pos[0] $pos[1] 250 145 "#F7F8F4" "#D9DDD4" | Out-Null
        AddText $s2 $clip.short_label ($pos[0] + 20) ($pos[1] + 18) 142 26 20 $colors[$i] $true | Out-Null
        AddText $s2 $clip.note ($pos[0] + 20) ($pos[1] + 52) 170 24 12 $muted $false | Out-Null
        AddText $s2 ("Nyquist：" + $clip.nyquist_khz + " kHz") ($pos[0] + 20) ($pos[1] + 82) 160 22 12 $ink $false | Out-Null
        AddText $s2 "点击播放" ($pos[0] + 178) ($pos[1] + 97) 58 18 9 $muted $false | Out-Null
        AddMedia $s2 $clip.file ($pos[0] + 184) ($pos[1] + 36) 40 40 | Out-Null
    }
    AddFooter $s2 "提示：所有音频已重建为 PPT 兼容的 44.1 kHz WAV，卡片标题表示被模拟的采样率。"

    # Slide 3
    $s3 = $pres.Slides.Add(3, $blank)
    $s3.Background.Fill.ForeColor.RGB = RgbInt "#F7F8F4"
    AddText $s3 "看见差异：高频能量随采样率下降而消失" 44 34 700 38 25 $ink $true | Out-Null
    AddText $s3 "采样率越低，Nyquist 频率越低；为了避免混叠，系统必须先滤掉更高频率。" 46 76 760 24 14 $muted $false | Out-Null
    $pic = $s3.Shapes.AddPicture($manifest.spectra_image, 0, -1, 36, 110, 888, 385)
    AddFooter $s3 "黑色竖线是每个采样率下可保留频带的大致边界。"

    # Slide 4
    $s4 = $pres.Slides.Add(4, $blank)
    $s4.Background.Fill.ForeColor.RGB = RgbInt "#FFFFFF"
    AddText $s4 "看见差异：采样点变稀疏，细节难以重建" 44 34 730 38 25 $ink $true | Out-Null
    AddText $s4 "这个图适合讲解「每秒采样点数」这一工程含义：不是点越多越玄学，而是点数决定能跟上的最高变化速度。" 46 76 800 24 14 $muted $false | Out-Null
    $pic2 = $s4.Shapes.AddPicture($manifest.waveform_image, 0, -1, 36, 112, 888, 390)
    AddFooter $s4 "低采样率并不只是音量变小，而是信息本身没有被记录下来。"

    # Slide 5
    $s5 = $pres.Slides.Add(5, $blank)
    $s5.Background.Fill.ForeColor.RGB = RgbInt "#F7F8F4"
    AddText $s5 "工程应用结论：录音系统怎么选采样率？" 44 38 730 38 27 $ink $true | Out-Null
    AddText $s5 "输入声波" 76 166 100 25 16 $ink $true | Out-Null
    AddText $s5 "抗混叠滤波" 248 166 130 25 16 $ink $true | Out-Null
    AddText $s5 "A/D 采样" 446 166 100 25 16 $ink $true | Out-Null
    AddText $s5 "数字音频" 615 166 120 25 16 $ink $true | Out-Null
    AddText $s5 "D/A 播放" 790 166 100 25 16 $ink $true | Out-Null
    $xs = @(60, 232, 428, 596, 772)
    foreach ($xv in $xs) {
        $shape = $s5.Shapes.AddShape(9, $xv, 118, 120, 42)
        $shape.Fill.ForeColor.RGB = RgbInt "#FFFFFF"
        $shape.Line.ForeColor.RGB = RgbInt "#D9DDD4"
    }
    foreach ($xv in @(182, 380, 548, 720)) {
        $arrow = $s5.Shapes.AddShape(33, $xv, 129, 42, 18)
        $arrow.Fill.ForeColor.RGB = RgbInt $accentTeal
        $arrow.Line.Visible = 0
    }
    AddCard $s5 70 250 380 160 "#FFFFFF" "#D9DDD4" | Out-Null
    AddText $s5 "满足可听范围" 94 274 180 24 18 $accentTeal $true | Out-Null
    AddText $s5 "若 fmax = 20 kHz，理论下限为 fs > 40 kHz。工程中选 44.1 kHz / 48 kHz，是为了给滤波器留出过渡带。" 94 312 315 74 14 $ink $false | Out-Null
    AddCard $s5 510 250 380 160 "#FFFFFF" "#D9DDD4" | Out-Null
    AddText $s5 "低采样率适用场景" 534 274 190 24 18 $accentAmber $true | Out-Null
    AddText $s5 "8 kHz 足够电话语音，但不适合保真音乐；4 kHz 会让乐器质感、空气感与泛音严重丢失。" 534 312 312 74 14 $ink $false | Out-Null
    AddFooter $s5 "一句话：采样率不是越高越好，而是必须高到能覆盖目标频带，并留出工程余量。"

    # Slide 6
    $s6 = $pres.Slides.Add(6, $blank)
    $s6.Background.Fill.ForeColor.RGB = RgbInt "#FFFFFF"
    AddText $s6 "可选补充：无抗混叠滤波会怎样？" 44 38 680 38 27 $ink $true | Out-Null
    AddText $s6 "这里把 8 kHz 采样分成两个版本：一个先滤波，一个直接采样。直接采样会把高频「折回」成错误的低频。" 46 84 795 30 15 $muted $false | Out-Null
    AddCard $s6 105 170 330 190 "#F7F8F4" "#D9DDD4" | Out-Null
    AddText $s6 "8 kHz，有抗混叠滤波" 132 203 210 28 19 $accentTeal $true | Out-Null
    AddText $s6 "声音更闷，但失真较可控。" 132 243 210 24 14 $ink $false | Out-Null
    AddText $s6 "点击播放" 132 293 84 20 11 $muted $false | Out-Null
    AddMedia $s6 $manifest.clips[3].file 245 282 46 46 | Out-Null
    AddCard $s6 525 170 330 190 "#FFF6F2" "#E3C9BD" | Out-Null
    AddText $s6 "8 kHz，无抗混叠滤波" 552 203 220 28 19 $accentRed $true | Out-Null
    AddText $s6 "高频折回，可能出现奇怪的毛刺和假低频。" 552 243 235 42 14 $ink $false | Out-Null
    AddText $s6 "点击播放" 552 293 84 20 11 $muted $false | Out-Null
    AddMedia $s6 $manifest.clips[5].file 665 282 46 46 | Out-Null
    AddFooter $s6 "课堂讲法：采样率不足会丢信息；没有抗混叠滤波，则会把错误信息写进数字信号。"

    # Slide 7
    $s7 = $pres.Slides.Add(7, $blank)
    $s7.Background.Fill.ForeColor.RGB = RgbInt "#F7F8F4"
    AddText $s7 "如何替换成你自己的《Melody》片段" 44 42 690 38 27 $ink $true | Out-Null
    AddText $s7 "为避免版权问题，当前 PPT 内嵌的是无版权测试乐句。你可以用自己拥有使用权的音频重新生成。" 46 86 800 30 15 $muted $false | Out-Null
    AddCard $s7 88 158 780 238 "#FFFFFF" "#D9DDD4" | Out-Null
    AddText $s7 "1. 准备 10–20 秒 WAV 文件，例如 Melody_authorized.wav" 122 190 700 28 16 $ink $false | Out-Null
    AddText $s7 "2. 放到本文件夹，或记下完整路径" 122 235 700 28 16 $ink $false | Out-Null
    AddText $s7 "3. 运行：" 122 280 120 28 16 $ink $false | Out-Null
    AddText $s7 "powershell -ExecutionPolicy Bypass -File outputs\\audio-sampling-demo\\build_audio_sampling_demo.ps1 -SourceWav .\\Melody_authorized.wav" 164 324 640 42 12 "#343A40" $false | Out-Null
    AddFooter $s7 "建议用 WAV/PCM 输入；MP3/M4A 可先用音频软件导出为 WAV。"

    $pres.SaveAs($workingPptx, 24)
    if (Test-Path $previewDir) {
        Remove-Item -Path $previewDir -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $previewDir | Out-Null
    $pres.Export($previewDir, "PNG", 1600, 900)
    Copy-Item -Path $workingPptx -Destination $OutputPptx -Force
    Write-Host "Created: $OutputPptx"
    Write-Host "Preview: $previewDir"
}
finally {
    if ($pres) {
        for ($try = 0; $try -lt 5; $try++) {
            try {
                $pres.Close() | Out-Null
                break
            } catch {
                Start-Sleep -Milliseconds 500
            }
        }
    }
    if ($pp) {
        for ($try = 0; $try -lt 5; $try++) {
            try {
                $pp.Quit() | Out-Null
                break
            } catch {
                Start-Sleep -Milliseconds 500
            }
        }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500
    if (Test-Path $workingPptx) {
        Remove-Item -Path $workingPptx -Force -ErrorAction SilentlyContinue
    }
}
