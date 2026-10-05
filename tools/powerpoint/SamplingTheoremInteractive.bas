Attribute VB_Name = "SamplingTheoremInteractive"
Option Explicit

Private dynId As Long
Private Const BASE_W As Double = 960#
Private Const BASE_H As Double = 540#

Public Sub BuildSamplingDemoOnSlide(Optional ByVal slideIndex As Long = 0)
    Dim s As Slide
    If slideIndex > 0 Then
        Set s = ActivePresentation.Slides(slideIndex)
    Else
        Set s = CurrentDemoSlide()
    End If
    ClearSlide s
    SetDefaults s
    BuildChrome s
    RedrawSamplingDemo
End Sub

Public Sub RedrawSamplingDemo()
    Dim s As Slide
    Set s = CurrentDemoSlide()
    If s Is Nothing Then Exit Sub
    If GetTag(s, "isDemo", "0") <> "1" Then Exit Sub

    DeleteDynamic s
    UpdateReadouts s
    DrawTimeChart s, 0
    DrawReconChart s
    DrawSpectrumChart s
    DrawTheoryPanel s
End Sub

Public Sub ResetDemo()
    Dim s As Slide
    Set s = CurrentDemoSlide()
    SetDefaults s
    RedrawSamplingDemo
End Sub

Public Sub F0Down(): AdjustNumber "f0", -20, 20, 500: End Sub
Public Sub F0Up(): AdjustNumber "f0", 20, 20, 500: End Sub
Public Sub FsDown(): AdjustNumber "fs", -50, 50, 2000: End Sub
Public Sub FsUp(): AdjustNumber "fs", 50, 50, 2000: End Sub
Public Sub PhaseDown(): AdjustNumber "phase", -30, -180, 180: End Sub
Public Sub PhaseUp(): AdjustNumber "phase", 30, -180, 180: End Sub

Public Sub ToggleMode()
    Dim s As Slide, v As Long
    Set s = CurrentDemoSlide()
    v = CLng(Val(GetTag(s, "mode", "0")))
    v = (v + 1) Mod 4
    SetTag s, "mode", CStr(v)
    RedrawSamplingDemo
End Sub

Public Sub ToggleMethod()
    Dim s As Slide, v As Long
    Set s = CurrentDemoSlide()
    v = CLng(Val(GetTag(s, "method", "0")))
    v = 1 - v
    SetTag s, "method", CStr(v)
    RedrawSamplingDemo
End Sub

Public Sub ToggleNoise()
    Dim s As Slide, v As Long
    Set s = CurrentDemoSlide()
    v = CLng(Val(GetTag(s, "noiseOn", "0")))
    v = 1 - v
    SetTag s, "noiseOn", CStr(v)
    RedrawSamplingDemo
End Sub

Private Sub AdjustNumber(ByVal key As String, ByVal delta As Double, ByVal mn As Double, ByVal mx As Double)
    Dim s As Slide, v As Double
    Set s = CurrentDemoSlide()
    If s Is Nothing Then Exit Sub
    v = Val(GetTag(s, key, "0")) + delta
    If v < mn Then v = mn
    If v > mx Then v = mx
    SetTag s, key, CStr(v)
    RedrawSamplingDemo
End Sub

Private Function CurrentDemoSlide() As Slide
    On Error Resume Next
    If SlideShowWindows.Count > 0 Then
        Set CurrentDemoSlide = SlideShowWindows(1).View.Slide
    End If
    If CurrentDemoSlide Is Nothing Then
        Set CurrentDemoSlide = ActiveWindow.View.Slide
    End If
    If CurrentDemoSlide Is Nothing Then
        Set CurrentDemoSlide = ActivePresentation.Slides(ActivePresentation.Slides.Count)
    End If
    On Error GoTo 0
End Function

Private Sub SetDefaults(ByVal s As Slide)
    SetTag s, "isDemo", "1"
    SetTag s, "f0", "120"
    SetTag s, "f1", "260"
    SetTag s, "fs", "400"
    SetTag s, "phase", "0"
    SetTag s, "mode", "0"
    SetTag s, "method", "0"
    SetTag s, "noiseOn", "0"
End Sub

Private Sub SetTag(ByVal s As Slide, ByVal key As String, ByVal value As String)
    On Error Resume Next
    s.Tags.Delete key
    s.Tags.Add key, value
    On Error GoTo 0
End Sub

Private Function GetTag(ByVal s As Slide, ByVal key As String, ByVal fallback As String) As String
    Dim v As String
    On Error Resume Next
    v = s.Tags(key)
    On Error GoTo 0
    If Len(v) = 0 Then v = fallback
    GetTag = v
End Function

Private Sub ClearSlide(ByVal s As Slide)
    Dim i As Long
    For i = s.Shapes.Count To 1 Step -1
        s.Shapes(i).Delete
    Next i
End Sub

Private Sub DeleteDynamic(ByVal s As Slide)
    Dim i As Long
    For i = s.Shapes.Count To 1 Step -1
        If Left$(s.Shapes(i).Name, 4) = "dyn_" Then s.Shapes(i).Delete
    Next i
    dynId = 0
End Sub

Private Function SX(ByVal s As Slide, ByVal x As Double) As Double
    SX = x * s.Parent.PageSetup.SlideWidth / BASE_W
End Function

Private Function SY(ByVal s As Slide, ByVal y As Double) As Double
    SY = y * s.Parent.PageSetup.SlideHeight / BASE_H
End Function

Private Function Sc(ByVal s As Slide) As Double
    Sc = s.Parent.PageSetup.SlideWidth / BASE_W
End Function

Private Function RgbHex(ByVal hex As String) As Long
    hex = Replace(hex, "#", "")
    RgbHex = RGB(CLng("&H" & Mid$(hex, 1, 2)), CLng("&H" & Mid$(hex, 3, 2)), CLng("&H" & Mid$(hex, 5, 2)))
End Function

Private Function AddBox(ByVal s As Slide, ByVal nm As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal fillHex As String, Optional ByVal lineHex As String = "#000000", Optional ByVal radius As Boolean = False) As Shape
    Dim shp As Shape
    Set shp = s.Shapes.AddShape(IIf(radius, msoShapeRoundedRectangle, msoShapeRectangle), SX(s, x), SY(s, y), SX(s, w), SY(s, h))
    shp.Name = nm
    shp.Fill.ForeColor.RGB = RgbHex(fillHex)
    shp.Line.ForeColor.RGB = RgbHex(lineHex)
    shp.Line.Weight = 1
    Set AddBox = shp
End Function

Private Function AddText(ByVal s As Slide, ByVal nm As String, ByVal txt As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal size As Double, ByVal colorHex As String, Optional ByVal bold As Boolean = False, Optional ByVal align As Long = ppAlignLeft) As Shape
    Dim shp As Shape
    Set shp = s.Shapes.AddTextbox(msoTextOrientationHorizontal, SX(s, x), SY(s, y), SX(s, w), SY(s, h))
    shp.Name = nm
    With shp.TextFrame
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoTrue
    End With
    With shp.TextFrame.TextRange
        .Text = txt
        .Font.Name = "Microsoft YaHei"
        .Font.NameFarEast = "Microsoft YaHei"
        .Font.Size = size * Sc(s)
        .Font.Bold = IIf(bold, msoTrue, msoFalse)
        .Font.Color.RGB = RgbHex(colorHex)
        .ParagraphFormat.Alignment = align
    End With
    Set AddText = shp
End Function

Private Function AddButton(ByVal s As Slide, ByVal nm As String, ByVal txt As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal macroName As String, Optional ByVal accent As String = "#3fe1d0") As Shape
    Dim shp As Shape
    Set shp = AddBox(s, nm, x, y, w, h, "#1d2a38", accent, True)
    With shp.TextFrame
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 2
        .MarginBottom = 2
    End With
    With shp.TextFrame.TextRange
        .Text = txt
        .Font.Name = "Microsoft YaHei"
        .Font.NameFarEast = "Microsoft YaHei"
        .Font.Size = 12 * Sc(s)
        .Font.Bold = msoTrue
        .Font.Color.RGB = RgbHex("#f2f7ff")
        .ParagraphFormat.Alignment = ppAlignCenter
    End With
    shp.ActionSettings(ppMouseClick).Action = ppActionRunMacro
    shp.ActionSettings(ppMouseClick).Run = macroName
    Set AddButton = shp
End Function

Private Sub BuildChrome(ByVal s As Slide)
    AddBox s, "static_bg", 0, 0, BASE_W, BASE_H, "#0b1017", "#0b1017"
    AddText s, "static_title", "采样定理互动演示器", 20, 14, 360, 28, 22, "#f2f7ff", True
    AddText s, "static_subtitle", "在放映模式中点击按钮，直接改变参数并重画采样、重建与频谱。", 22, 42, 520, 18, 10.5, "#a9b8c9"
    AddBox s, "static_sidebar", 16, 72, 214, 448, "#141d28", "#304052", True
    AddText s, "static_controls_title", "参数控制", 32, 88, 120, 20, 14, "#f2f7ff", True

    AddText s, "static_f0_label", "主频 f0", 32, 124, 70, 16, 10.5, "#dbe8f5", True
    AddButton s, "btn_f0_down", "-20", 108, 118, 44, 26, "F0Down"
    AddButton s, "btn_f0_up", "+20", 158, 118, 44, 26, "F0Up"
    AddText s, "val_f0", "", 32, 148, 170, 18, 12, "#3fe1d0", True

    AddText s, "static_fs_label", "采样频率 fs", 32, 182, 92, 16, 10.5, "#dbe8f5", True
    AddButton s, "btn_fs_down", "-50", 108, 176, 44, 26, "FsDown", "#ffae2b"
    AddButton s, "btn_fs_up", "+50", 158, 176, 44, 26, "FsUp", "#ffae2b"
    AddText s, "val_fs", "", 32, 206, 170, 18, 12, "#ffae2b", True

    AddText s, "static_phase_label", "相位 φ", 32, 240, 70, 16, 10.5, "#dbe8f5", True
    AddButton s, "btn_phase_down", "-30°", 108, 234, 44, 26, "PhaseDown"
    AddButton s, "btn_phase_up", "+30°", 158, 234, 44, 26, "PhaseUp"
    AddText s, "val_phase", "", 32, 264, 170, 18, 12, "#3fe1d0", True

    AddText s, "static_mode_label", "波形 / 重建", 32, 300, 110, 16, 10.5, "#dbe8f5", True
    AddButton s, "btn_mode", "切换波形", 32, 324, 82, 30, "ToggleMode"
    AddButton s, "btn_method", "重建方式", 122, 324, 82, 30, "ToggleMethod"
    AddText s, "val_mode", "", 32, 362, 170, 18, 11.5, "#f2f7ff", True
    AddText s, "val_method", "", 32, 382, 170, 18, 11.5, "#e64b9b", True

    AddButton s, "btn_noise", "噪声开/关", 32, 418, 82, 30, "ToggleNoise", "#e64b9b"
    AddButton s, "btn_reset", "重置", 122, 418, 82, 30, "ResetDemo", "#7bd66e"
    AddButton s, "btn_redraw", "初始化 / 重画", 32, 464, 172, 34, "RedrawSamplingDemo", "#3fe1d0"

    AddBox s, "statusBox", 748, 18, 190, 32, "#101722", "#304052", True
    AddText s, "statusText", "", 764, 25, 160, 18, 11, "#7bd66e", True, ppAlignCenter

    AddText s, "static_chart1", "1. 连续信号与采样点", 252, 66, 240, 18, 12.5, "#f2f7ff", True
    AddText s, "static_chart2", "2. 重建信号", 252, 180, 180, 18, 12.5, "#f2f7ff", True
    AddText s, "static_chart3", "3. 频谱复制与混叠", 252, 294, 220, 18, 12.5, "#f2f7ff", True
End Sub

Private Sub UpdateReadouts(ByVal s As Slide)
    Dim f0 As Double, fs As Double, phase As Double, fmax As Double, fn As Double, ok As Boolean
    f0 = Val(GetTag(s, "f0", "120"))
    fs = Val(GetTag(s, "fs", "400"))
    phase = Val(GetTag(s, "phase", "0"))
    fmax = CurrentFMax(s)
    fn = fs / 2
    ok = (fs >= 2 * fmax)
    SetText s, "val_f0", FormatHz(f0)
    SetText s, "val_fs", FormatHz(fs) & "   fN = " & FormatHz(fn)
    SetText s, "val_phase", CStr(CLng(phase)) & "°"
    SetText s, "val_mode", "波形：" & ModeName(CLng(Val(GetTag(s, "mode", "0"))))
    SetText s, "val_method", "重建：" & IIf(CLng(Val(GetTag(s, "method", "0"))) = 0, "Sinc 理想重建", "零阶保持")
    If ok Then
        SetText s, "statusText", "满足 Nyquist"
        s.Shapes("statusText").TextFrame.TextRange.Font.Color.RGB = RgbHex("#7bd66e")
        s.Shapes("statusBox").Line.ForeColor.RGB = RgbHex("#7bd66e")
    Else
        SetText s, "statusText", "发生混叠"
        s.Shapes("statusText").TextFrame.TextRange.Font.Color.RGB = RgbHex("#ff4f5f")
        s.Shapes("statusBox").Line.ForeColor.RGB = RgbHex("#ff4f5f")
    End If
End Sub

Private Sub SetText(ByVal s As Slide, ByVal name As String, ByVal txt As String)
    On Error Resume Next
    s.Shapes(name).TextFrame.TextRange.Text = txt
    On Error GoTo 0
End Sub

Private Function ModeName(ByVal mode As Long) As String
    Select Case mode
        Case 1: ModeName = "双频叠加"
        Case 2: ModeName = "三/五次谐波"
        Case 3: ModeName = "调幅信号"
        Case Else: ModeName = "单频正弦"
    End Select
End Function

Private Function FormatHz(ByVal f As Double) As String
    If f >= 1000 Then
        FormatHz = Format(f / 1000, "0.0") & " kHz"
    Else
        FormatHz = Format(f, "0") & " Hz"
    End If
End Function

Private Function SignalAt(ByVal t As Double, ByVal f0 As Double, ByVal f1 As Double, ByVal phaseDeg As Double, ByVal mode As Long) As Double
    Dim p As Double, fm As Double
    p = phaseDeg * 3.14159265358979 / 180#
    Select Case mode
        Case 1
            SignalAt = 0.78 * Sin(2 * 3.14159265358979 * f0 * t + p) + 0.46 * Sin(2 * 3.14159265358979 * f1 * t - p * 0.45)
        Case 2
            SignalAt = 0.78 * Sin(2 * 3.14159265358979 * f0 * t + p) + 0.28 * Sin(2 * 3.14159265358979 * 3 * f0 * t + p * 0.35) + 0.16 * Sin(2 * 3.14159265358979 * 5 * f0 * t - p * 0.25)
        Case 3
            fm = f1
            If fm > f0 * 0.9 Then fm = f0 * 0.9
            If fm < 10 Then fm = 10
            SignalAt = (0.72 + 0.28 * Cos(2 * 3.14159265358979 * fm * t)) * Sin(2 * 3.14159265358979 * f0 * t + p)
        Case Else
            SignalAt = Sin(2 * 3.14159265358979 * f0 * t + p)
    End Select
End Function

Private Function CurrentFMax(ByVal s As Slide) As Double
    Dim f0 As Double, f1 As Double, mode As Long, fm As Double
    f0 = Val(GetTag(s, "f0", "120"))
    f1 = Val(GetTag(s, "f1", "260"))
    mode = CLng(Val(GetTag(s, "mode", "0")))
    Select Case mode
        Case 1: CurrentFMax = IIf(f1 > f0, f1, f0)
        Case 2: CurrentFMax = 5 * f0
        Case 3
            fm = f1
            If fm > f0 * 0.9 Then fm = f0 * 0.9
            If fm < 10 Then fm = 10
            CurrentFMax = f0 + fm
        Case Else: CurrentFMax = f0
    End Select
End Function

Private Function Sinc(ByVal x As Double) As Double
    If Abs(x) < 0.0000001 Then
        Sinc = 1
    Else
        Sinc = Sin(3.14159265358979 * x) / (3.14159265358979 * x)
    End If
End Function

Private Function NextDynName() As String
    dynId = dynId + 1
    NextDynName = "dyn_" & CStr(dynId)
End Function

Private Sub AddDynLine(ByVal s As Slide, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal colorHex As String, Optional ByVal weight As Double = 1.2, Optional ByVal dash As Boolean = False)
    Dim shp As Shape
    Set shp = s.Shapes.AddLine(SX(s, x1), SY(s, y1), SX(s, x2), SY(s, y2))
    shp.Name = NextDynName()
    shp.Line.ForeColor.RGB = RgbHex(colorHex)
    shp.Line.Weight = weight * Sc(s)
    If dash Then shp.Line.DashStyle = msoLineDash
End Sub

Private Sub AddDynText(ByVal s As Slide, ByVal txt As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal size As Double, ByVal colorHex As String, Optional ByVal bold As Boolean = False, Optional ByVal align As Long = ppAlignLeft)
    Dim shp As Shape
    Set shp = AddText(s, NextDynName(), txt, x, y, w, h, size, colorHex, bold, align)
End Sub

Private Sub AddDynDot(ByVal s As Slide, ByVal cx As Double, ByVal cy As Double, ByVal r As Double, ByVal outlineHex As String)
    Dim shp As Shape
    Set shp = s.Shapes.AddShape(msoShapeOval, SX(s, cx - r), SY(s, cy - r), SX(s, 2 * r), SY(s, 2 * r))
    shp.Name = NextDynName()
    shp.Fill.ForeColor.RGB = RgbHex("#ffffff")
    shp.Line.ForeColor.RGB = RgbHex(outlineHex)
    shp.Line.Weight = 1.4 * Sc(s)
End Sub

Private Sub DrawFrame(ByVal s As Slide, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal yMin As Double, ByVal yMax As Double, Optional ByVal bandFn As Double = -1, Optional ByVal fRange As Double = 1)
    Dim shp As Shape, i As Long, yy As Double, xx As Double
    Set shp = AddBox(s, NextDynName(), x, y, w, h, "#f8fbff", "#c8d2de", False)
    If bandFn > 0 Then
        Dim bx1 As Double, bx2 As Double
        bx1 = x + (0.5 - bandFn / (2 * fRange)) * w
        bx2 = x + (0.5 + bandFn / (2 * fRange)) * w
        Set shp = AddBox(s, NextDynName(), bx1, y, bx2 - bx1, h, "#dff9f6", "#dff9f6", False)
        shp.Fill.Transparency = 0.15
        AddDynText s, "理想低通通带", bx1 + 40, y + 6, bx2 - bx1 - 80, 14, 8.5, "#5d6b78", False, ppAlignCenter
    End If
    For i = 1 To 4
        yy = y + h * i / 5
        AddDynLine s, x, yy, x + w, yy, "#dce3eb", 0.6, True
    Next i
    For i = 1 To 5
        xx = x + w * i / 6
        AddDynLine s, xx, y, xx, y + h, "#dce3eb", 0.6, True
    Next i
    If yMin < 0 And yMax > 0 Then
        yy = y + (yMax / (yMax - yMin)) * h
        AddDynLine s, x, yy, x + w, yy, "#9aa7b4", 0.8, False
    End If
End Sub

Private Function MapY(ByVal yVal As Double, ByVal top As Double, ByVal h As Double, ByVal yMin As Double, ByVal yMax As Double) As Double
    MapY = top + (yMax - yVal) / (yMax - yMin) * h
End Function

Private Sub DrawTimeChart(ByVal s As Slide, ByVal reconMode As Long)
    Dim f0 As Double, f1 As Double, fs As Double, phase As Double, mode As Long
    Dim x As Double, y As Double, w As Double, h As Double, dur As Double
    Dim i As Long, steps As Long, t1 As Double, t2 As Double, v1 As Double, v2 As Double
    Dim px1 As Double, py1 As Double, px2 As Double, py2 As Double, n As Long, tn As Double, yn As Double
    f0 = Val(GetTag(s, "f0", "120")): f1 = Val(GetTag(s, "f1", "260")): fs = Val(GetTag(s, "fs", "400"))
    phase = Val(GetTag(s, "phase", "0")): mode = CLng(Val(GetTag(s, "mode", "0")))
    x = 252: y = 88: w = 670: h = 80: dur = 0.1
    DrawFrame s, x, y, w, h, -1.35, 1.35
    steps = 96
    For i = 0 To steps - 1
        t1 = dur * i / steps: t2 = dur * (i + 1) / steps
        v1 = SignalAt(t1, f0, f1, phase, mode): v2 = SignalAt(t2, f0, f1, phase, mode)
        px1 = x + t1 / dur * w: px2 = x + t2 / dur * w
        py1 = MapY(v1, y, h, -1.35, 1.35): py2 = MapY(v2, y, h, -1.35, 1.35)
        AddDynLine s, px1, py1, px2, py2, "#2867d6", 1.4
    Next i
    For n = 0 To CLng(dur * fs)
        tn = n / fs
        If tn >= 0 And tn <= dur Then
            yn = SignalAt(tn, f0, f1, phase, mode)
            AddDynDot s, x + tn / dur * w, MapY(yn, y, h, -1.35, 1.35), 3, "#ff3348"
        End If
    Next n
    AddDynText s, "时间 t (s)", x + w / 2 - 36, y + h + 8, 72, 14, 8, "#263241", False, ppAlignCenter
    AddDynText s, "采样周期 Ts = " & Format(1000 / fs, "0.000") & " ms", 782, 84, 132, 18, 9.5, "#101820", True
    AddDynText s, "窗口点数 = " & CStr(CLng(dur * fs) + 1), 782, 110, 132, 18, 9.5, "#e64b9b", True
    AddDynText s, "fmax = " & FormatHz(CurrentFMax(s)), 782, 136, 132, 18, 9.5, "#e64b9b", True
End Sub

Private Function ReconAt(ByVal t As Double, ByRef ts() As Double, ByRef ys() As Double, ByVal fs As Double, ByVal method As Long, ByVal count As Long) As Double
    Dim i As Long, sum As Double, best As Double
    If method = 1 Then
        best = ys(0)
        For i = 0 To count - 1
            If ts(i) <= t Then best = ys(i)
        Next i
        ReconAt = best
        Exit Function
    End If
    For i = 0 To count - 1
        If Abs((t - ts(i)) * fs) <= 12 Then
            sum = sum + ys(i) * Sinc((t - ts(i)) * fs)
        End If
    Next i
    ReconAt = sum
End Function

Private Sub DrawReconChart(ByVal s As Slide)
    Dim f0 As Double, f1 As Double, fs As Double, phase As Double, mode As Long, method As Long
    Dim x As Double, y As Double, w As Double, h As Double, dur As Double
    Dim nStart As Long, nEnd As Long, n As Long, idx As Long, count As Long
    Dim ts() As Double, ys() As Double, t1 As Double, t2 As Double, v1 As Double, v2 As Double
    Dim i As Long, steps As Long, truth1 As Double, truth2 As Double
    f0 = Val(GetTag(s, "f0", "120")): f1 = Val(GetTag(s, "f1", "260")): fs = Val(GetTag(s, "fs", "400"))
    phase = Val(GetTag(s, "phase", "0")): mode = CLng(Val(GetTag(s, "mode", "0"))): method = CLng(Val(GetTag(s, "method", "0")))
    x = 252: y = 202: w = 670: h = 80: dur = 0.1
    DrawFrame s, x, y, w, h, -1.35, 1.35
    nStart = -12: nEnd = CLng(dur * fs) + 12
    count = nEnd - nStart + 1
    ReDim ts(0 To count - 1): ReDim ys(0 To count - 1)
    idx = 0
    For n = nStart To nEnd
        ts(idx) = n / fs
        ys(idx) = SignalAt(ts(idx), f0, f1, phase, mode)
        idx = idx + 1
    Next n
    steps = 84
    For i = 0 To steps - 1
        t1 = dur * i / steps: t2 = dur * (i + 1) / steps
        truth1 = SignalAt(t1, f0, f1, phase, mode): truth2 = SignalAt(t2, f0, f1, phase, mode)
        AddDynLine s, x + t1 / dur * w, MapY(truth1, y, h, -1.35, 1.35), x + t2 / dur * w, MapY(truth2, y, h, -1.35, 1.35), "#2867d6", 1.1
        v1 = ReconAt(t1, ts, ys, fs, method, count): v2 = ReconAt(t2, ts, ys, fs, method, count)
        AddDynLine s, x + t1 / dur * w, MapY(v1, y, h, -1.35, 1.35), x + t2 / dur * w, MapY(v2, y, h, -1.35, 1.35), "#e64b9b", 1.3, True
    Next i
    For n = 0 To CLng(dur * fs)
        AddDynDot s, x + (n / fs) / dur * w, MapY(SignalAt(n / fs, f0, f1, phase, mode), y, h, -1.35, 1.35), 3, "#ff3348"
    Next n
    AddDynText s, IIf(method = 0, "重建：Sinc 插值", "重建：零阶保持"), 770, 206, 150, 22, 9.2, "#101820", True
End Sub

Private Sub DrawPeak(ByVal s As Slide, ByVal f As Double, ByVal amp As Double, ByVal range As Double, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal colorHex As String, Optional ByVal label As String = "", Optional ByVal pale As Boolean = False)
    If Abs(f) > range Then Exit Sub
    Dim px As Double, py As Double
    px = x + (f + range) / (2 * range) * w
    py = y + h * (1 - amp / 1.15)
    AddDynLine s, px, y + h, px, py, colorHex, IIf(pale, 0.8, 1.6)
    If Not pale Then AddDynDot s, px, py, 2.6, colorHex
    If Len(label) > 0 Then AddDynText s, label, px - 28, py - 18, 56, 14, 8, colorHex, True, ppAlignCenter
End Sub

Private Sub DrawSpectrumChart(ByVal s As Slide)
    Dim f0 As Double, f1 As Double, fs As Double, mode As Long, fmax As Double, fn As Double, range As Double
    Dim x As Double, y As Double, w As Double, h As Double, k As Long, aliasF As Double, nearestK As Long
    f0 = Val(GetTag(s, "f0", "120")): f1 = Val(GetTag(s, "f1", "260")): fs = Val(GetTag(s, "fs", "400"))
    mode = CLng(Val(GetTag(s, "mode", "0")))
    fmax = CurrentFMax(s): fn = fs / 2
    range = 500
    If fs * 1.15 > range Then range = fs * 1.15
    If fmax * 1.35 > range Then range = fmax * 1.35
    x = 252: y = 316: w = 670: h = 80
    DrawFrame s, x, y, w, h, 0, 1.15, fn, range
    For k = -2 To 2
        If k <> 0 Then
            DrawPeak s, k * fs + f0, 0.42, range, x, y, w, h, "#ffae2b", "", True
            DrawPeak s, k * fs - f0, 0.42, range, x, y, w, h, "#ffae2b", "", True
        End If
    Next k
    DrawPeak s, f0, 1, range, x, y, w, h, "#2867d6", "f0"
    DrawPeak s, -f0, 1, range, x, y, w, h, "#2867d6"
    DrawPeak s, fn, 1.08, range, x, y, w, h, "#ffae2b", "fN"
    DrawPeak s, -fn, 1.08, range, x, y, w, h, "#ffae2b"
    aliasF = AliasOf(f0, fs, nearestK)
    If f0 > fn Then
        DrawPeak s, aliasF, 1.04, range, x, y, w, h, "#e64b9b", FormatHz(aliasF)
        DrawPeak s, -aliasF, 1.04, range, x, y, w, h, "#e64b9b"
    End If
    AddDynText s, "频谱间隔 = " & FormatHz(fs), 780, 320, 140, 18, 9.2, "#e64b9b", True
    If f0 > fn Then
        AddDynText s, "混叠：" & FormatHz(f0) & " → " & FormatHz(aliasF), 780, 346, 140, 18, 9.2, "#ff4f5f", True
    Else
        AddDynText s, "副本未重叠", 780, 346, 140, 18, 9.2, "#008a3d", True
    End If
End Sub

Private Function AliasOf(ByVal f As Double, ByVal fs As Double, ByRef kOut As Long) As Double
    Dim k As Long, a As Double
    If f >= 0 Then
        k = CLng(Int(f / fs + 0.5))
    Else
        k = CLng(-Int(Abs(f / fs) + 0.5))
    End If
    a = Abs(f - k * fs)
    Do While a > fs / 2
        a = Abs(a - fs)
    Loop
    kOut = k
    AliasOf = a
End Function

Private Sub DrawTheoryPanel(ByVal s As Slide)
    Dim f0 As Double, fs As Double, fmax As Double, fn As Double, aliasF As Double, k As Long
    f0 = Val(GetTag(s, "f0", "120")): fs = Val(GetTag(s, "fs", "400")): fmax = CurrentFMax(s): fn = fs / 2
    aliasF = AliasOf(f0, fs, k)
    AddBox s, NextDynName(), 252, 414, 670, 92, "#101722", "#304052", True
    AddDynText s, "采样定理机理", 274, 428, 150, 18, 12.5, "#f2f7ff", True
    AddDynText s, "时域：xs(t)=x(t)·Σδ(t-nTs)    频域：Xs(f)=fs·ΣX(f-kfs)    重建：x_hat(t)=Σx[n]sinc(fs t-n)", 274, 456, 620, 18, 10.5, "#ffe1a3", True
    If fs >= 2 * fmax Then
        AddDynText s, "当前 " & FormatHz(fs) & " ≥ 2×" & FormatHz(fmax) & "，中心频谱副本与相邻副本分开，理想低通可以还原原信号。", 274, 482, 610, 18, 10.5, "#d6e2ef"
    Else
        AddDynText s, "当前 " & FormatHz(fs) & " < 2×" & FormatHz(fmax) & "；例如 |" & FormatHz(f0) & " - " & CStr(k) & "×" & FormatHz(fs) & "| = " & FormatHz(aliasF) & "，重建会出现假频率。", 274, 482, 610, 18, 10.5, "#ffd7de"
    End If
End Sub
