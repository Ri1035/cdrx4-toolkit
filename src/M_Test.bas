Attribute VB_Name = "M_Test"
Option Explicit

'==========================================================
' 自检
'   新建一个临时文档，把每个功能的「纯逻辑入口」跑一遍，
'   结果写进 %TEMP%\cdrx4_selftest.log
'
'   这个模块不挂工具栏按钮，只给构建 / 冒烟测试用。
'   跑完会把临时文档丢掉，不动用户正在编辑的文档。
'==========================================================

Public Sub SelfTest()
    Dim fso As Object
    Dim doc As Document
    Dim pg As Page
    Dim lay As Object
    Dim r As Shape
    Dim e As Shape
    Dim t As Shape
    Dim pathShape As Shape
    Dim objShape As Shape
    Dim c1 As Color
    Dim c2 As Color
    Dim n As Long
    Dim log As String
    Dim logPath As String
    Dim jpgDir As String
    ' 1b 转曲范围测试用的临时文档
    Dim d2 As Document
    Dim l2 As Object
    Dim nT As Long
    Dim nS As Long
    Dim nA As Long

    logPath = Environ$("TEMP") & "\cdrx4_selftest.log"
    log = "CDRX4Toolkit selftest " & Now & vbCrLf

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub

    On Error Resume Next
    Err.Clear
    Set doc = CorelDRAW.CreateDocument
    log = log & "0 create doc   err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    If doc Is Nothing Then
        WriteLog fso, logPath, log
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear
    Set pg = doc.ActivePage
    Set lay = pg.ActiveLayer
    log = log & "0 page/layer   err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    If lay Is Nothing Then
        WriteLog fso, logPath, log
        Exit Sub
    End If

    ' --- 造测试内容 ---
    On Error Resume Next
    Err.Clear
    Set r = lay.CreateRectangle2(0, 0, 40, 30)
    log = log & "0 rect          err=" & Err.Number & vbCrLf
    Err.Clear
    Set e = lay.CreateEllipse2(70, 15, 15, 15)
    log = log & "0 ellipse       err=" & Err.Number & vbCrLf
    Err.Clear
    Set t = lay.CreateArtisticTextWide(0, 60, "Test 123")
    log = log & "0 text          err=" & Err.Number & vbCrLf
    Err.Clear
    Set c1 = MakeRGB(200, 30, 40)
    If Not r Is Nothing Then r.Fill.ApplyUniformFill c1
    log = log & "0 fill          err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 1 全部转曲 ---
    On Error Resume Next
    Err.Clear
    n = M_Curves.ConvertAllToCurvesCore(True)
    log = log & "1 curves        n=" & n & " err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 2 转CMYK ---
    On Error Resume Next
    Err.Clear
    n = M_CMYK.ConvertToCMYKCore(True)
    log = log & "2 cmyk          n=" & n & " err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 3 颜色替换 ---
    On Error Resume Next
    Err.Clear
    Set r = lay.CreateRectangle2(120, 0, 20, 20)
    Set e = lay.CreateRectangle2(150, 0, 20, 20)
    Set c1 = MakeRGB(200, 30, 40)
    Set c2 = MakeRGB(10, 200, 40)
    If Not r Is Nothing Then r.Fill.ApplyUniformFill c1
    If Not e Is Nothing Then e.Fill.ApplyUniformFill c2
    doc.ClearSelection
    If Not r Is Nothing Then r.AddToSelection
    If Not e Is Nothing Then e.AddToSelection
    log = log & "3 select        err=" & Err.Number & vbCrLf
    Err.Clear
    n = M_Color.ReplaceColorCore(c1, c2)
    log = log & "3 color         n=" & n & " err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 4 标准矩形 ---
    On Error Resume Next
    Err.Clear
    doc.ClearSelection
    If Not r Is Nothing Then r.AddToSelection
    n = M_Rect.StdRectangleCore()
    log = log & "4 rect          n=" & n & " err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 5 插入页码 ---
    On Error Resume Next
    Err.Clear
    n = M_PageNo.InsertPageNumberCore()
    log = log & "5 pageno        n=" & n & " err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 6 日历 ---
    On Error Resume Next
    Err.Clear
    n = M_Calendar.CreateCalendarCore(2026, 9)
    log = log & "6 calendar      n=" & n & " err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 7 印章 ---
    On Error Resume Next
    Err.Clear
    n = M_Seal.CreateSealCore("测试单位名称", "财务专用章")
    log = log & "7 seal          n=" & n & " err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 8 对象适合路径 ---
    On Error Resume Next
    Err.Clear
    Set pathShape = lay.CreateEllipse2(220, 120, 40, 25)
    If Not pathShape Is Nothing Then pathShape.ConvertToCurves
    Set objShape = lay.CreateRectangle2(0, 200, 10, 10)
    doc.ClearSelection
    If Not pathShape Is Nothing Then pathShape.AddToSelection
    If Not objShape Is Nothing Then objShape.AddToSelection
    log = log & "8 select        err=" & Err.Number & vbCrLf
    Err.Clear
    n = M_FitPath.FitToPathCore(8)
    log = log & "8 fitpath       n=" & n & " err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 9 JPG 批量导出 ---
    jpgDir = Environ$("TEMP") & "\cdrx4_selftest_jpg"
    On Error Resume Next
    Err.Clear
    If Not fso.FolderExists(jpgDir) Then fso.CreateFolder jpgDir
    Err.Clear
    n = M_JPG.ExportJPGTo(jpgDir, 150)
    log = log & "9 jpg           n=" & n & " err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- 1b 转曲范围：仅文字 / 仅图形 / 全部 ---
    ' 每个范围都要一份「还没转过」的对象，所以另开一个临时文档。
    ' 三步的期望值互相印证：
    '   仅文字 → 文字变曲线，转 1 个
    '   仅图形 → 转 3 个（矩形 / 椭圆 / 多边形；上一步那根曲线应被跳过）
    '   全部   → 此时已全是曲线，转 0 个
    On Error Resume Next
    Err.Clear
    Set d2 = CorelDRAW.CreateDocument
    If Err.Number = 0 And Not d2 Is Nothing Then
        Set l2 = d2.ActivePage.ActiveLayer
    End If
    Err.Clear
    On Error GoTo 0

    If l2 Is Nothing Then
        log = log & "1b scope        SKIPPED (no doc)" & vbCrLf
    Else
        On Error Resume Next
        Err.Clear
        l2.CreateArtisticTextWide 0, 0, "Scope"
        l2.CreateRectangle2 0, 20, 20, 10
        l2.CreateEllipse2 40, 25, 5, 5
        l2.CreatePolygon2 60, 25, 5, 6
        log = log & "1b build        err=" & Err.Number & vbCrLf
        Err.Clear
        On Error GoTo 0

        nT = M_Curves.ConvertAllToCurvesCore(False, 1)
        log = log & "1b scope text   n=" & nT & " want=1" & vbCrLf
        nS = M_Curves.ConvertAllToCurvesCore(False, 2)
        log = log & "1b scope shape  n=" & nS & " want=3" & vbCrLf
        nA = M_Curves.ConvertAllToCurvesCore(False, 0)
        log = log & "1b scope all    n=" & nA & " want=0" & vbCrLf

        On Error Resume Next
        d2.Dirty = False
        d2.Close
        On Error GoTo 0
    End If

    log = log & "done" & vbCrLf
    WriteLog fso, logPath, log

    ' 丢掉临时文档，别打扰用户正在编辑的文件
    On Error Resume Next
    doc.Dirty = False
    doc.Close
    On Error GoTo 0
End Sub

Private Sub WriteLog(ByVal fso As Object, ByVal p As String, ByVal s As String)
    Dim ts As Object

    On Error Resume Next
    Set ts = fso.CreateTextFile(p, True, True)
    If Not ts Is Nothing Then
        ts.Write s
        ts.Close
    End If
    On Error GoTo 0
End Sub

'==========================================================
' 详细探针：把「编译能过但运行结果不对」的那几处语义问清楚
' 结果写 %TEMP%\cdrx4_probe.log
'==========================================================
Public Sub Probe()
    Dim fso As Object
    Dim doc As Document
    Dim pg As Page
    Dim lay As Object
    Dim sr As Object
    Dim sh As Shape
    Dim r As Shape
    Dim e As Shape
    Dim t As Shape
    Dim r2 As Shape
    Dim t2 As Shape
    Dim p2 As Shape
    Dim c As Color
    Dim ro As Object
    Dim cv As Object
    Dim o As Object
    Dim log As String
    Dim logPath As String
    Dim i As Long

    logPath = Environ$("TEMP") & "\cdrx4_probe.log"
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub

    log = "CDRX4Toolkit probe " & Now & vbCrLf
    log = log & "CONST uniformFill=" & cdrUniformFill & " text=" & cdrTextShape & _
          " rect=" & cdrRectangleShape & " ellipse=" & cdrEllipseShape & _
          " poly=" & cdrPolygonShape & " curve=" & cdrCurveShape & _
          " group=" & cdrGroupShape & " centerAlign=" & cdrCenterAlignment & vbCrLf
    log = log & "CONST curPage=" & cdrCurrentPage & " jpeg=" & cdrJPEG & _
          " rgbImg=" & cdrRGBColorImage & " aa=" & cdrNormalAntiAliasing & vbCrLf

    On Error Resume Next
    Err.Clear
    Set doc = CorelDRAW.CreateDocument
    log = log & "createDoc err=" & Err.Number & vbCrLf
    Err.Clear
    Set pg = doc.ActivePage
    Set lay = pg.ActiveLayer
    log = log & "page err=" & Err.Number & " pages=" & doc.Pages.Count & vbCrLf
    Err.Clear
    On Error GoTo 0

    If lay Is Nothing Then
        WriteLog fso, logPath, log
        Exit Sub
    End If

    ' --- A. 造形状 + Type ---
    On Error Resume Next
    Err.Clear
    Set r = lay.CreateRectangle2(0, 0, 40, 30)
    log = log & "A rect    err=" & Err.Number & " type=" & r.Type & vbCrLf
    Err.Clear
    Set e = lay.CreateEllipse2(70, 15, 15, 15)
    log = log & "A ellipse err=" & Err.Number & " type=" & e.Type & vbCrLf
    Err.Clear
    Set t = lay.CreateArtisticTextWide(0, 60, "Test 123")
    log = log & "A text    err=" & Err.Number & " type=" & t.Type & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- B. Shapes.All 的 Count / Item ---
    On Error Resume Next
    Err.Clear
    Set sr = pg.Shapes.All
    log = log & "B Shapes.All typename=" & TypeName(sr) & " err=" & Err.Number & vbCrLf
    Err.Clear
    log = log & "B sr.Count=" & sr.Count & " err=" & Err.Number & vbCrLf
    Err.Clear
    For i = 1 To sr.Count
        Set sh = sr.Item(i)
        log = log & "B   item " & i & " err=" & Err.Number & " type=" & sh.Type & vbCrLf
        Err.Clear
    Next i
    On Error GoTo 0

    ' --- C. 填充 + 转 CMYK ---
    On Error Resume Next
    Err.Clear
    Set c = MakeRGB(200, 30, 40)
    log = log & "C makeRGB err=" & Err.Number & " type=" & c.Type & vbCrLf
    Err.Clear
    r.Fill.ApplyUniformFill c
    log = log & "C applyUniformFill err=" & Err.Number & vbCrLf
    Err.Clear
    Set o = r
    log = log & "C fill.type=" & o.Fill.Type & " err=" & Err.Number & vbCrLf
    Err.Clear
    Set c = Nothing
    Set c = o.Fill.UniformColor
    log = log & "C uniformColor err=" & Err.Number & " nothing=" & (c Is Nothing) & vbCrLf
    Err.Clear
    If Not c Is Nothing Then
        log = log & "C color.type=" & c.Type & " err=" & Err.Number & vbCrLf
        Err.Clear
        c.ConvertToCMYK
        log = log & "C convertToCMYK err=" & Err.Number & " typeAfter=" & c.Type & vbCrLf
        Err.Clear
        o.Fill.ApplyUniformFill c
        log = log & "C reapply err=" & Err.Number & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0

    ' --- D. ConvertToCurves ---
    On Error Resume Next
    Err.Clear
    r.ConvertToCurves
    log = log & "D rect ConvertToCurves err=" & Err.Number & " typeAfter=" & r.Type & vbCrLf
    Err.Clear
    t.ConvertToCurves
    log = log & "D text ConvertToCurves err=" & Err.Number & " typeAfter=" & t.Type & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- E. 矩形圆角属性 ---
    On Error Resume Next
    Err.Clear
    Set r2 = lay.CreateRectangle2(200, 0, 30, 30)
    log = log & "E CreateRectangle2 err=" & Err.Number & " type=" & r2.Type & vbCrLf
    Err.Clear
    Set ro = Nothing
    Set ro = r2.Rectangle
    log = log & "E .Rectangle err=" & Err.Number & " nothing=" & (ro Is Nothing) & vbCrLf
    Err.Clear
    If Not ro Is Nothing Then
        ro.RadiusUpperLeft = 0
        log = log & "E set RadiusUpperLeft err=" & Err.Number & vbCrLf
        Err.Clear
        ro.RadiusUpperRight = 5
        log = log & "E set RadiusUpperRight=5 err=" & Err.Number & _
              " read=" & ro.RadiusUpperRight & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0

    ' --- F. 选择集 ---
    On Error Resume Next
    Err.Clear
    doc.ClearSelection
    r2.AddToSelection
    log = log & "F AddToSelection err=" & Err.Number & vbCrLf
    Err.Clear
    Set sr = Nothing
    Set sr = CorelDRAW.ActiveSelectionRange
    log = log & "F ActiveSelectionRange err=" & Err.Number & _
          " nothing=" & (sr Is Nothing) & vbCrLf
    Err.Clear
    If Not sr Is Nothing Then
        log = log & "F sr.Count=" & sr.Count & " err=" & Err.Number & vbCrLf
        Err.Clear
        log = log & "F item1 type=" & sr.Item(1).Type & " err=" & Err.Number & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0

    ' --- G. CenterAt 语义 ---
    On Error Resume Next
    Err.Clear
    log = log & "G before L=" & r2.LeftX & " B=" & r2.BottomY & _
          " W=" & r2.SizeWidth & " H=" & r2.SizeHeight & vbCrLf
    Err.Clear
    CenterAt r2, 500, 400
    log = log & "G after  L=" & r2.LeftX & " B=" & r2.BottomY & _
          " W=" & r2.SizeWidth & " H=" & r2.SizeHeight & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- H. 十边形 → 转曲 → 节点定位 ---
    On Error Resume Next
    Err.Clear
    Set p2 = lay.CreatePolygon2(300, 300, 50, 10)
    log = log & "H CreatePolygon2 err=" & Err.Number & " nothing=" & (p2 Is Nothing) & vbCrLf
    Err.Clear
    If Not p2 Is Nothing Then
        p2.ConvertToCurves
        log = log & "H ConvertToCurves err=" & Err.Number & vbCrLf
        Err.Clear
        Set cv = ShapeCurve(p2)
        log = log & "H ShapeCurve nothing=" & (cv Is Nothing) & " err=" & Err.Number & vbCrLf
        Err.Clear
        If Not cv Is Nothing Then
            log = log & "H nodes=" & cv.Nodes.Count & " err=" & Err.Number & vbCrLf
            Err.Clear
            log = log & "H node1 X=" & cv.Nodes(1).PositionX & _
                  " Y=" & cv.Nodes(1).PositionY & " err=" & Err.Number & vbCrLf
            Err.Clear
            cv.Nodes(1).SetPosition 300, 350
            log = log & "H node setpos err=" & Err.Number & vbCrLf
            Err.Clear
        End If
    End If
    On Error GoTo 0

    ' --- I. 文本适合路径 ---
    On Error Resume Next
    Err.Clear
    Set t2 = lay.CreateArtisticTextWide(0, 0, "AB")
    log = log & "I text err=" & Err.Number & vbCrLf
    Err.Clear
    t2.Text.FitToPath e
    log = log & "I FitToPath err=" & Err.Number & " type=" & t2.Type & vbCrLf
    Err.Clear
    On Error GoTo 0

    ' --- J. 轮廓宽度 / 无填充 ---
    On Error Resume Next
    Err.Clear
    r2.Outline.Width = 0
    log = log & "J outline.width=0 err=" & Err.Number & vbCrLf
    Err.Clear
    r2.Fill.ApplyNoFill
    log = log & "J applyNoFill err=" & Err.Number & vbCrLf
    Err.Clear
    On Error GoTo 0

    log = log & "done" & vbCrLf
    WriteLog fso, logPath, log

    On Error Resume Next
    doc.Dirty = False
    doc.Close
    On Error GoTo 0
End Sub

'==========================================================
' 深挖：为什么「全部转曲 / 转CMYK / 标准矩形」返回 0
' 结果写 %TEMP%\cdrx4_diag.log
'
' 这里每个取值都走 S()，因为 log = log & ... & obj.Prop 这种写法
' 一旦某个属性取值抛错，整个赋值会被中止 —— 整行日志凭空消失，
' 之前那轮 probe 就是这么把 B 段的逐项输出吃掉的。
'==========================================================
Public Sub Diag()
    Dim fso As Object
    Dim d As Document
    Dim adoc As Object
    Dim pg As Page
    Dim lay As Object
    Dim sr As Object
    Dim sh As Shape
    Dim r As Shape
    Dim e As Shape
    Dim t As Shape
    Dim ro As Object
    Dim log As String
    Dim logPath As String
    Dim i As Long
    Dim n As Long

    logPath = Environ$("TEMP") & "\cdrx4_diag.log"
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub

    log = "diag " & Now & vbCrLf

    On Error Resume Next
    Err.Clear
    Set d = CorelDRAW.CreateDocument
    log = log & "D1  createDoc      nothing=" & (d Is Nothing) & _
          "  err=" & S(Err.Number) & " " & S(Err.Description) & vbCrLf
    Err.Clear

    Set adoc = Nothing
    Set adoc = CorelDRAW.ActiveDocument
    log = log & "D2  ActiveDocument nothing=" & (adoc Is Nothing) & _
          "  same=" & S(d Is adoc) & "  err=" & S(Err.Number) & vbCrLf
    Err.Clear
    log = log & "D3  Documents.Count=" & S(CorelDRAW.Documents.Count) & _
          "  err=" & S(Err.Number) & vbCrLf
    Err.Clear
    log = log & "D4  d.Pages.Count=" & S(d.Pages.Count) & _
          "  err=" & S(Err.Number) & vbCrLf
    Err.Clear

    Set pg = Nothing
    Set pg = d.ActivePage
    log = log & "D5  d.ActivePage   nothing=" & (pg Is Nothing) & _
          "  err=" & S(Err.Number) & vbCrLf
    Err.Clear
    Set lay = Nothing
    Set lay = pg.ActiveLayer
    log = log & "D6  ActiveLayer    nothing=" & (lay Is Nothing) & _
          "  err=" & S(Err.Number) & vbCrLf
    Err.Clear

    If lay Is Nothing Then
        log = log & "!! 没有图层，到此为止" & vbCrLf
        WriteLog fso, logPath, log
        Exit Sub
    End If

    Set r = lay.CreateRectangle2(0, 0, 40, 30)
    log = log & "D7  rect           nothing=" & (r Is Nothing) & _
          "  type=" & S(r.Type) & "  err=" & S(Err.Number) & vbCrLf
    Err.Clear
    Set e = lay.CreateEllipse2(70, 15, 15, 15)
    log = log & "D8  ellipse        nothing=" & (e Is Nothing) & _
          "  type=" & S(e.Type) & "  err=" & S(Err.Number) & vbCrLf
    Err.Clear
    Set t = lay.CreateArtisticTextWide(0, 60, "AB")
    log = log & "D9  text           nothing=" & (t Is Nothing) & _
          "  type=" & S(t.Type) & "  err=" & S(Err.Number) & vbCrLf
    Err.Clear

    ' --- For Each 走一遍 Pages ---
    n = 0
    For Each pg In d.Pages
        n = n + 1
        log = log & "D10 page " & n & "  shapes=" & S(pg.Shapes.All.Count) & _
              "  err=" & S(Err.Number) & vbCrLf
        Err.Clear
    Next pg
    log = log & "D11 遍历页数=" & n & "  err=" & S(Err.Number) & vbCrLf
    Err.Clear

    ' --- ShapeRange 的两种取法 ---
    Set pg = d.ActivePage
    Set sr = Nothing
    Set sr = pg.Shapes.All
    log = log & "D12 Shapes.All     nothing=" & (sr Is Nothing) & _
          "  count=" & S(sr.Count) & "  err=" & S(Err.Number) & vbCrLf
    Err.Clear
    If Not sr Is Nothing Then
        For i = 1 To S(sr.Count)
            Set sh = Nothing
            Set sh = sr.Item(i)
            log = log & "D13  .Item(" & i & ")     nothing=" & (sh Is Nothing) & _
                  "  type=" & S(sh.Type) & "  err=" & S(Err.Number) & _
                  " " & S(Err.Description)
            Err.Clear
            Set sh = Nothing
            Set sh = sr.Shapes(i)
            log = log & "   | .Shapes(" & i & ") nothing=" & (sh Is Nothing) & _
                  "  type=" & S(sh.Type) & "  err=" & S(Err.Number) & vbCrLf
            Err.Clear
        Next i
    End If

    ' --- 三个 Core 各自返回什么 ---
    n = M_Curves.ConvertAllToCurvesCore(True)
    log = log & "D14 curvesCore     n=" & n & "  err=" & S(Err.Number) & _
          " " & S(Err.Description) & vbCrLf
    Err.Clear
    log = log & "D15 转曲后 shapes=" & S(d.ActivePage.Shapes.All.Count) & _
          "  err=" & S(Err.Number) & vbCrLf
    Err.Clear

    n = M_CMYK.ConvertToCMYKCore(True)
    log = log & "D16 cmykCore       n=" & n & "  err=" & S(Err.Number) & _
          " " & S(Err.Description) & vbCrLf
    Err.Clear

    d.ClearSelection
    log = log & "D17 ClearSelection err=" & S(Err.Number) & vbCrLf
    Err.Clear
    If Not r Is Nothing Then r.AddToSelection
    log = log & "D18 AddToSelection err=" & S(Err.Number) & vbCrLf
    Err.Clear
    Set sr = Nothing
    Set sr = CorelDRAW.ActiveSelectionRange
    log = log & "D19 SelectionRange nothing=" & (sr Is Nothing) & _
          "  count=" & S(sr.Count) & "  err=" & S(Err.Number) & vbCrLf
    Err.Clear
    If Not sr Is Nothing Then
        Set sh = Nothing
        Set sh = sr.Item(1)
        log = log & "D20  sel.Item(1)   nothing=" & (sh Is Nothing) & _
              "  type=" & S(sh.Type) & "  err=" & S(Err.Number) & vbCrLf
        Err.Clear
        Set ro = Nothing
        Set ro = sh.Rectangle
        log = log & "D21  .Rectangle    nothing=" & (ro Is Nothing) & _
              "  err=" & S(Err.Number) & vbCrLf
        Err.Clear
    End If
    n = M_Rect.StdRectangleCore()
    log = log & "D22 rectCore       n=" & n & "  err=" & S(Err.Number) & _
          " " & S(Err.Description) & vbCrLf
    Err.Clear
    On Error GoTo 0

    log = log & "done" & vbCrLf
    WriteLog fso, logPath, log

    On Error Resume Next
    d.Dirty = False
    d.Close
    On Error GoTo 0
End Sub

' 安全转字符串：取值失败变成 <errNN>，绝不让整行日志消失
Private Function S(ByVal v As Variant) As String
    On Error Resume Next
    Err.Clear
    S = CStr(v)
    If Err.Number <> 0 Then S = "<err" & Err.Number & ">"
    Err.Clear
    On Error GoTo 0
End Function