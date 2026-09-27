Attribute VB_Name = "M_Seal"
Option Explicit

'==========================================================
' 印章制作
'   生成圆形印章：外圆 + 环形文字 + 五角星 + 下部文字
'   统一用红色，按页面中心摆放
'
'   X4 实测：
'     CreateEllipse2(cx, cy, rx, ry)  (cx,cy)=中心 rx/ry=半径
'     CreatePolygon / CreatePolygon2 不能直接设尖锐度
'     五角星做法 = 十边形 → ConvertToCurves → 隔一个顶点拉向圆心
'     变量名不能叫 circle（撞保留字）
'==========================================================

' 带对话框的入口
Public Sub CreateSeal()
    Dim txt1 As String
    Dim txt2 As String
    Dim n As Long

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "印章制作"
        Exit Sub
    End If

    txt1 = InputBox("输入印章上方的环形文字（单位名称）。", "印章制作", "某某市某某有限公司")
    If Len(txt1) = 0 Then Exit Sub
    txt2 = InputBox("输入印章下方的文字，可留空。", "印章制作", "财务专用章")

    n = CreateSealCore(txt1, txt2)

    If n = 0 Then
        MsgBox "印章没生成成功，请把报错反馈给我。", vbExclamation, "印章制作"
    Else
        MsgBox "印章已生成，共 " & n & " 个对象，摆在本页中心。", vbInformation, "印章制作"
    End If
End Sub

' 纯逻辑，供自检调用
Public Function CreateSealCore(ByVal txt1 As String, ByVal txt2 As String) As Long
    Dim doc As Document
    Dim pg As Page
    Dim lay As Object
    Dim lo As Object
    Dim w As Double
    Dim h As Double
    Dim cx As Double
    Dim cy As Double
    Dim r As Double
    Dim cxD As Double
    Dim cyD As Double
    Dim rD As Double
    Dim red As Color
    Dim outerCircle As Shape
    Dim t As Shape
    Dim star As Shape
    Dim n As Long

    Set doc = CorelDRAW.ActiveDocument
    If doc Is Nothing Then Exit Function

    Set pg = ActivePageSafe()
    If pg Is Nothing Then Exit Function
    Set lay = pg.ActiveLayer
    Set lo = lay

    PageWHmm pg, w, h
    If w <= 0 Then w = 210
    If h <= 0 Then h = 297

    cx = w / 2
    cy = h / 2
    r = 20                       ' 印章半径 20mm（直径 40mm）

    cxD = MmToDoc(cx)
    cyD = MmToDoc(cy)
    rD = MmToDoc(r)

    Set red = MakeCMYK(0, 100, 100, 0)
    If red Is Nothing Then Set red = MakeRGB(200, 0, 0)

    CorelDRAW.Optimization = True
    doc.BeginCommandGroup "印章制作"

    ' --- 1. 外圆：无填充 + 红色轮廓 ---
    On Error Resume Next
    Err.Clear
    Set outerCircle = lo.CreateEllipse2(cxD, cyD, rD, rD)
    If Err.Number <> 0 Then Set outerCircle = Nothing
    Err.Clear
    On Error GoTo 0

    If outerCircle Is Nothing Then
        doc.EndCommandGroup
        CorelDRAW.Optimization = False
        Exit Function
    End If

    On Error Resume Next
    outerCircle.Fill.ApplyNoFill
    outerCircle.Outline.Color.CopyAssign red
    outerCircle.Outline.Width = MmToDoc(1.2)
    Err.Clear
    On Error GoTo 0
    n = n + 1

    ' --- 2. 环形文字 ---
    If Len(txt1) > 0 Then
        Set t = Nothing
        On Error Resume Next
        Err.Clear
        Set t = lo.CreateArtisticTextWide(cxD, cyD, txt1)
        If Err.Number <> 0 Then Set t = Nothing
        Err.Clear
        On Error GoTo 0

        If Not t Is Nothing Then
            SetTextStyle t, "宋体", 10, True
            On Error Resume Next
            Err.Clear
            t.Text.FitToPath outerCircle
            t.Fill.ApplyUniformFill red
            Err.Clear
            On Error GoTo 0
            n = n + 1
        End If
    End If

    ' --- 3. 下部文字 ---
    If Len(txt2) > 0 Then
        Set t = MakeCenteredText(lay, txt2, cx, cy - r * 0.45, 12, "宋体")
        If Not t Is Nothing Then
            On Error Resume Next
            t.Fill.ApplyUniformFill red
            Err.Clear
            On Error GoTo 0
            n = n + 1
        End If
    End If

    ' --- 4. 中间的五角星 ---
    Set star = MakeStar(lo, cxD, cyD, rD * 0.34, red)
    If Not star Is Nothing Then n = n + 1

    doc.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh

    CreateSealCore = n
End Function

' 五角星：十边形转曲后，把隔一个的顶点拉向圆心
Private Function MakeStar(ByVal lo As Object, ByVal cxD As Double, ByVal cyD As Double, _
                          ByVal rOuterD As Double, ByVal red As Color) As Shape
    Dim sh As Shape
    Dim cv As Object
    Dim j As Long
    Dim cnt As Long
    Dim px As Double
    Dim py As Double
    Dim ang As Double
    Dim rr As Double
    Dim rIn As Double
    Dim ok As Boolean

    If lo Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    Set sh = lo.CreatePolygon2(cxD, cyD, rOuterD, 10)
    ok = (Err.Number = 0) And (Not sh Is Nothing)
    Err.Clear
    If Not ok Then
        ' 退路：老式 CreatePolygon（第 5 个参数才是边数）
        Set sh = lo.CreatePolygon(cxD, cyD, rOuterD, 0, 10)
        ok = (Err.Number = 0) And (Not sh Is Nothing)
        Err.Clear
    End If
    On Error GoTo 0
    If Not ok Then Exit Function

    On Error Resume Next
    sh.ConvertToCurves
    Err.Clear
    On Error GoTo 0

    Set cv = ShapeCurve(sh)
    If cv Is Nothing Then
        ' 转曲失败：把十边形留着，至少不报错
        On Error Resume Next
        sh.Fill.ApplyUniformFill red
        sh.Outline.Width = 0
        Err.Clear
        On Error GoTo 0
        Set MakeStar = sh
        Exit Function
    End If

    rIn = rOuterD * 0.382           ' 正五角星内外半径比
    On Error Resume Next
    cnt = cv.Nodes.Count
    For j = 1 To cnt
        px = cv.Nodes(j).PositionX
        py = cv.Nodes(j).PositionY
        ang = Atan2(py - cyD, px - cxD)
        If (j Mod 2) = 0 Then
            rr = rIn
        Else
            rr = rOuterD
        End If
        cv.Nodes(j).SetPosition cxD + rr * Cos(ang), cyD + rr * Sin(ang)
        Err.Clear
    Next j

    sh.Fill.ApplyUniformFill red
    sh.Outline.Width = 0
    Err.Clear
    On Error GoTo 0

    Set MakeStar = sh
End Function