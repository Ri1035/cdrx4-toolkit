Attribute VB_Name = "M_Seal"
Option Explicit

'==========================================================
' 印章制作
'   生成圆形印章：外圆 + 环形文字 + 五角星 + 下部文字
'   统一用红色，按页面中心摆放
'==========================================================

Public Sub CreateSeal()
    Dim doc As Document
    Dim pg As Page
    Dim lay As Layer
    Dim w As Double
    Dim h As Double
    Dim txt1 As String
    Dim txt2 As String
    Dim cx As Double
    Dim cy As Double
    Dim r As Double
    Dim red As Color
    Dim circle As Shape
    Dim t As Shape
    Dim star As Shape

    If Not HasDocument() Then Exit Sub

    txt1 = InputBox("输入印章上方的环形文字（单位名称）。", "印章制作", "某某市某某有限公司")
    If Len(txt1) = 0 Then Exit Sub
    txt2 = InputBox("输入印章下方的文字，可留空。", "印章制作", "财务专用章")

    Set doc = CorelDRAW.ActiveDocument
    Set pg = doc.ActivePage
    Set lay = pg.ActiveLayer

    PageWH pg, w, h
    If w <= 0 Then w = 210
    If h <= 0 Then h = 297

    cx = w / 2
    cy = h / 2
    r = 20

    Set red = CreateRGBColor(255, 0, 0)

    Optimization = True
    doc.BeginCommandGroup "印章制作"

    ' 1. 外圆：无填充 + 红色细轮廓
    Set circle = Nothing
    On Error Resume Next
    Set circle = lay.CreateEllipse2(cx, cy, r, r)
    On Error GoTo 0

    If circle Is Nothing Then
        MsgBox "创建外圆失败，请把报错反馈给我。", vbExclamation, "印章制作"
        doc.EndCommandGroup
        Optimization = False
        Exit Sub
    End If

    ' 外圆本身不填色，只留红色轮廓
    On Error Resume Next
    circle.Fill.ApplyNoFill
    circle.Outline.Color.CopyAssign red
    circle.Outline.Width = 1.2
    On Error GoTo 0

    ' 2. 环形文字
    On Error Resume Next
    Set t = lay.CreateArtisticTextWide(cx, cy, txt1)
    If Not t Is Nothing Then
        t.Text.Story.Font = "微软雅黑"
        t.Text.Story.Size = 10
        t.Text.FitToPath circle
        t.Fill.ApplyUniformFill red
    End If
    Set t = Nothing
    On Error GoTo 0

    ' 3. 下部文字
    If Len(txt2) > 0 Then
        On Error Resume Next
        Set t = lay.CreateArtisticTextWide(cx, cy - 8, txt2)
        If Not t Is Nothing Then
            t.Text.Story.Font = "微软雅黑"
            t.Text.Story.Size = 12
            t.Text.Story.Alignment = cdrCenterAlignment
            t.SetPosition cx, cy - 8
            t.Fill.ApplyUniformFill red
        End If
        Set t = Nothing
        On Error GoTo 0
    End If

    ' 4. 中间的五角星
    On Error Resume Next
    Set star = lay.CreatePolygon(cx, cy, 5, 5)
    If Not star Is Nothing Then
        star.SetPolygonProperties 5, 53
        star.Fill.ApplyUniformFill red
        star.Outline.Width = 0
    End If
    Set star = Nothing
    On Error GoTo 0

    doc.EndCommandGroup
    Optimization = False
    DoRefresh
End Sub