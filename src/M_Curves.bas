Attribute VB_Name = "M_Curves"
Option Explicit

'==========================================================
' 全部转曲
'   把所有页面的文本与图形都转成曲线
'   可选是否深入 PowerClip 容器内部
'==========================================================

Public Sub ConvertAllToCurves()
    Dim doc As Document
    Dim pg As Page
    Dim ans As VbMsgBoxResult
    Dim deep As Boolean

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "全部转曲"
        Exit Sub
    End If

    ans = MsgBox("是否全部转曲？" & vbCrLf & vbCrLf & _
                 "是 = 转曲，并深入 PowerClip 容器内部一起处理" & vbCrLf & _
                 "否 = 只转曲顶层对象，不动容器内部" & vbCrLf & _
                 "取消 = 什么也不做", _
                 vbYesNoCancel + vbQuestion, "全部转曲")
    If ans = vbCancel Then Exit Sub
    deep = (ans = vbYes)

    Set doc = CorelDRAW.ActiveDocument

    Optimization = True
    doc.BeginCommandGroup "全部转曲"
    For Each pg In doc.Pages
        ConvertRange pg.Shapes.All, deep
    Next pg
    doc.EndCommandGroup
    Optimization = False
    DoRefresh
End Sub

' 递归处理一个形状范围
Private Sub ConvertRange(ByVal sr As ShapeRange, ByVal deep As Boolean)
    Dim sh As Shape
    Dim pc As ShapeRange

    If sr Is Nothing Then Exit Sub

    For Each sh In sr
        Select Case sh.Type
            Case cdrGroupShape
                ConvertRange sh.Shapes.All, deep

            Case cdrTextShape, cdrRectangleShape, cdrEllipseShape, _
                 cdrPolygonShape
                On Error Resume Next
                sh.ConvertToCurves
                On Error GoTo 0
        End Select

        ' 深入 PowerClip 内部
        If deep Then
            Set pc = Nothing
            On Error Resume Next
            Set pc = sh.PowerClip.Shapes.All
            On Error GoTo 0
            If Not pc Is Nothing Then ConvertRange pc, True
        End If
    Next sh
End Sub