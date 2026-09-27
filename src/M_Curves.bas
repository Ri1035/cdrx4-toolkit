Attribute VB_Name = "M_Curves"
Option Explicit

'==========================================================
' 全部转曲
'   把所有页面的文本与图形都转成曲线
'   可选是否深入 PowerClip 容器内部
'
'   X4 实测：Shape.ConvertToCurves 是「语句」，不能 Set 返回值。
'==========================================================

' 带对话框的入口
Public Sub ConvertAllToCurves()
    Dim ans As VbMsgBoxResult
    Dim deep As Boolean
    Dim n As Long

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

    n = ConvertAllToCurvesCore(deep)

    MsgBox "全部转曲完成，共转换 " & n & " 个对象。", vbInformation, "全部转曲"
End Sub

' 纯逻辑，供自检调用
Public Function ConvertAllToCurvesCore(ByVal deep As Boolean) As Long
    Dim doc As Document
    Dim pg As Page
    Dim n As Long

    Set doc = CorelDRAW.ActiveDocument
    If doc Is Nothing Then Exit Function

    CorelDRAW.Optimization = True
    doc.BeginCommandGroup "全部转曲"

    For Each pg In doc.Pages
        n = n + ConvertRange(pg.Shapes.All, deep)
    Next pg

    doc.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh

    ConvertAllToCurvesCore = n
End Function

' 递归处理一个形状范围，返回转换成功的个数
Private Function ConvertRange(ByVal sr As Object, ByVal deep As Boolean) As Long
    Dim i As Long
    Dim sh As Shape
    Dim kids As Object
    Dim n As Long

    If sr Is Nothing Then Exit Function

    On Error Resume Next
    For i = 1 To sr.Count
        Set sh = Nothing
        Set sh = sr.Item(i)
        If Err.Number <> 0 Or sh Is Nothing Then
            ' Item 在部分 ShapeRange 上取不到，退回 Shapes(i)
            Err.Clear
            Set sh = sr.Shapes(i)
        End If
        If Err.Number = 0 And Not sh Is Nothing Then
            Select Case sh.Type
                Case cdrGroupShape
                    ' 组内对象逐个处理
                    Set kids = sh.Shapes.All
                    If Err.Number = 0 And Not kids Is Nothing Then
                        n = n + ConvertRange(kids, deep)
                    End If
                    Err.Clear

                Case cdrTextShape, cdrRectangleShape, cdrEllipseShape, _
                     cdrPolygonShape
                    ' ConvertToCurves 是语句，不能 Set
                    sh.ConvertToCurves
                    If Err.Number = 0 Then n = n + 1
                    Err.Clear
            End Select

            ' 深入 PowerClip 内部
            If deep Then
                Set kids = PowerClipRange(sh)
                If Not kids Is Nothing Then n = n + ConvertRange(kids, True)
            End If
        End If
        Err.Clear
    Next i
    On Error GoTo 0

    ConvertRange = n
End Function

' 取 PowerClip 内部的对象集合，没有就返回 Nothing
Private Function PowerClipRange(ByVal sh As Shape) As Object
    Dim o As Object
    Dim pc As Object
    Dim sr As Object

    On Error Resume Next
    Err.Clear
    Set o = sh
    Set pc = o.PowerClip
    If Err.Number = 0 And Not pc Is Nothing Then
        Set sr = pc.Shapes.All
        If Err.Number <> 0 Then
            Err.Clear
            Set sr = pc.Contents.Shapes.All
        End If
    End If
    Err.Clear
    On Error GoTo 0

    Set PowerClipRange = sr
End Function