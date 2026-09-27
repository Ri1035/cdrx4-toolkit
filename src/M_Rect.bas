Attribute VB_Name = "M_Rect"
Option Explicit

'==========================================================
' 标准矩形
'   把选中矩形的圆角 / 扇形角 / 倒棱角重置为直角
'==========================================================

Public Sub StdRectangle()
    Dim sh As Shape
    Dim n As Long
    Dim total As Long

    If Not HasDocument() Then Exit Sub

    total = CorelDRAW.ActiveSelection.Count
    If total = 0 Then
        MsgBox "请先选中至少一个矩形。", vbExclamation, "标准矩形"
        Exit Sub
    End If

    Optimization = True
    CorelDRAW.ActiveDocument.BeginCommandGroup "标准矩形"

    For Each sh In CorelDRAW.ActiveSelection.Shapes
        If sh.Type = cdrRectangleShape Then
            On Error Resume Next
            With sh.Rectangle
                .EqualCorners = True
                .Radius = 0
                .RadiusUpperLeft = 0
                .RadiusUpperRight = 0
                .RadiusLowerLeft = 0
                .RadiusLowerRight = 0
            End With
            If Err.Number = 0 Then n = n + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next sh

    CorelDRAW.ActiveDocument.EndCommandGroup
    Optimization = False
    DoRefresh

    If n = 0 Then
        MsgBox "选中的对象里没有矩形，什么都没改。", vbExclamation, "标准矩形"
    Else
        MsgBox "已处理 " & n & " 个矩形，四角都重置为直角了。", vbInformation, "标准矩形"
    End If
End Sub