Attribute VB_Name = "M_Rect"
Option Explicit

'==========================================================
' 标准矩形
'   把选中矩形的圆角 / 扇形角 / 倒棱角重置为直角
'
'   X4 实测：Rectangle.RadiusUpperLeft 等属性在早期绑定下
'   会被判成「只读」而编译报错，必须用 Object 晚期绑定。
'==========================================================

' 带对话框的入口
Public Sub StdRectangle()
    Dim n As Long

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "标准矩形"
        Exit Sub
    End If

    n = StdRectangleCore()

    If n = 0 Then
        MsgBox "选中的对象里没有矩形，什么都没改。" & vbCrLf & vbCrLf & _
               "先选中要处理的矩形再点一次。", vbExclamation, "标准矩形"
    Else
        MsgBox "已处理 " & n & " 个矩形，四角都重置为直角了。", vbInformation, "标准矩形"
    End If
End Sub

' 纯逻辑，供自检调用
Public Function StdRectangleCore() As Long
    Dim sr As Object
    Dim sh As Shape
    Dim n As Long
    Dim i As Long

    Set sr = Nothing
    On Error Resume Next
    Err.Clear
    Set sr = CorelDRAW.ActiveSelectionRange
    If Err.Number <> 0 Then Set sr = Nothing
    Err.Clear
    On Error GoTo 0

    If sr Is Nothing Then Exit Function

    CorelDRAW.Optimization = True
    CorelDRAW.ActiveDocument.BeginCommandGroup "标准矩形"

    For i = 1 To sr.Count
        Set sh = Nothing
        On Error Resume Next
        Err.Clear
        Set sh = sr.Item(i)
        If Err.Number <> 0 Or sh Is Nothing Then
            ' Item 在部分 ShapeRange 上取不到，退回 Shapes(i)
            Err.Clear
            Set sh = sr.Shapes(i)
        End If
        If Err.Number = 0 And Not sh Is Nothing Then
            If sh.Type = cdrRectangleShape Then
                If SharpRect(sh) Then n = n + 1
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next i

    CorelDRAW.ActiveDocument.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh

    StdRectangleCore = n
End Function

' 四个角半径清零。
' 每个属性单独判错：X4 里个别 Radius* 可能不存在，
' 不能让最后一个赋值把前面成功的结果一起判死。
Private Function SharpRect(ByVal sh As Shape) As Boolean
    Dim ro As Object
    Dim ok As Boolean

    If sh Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    Set ro = sh.Rectangle
    If Err.Number = 0 And Not ro Is Nothing Then
        ok = True

        ro.RadiusUpperLeft = 0
        If Err.Number <> 0 Then ok = False
        Err.Clear

        ro.RadiusUpperRight = 0
        If Err.Number <> 0 Then ok = False
        Err.Clear

        ro.RadiusLowerLeft = 0
        If Err.Number <> 0 Then ok = False
        Err.Clear

        ro.RadiusLowerRight = 0
        If Err.Number <> 0 Then ok = False
        Err.Clear
    End If
    On Error GoTo 0

    SharpRect = ok
End Function