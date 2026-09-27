Attribute VB_Name = "M_Color"
Option Explicit

'==========================================================
' 颜色替换
'   先选中 2 个对象：第 1 个取原色，第 2 个取新色
'   然后把全文档所有填充 / 轮廓里的原色换成新色
'
'   X4 实测：CorelDRAW.ActiveSelection 返回的是 Shape，
'   没有 .Count / .Item；要用 ActiveSelectionRange。
'==========================================================

' 带对话框的入口
Public Sub ReplaceColor()
    Dim sr As Object
    Dim src As Color
    Dim dst As Color
    Dim n As Long

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "颜色替换"
        Exit Sub
    End If

    Set sr = Nothing
    On Error Resume Next
    Err.Clear
    Set sr = CorelDRAW.ActiveSelectionRange
    If Err.Number <> 0 Then Set sr = Nothing
    Err.Clear
    On Error GoTo 0

    If sr Is Nothing Then
        MsgBox "请先选中 2 个对象。" & vbCrLf & vbCrLf & _
               "第 1 个 = 原色（要被替换掉的）" & vbCrLf & _
               "第 2 个 = 新色（换成什么）", vbExclamation, "颜色替换"
        Exit Sub
    End If

    If sr.Count <> 2 Then
        MsgBox "请先选中 2 个对象，现在选中的是 " & sr.Count & " 个。" & vbCrLf & vbCrLf & _
               "第 1 个 = 原色（要被替换掉的）" & vbCrLf & _
               "第 2 个 = 新色（换成什么）", vbExclamation, "颜色替换"
        Exit Sub
    End If

    Set src = Nothing
    Set dst = Nothing
    On Error Resume Next
    Err.Clear
    Set src = sr.Item(1).Fill.UniformColor
    Set dst = sr.Item(2).Fill.UniformColor
    Err.Clear
    On Error GoTo 0

    If src Is Nothing Or dst Is Nothing Then
        MsgBox "取色失败：这两个对象都要有纯色填充。", vbExclamation, "颜色替换"
        Exit Sub
    End If

    n = ReplaceColorCore(src, dst)

    MsgBox "颜色替换完成，共替换 " & n & " 处颜色。", vbInformation, "颜色替换"
End Sub

' 纯逻辑，供自检调用
Public Function ReplaceColorCore(ByVal src As Color, ByVal dst As Color) As Long
    Dim doc As Document
    Dim pg As Page
    Dim sh As Shape
    Dim n As Long

    Set doc = CorelDRAW.ActiveDocument
    If doc Is Nothing Then Exit Function

    CorelDRAW.Optimization = True
    doc.BeginCommandGroup "颜色替换"

    For Each pg In doc.Pages
        For Each sh In pg.Shapes.All
            n = n + ReplaceInShape(sh, src, dst)
        Next sh
    Next pg

    doc.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh

    ReplaceColorCore = n
End Function

' 处理一个形状（组会递归进去），返回替换的处数
Private Function ReplaceInShape(ByVal sh As Shape, ByVal src As Color, ByVal dst As Color) As Long
    Dim c As Color
    Dim kids As Object
    Dim n As Long

    If sh Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear

    ' 组：递归进去
    If sh.Type = cdrGroupShape Then
        Set kids = sh.Shapes.All
        If Err.Number = 0 And Not kids Is Nothing Then
            n = n + ReplaceInRange(kids, src, dst)
        End If
        Err.Clear
        On Error GoTo 0
        ReplaceInShape = n
        Exit Function
    End If

    ' 填充
    Set c = Nothing
    If sh.Fill.Type = cdrUniformFill Then
        Set c = sh.Fill.UniformColor
        If Not c Is Nothing Then
            If SameColor(c, src) Then
                sh.Fill.ApplyUniformFill dst
                If Err.Number = 0 Then n = n + 1
                Err.Clear
            End If
        End If
    End If
    Err.Clear

    ' 轮廓
    Set c = Nothing
    Set c = sh.Outline.Color
    If Not c Is Nothing Then
        If SameColor(c, src) Then
            sh.Outline.Color.CopyAssign dst
            If Err.Number = 0 Then n = n + 1
            Err.Clear
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ReplaceInShape = n
End Function

Private Function ReplaceInRange(ByVal sr As Object, ByVal src As Color, ByVal dst As Color) As Long
    Dim i As Long
    Dim sh As Shape
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
            n = n + ReplaceInShape(sh, src, dst)
        End If
        Err.Clear
    Next i
    On Error GoTo 0

    ReplaceInRange = n
End Function