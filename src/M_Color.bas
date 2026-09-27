Attribute VB_Name = "M_Color"
Option Explicit

'==========================================================
' 颜色替换
'   先选中 2 个对象：第 1 个取原色，第 2 个取新色
'   然后把全文档所有填充 / 轮廓里的原色换成新色
'==========================================================

Public Sub ReplaceColor()
    Dim sr As ShapeRange
    Dim src As Color
    Dim dst As Color
    Dim doc As Document
    Dim pg As Page
    Dim sh As Shape

    If Not HasDocument() Then Exit Sub

    Set sr = CorelDRAW.ActiveSelection
    If sr.Count <> 2 Then
        MsgBox "请先选中 2 个对象。" & vbCrLf & vbCrLf & _
               "第 1 个 = 原色（要被替换掉的），第 2 个 = 新色（换成什么）", vbExclamation, "颜色替换"
        Exit Sub
    End If

    Set src = Nothing
    Set dst = Nothing
    On Error Resume Next
    Set src = sr(1).Fill.UniformColor
    Set dst = sr(2).Fill.UniformColor
    On Error GoTo 0

    If src Is Nothing Or dst Is Nothing Then
        MsgBox "取色失败：这两个对象都要有纯色填充。", vbExclamation, "颜色替换"
        Exit Sub
    End If

    Set doc = CorelDRAW.ActiveDocument

    Optimization = True
    doc.BeginCommandGroup "颜色替换"
    For Each pg In doc.Pages
        For Each sh In pg.Shapes.All
            ReplaceInShape sh, src, dst
        Next sh
    Next pg
    doc.EndCommandGroup
    Optimization = False
    DoRefresh

    MsgBox "颜色替换完成。", vbInformation, "颜色替换"
End Sub

Private Sub ReplaceInShape(ByVal sh As Shape, ByVal src As Color, ByVal dst As Color)
    Dim c As Color

    On Error Resume Next

    ' 填充
    If sh.Fill.Type = cdrUniformFill Then
        Set c = sh.Fill.UniformColor
        If Not c Is Nothing Then
            If c.IsSame(src) Then sh.Fill.ApplyUniformFill dst
        End If
    End If

    ' 轮廓
    Set c = sh.Outline.Color
    If Not c Is Nothing Then
        If c.IsSame(src) Then sh.Outline.Color.CopyAssign dst
    End If

    On Error GoTo 0
End Sub