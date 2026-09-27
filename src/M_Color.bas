Attribute VB_Name = "M_Color"
Option Explicit

'==========================================================
' ?????I
'   ????????????? 1 ?????????????? 2 ???????????
'   ?????????????????????? / ??????I????
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
        MsgBox "?????????????????" & vbCrLf & vbCrLf & _
               "?? 1 ?? = ?????? 2 ?? = ???", vbExclamation, "?????I"
        Exit Sub
    End If

    Set src = Nothing
    Set dst = Nothing
    On Error Resume Next
    Set src = sr(1).Fill.UniformColor
    Set dst = sr(2).Fill.UniformColor
    On Error GoTo 0

    If src Is Nothing Or dst Is Nothing Then
        MsgBox "???????????????????", vbExclamation, "?????I"
        Exit Sub
    End If

    Set doc = CorelDRAW.ActiveDocument

    Optimization = True
    doc.BeginCommandGroup "?????I"
    For Each pg In doc.Pages
        For Each sh In pg.Shapes.All
            ReplaceInShape sh, src, dst
        Next sh
    Next pg
    doc.EndCommandGroup
    Optimization = False
    DoRefresh

    MsgBox "?????I?????", vbInformation, "?????I"
End Sub

Private Sub ReplaceInShape(ByVal sh As Shape, ByVal src As Color, ByVal dst As Color)
    Dim c As Color

    On Error Resume Next

    ' ???
    If sh.Fill.Type = cdrUniformFill Then
        Set c = sh.Fill.UniformColor
        If Not c Is Nothing Then
            If c.IsSame(src) Then sh.Fill.ApplyUniformFill dst
        End If
    End If

    ' ????
    Set c = sh.Outline.Color
    If Not c Is Nothing Then
        If c.IsSame(src) Then sh.Outline.Color.CopyAssign dst
    End If

    On Error GoTo 0
End Sub