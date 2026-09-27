Attribute VB_Name = "M_Curves"
Option Explicit

'==========================================================
' ??????
'   ??????????????????????????????
'   ???????????? PowerClip ???????
'==========================================================

Public Sub ConvertAllToCurves()
    Dim doc As Document
    Dim pg As Page
    Dim ans As VbMsgBoxResult
    Dim deep As Boolean

    If Not HasDocument() Then
        MsgBox "?????????????", vbExclamation, "??????"
        Exit Sub
    End If

    ans = MsgBox("?????????" & vbCrLf & vbCrLf & _
                 "???????? + ??? + PowerClip ??????????????" & vbCrLf & _
                 "????????? + ???????????????" & vbCrLf & _
                 "??????????", _
                 vbYesNoCancel + vbQuestion, "??????")
    If ans = vbCancel Then Exit Sub
    deep = (ans = vbYes)

    Set doc = CorelDRAW.ActiveDocument

    Optimization = True
    doc.BeginCommandGroup "??????"
    For Each pg In doc.Pages
        ConvertRange pg.Shapes.All, deep
    Next pg
    doc.EndCommandGroup
    Optimization = False
    DoRefresh
End Sub

' ????????????
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

        ' ???? PowerClip ????
        If deep Then
            Set pc = Nothing
            On Error Resume Next
            Set pc = sh.PowerClip.Shapes.All
            On Error GoTo 0
            If Not pc Is Nothing Then ConvertRange pc, True
        End If
    Next sh
End Sub