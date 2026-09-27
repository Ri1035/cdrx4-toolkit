Attribute VB_Name = "M_Rect"
Option Explicit

'==========================================================
' ???????
'   ?????????????????/??????/?????????????????
'==========================================================

Public Sub StdRectangle()
    Dim sh As Shape
    Dim n As Long
    Dim total As Long

    If Not HasDocument() Then Exit Sub

    total = CorelDRAW.ActiveSelection.Count
    If total = 0 Then
        MsgBox "?????????????", vbExclamation, "???????"
        Exit Sub
    End If

    Optimization = True
    CorelDRAW.ActiveDocument.BeginCommandGroup "???????"

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
        MsgBox "???????????????????", vbExclamation, "???????"
    Else
        MsgBox "??? " & n & " ??????????????????????", vbInformation, "???????"
    End If
End Sub