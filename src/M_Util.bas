Attribute VB_Name = "M_Util"
Option Explicit

'==========================================================
' CDR X4 ???????? ???? ????????
' ?????????? CDRX4Toolkit
'==========================================================

' ?????????
Public Function HasDocument() As Boolean
    On Error Resume Next
    HasDocument = (CorelDRAW.Documents.Count > 0)
    On Error GoTo 0
End Function

' ??????????????????????? mm??
Public Sub PageWH(ByVal pg As Page, ByRef w As Double, ByRef h As Double)
    w = 0
    h = 0
    On Error Resume Next
    w = pg.SizeWidth
    h = pg.SizeHeight
    On Error GoTo 0
End Sub

' ????????
Public Function DesktopPath() As String
    On Error Resume Next
    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0
End Function

' ??????????????????????
Public Function BaseName(ByVal p As String) As String
    Dim n As String
    Dim i As Long

    If Len(p) = 0 Then
        BaseName = "??????"
        Exit Function
    End If

    n = p
    i = InStrRev(n, "\")
    If i > 0 Then n = Mid$(n, i + 1)
    i = InStrRev(n, ".")
    If i > 0 Then n = Left$(n, i - 1)
    BaseName = n
End Function

' ??????
Public Sub DoRefresh()
    On Error Resume Next
    CorelDRAW.Refresh
    On Error GoTo 0
End Sub