Attribute VB_Name = "M_PageNo"
Option Explicit

'==========================================================
' 插入页码
'   为每一页在底部居中插入「第 X 页 / 共 Y 页」
'==========================================================

Public Sub InsertPageNumber()
    Dim doc As Document
    Dim pg As Page
    Dim w As Double
    Dim h As Double
    Dim i As Long
    Dim n As Long
    Dim s As Shape
    Dim txt As String

    If Not HasDocument() Then Exit Sub
    Set doc = CorelDRAW.ActiveDocument
    n = doc.Pages.Count

    Optimization = True
    doc.BeginCommandGroup "插入页码"

    For i = 1 To n
        Set pg = doc.Pages(i)
        PageWH pg, w, h
        If w <= 0 Then w = 210
        If h <= 0 Then h = 297

        txt = "第 " & i & " 页 / 共 " & n & " 页"

        Set s = Nothing
        On Error Resume Next
        Set s = pg.ActiveLayer.CreateArtisticTextWide(w / 2, h - 8, txt)
        If Not s Is Nothing Then
            s.Text.Story.Font = "微软雅黑"
            s.Text.Story.Size = 9
            s.Text.Story.Alignment = cdrCenterAlignment
            s.SetPosition w / 2, h - 8
        End If
        On Error GoTo 0
    Next i

    doc.EndCommandGroup
    Optimization = False
    DoRefresh
End Sub