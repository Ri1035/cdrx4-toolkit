Attribute VB_Name = "M_PageNo"
Option Explicit

'==========================================================
' 插入页码
'   为每一页在底部居中插入「第 X 页 / 共 Y 页」
'   位置按毫米算，再换算成当前文档单位
'==========================================================

' 带对话框的入口
Public Sub InsertPageNumber()
    Dim n As Long

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "插入页码"
        Exit Sub
    End If

    n = InsertPageNumberCore()

    MsgBox "已为 " & n & " 页插入页码。", vbInformation, "插入页码"
End Sub

' 纯逻辑，供自检调用
Public Function InsertPageNumberCore() As Long
    Dim doc As Document
    Dim pg As Page
    Dim w As Double
    Dim h As Double
    Dim idx As Long
    Dim total As Long
    Dim s As Shape
    Dim txt As String

    Set doc = CorelDRAW.ActiveDocument
    If doc Is Nothing Then Exit Function

    total = doc.Pages.Count

    CorelDRAW.Optimization = True
    doc.BeginCommandGroup "插入页码"

    idx = 0
    For Each pg In doc.Pages
        idx = idx + 1

        PageWHmm pg, w, h
        If w <= 0 Then w = 210
        If h <= 0 Then h = 297

        txt = "第 " & idx & " 页 / 共 " & total & " 页"

        Set s = MakeCenteredText(pg.ActiveLayer, txt, w / 2, h - 10, 9, "宋体")
        If Not s Is Nothing Then InsertPageNumberCore = InsertPageNumberCore + 1
    Next pg

    doc.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh
End Function