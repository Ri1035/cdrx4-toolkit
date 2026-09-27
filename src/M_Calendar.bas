Attribute VB_Name = "M_Calendar"
Option Explicit

'==========================================================
' 日历创建
'   输入 2026-09，生成该月的文字日历
'   内容：年 + 月标题 + 星期表头 + 日期数字
'==========================================================

Public Sub CreateCalendar()
    Dim doc As Document
    Dim pg As Page
    Dim lay As Layer
    Dim s As String
    Dim parts As Variant
    Dim y As Long
    Dim m As Long
    Dim firstDay As Date
    Dim days As Long
    Dim startCol As Long
    Dim w As Double
    Dim h As Double
    Dim cw As Double
    Dim ch As Double
    Dim x0 As Double
    Dim yTop As Double
    Dim i As Long
    Dim col As Long
    Dim row As Long
    Dim wd As Variant

    If Not HasDocument() Then Exit Sub

    s = InputBox("输入要生成的年月，例如 2026-09", "日历创建", Format$(Date, "yyyy-mm"))
    If Len(s) = 0 Then Exit Sub

    parts = Split(s, "-")
    If UBound(parts) < 1 Then
        MsgBox "格式不对，应该像 2026-09 这样。", vbExclamation, "日历创建"
        Exit Sub
    End If

    y = Val(parts(0))
    m = Val(parts(1))
    If y < 1900 Or y > 2200 Or m < 1 Or m > 12 Then
        MsgBox "年份或月份超出范围了。", vbExclamation, "日历创建"
        Exit Sub
    End If

    firstDay = DateSerial(y, m, 1)
    days = Day(DateSerial(y, m + 1, 0))
    startCol = Weekday(firstDay, vbMonday) - 1

    Set doc = CorelDRAW.ActiveDocument
    Set pg = doc.ActivePage
    Set lay = pg.ActiveLayer

    PageWH pg, w, h
    If w <= 0 Then w = 210
    If h <= 0 Then h = 297

    cw = 20
    ch = 14
    x0 = 20
    yTop = h - 20

    wd = Array("一", "二", "三", "四", "五", "六", "日")

    Optimization = True
    doc.BeginCommandGroup "日历创建"

    ' 标题
    AddText lay, x0 + cw * 3, yTop, y & " 年 " & m & " 月", 16

    ' 星期表头
    For i = 0 To 6
        AddText lay, x0 + cw * i + cw / 2, yTop - 12, wd(i), 11
    Next i

    ' 日期
    For i = 1 To days
        col = (startCol + i - 1) Mod 7
        row = (startCol + i - 1) \ 7
        AddText lay, x0 + cw * col + cw / 2, yTop - 12 - ch * (row + 1), CStr(i), 11
    Next i

    doc.EndCommandGroup
    Optimization = False
    DoRefresh
End Sub

Private Sub AddText(ByVal lay As Layer, ByVal x As Double, ByVal y As Double, _
                    ByVal txt As String, ByVal sz As Long)
    Dim t As Shape

    On Error Resume Next
    Set t = lay.CreateArtisticTextWide(x, y, txt)
    If Not t Is Nothing Then
        t.Text.Story.Font = "微软雅黑"
        t.Text.Story.Size = sz
        t.Text.Story.Alignment = cdrCenterAlignment
        t.SetPosition x, y
    End If
    On Error GoTo 0
End Sub