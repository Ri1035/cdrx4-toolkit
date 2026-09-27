Attribute VB_Name = "M_Calendar"
Option Explicit

'==========================================================
' 日历创建
'   输入 2026-09，生成该月的文字日历
'   内容：年 + 月标题 + 星期表头 + 日期数字
'   尺寸按毫米算，再换算成当前文档单位
'==========================================================

' 带对话框的入口
Public Sub CreateCalendar()
    Dim s As String
    Dim parts As Variant
    Dim y As Long
    Dim m As Long
    Dim n As Long

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "日历创建"
        Exit Sub
    End If

    s = InputBox("输入要生成的年月，例如 2026-09", "日历创建", Format$(Date, "yyyy-mm"))
    If Len(s) = 0 Then Exit Sub

    parts = Split(Replace(s, "/", "-"), "-")
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

    n = CreateCalendarCore(y, m)

    MsgBox "日历已生成，共 " & n & " 个文字对象。", vbInformation, "日历创建"
End Sub

' 纯逻辑，供自检调用
Public Function CreateCalendarCore(ByVal y As Long, ByVal m As Long) As Long
    Dim doc As Document
    Dim pg As Page
    Dim lay As Object
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
    Dim n As Long

    If y < 1900 Or y > 2200 Or m < 1 Or m > 12 Then Exit Function

    Set doc = CorelDRAW.ActiveDocument
    If doc Is Nothing Then Exit Function

    Set pg = ActivePageSafe()
    If pg Is Nothing Then Exit Function
    Set lay = pg.ActiveLayer

    PageWHmm pg, w, h
    If w <= 0 Then w = 210
    If h <= 0 Then h = 297

    firstDay = DateSerial(y, m, 1)
    days = Day(DateSerial(y, m + 1, 0))
    startCol = Weekday(firstDay, vbMonday) - 1

    cw = 20
    ch = 14
    x0 = 20
    yTop = h - 25

    wd = Array("一", "二", "三", "四", "五", "六", "日")

    CorelDRAW.Optimization = True
    doc.BeginCommandGroup "日历创建"

    ' 标题
    If AddCalText(lay, y & " 年 " & m & " 月", x0 + cw * 3, yTop, 16) Then n = n + 1

    ' 星期表头
    For i = 0 To 6
        If AddCalText(lay, CStr(wd(i)), x0 + cw * i + cw / 2, yTop - 12, 11) Then n = n + 1
    Next i

    ' 日期
    For i = 1 To days
        col = (startCol + i - 1) Mod 7
        row = (startCol + i - 1) \ 7
        If AddCalText(lay, CStr(i), x0 + cw * col + cw / 2, _
                      yTop - 12 - ch * (row + 1), 11) Then n = n + 1
    Next i

    doc.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh

    CreateCalendarCore = n
End Function

Private Function AddCalText(ByVal lay As Object, ByVal txt As String, _
                            ByVal xmm As Double, ByVal ymm As Double, _
                            ByVal sizePt As Double) As Boolean
    Dim sh As Shape

    Set sh = MakeCenteredText(lay, txt, xmm, ymm, sizePt, "宋体")
    AddCalText = Not (sh Is Nothing)
End Function