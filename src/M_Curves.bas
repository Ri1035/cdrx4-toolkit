Attribute VB_Name = "M_Curves"
Option Explicit

'==========================================================
' 全部转曲
'   把所有页面的文本与图形转成曲线。
'   弹窗两步：① 范围（1 全部 / 2 仅文字 / 3 仅图形）
'             ② 是否深入 PowerClip 容器内部
'   两步都能「取消」= 什么都不做，且没有任何按钮被借去表示第二种含义。
'
'   X4 实测：Shape.ConvertToCurves 是「语句」，不能 Set 返回值。
'
'   【不要引用没验证过的 cdr* 常量】一个常量在 X4 里不存在 → 整个工程编译不过
'   → 所有宏全废（PLAN §3.5 的教训）。下面用到的常量都由 M_Test.Probe 实际印出过。
'==========================================================

' 转曲范围
Private Const SCOPE_ALL As Long = 0
Private Const SCOPE_TEXT As Long = 1
Private Const SCOPE_SHAPE As Long = 2

' 带对话框的入口
Public Sub ConvertAllToCurves()
    Dim scope As Long
    Dim deep As Long
    Dim n As Long

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "全部转曲"
        Exit Sub
    End If

    ' ① 范围：编号选择。原来的做法把「取消」借来表示「仅图形」，
    '    用户得先读说明才敢点；现在三个范围各有一个编号，「取消」就是取消。
    scope = AskScope()
    If scope < 0 Then Exit Sub

    ' ② 深度
    deep = AskDeep()
    If deep < 0 Then Exit Sub

    n = ConvertAllToCurvesCore(deep <> 0, scope)

    If n = 0 Then
        MsgBox "没有转换任何对象（" & ScopeName(scope) & "）。" & vbCrLf & vbCrLf & _
               "可能的原因：" & vbCrLf & _
               "  · 这个范围下本来就没有可转的对象" & vbCrLf & _
               "  · 对象已经是曲线了" & vbCrLf & _
               "  · 容器内部的还没转（第 2 步选了「只处理顶层」）", _
               vbInformation, "全部转曲"
    Else
        MsgBox "转曲完成，共转换 " & n & " 个对象（" & ScopeName(scope) & "）。", _
               vbInformation, "全部转曲"
    End If
End Sub

' 询问转曲范围。返回 -1 表示放弃。
Private Function AskScope() As Long
    Dim s As String

    s = InputBox( _
        "第 1 / 2 步：选择转曲范围" & vbCrLf & vbCrLf & _
        "1 = 文字 + 图形（全部）" & vbCrLf & _
        "2 = 只转文字" & vbCrLf & _
        "3 = 只转图形" & vbCrLf & vbCrLf & _
        "填序号后点「确定」；点「取消」则什么都不做。", _
        "全部转曲", "1")

    If Len(Trim(s)) = 0 Then
        AskScope = -1
        Exit Function
    End If

    Select Case Left(Trim(s), 1)
        Case "2": AskScope = SCOPE_TEXT
        Case "3": AskScope = SCOPE_SHAPE
        Case Else: AskScope = SCOPE_ALL
    End Select
End Function

' 询问是否深入容器。返回 -1 表示放弃。
Private Function AskDeep() As Long
    Dim ans As VbMsgBoxResult

    ans = MsgBox( _
        "第 2 / 2 步：是否深入 PowerClip 容器内部？" & vbCrLf & vbCrLf & _
        "是 = 容器内部的文字 / 图形也一起转" & vbCrLf & _
        "否 = 只处理顶层对象" & vbCrLf & _
        "取消 = 放弃，什么都不做", _
        vbYesNoCancel + vbQuestion, "全部转曲")

    Select Case ans
        Case vbYes: AskDeep = 1
        Case vbNo:  AskDeep = 0
        Case Else:  AskDeep = -1
    End Select
End Function

' 纯逻辑，供自检调用。
' scope 省略时按「全部」处理，所以旧的 ConvertAllToCurvesCore(True) 调用仍然有效。
Public Function ConvertAllToCurvesCore(ByVal deep As Boolean, _
                                       Optional ByVal scope As Long = SCOPE_ALL) As Long
    Dim doc As Document
    Dim pg As Page
    Dim n As Long

    Set doc = CorelDRAW.ActiveDocument
    If doc Is Nothing Then Exit Function

    CorelDRAW.Optimization = True
    doc.BeginCommandGroup "转曲（" & ScopeName(scope) & "）"

    For Each pg In doc.Pages
        n = n + ConvertRange(pg.Shapes.All, deep, scope)
    Next pg

    doc.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh

    ConvertAllToCurvesCore = n
End Function

' 递归处理一个形状范围，返回转换成功的个数
Private Function ConvertRange(ByVal sr As Object, ByVal deep As Boolean, _
                              ByVal scope As Long) As Long
    Dim i As Long
    Dim sh As Shape
    Dim kids As Object
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
            If sh.Type = cdrGroupShape Then
                ' 组内对象逐个处理
                Set kids = sh.Shapes.All
                If Err.Number = 0 And Not kids Is Nothing Then
                    n = n + ConvertRange(kids, deep, scope)
                End If
                Err.Clear
            ElseIf WantsCurves(sh.Type, scope) Then
                ' ConvertToCurves 是语句，不能 Set
                sh.ConvertToCurves
                If Err.Number = 0 Then n = n + 1
                Err.Clear
            End If

            ' 深入 PowerClip 内部
            If deep Then
                Set kids = PowerClipRange(sh)
                If Not kids Is Nothing Then n = n + ConvertRange(kids, True, scope)
            End If
        End If
        Err.Clear
    Next i
    On Error GoTo 0

    ConvertRange = n
End Function

' 这个类型在当前范围下要不要转曲
Private Function WantsCurves(ByVal t As Long, ByVal scope As Long) As Boolean
    If scope = SCOPE_TEXT Then
        WantsCurves = (t = cdrTextShape)
        Exit Function
    End If

    ' 「仅图形」与「全部」的差别只在文本；其余非文本矢量对象一律转。
    Select Case t
        Case cdrTextShape
            WantsCurves = (scope = SCOPE_ALL)
        Case cdrCurveShape
            WantsCurves = False          ' 本来就是曲线，转了也是白转
        Case Else
            ' 矩形 / 椭圆 / 多边形 / 符号实例 / 自定义形状 / 连接线 …
            ' 位图等转不了的会在这里报错，靠上面的 Err 判定跳过，不计入个数
            WantsCurves = True
    End Select
End Function

Private Function ScopeName(ByVal scope As Long) As String
    Select Case scope
        Case SCOPE_TEXT:  ScopeName = "仅文字"
        Case SCOPE_SHAPE: ScopeName = "仅图形"
        Case Else:        ScopeName = "文字 + 图形"
    End Select
End Function

' 取 PowerClip 内部的对象集合，没有就返回 Nothing
Private Function PowerClipRange(ByVal sh As Shape) As Object
    Dim o As Object
    Dim pc As Object
    Dim sr As Object

    On Error Resume Next
    Err.Clear
    Set o = sh
    Set pc = o.PowerClip
    If Err.Number = 0 And Not pc Is Nothing Then
        Set sr = pc.Shapes.All
        If Err.Number <> 0 Then
            Err.Clear
            Set sr = pc.Contents.Shapes.All
        End If
    End If
    Err.Clear
    On Error GoTo 0

    Set PowerClipRange = sr
End Function