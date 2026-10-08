Attribute VB_Name = "M_Color"
Option Explicit

'==========================================================
' 颜色替换
'   先选中 2 个对象：第 1 个取原色，第 2 个取新色
'   然后把全文档所有填充 / 轮廓里的原色换成新色
'
'   X4 实测：CorelDRAW.ActiveSelection 返回的是 Shape，
'   没有 .Count / .Item；要用 ActiveSelectionRange。
'
'   【对话框约定，与 M_Curves / M_CMYK 保持一致】
'     先用 InputBox 编号选方向，再干活；「取消」= 什么都不做。
'     不给「取消」按钮借第二种含义。
'
'   【不要引用没验证过的 cdr* 常量 / Color 成员】一个不存在的成员
'   → 整个工程编译不过 → 所有宏全废（PLAN §3.5）。下面只用到已实测的：
'   cdrUniformFill、Color.RGBRed/RGBGreen/RGBBlue（M_Util 在用）。
'==========================================================

' 带对话框的入口
Public Sub ReplaceColor()
    Dim sr As Object
    Dim src As Color
    Dim dst As Color
    Dim tmp As Color
    Dim n As Long
    Dim cnt As Long
    Dim pick As Long
    Dim s As String

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "颜色替换"
        Exit Sub
    End If

    ' ---- 取当前选择。VBA 的 Or 不短路，所以先落成 cnt 再判 ----
    Set sr = Nothing
    On Error Resume Next
    Err.Clear
    Set sr = CorelDRAW.ActiveSelectionRange
    If Err.Number <> 0 Then Set sr = Nothing
    Err.Clear
    On Error GoTo 0

    cnt = 0
    If Not sr Is Nothing Then cnt = sr.Count

    If cnt <> 2 Then
        MsgBox "颜色替换需要先选中 2 个对象，现在选中的是 " & cnt & " 个。" & vbCrLf & vbCrLf & _
               "第 1 个 = 原色（要被替换掉的）" & vbCrLf & _
               "第 2 个 = 新色（换成什么）" & vbCrLf & vbCrLf & _
               "对象只要有纯色填充，或者有轮廓色，就能当作色源。", _
               vbExclamation, "颜色替换"
        Exit Sub
    End If

    ' ---- 取两个颜色：优先纯色填充，没有纯色填充就退回轮廓色 ----
    Set src = PickColor(sr.Item(1))
    Set dst = PickColor(sr.Item(2))

    If src Is Nothing Or dst Is Nothing Then
        MsgBox "取色失败：两个对象都要有纯色填充，或者有轮廓色。" & vbCrLf & vbCrLf & _
               "既没有纯色填充、又没有轮廓色的对象（透明矩形、位图等）不能当色源。", _
               vbExclamation, "颜色替换"
        Exit Sub
    End If

    If SameColor(src, dst) Then
        MsgBox "这两个对象的颜色相同，没有可替换的内容。" & vbCrLf & vbCrLf & _
               "颜色：" & DescribeColor(src), vbInformation, "颜色替换"
        Exit Sub
    End If

    ' ---- 确认方向。选中顺序不一定等于点击顺序，让用户明确一次 ----
    s = InputBox( _
        "即将把全文档里的颜色替换掉：" & vbCrLf & vbCrLf & _
        "  原色 = " & DescribeColor(src) & "（第 1 个对象）" & vbCrLf & _
        "  新色 = " & DescribeColor(dst) & "（第 2 个对象）" & vbCrLf & vbCrLf & _
        "1 = 按上面这样替换" & vbCrLf & _
        "2 = 反过来（把第 1 个当新色、第 2 个当原色）" & vbCrLf & vbCrLf & _
        "填序号后点「确定」；点「取消」则什么都不做。", _
        "颜色替换", "1")

    If Len(Trim(s)) = 0 Then Exit Sub
    pick = 1
    If Left(Trim(s), 1) = "2" Then pick = 2

    If pick = 2 Then
        Set tmp = src
        Set src = dst
        Set dst = tmp
    End If

    n = ReplaceColorCore(src, dst)

    If n = 0 Then
        MsgBox "没有找到 " & DescribeColor(src) & " 这个颜色。" & vbCrLf & vbCrLf & _
               "可能的原因：" & vbCrLf & _
               "  · 文档里本来就没有这个颜色" & vbCrLf & _
               "  · 用了渐变 / 位图 / 图案填充（这些不在替换范围内）" & vbCrLf & _
               "  · 该颜色只在别的页面或容器内部", _
               vbInformation, "颜色替换"
    Else
        MsgBox "颜色替换完成，共替换 " & n & " 处。" & vbCrLf & vbCrLf & _
               DescribeColor(src) & "  →  " & DescribeColor(dst), _
               vbInformation, "颜色替换"
    End If
End Sub

' 从一个对象取「代表色」：优先纯色填充，没有就退回轮廓色（宽度 > 0 才算）
' 取不到返回 Nothing
Private Function PickColor(ByVal sh As Shape) As Color
    Dim c As Color

    If sh Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear

    If sh.Fill.Type = cdrUniformFill Then
        Set c = sh.Fill.UniformColor
    End If

    If Err.Number <> 0 Or c Is Nothing Then
        Err.Clear
        ' 无纯色填充：退回轮廓色，但只有真正有轮廓（宽度 > 0）才当色源
        If sh.Outline.Width > 0 Then Set c = sh.Outline.Color
    End If

    Err.Clear
    On Error GoTo 0

    Set PickColor = c
End Function

' 把颜色描述成 RGB 三元组，供对话框显示。只读已实测的 RGB* 成员。
Private Function DescribeColor(ByVal c As Color) As String
    Dim r As Long
    Dim g As Long
    Dim b As Long

    If c Is Nothing Then
        DescribeColor = "（无）"
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    r = c.RGBRed
    g = c.RGBGreen
    b = c.RGBBlue
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        DescribeColor = "（颜色）"
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    DescribeColor = "RGB(" & r & ", " & g & ", " & b & ")"
End Function

' 纯逻辑，供自检调用
Public Function ReplaceColorCore(ByVal src As Color, ByVal dst As Color) As Long
    Dim doc As Document
    Dim pg As Page
    Dim sh As Shape
    Dim n As Long

    Set doc = CorelDRAW.ActiveDocument
    If doc Is Nothing Then Exit Function

    CorelDRAW.Optimization = True
    doc.BeginCommandGroup "颜色替换"

    For Each pg In doc.Pages
        For Each sh In pg.Shapes.All
            n = n + ReplaceInShape(sh, src, dst)
        Next sh
    Next pg

    doc.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh

    ReplaceColorCore = n
End Function

' 处理一个形状（组会递归进去），返回替换的处数
Private Function ReplaceInShape(ByVal sh As Shape, ByVal src As Color, ByVal dst As Color) As Long
    Dim c As Color
    Dim kids As Object
    Dim n As Long

    If sh Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear

    ' 组：递归进去
    If sh.Type = cdrGroupShape Then
        Set kids = sh.Shapes.All
        If Err.Number = 0 And Not kids Is Nothing Then
            n = n + ReplaceInRange(kids, src, dst)
        End If
        Err.Clear
        On Error GoTo 0
        ReplaceInShape = n
        Exit Function
    End If

    ' 填充
    Set c = Nothing
    If sh.Fill.Type = cdrUniformFill Then
        Set c = sh.Fill.UniformColor
        If Not c Is Nothing Then
            If SameColor(c, src) Then
                sh.Fill.ApplyUniformFill dst
                If Err.Number = 0 Then n = n + 1
                Err.Clear
            End If
        End If
    End If
    Err.Clear

    ' 轮廓
    Set c = Nothing
    Set c = sh.Outline.Color
    If Not c Is Nothing Then
        If SameColor(c, src) Then
            sh.Outline.Color.CopyAssign dst
            If Err.Number = 0 Then n = n + 1
            Err.Clear
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ReplaceInShape = n
End Function

Private Function ReplaceInRange(ByVal sr As Object, ByVal src As Color, ByVal dst As Color) As Long
    Dim i As Long
    Dim sh As Shape
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
            n = n + ReplaceInShape(sh, src, dst)
        End If
        Err.Clear
    Next i
    On Error GoTo 0

    ReplaceInRange = n
End Function