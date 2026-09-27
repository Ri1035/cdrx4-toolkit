Attribute VB_Name = "M_Util"
Option Explicit

'==========================================================
' CDR X4 增强工具包 —— 公共函数
' 供其余模块调用，本身不对应工具栏按钮
'
' 约定：所有涉及「可能只读的属性」都用晚期绑定（As Object），
'       避免早期绑定在编译期就报「不能给只读属性赋值」，
'       那样会让整个 VBA 工程编译不过、所有宏一起失效。
'==========================================================

' 当前是否有打开的文档
Public Function HasDocument() As Boolean
    On Error Resume Next
    HasDocument = (CorelDRAW.Documents.Count > 0)
    On Error GoTo 0
End Function

'==========================================================
' 单位换算
' X4 实测 doc.Unit：1=英寸 2=英尺 3=毫米 4=厘米 5=像素
'                   6=英里 7=米 8=千米 9=迪多点 11=码 12=派卡
' 页面尺寸（SizeWidth/SizeHeight）跟着文档单位走，
' 所以生成类功能必须先换算，否则英寸文档里会差 25 倍。
'==========================================================

' 当前文档单位 -> 毫米 的换算系数
Public Function MmPerUnit() As Double
    Dim u As Long
    u = 3
    On Error Resume Next
    u = CorelDRAW.ActiveDocument.Unit
    On Error GoTo 0

    Select Case u
        Case 1:   MmPerUnit = 25.4              ' 英寸
        Case 2:   MmPerUnit = 304.8             ' 英尺
        Case 3:   MmPerUnit = 1                 ' 毫米
        Case 4:   MmPerUnit = 10                ' 厘米
        Case 5:   MmPerUnit = 25.4 / 300        ' 像素（按 300dpi 折算）
        Case 6:   MmPerUnit = 1609344           ' 英里
        Case 7:   MmPerUnit = 1000              ' 米
        Case 8:   MmPerUnit = 1000000           ' 千米
        Case 9:   MmPerUnit = 0.3759259         ' 迪多点
        Case 11:  MmPerUnit = 914.4             ' 码
        Case 12:  MmPerUnit = 4.2333333         ' 派卡
        Case Else: MmPerUnit = 1                ' 兜底当毫米
    End Select
End Function

' 毫米 -> 当前文档单位
Public Function MmToDoc(ByVal v As Double) As Double
    MmToDoc = v / MmPerUnit()
End Function

' 页面宽高，统一返回毫米
Public Sub PageWHmm(ByVal pg As Page, ByRef w As Double, ByRef h As Double)
    w = 0
    h = 0
    On Error Resume Next
    w = pg.SizeWidth * MmPerUnit()
    h = pg.SizeHeight * MmPerUnit()
    On Error GoTo 0
End Sub

' 当前页，取不到就退回第一页
Public Function ActivePageSafe() As Page
    Dim doc As Document
    Dim pg As Page

    On Error Resume Next
    Set doc = CorelDRAW.ActiveDocument
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Set pg = doc.ActivePage
    If Err.Number <> 0 Or pg Is Nothing Then
        Err.Clear
        For Each pg In doc.Pages
            Exit For
        Next pg
    End If
    Err.Clear
    On Error GoTo 0
    Set ActivePageSafe = pg
End Function

'==========================================================
' 摆放 / 样式
'==========================================================

' 把形状的包围盒中心移到 (x, y)，坐标是「当前文档单位」
' 先试 CenterX/CenterY（晚期绑定，只读也不会编译报错），
' 不行再按包围盒算偏移用 Move 兜底。
Public Sub CenterAt(ByVal sh As Shape, ByVal x As Double, ByVal y As Double)
    Dim o As Object
    Dim dx As Double
    Dim dy As Double
    Dim ok As Boolean

    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Err.Clear
    Set o = sh
    o.CenterX = x
    o.CenterY = y
    ok = (Err.Number = 0)
    If Not ok Then
        Err.Clear
        dx = x - (o.LeftX + o.SizeWidth / 2)
        dy = y - (o.BottomY + o.SizeHeight / 2)
        o.Move dx, dy
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' 设文字样式；字体名传空就只改字号/对齐
Public Sub SetTextStyle(ByVal sh As Shape, ByVal fontName As String, _
                        ByVal sizePt As Double, ByVal alignCenter As Boolean)
    Dim o As Object

    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Err.Clear
    Set o = sh
    If Len(fontName) > 0 Then o.Text.Story.Font = fontName
    If sizePt > 0 Then o.Text.Story.Size = sizePt
    If alignCenter Then o.Text.Story.Alignment = cdrCenterAlignment
    Err.Clear
    On Error GoTo 0
End Sub

' 在指定层建一段居中对齐的文本，中心放在 (xmm, ymm)（毫米）
' 返回建好的 Shape，失败返回 Nothing
Public Function MakeCenteredText(ByVal lay As Object, ByVal txt As String, _
                                 ByVal xmm As Double, ByVal ymm As Double, _
                                 ByVal sizePt As Double, ByVal fontName As String) As Shape
    Dim sh As Shape
    Dim ok As Boolean

    If lay Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    Set sh = lay.CreateArtisticTextWide(0, 0, txt)
    ok = (Err.Number = 0) And (Not sh Is Nothing)
    Err.Clear
    On Error GoTo 0

    If Not ok Then Exit Function

    SetTextStyle sh, fontName, sizePt, True
    CenterAt sh, MmToDoc(xmm), MmToDoc(ymm)

    Set MakeCenteredText = sh
End Function

'==========================================================
' 颜色小工具
'==========================================================

' 造一个 CMYK 颜色（X4 里 CreateCMYKColor 只能从 Application 上取）
Public Function MakeCMYK(ByVal cc As Long, ByVal mm As Long, _
                         ByVal yy As Long, ByVal kk As Long) As Color
    Dim col As Color
    On Error Resume Next
    Set col = CorelDRAW.CreateCMYKColor(cc, mm, yy, kk)
    On Error GoTo 0
    Set MakeCMYK = col
End Function

' 造一个 RGB 颜色
Public Function MakeRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Color
    Dim col As Color
    On Error Resume Next
    Set col = CorelDRAW.CreateRGBColor(r, g, b)
    On Error GoTo 0
    Set MakeRGB = col
End Function

' 两个颜色是否一样
Public Function SameColor(ByVal a As Color, ByVal b As Color) As Boolean
    Dim ok As Boolean
    If a Is Nothing Or b Is Nothing Then
        SameColor = False
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    ok = a.IsSame(b)
    If Err.Number <> 0 Then
        Err.Clear
        ok = (a.RGBRed = b.RGBRed) And (a.RGBGreen = b.RGBGreen) And (a.RGBBlue = b.RGBBlue)
        Err.Clear
    End If
    On Error GoTo 0
    SameColor = ok
End Function

'==========================================================
' 杂项
'==========================================================

' 桌面路径
Public Function DesktopPath() As String
    On Error Resume Next
    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0
End Function

' 去掉路径和扩展名，只留文件名
Public Function BaseName(ByVal p As String) As String
    Dim n As String
    Dim i As Long

    If Len(p) = 0 Then
        BaseName = "未命名"
        Exit Function
    End If

    n = p
    i = InStrRev(n, "\")
    If i > 0 Then n = Mid$(n, i + 1)
    i = InStrRev(n, ".")
    If i > 0 Then n = Left$(n, i - 1)
    BaseName = n
End Function

' 刷新界面
Public Sub DoRefresh()
    On Error Resume Next
    CorelDRAW.Refresh
    On Error GoTo 0
End Sub

' 让一个颜色就地变成 CMYK，成功返回 True
Public Function ToCMYK(ByVal c As Color) As Boolean
    Dim ok As Boolean
    If c Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    c.ConvertToCMYK
    ok = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    ToCMYK = ok
End Function

' VBA 没有 Atan2，自己来
Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    Const PI As Double = 3.14159265358979
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atan2 = Atn(y / x) + PI
        Else
            Atan2 = Atn(y / x) - PI
        End If
    Else
        If y > 0 Then
            Atan2 = PI / 2
        ElseIf y < 0 Then
            Atan2 = -PI / 2
        Else
            Atan2 = 0
        End If
    End If
End Function

' 取一个形状的曲线对象（不是曲线返回 Nothing）
Public Function ShapeCurve(ByVal sh As Shape) As Object
    Dim o As Object
    Dim cv As Object

    If sh Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    Set o = sh
    Set cv = o.Curve
    If Err.Number <> 0 Then Set cv = Nothing
    Err.Clear
    On Error GoTo 0
    Set ShapeCurve = cv
End Function