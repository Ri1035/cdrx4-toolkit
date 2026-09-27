Attribute VB_Name = "M_FitPath"
Option Explicit

'==========================================================
' 对象适合路径
'   用法：选中 2 个对象
'         第 1 个 = 路径（曲线；不是曲线会复制一份转曲来读节点）
'         第 2 个 = 要沿路径排的对象（会被复制若干份）
'
'   说明：X4 的 Shape.CreateBlend 一律 [13] 类型不匹配，
'         Application / Document / Layer 上都没有 CreateBlend，
'         所以不走「混合沿路径」，改成把对象副本沿路径节点折线
'         均匀插值排布。
'==========================================================

' 带对话框的入口
Public Sub FitObjectsToPath()
    Dim cnt As Long
    Dim n As Long

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "对象适合路径"
        Exit Sub
    End If

    cnt = Val(InputBox("沿路径排几个？默认 8", "对象适合路径", "8"))
    If cnt < 2 Then cnt = 8
    If cnt > 200 Then cnt = 200

    n = FitToPathCore(cnt)

    If n = 0 Then
        MsgBox "没排成功。请先选中 2 个对象：" & vbCrLf & vbCrLf & _
               "第 1 个 = 路径（曲线）" & vbCrLf & _
               "第 2 个 = 要沿路径排的对象", vbExclamation, "对象适合路径"
    Else
        MsgBox "已沿路径排布 " & n & " 个对象。", vbInformation, "对象适合路径"
    End If
End Sub

' 纯逻辑，供自检调用
Public Function FitToPathCore(ByVal copies As Long) As Long
    Dim sr As Object
    Dim doc As Document
    Dim pathShape As Shape
    Dim objShape As Shape
    Dim tempDup As Shape
    Dim cv As Object
    Dim xs() As Double
    Dim ys() As Double
    Dim acc() As Double
    Dim cnt As Long
    Dim j As Long
    Dim k As Long
    Dim seg As Long
    Dim total As Double
    Dim d As Double
    Dim t As Double
    Dim px As Double
    Dim py As Double
    Dim dup As Shape
    Dim i As Long
    Dim n As Long

    If copies < 2 Then copies = 2

    Set sr = Nothing
    On Error Resume Next
    Err.Clear
    Set sr = CorelDRAW.ActiveSelectionRange
    If Err.Number <> 0 Then Set sr = Nothing
    Err.Clear
    On Error GoTo 0
    If sr Is Nothing Then Exit Function
    If sr.Count < 2 Then Exit Function

    Set pathShape = Nothing
    Set objShape = Nothing
    On Error Resume Next
    Err.Clear
    Set pathShape = sr.Item(1)
    Set objShape = sr.Item(2)
    Err.Clear
    On Error GoTo 0
    If pathShape Is Nothing Or objShape Is Nothing Then Exit Function

    Set doc = CorelDRAW.ActiveDocument
    If doc Is Nothing Then Exit Function

    ' --- 取路径的节点 ---
    Set tempDup = Nothing
    Set cv = ShapeCurve(pathShape)
    If cv Is Nothing Then
        ' 路径不是曲线：复制一份、转曲、读节点，读完删掉复制件
        On Error Resume Next
        Err.Clear
        Set tempDup = pathShape.Duplicate
        If Err.Number = 0 And Not tempDup Is Nothing Then
            tempDup.ConvertToCurves
            Set cv = ShapeCurve(tempDup)
        End If
        Err.Clear
        On Error GoTo 0
    End If

    If cv Is Nothing Then Exit Function

    cnt = 0
    On Error Resume Next
    Err.Clear
    cnt = cv.Nodes.Count
    Err.Clear
    On Error GoTo 0
    If cnt < 2 Then
        If Not tempDup Is Nothing Then DeleteShape tempDup
        Exit Function
    End If

    ReDim xs(1 To cnt)
    ReDim ys(1 To cnt)
    On Error Resume Next
    For j = 1 To cnt
        xs(j) = cv.Nodes(j).PositionX
        ys(j) = cv.Nodes(j).PositionY
    Next j
    Err.Clear
    On Error GoTo 0

    ' 复制件用完就删
    If Not tempDup Is Nothing Then DeleteShape tempDup

    ' --- 按节点折线累计弧长 ---
    ReDim acc(1 To cnt)
    acc(1) = 0
    For j = 2 To cnt
        acc(j) = acc(j - 1) + Sqr((xs(j) - xs(j - 1)) ^ 2 + (ys(j) - ys(j - 1)) ^ 2)
    Next j
    total = acc(cnt)
    If total <= 0 Then Exit Function

    CorelDRAW.Optimization = True
    doc.BeginCommandGroup "对象适合路径"

    For i = 1 To copies
        d = (i - 1) / (copies - 1) * total

        seg = cnt - 1
        For k = 2 To cnt
            If acc(k) >= d Then
                seg = k - 1
                Exit For
            End If
        Next k
        If seg < 1 Then seg = 1
        If seg > cnt - 1 Then seg = cnt - 1

        If acc(seg + 1) > acc(seg) Then
            t = (d - acc(seg)) / (acc(seg + 1) - acc(seg))
        Else
            t = 0
        End If

        px = xs(seg) + (xs(seg + 1) - xs(seg)) * t
        py = ys(seg) + (ys(seg + 1) - ys(seg)) * t

        Set dup = Nothing
        On Error Resume Next
        Err.Clear
        If i = 1 Then
            Set dup = objShape
        Else
            Set dup = objShape.Duplicate
        End If
        Err.Clear
        On Error GoTo 0

        If Not dup Is Nothing Then
            CenterAt dup, px, py
            n = n + 1
        End If
    Next i

    doc.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh

    FitToPathCore = n
End Function

Private Sub DeleteShape(ByVal sh As Shape)
    On Error Resume Next
    sh.Delete
    On Error GoTo 0
End Sub