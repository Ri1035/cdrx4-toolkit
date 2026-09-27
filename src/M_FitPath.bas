Attribute VB_Name = "M_FitPath"
Option Explicit

'==========================================================
' 对象适合路径（v1 简化版）
'
' 用法：选中 1 条曲线 + 2 个对象
'       曲线当路径，两个对象用混合沿路径分布
'
' 说明：CorelDRAW X4 没有对应的原生 API，
'       本版走「混合沿路径」，
'       结果是可调整的混合组，不是拆散后的独立对象。
'==========================================================

Public Sub FitObjectsToPath()
    Dim sr As ShapeRange
    Dim i As Long
    Dim cntPath As Long
    Dim pathShape As Shape
    Dim a As Shape
    Dim b As Shape
    Dim eff As Object

    If Not HasDocument() Then Exit Sub

    Set sr = CorelDRAW.ActiveSelection
    If sr.Count < 3 Then
        MsgBox "请选中 1 条曲线 + 至少 2 个对象。", vbExclamation, "对象适合路径"
        Exit Sub
    End If

    ' 先数一数选中里有几条曲线
    For i = 1 To sr.Count
        If sr(i).Type = cdrCurveShape Then
            cntPath = cntPath + 1
            If cntPath = 1 Then Set pathShape = sr(i)
        End If
    Next i

    If cntPath <> 1 Then
        MsgBox "选中的曲线有 " & cntPath & " 条，必须正好 1 条当路径。", _
               vbExclamation, "对象适合路径"
        Exit Sub
    End If

    ' 另外两个对象当作混合的起止
    For i = 1 To sr.Count
        If Not sr(i) Is pathShape Then
            If a Is Nothing Then
                Set a = sr(i)
            ElseIf b Is Nothing Then
                Set b = sr(i)
            End If
        End If
    Next i

    If a Is Nothing Or b Is Nothing Then
        MsgBox "除曲线外还需要 2 个对象。", vbExclamation, "对象适合路径"
        Exit Sub
    End If

    Set eff = Nothing
    On Error Resume Next
    Set eff = a.CreateBlend(b)
    If Not eff Is Nothing Then eff.Blend.Path = pathShape
    On Error GoTo 0

    If eff Is Nothing Then
        MsgBox "建立混合失败，请换两个对象再试。", vbExclamation, "对象适合路径"
        Exit Sub
    End If

    DoRefresh
End Sub