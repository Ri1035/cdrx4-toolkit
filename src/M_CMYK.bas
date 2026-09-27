Attribute VB_Name = "M_CMYK"
Option Explicit

'==========================================================
' 转CMYK
'   把文档里的 RGB 等填充 / 轮廓转成 CMYK 颜色
'   可选给纯黑加上叠印
'
'   X4 实测：
'     Color.ConvertToCMYK 是「语句」，就地转换，返回值是 Nothing
'     CorelDRAW.CreateCMYKColor(c,m,y,k) 造色
'     叠印写在 Shape 上（sh.OverprintFill / sh.OverprintOutline）
'==========================================================

' 带对话框的入口
Public Sub ConvertToCMYK()
    Dim ans As VbMsgBoxResult
    Dim withBlack As Boolean
    Dim n As Long

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "转CMYK"
        Exit Sub
    End If

    ans = MsgBox("转 CMYK 选项：" & vbCrLf & vbCrLf & _
                 "是 = 转 CMYK，并给纯黑加叠印" & vbCrLf & _
                 "否 = 只转 CMYK，不动叠印" & vbCrLf & _
                 "取消 = 什么也不做", _
                 vbYesNoCancel + vbQuestion, "转CMYK")
    If ans = vbCancel Then Exit Sub
    withBlack = (ans = vbYes)

    n = ConvertToCMYKCore(withBlack)

    MsgBox "转 CMYK 完成，共处理 " & n & " 处颜色。", vbInformation, "转CMYK"
End Sub

' 纯逻辑，供自检调用
Public Function ConvertToCMYKCore(ByVal withBlack As Boolean) As Long
    Dim doc As Document
    Dim pg As Page
    Dim n As Long

    Set doc = CorelDRAW.ActiveDocument
    If doc Is Nothing Then Exit Function

    CorelDRAW.Optimization = True
    doc.BeginCommandGroup "转CMYK"

    For Each pg In doc.Pages
        n = n + CMYKRange(pg.Shapes.All, withBlack)
    Next pg

    doc.EndCommandGroup
    CorelDRAW.Optimization = False
    DoRefresh

    ConvertToCMYKCore = n
End Function

Private Function CMYKRange(ByVal sr As Object, ByVal withBlack As Boolean) As Long
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
                Set kids = sh.Shapes.All
                If Err.Number = 0 And Not kids Is Nothing Then
                    n = n + CMYKRange(kids, withBlack)
                End If
                Err.Clear
            Else
                n = n + CMYKShape(sh, withBlack)
            End If
        End If
        Err.Clear
    Next i
    On Error GoTo 0

    CMYKRange = n
End Function

' 处理单个形状，返回改过的颜色处数
Private Function CMYKShape(ByVal sh As Shape, ByVal withBlack As Boolean) As Long
    Dim c As Color
    Dim o As Object
    Dim n As Long

    On Error Resume Next
    Set o = sh

    '--- 填充 ---
    Err.Clear
    Set c = Nothing
    If o.Fill.Type = cdrUniformFill Then
        Set c = o.Fill.UniformColor
    End If
    If Err.Number = 0 And Not c Is Nothing Then
        c.ConvertToCMYK
        If Err.Number = 0 Then
            o.Fill.ApplyUniformFill c
            If Err.Number = 0 Then
                n = n + 1
                ' 叠印写在 Shape 上，Fill 上没有这个属性
                If withBlack Then
                    If IsPureBlack(c) Then
                        o.OverprintFill = True
                        Err.Clear
                    End If
                End If
            End If
        End If
        Err.Clear
    End If
    Err.Clear

    '--- 轮廓 ---
    Set c = Nothing
    Set c = o.Outline.Color
    If Err.Number = 0 And Not c Is Nothing Then
        c.ConvertToCMYK
        If Err.Number = 0 Then
            o.Outline.Color.CopyAssign c
            If Err.Number = 0 Then n = n + 1
        End If
        Err.Clear
    End If
    Err.Clear
    On Error GoTo 0

    CMYKShape = n
End Function

' 是否纯黑（C0 M0 Y0 K100）
Private Function IsPureBlack(ByVal c As Color) As Boolean
    Dim blk As Color
    Dim ok As Boolean

    On Error Resume Next
    Err.Clear
    Set blk = CorelDRAW.CreateCMYKColor(0, 0, 0, 100)
    If Err.Number = 0 And Not blk Is Nothing Then
        ok = SameColor(c, blk)
    End If
    Err.Clear
    On Error GoTo 0

    IsPureBlack = ok
End Function