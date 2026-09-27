Attribute VB_Name = "M_CMYK"
Option Explicit

'==========================================================
' 转CMYK
'   把文档里的 RGB 等填充 / 轮廓转成 CMYK 颜色
'   可选给纯黑加上叠印
'
'   说明：原版插件的对话框还有分辨率 / 反锯齿 /
'         ICC 配置文件等选项，那些是导出参数；
'         本版只做色彩模型转换 + 纯黑叠印。
'==========================================================

Public Sub ConvertToCMYK()
    Dim doc As Document
    Dim pg As Page
    Dim sh As Shape
    Dim ans As VbMsgBoxResult
    Dim withBlack As Boolean
    Dim cnt As Long

    If Not HasDocument() Then Exit Sub
    Set doc = CorelDRAW.ActiveDocument

    ans = MsgBox("转 CMYK 选项：" & vbCrLf & vbCrLf & _
                 "是 = 转 CMYK，并给纯黑加叠印" & vbCrLf & _
                 "否 = 只转 CMYK，不动叠印" & vbCrLf & _
                 "取消 = 什么也不做", _
                 vbYesNoCancel + vbQuestion, "转CMYK")
    If ans = vbCancel Then Exit Sub
    withBlack = (ans = vbYes)

    Optimization = True
    doc.BeginCommandGroup "转CMYK"
    For Each pg In doc.Pages
        For Each sh In pg.Shapes.All
            cnt = cnt + CMYKShape(sh, withBlack)
        Next sh
    Next pg
    doc.EndCommandGroup
    Optimization = False
    DoRefresh

    MsgBox "转 CMYK 完成，共处理 " & cnt & " 个填充。", vbInformation, "转CMYK"
End Sub

' 处理单个形状，返回 1 表示改过填充
Private Function CMYKShape(ByVal sh As Shape, ByVal withBlack As Boolean) As Long
    Dim c As Color
    Dim c2 As Color
    Dim blk As Color

    On Error Resume Next

    If sh.Fill.Type = cdrUniformFill Then
        Set c = sh.Fill.UniformColor
        If Not c Is Nothing Then
            Set c2 = Nothing
            Set c2 = c.ConvertToCMYK
            If c2 Is Nothing Then
                c.ConvertToCMYK
                Set c2 = c
            End If
            If Not c2 Is Nothing Then
                sh.Fill.ApplyUniformFill c2
                CMYKShape = 1
                If withBlack Then
                    Set blk = CreateCMYKColor(0, 0, 0, 100)
                    If c2.IsSame(blk) Then sh.Fill.OverprintFill = True
                End If
            End If
        End If
    End If

    Set c = Nothing
    Set c2 = Nothing
    Set blk = Nothing
    On Error GoTo 0
End Function