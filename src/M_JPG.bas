Attribute VB_Name = "M_JPG"
Option Explicit

'==========================================================
' JPG 批量导出
'   按页导出 JPG，可指定文件夹与 dpi
'
'   X4 实测（依据自带 FileConverter.gms 源码）：
'     doc.ExportEx(文件名, cdrJPEG, cdrCurrentPage, se, pal) 是 5 个参数
'     返回 ExportFilter，必须再调 ex.Finish 才会真正写文件
'     选项属性是 AntiAliasingType，不是 AntiAliasing
'==========================================================

' 带对话框的入口
Public Sub BatchExportJPG()
    Dim doc As Document
    Dim folder As String
    Dim dpi As Long
    Dim n As Long

    If Not HasDocument() Then
        MsgBox "当前没有打开的文档。", vbExclamation, "JPG批量导出"
        Exit Sub
    End If
    Set doc = CorelDRAW.ActiveDocument

    folder = InputBox("导出到哪个文件夹？留空 = 桌面", "JPG批量导出", DesktopPath())
    If Len(folder) = 0 Then Exit Sub
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    If Dir(folder, vbDirectory) = "" Then
        MsgBox "找不到这个文件夹：" & folder, vbExclamation, "JPG批量导出"
        Exit Sub
    End If

    dpi = Val(InputBox("输入导出 dpi，默认 300。", "JPG批量导出", "300"))
    If dpi <= 0 Then dpi = 300

    n = ExportJPGTo(folder, dpi)

    If n = 0 Then
        MsgBox "一页都没导出成功，请把报错反馈给我。", vbExclamation, "JPG批量导出"
    Else
        MsgBox "已导出 " & n & " 页 JPG，位置：" & vbCrLf & folder, vbInformation, "JPG批量导出"
    End If
End Sub

' 纯逻辑，供自检调用：把每一页导出成 folder\名字_第N页.jpg
Public Function ExportJPGTo(ByVal folder As String, ByVal dpi As Long) As Long
    Dim doc As Document
    Dim pg As Page
    Dim se As Object
    Dim pal As Object
    Dim ex As Object
    Dim f As String
    Dim base As String
    Dim idx As Long
    Dim n As Long

    If Not HasDocument() Then Exit Function
    Set doc = CorelDRAW.ActiveDocument
    If Len(folder) = 0 Then Exit Function
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    If dpi <= 0 Then dpi = 300

    Set se = Nothing
    Set pal = Nothing
    On Error Resume Next
    Err.Clear
    Set se = CorelDRAW.CreateStructExportOptions
    Set pal = CorelDRAW.CreateStructPaletteOptions
    Err.Clear
    On Error GoTo 0

    If se Is Nothing Or pal Is Nothing Then Exit Function

    On Error Resume Next
    se.ImageType = cdrRGBColorImage
    se.ResolutionX = dpi
    se.ResolutionY = dpi
    se.AntiAliasingType = cdrNormalAntiAliasing
    Err.Clear
    On Error GoTo 0

    base = BaseName(doc.FileName)

    CorelDRAW.Optimization = True

    idx = 0
    For Each pg In doc.Pages
        idx = idx + 1

        On Error Resume Next
        pg.Activate
        Err.Clear

        f = folder & base & "_第" & idx & "页.jpg"

        Set ex = Nothing
        Set ex = doc.ExportEx(f, cdrJPEG, cdrCurrentPage, se, pal)
        If Err.Number = 0 And Not ex Is Nothing Then
            ex.Compression = 90
            ex.Finish
            If Err.Number = 0 Then n = n + 1
        End If
        Err.Clear
        On Error GoTo 0
    Next pg

    CorelDRAW.Optimization = False

    ExportJPGTo = n
End Function