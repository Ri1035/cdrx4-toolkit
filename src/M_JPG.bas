Attribute VB_Name = "M_JPG"
Option Explicit

'==========================================================
' JPG 批量导出
'   按页导出 JPG，可指定文件夹与 dpi
'==========================================================

Public Sub BatchExportJPG()
    Dim doc As Document
    Dim pg As Page
    Dim folder As String
    Dim dpi As Long
    Dim i As Long
    Dim opt As StructExportOptions
    Dim f As String
    Dim base As String

    If Not HasDocument() Then Exit Sub
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

    ' 导出选项，取不到就退回默认
    On Error Resume Next
    Set opt = CreateStructExportOptions()
    If opt Is Nothing Then Set opt = New StructExportOptions
    On Error GoTo 0

    If opt Is Nothing Then
        MsgBox "无法创建导出选项，导出取消。", vbExclamation, "JPG批量导出"
        Exit Sub
    End If

    opt.ImageType = cdrRGBColorImage
    opt.ResolutionX = dpi
    opt.ResolutionY = dpi
    opt.AntiAliasing = cdrNormalAntiAliasing

    base = BaseName(doc.FileName)

    Optimization = True
    doc.BeginCommandGroup "JPG批量导出"
    For i = 1 To doc.Pages.Count
        Set pg = doc.Pages(i)
        On Error Resume Next
        pg.Activate
        On Error GoTo 0

        f = folder & base & "_第" & i & "页.jpg"

        On Error Resume Next
        doc.ExportEx f, cdrJPEG, cdrCurrentPage, opt
        On Error GoTo 0
    Next i
    doc.EndCommandGroup
    Optimization = False

    MsgBox "已导出 " & doc.Pages.Count & " 页 JPG，位置：" & vbCrLf & folder, _
           vbInformation, "JPG批量导出"
End Sub