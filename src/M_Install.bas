Attribute VB_Name = "M_Install"
Option Explicit

'==========================================================
' CDR X4 增强工具包 —— 工具栏安装 / 卸载
'
' 重要：VBA 项目名必须为 CDRX4Toolkit，
'       否则下面拼出来的命令路径会失效。
'==========================================================

Public Const PRJ_NAME As String = "CDRX4Toolkit"
Public Const TOOLBAR_NAME As String = "增强工具"

' 命令清单：模块.过程 | 按钮名 | 提示
Private Function CmdList() As Variant
    CmdList = Array( _
        Array("M_Curves.ConvertAllToCurves", "全部转曲", "所有页面的文本与图形转换为曲线"), _
        Array("M_Rect.StdRectangle", "标准矩形", "把选中的圆角矩形重置为直角矩形"), _
        Array("M_Color.ReplaceColor", "颜色替换", "选中两个对象：用第一个的颜色替换第二个的"), _
        Array("M_CMYK.ConvertToCMYK", "转CMYK", "把文档中的 RGB 填充与轮廓转换为 CMYK"), _
        Array("M_FitPath.FitObjectsToPath", "对象适合路径", "把选中的对象沿一条路径分布"), _
        Array("M_JPG.BatchExportJPG", "JPG批量导出", "按页批量导出 JPG"), _
        Array("M_PageNo.InsertPageNumber", "插入页码", "为每一页在底部居中插入页码"), _
        Array("M_Calendar.CreateCalendar", "日历创建", "生成指定年月的日历"), _
        Array("M_Seal.CreateSeal", "印章制作", "生成圆形印章：环形文字 + 五角星"))
End Function

' 安装工具栏（带提示，手动执行用）
Public Sub InstallToolbar()
    InstallCore True
End Sub

' 静默安装（启动时自动调用）
Public Sub InstallToolbarSilent()
    InstallCore False
End Sub

Private Sub InstallCore(ByVal showMsg As Boolean)
    Dim items As Variant
    Dim cb As Object
    Dim btn As Object
    Dim i As Long
    Dim created As Boolean

    items = CmdList()

    ' 注册插件命令（重复注册会报错，忽略）
    On Error Resume Next
    For i = LBound(items) To UBound(items)
        CorelDRAW.AddPluginCommand PRJ_NAME & "." & items(i)(0), items(i)(1), items(i)(2)
    Next i
    On Error GoTo 0

    ' 已经装过就不再重建，避免每次启动都弄脏工作区；但把按钮文字刷成中文
    If ToolbarExists() Then
        RefreshCaptions items
        If showMsg Then MsgBox "工具栏「" & TOOLBAR_NAME & "」已经安装过了，按钮文字已刷新。", _
                               vbInformation, "增强工具"
        Exit Sub
    End If

    Set cb = Nothing
    On Error Resume Next
    Set cb = CorelDRAW.CommandBars.Add(TOOLBAR_NAME)
    On Error GoTo 0

    If cb Is Nothing Then
        If showMsg Then MsgBox "创建工具栏失败，请把报错反馈给我。", vbExclamation, "增强工具"
        Exit Sub
    End If

    cb.Visible = True
    created = True

    On Error Resume Next
    For i = LBound(items) To UBound(items)
        Set btn = cb.Controls.AddCustomButton(cdrCmdCategoryMacros, PRJ_NAME & "." & items(i)(0))
        ' 直接写中文标题，免去重启后才刷新的等待
        If Not (btn Is Nothing) Then btn.Caption = items(i)(1)
    Next i
    On Error GoTo 0

    If showMsg Then
        MsgBox "工具栏「" & TOOLBAR_NAME & "」安装完成，共 " & (UBound(items) + 1) & " 个功能按钮。", _
               vbInformation, "增强工具"
    End If
End Sub

Private Function ToolbarExists() As Boolean
    Dim cb As Object
    Set cb = Nothing
    On Error Resume Next
    Set cb = CorelDRAW.CommandBars(TOOLBAR_NAME)
    On Error GoTo 0
    ToolbarExists = Not (cb Is Nothing)
End Function

' 工具栏已存在时，把按钮文字刷成中文（兼容早期版本建出来的工具栏）
Private Sub RefreshCaptions(ByVal items As Variant)
    Dim cb As Object
    Dim i As Long
    Dim ok As Boolean

    Set cb = Nothing
    On Error Resume Next
    Set cb = CorelDRAW.CommandBars(TOOLBAR_NAME)
    ok = Not (cb Is Nothing)
    If ok Then ok = (cb.Controls.Count = (UBound(items) - LBound(items) + 1))
    If ok Then
        For i = LBound(items) To UBound(items)
            cb.Controls.Item(i - LBound(items) + 1).Caption = items(i)(1)
        Next i
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' 删除工具栏
Public Sub DeleteToolbar()
    On Error Resume Next
    CorelDRAW.CommandBars(TOOLBAR_NAME).Delete
    On Error GoTo 0
End Sub

' 卸载工具栏
Public Sub UninstallToolbar()
    DeleteToolbar
    MsgBox "工具栏「" & TOOLBAR_NAME & "」已卸载。", vbInformation, "增强工具"
End Sub