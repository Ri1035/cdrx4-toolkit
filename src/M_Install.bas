Attribute VB_Name = "M_Install"
Option Explicit

'==========================================================
' CDR X4 增强工具包 —— 工具栏安装 / 卸载
'
' 重要：VBA 项目名必须为 CDRX4Toolkit，
'       否则下面拼出来的命令路径会失效。
'
' 【编译红线】绝对不要写
'       Dim app As Object
'       Set app = CorelDRAW
' X4 的 VBA 把 CorelDRAW 这个全局对象整体赋给变量，会直接报
' 「编译错误：类型不匹配」（VBE 会高亮 CorelDRAW 这个词）。
' 而 VBA 是「一处编译不过 → 整个工程所有宏全废」，于是每次启动
' CorelDRAW 都弹编译错误框、工具栏永远装不上。
'
' 正确写法是直接用 `CorelDRAW.成员`；要晚期绑定就把「取回来的成员」
' 放进 Object 变量，例如：
'       Dim cb As Object
'       Set cb = CorelDRAW.CommandBars(TOOLBAR_NAME)
' 这样既能拿到晚期绑定的好处，又不会踩上面那条红线。
'==========================================================

Public Const PRJ_NAME As String = "CDRX4Toolkit"
Public Const TOOLBAR_NAME As String = "增强工具"

' 命令清单：模块.过程 | 按钮名 | 提示
' 顺序按用户要求：转曲 → CMYK → 颜色替换 → 其余
Private Function CmdList() As Variant
    CmdList = Array( _
        Array("M_Curves.ConvertAllToCurves", "全部转曲", "所有页面的文本与图形转换为曲线"), _
        Array("M_CMYK.ConvertToCMYK", "转CMYK", "把文档中的 RGB 填充与轮廓转换为 CMYK"), _
        Array("M_Color.ReplaceColor", "颜色替换", "选中两个对象：用第一个的颜色替换第二个的"), _
        Array("M_Rect.StdRectangle", "标准矩形", "把选中的圆角矩形重置为直角矩形"), _
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

    items = CmdList()

    ' 注册插件命令（重复注册会报错，忽略）
    On Error Resume Next
    For i = LBound(items) To UBound(items)
        Err.Clear
        CorelDRAW.AddPluginCommand PRJ_NAME & "." & items(i)(0), items(i)(1), items(i)(2)
        Err.Clear
    Next i
    On Error GoTo 0

    ' 先删掉旧的再重建。早期版本建出来的工具栏按钮顺序和现在不同，
    ' 只刷新 Caption 会让「按钮名」和「实际动作」错位，所以宁可重建。
    On Error Resume Next
    Err.Clear
    CorelDRAW.CommandBars(TOOLBAR_NAME).Delete
    Err.Clear
    On Error GoTo 0

    Set cb = Nothing
    On Error Resume Next
    Err.Clear
    Set cb = CorelDRAW.CommandBars.Add(TOOLBAR_NAME)
    If Err.Number <> 0 Then Set cb = Nothing
    Err.Clear
    On Error GoTo 0

    If cb Is Nothing Then
        If showMsg Then MsgBox "创建工具栏失败，请把报错反馈给我。", vbExclamation, "增强工具"
        Exit Sub
    End If

    On Error Resume Next
    cb.Visible = True
    On Error GoTo 0

    On Error Resume Next
    For i = LBound(items) To UBound(items)
        Set btn = Nothing
        Err.Clear
        Set btn = cb.Controls.AddCustomButton(cdrCmdCategoryMacros, PRJ_NAME & "." & items(i)(0))
        If Err.Number = 0 And Not btn Is Nothing Then
            ' 直接写中文标题，免去重启后才刷新的等待
            btn.Caption = items(i)(1)
            btn.TooltipText = items(i)(2)
        End If
        Err.Clear
    Next i
    On Error GoTo 0

    If showMsg Then
        MsgBox "工具栏「" & TOOLBAR_NAME & "」安装完成，共 " & (UBound(items) + 1) & " 个功能按钮。", _
               vbInformation, "增强工具"
    End If
End Sub

' 删除工具栏
Public Sub DeleteToolbar()
    On Error Resume Next
    Err.Clear
    CorelDRAW.CommandBars(TOOLBAR_NAME).Delete
    Err.Clear
    On Error GoTo 0
End Sub

' 卸载工具栏
Public Sub UninstallToolbar()
    DeleteToolbar
    MsgBox "工具栏「" & TOOLBAR_NAME & "」已卸载。", vbInformation, "增强工具"
End Sub

'==========================================================
' 工具栏自检：装一遍，再把「真实建出来的按钮」逐项写进日志
' 结果写 %TEMP%\cdrx4_toolbar.log
'
' 为什么不让外面的冒烟脚本查：X4 不把 CommandBars 交给外部自动化
' 客户端，app.CommandBars(名字) 只会返回 Nothing + err=13 类型不匹配，
' 所以工具栏的验收只能由 VBA 自己做完再落盘。
'==========================================================
Public Sub DiagToolbar()
    Dim fso As Object
    Dim items As Variant
    Dim cb As Object
    Dim btn As Object
    Dim i As Long
    Dim n As Long
    Dim cap As String
    Dim ttl As String
    Dim log As String
    Dim logPath As String
    Dim errNo As Long

    logPath = Environ$("TEMP") & "\cdrx4_toolbar.log"
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Sub

    items = CmdList()
    log = "CDRX4Toolkit toolbar " & Now & vbCrLf
    log = log & "expect       n=" & (UBound(items) + 1) & vbCrLf

    ' 期望值也写进同一份日志：外面只需要读一个文件、一种编码就能比对，
    ' 不用再去解析 src\M_Install.bas 的 UTF-8 源码。
    For i = LBound(items) To UBound(items)
        log = log & "exp " & (i + 1) & " caption=[" & items(i)(1) & "] proc=[" & items(i)(0) & "]" & vbCrLf
    Next i

    ' 装一遍（先删旧的重建，幂等）
    On Error Resume Next
    Err.Clear
    InstallCore False
    errNo = Err.Number
    Err.Clear
    On Error GoTo 0
    log = log & "install      err=" & errNo & vbCrLf

    ' 再从工具栏里取回来逐项核对
    Set cb = Nothing
    On Error Resume Next
    Err.Clear
    Set cb = CorelDRAW.CommandBars(TOOLBAR_NAME)
    errNo = Err.Number
    Err.Clear
    On Error GoTo 0

    log = log & "toolbar      found=" & CStr(Not (cb Is Nothing)) & "  err=" & errNo & vbCrLf

    If Not cb Is Nothing Then
        n = 0
        On Error Resume Next
        n = cb.Controls.Count
        Err.Clear
        On Error GoTo 0
        log = log & "controls     n=" & n & vbCrLf

        For i = 1 To n
            cap = ""
            ttl = ""
            On Error Resume Next
            Err.Clear
            Set btn = Nothing
            Set btn = cb.Controls.Item(i)
            If Err.Number = 0 And Not btn Is Nothing Then
                cap = btn.Caption
                ttl = btn.TooltipText
            End If
            errNo = Err.Number
            Err.Clear
            On Error GoTo 0
            log = log & "btn " & i & " caption=[" & cap & "] tooltip=[" & ttl & "] err=" & errNo & vbCrLf
        Next i
    End If

    log = log & "done" & vbCrLf
    WriteText fso, logPath, log
End Sub

Private Sub WriteText(ByVal fso As Object, ByVal p As String, ByVal s As String)
    Dim ts As Object

    On Error Resume Next
    Set ts = fso.CreateTextFile(p, True, True)
    If Not ts Is Nothing Then
        ts.Write s
        ts.Close
    End If
    On Error GoTo 0
End Sub