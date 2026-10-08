Attribute VB_Name = "M_Install"
Option Explicit

'==========================================================
' CDR X4 增强工具包 —— 按钮清单（唯一来源）+ 自检导出
'
' 【本模块不碰任何 UI，也不碰任何 CorelDRAW 对象】
'
' v1.0.2 曾经在这里用 CommandBars API（Delete / Add / Visible /
' AddCustomButton）建工具栏，并从 GlobalMacroStorage_Start 启动回调里
' 调用它 —— 每次启动 CorelDRAW 必崩（CrlFrmWk.dll，0xc0000005 →
' 0xc000041d，留下杀不掉的 0 线程僵尸）。外部自动化在「工作区里还没有
' 这条工具栏」时建同样会崩（已复现两次）。结论：这条路本身不行，彻底
' 放弃。PLAN.md §10 有事件日志证据和对照实验。
'
' 现在的工具栏是**纯工作区 XML**：安装器在 CorelDRAW 关闭时，把
' workspace_markup.txt 的三段标记（10 个 itemData + 2 个 commandBarData
' + 1 条可见性）写进 DRAWUIConfig.xml，重启后 CorelDRAW 自己把它画出来。
' 所以本模块只剩一件事 —— CmdList()，它既是按钮清单的唯一来源，
' 也是构建 / 冒烟用来核对「GMS 里编译出来的清单」与「装进工作区的标记」
' 是否一致的依据（DumpItems 把它导成文件给外面读）。
'
' 【编译红线】绝对不要写
'       Dim app As Object
'       Set app = CorelDRAW
' X4 的 VBA 把 CorelDRAW 这个全局对象整体赋给变量会直接报
' 「编译错误：类型不匹配」（VBE 高亮 CorelDRAW 这个词）。VBA 是
' 「一处编译不过 → 整个工程所有宏全废」，所以这条红线必须守住。
' 本模块现在连 CorelDRAW.成员 都不需要，只用 ADODB 写文件。
'==========================================================

Public Const PRJ_NAME As String = "CDRX4Toolkit"
Public Const TOOLBAR_NAME As String = "增强工具"

' 命令清单：模块.过程 | 按钮名 | 提示
' 顺序按用户要求：转曲 → CMYK → 颜色替换 → 其余
'
' 展开按钮（flyout opener）不在这里 —— 它不绑宏，只负责展开菜单，
' 由 workspace_markup.txt 单独声明。
Private Function CmdList() As Variant
    CmdList = Array( _
        Array("M_Curves.ConvertAllToCurves", "全部转曲", "所有页面的文本与图形转换为曲线（可选只转文字或只转图形）"), _
        Array("M_CMYK.ConvertToCMYK", "转CMYK", "把文档中的 RGB 填充与轮廓转换为 CMYK"), _
        Array("M_Color.ReplaceColor", "颜色替换", "选中两个对象：用第一个的颜色替换第二个的"), _
        Array("M_Rect.StdRectangle", "标准矩形", "把选中的圆角矩形重置为直角矩形"), _
        Array("M_FitPath.FitObjectsToPath", "对象适合路径", "把选中的对象沿一条路径分布"), _
        Array("M_JPG.BatchExportJPG", "JPG批量导出", "按页批量导出 JPG"), _
        Array("M_PageNo.InsertPageNumber", "插入页码", "为每一页在底部居中插入页码"), _
        Array("M_Calendar.CreateCalendar", "日历创建", "生成指定年月的日历"), _
        Array("M_Seal.CreateSeal", "印章制作", "生成圆形印章：环形文字 + 五角星"))
End Function

'==========================================================
' 把按钮清单导出到 %TEMP%\cdrx4_toolbar_items.txt（UTF-8），
' 一行一个：proc|caption|tooltip
'
' 为什么需要：workspace_markup.txt 里写死了中文按钮名与宏路径，它必须
' 和 CmdList() 保持一致。构建 / 冒烟拿这份导出文件跟装进工作区的
' DRAWUIConfig.xml 逐项对（宏路径是纯 ASCII，便于比对），就能保证两边
' 不漂移。顺带它也是 M_Install 的**编译检查** —— RunMacro 能调通、
' 文件能写出来，就说明这个模块编译通过（§3.5 那条红线的回归测试）。
'==========================================================
Public Sub DumpItems()
    Dim items As Variant
    Dim st As Object
    Dim raw As Variant
    Dim i As Long
    Dim s As String

    items = CmdList()
    For i = LBound(items) To UBound(items)
        s = s & items(i)(0) & "|" & items(i)(1) & "|" & items(i)(2) & vbLf
    Next i

    ' UTF-8 WITHOUT a BOM: ADODB.Stream in text mode prepends one, and a
    ' BOM would ride along on the first line and break every string
    ' comparison the build / smoke / verify scripts do against this file.
    On Error Resume Next
    Err.Clear
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1
    st.Position = 3          ' skip the BOM ADODB just wrote
    raw = st.Read
    st.Close
    If Err.Number = 0 Then
        Set st = CreateObject("ADODB.Stream")
        st.Type = 1
        st.Open
        st.Write raw
        st.SaveToFile Environ$("TEMP") & "\cdrx4_toolbar_items.txt", 2
        st.Close
    End If
    Err.Clear
    On Error GoTo 0
End Sub