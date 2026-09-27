Attribute VB_Name = "M_Util"
Option Explicit

'==========================================================
' CDR X4 增强工具包 —— 公共函数
' 供其余模块调用，本身不对应工具栏按钮
'==========================================================

' 当前是否有打开的文档
Public Function HasDocument() As Boolean
    On Error Resume Next
    HasDocument = (CorelDRAW.Documents.Count > 0)
    On Error GoTo 0
End Function

' 取页面宽高（单位 mm），取不到就返回 0
Public Sub PageWH(ByVal pg As Page, ByRef w As Double, ByRef h As Double)
    w = 0
    h = 0
    On Error Resume Next
    w = pg.SizeWidth
    h = pg.SizeHeight
    On Error GoTo 0
End Sub

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