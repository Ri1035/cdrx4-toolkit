Option Explicit

'==========================================================
'  CorelDRAW X4 probe #18 - make project "CDRX4Toolkit" the
'  ACTIVE project in the VBE (that is what File > Save acts
'  on).  Read back both ActiveVBProject.Name and the caption
'  of the File menu's Save item after every strategy.
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe18.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

Dim app, vbe, p, c, cp, fm, ctl
Dim i, cnt, matchIdx, nm, fn

' caption of File menu item 3 (the Save item) tells us which
' project the menu would save
Function SaveCaption()
  Dim r
  r = "?"
  On Error Resume Next
  Err.Clear
  Set fm = vbe.CommandBars.Item(1).Controls.Item(1)
  If Err.Number = 0 And Not fm Is Nothing Then
    Set ctl = fm.Controls.Item(3)
    If Err.Number = 0 And Not ctl Is Nothing Then r = ctl.Caption
  End If
  Err.Clear
  On Error GoTo 0
  SaveCaption = r
End Function

Sub Snapshot(ByVal tag)
  Dim a
  a = "?"
  On Error Resume Next
  Err.Clear
  a = vbe.ActiveVBProject.Name
  Err.Clear
  On Error GoTo 0
  W "   >> " & tag & "  active=[" & a & "]  saveItem=[" & SaveCaption() & "]"
End Sub

W "probe18 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
W "  boot err " & Err.Number
Err.Clear
Set vbe = app.VBE
On Error GoTo 0
If vbe Is Nothing Then
  W "!! no VBE"
  logFile.Close
  WScript.Quit 1
End If

On Error Resume Next
cnt = vbe.VBProjects.Count
Err.Clear
On Error GoTo 0

matchIdx = 0
For i = 1 To cnt
  On Error Resume Next
  Err.Clear
  Set p = Nothing
  Set p = vbe.VBProjects.Item(i)
  If Err.Number = 0 And Not p Is Nothing Then
    nm = ""
    fn = ""
    Err.Clear
    nm = p.Name
    Err.Clear
    fn = p.FileName
    Err.Clear
    If InStr(1, fn, "CDRX4Toolkit.gms", vbTextCompare) > 0 Then
      matchIdx = i
      W "  match at " & i & " name=[" & nm & "] file=[" & fn & "]"
    End If
  End If
  Err.Clear
  On Error GoTo 0
Next
W "matchIdx = " & matchIdx
If matchIdx = 0 Then
  W "!! not loaded"
  logFile.Close
  WScript.Quit 1
End If

On Error Resume Next
Err.Clear
Set p = vbe.VBProjects.Item(matchIdx)
Err.Clear
On Error GoTo 0

W ""
W "A. baseline"
Snapshot "baseline"

W ""
W "B. MainWindow.Visible = True"
On Error Resume Next
Err.Clear
vbe.MainWindow.Visible = True
W "   err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0
Snapshot "after visible"

W ""
W "C. p.VBComponents(1).Activate  (ThisDocument)"
On Error Resume Next
Err.Clear
Set c = p.VBComponents.Item(1)
W "   comp name=[" & c.Name & "] type=" & c.Type & " err=" & Err.Number
Err.Clear
c.Activate
W "   activate err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0
WScript.Sleep 1500
Snapshot "after comp.Activate"

W ""
W "D. CodePane.Show + Window.SetFocus (ThisDocument)"
On Error Resume Next
Err.Clear
Set cp = Nothing
Set cp = p.VBComponents.Item(1).CodeModule.CodePane
W "   pane err=" & Err.Number & " nothing=" & (cp Is Nothing)
Err.Clear
cp.Show
W "   show err=" & Err.Number & " " & Err.Description
Err.Clear
cp.Window.SetFocus
W "   setfocus err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0
WScript.Sleep 1500
Snapshot "after pane show"

' which component is currently shown?
On Error Resume Next
Err.Clear
W "   vbe.ActiveCodePane err=" & Err.Number
Err.Clear
Dim acp
Set acp = Nothing
Set acp = vbe.ActiveCodePane
If Err.Number = 0 And Not acp Is Nothing Then
  Dim pn
  pn = ""
  pn = acp.Parent.Name
  W "   active codepane parent=[" & pn & "]"
Else
  W "   no active codepane err=" & Err.Number
End If
Err.Clear
On Error GoTo 0

W ""
W "E. activate a standard module (M_Install) then show its pane"
Dim j, target
target = 0
On Error Resume Next
For j = 1 To p.VBComponents.Count
  Err.Clear
  Set c = Nothing
  Set c = p.VBComponents.Item(j)
  If Err.Number = 0 And Not c Is Nothing Then
    If c.Name = "M_Install" Then target = j
  End If
  Err.Clear
Next
On Error GoTo 0
W "   M_Install index = " & target
If target > 0 Then
  On Error Resume Next
  Err.Clear
  Set c = p.VBComponents.Item(target)
  c.Activate
  W "   activate err=" & Err.Number
  Err.Clear
  Set cp = Nothing
  Set cp = c.CodeModule.CodePane
  cp.Show
  W "   show err=" & Err.Number
  Err.Clear
  cp.Window.SetFocus
  W "   setfocus err=" & Err.Number
  Err.Clear
  On Error GoTo 0
  WScript.Sleep 1500
  Snapshot "after M_Install pane"
End If

W ""
W "F. vbe.MainWindow.SetFocus then comp.Activate"
On Error Resume Next
Err.Clear
vbe.MainWindow.SetFocus
W "   main setfocus err=" & Err.Number
Err.Clear
Set c = p.VBComponents.Item(1)
c.Activate
W "   comp activate err=" & Err.Number
Err.Clear
On Error GoTo 0
WScript.Sleep 1500
Snapshot "after mainwindow focus"

W ""
W "done"
logFile.Close