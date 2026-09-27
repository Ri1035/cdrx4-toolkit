Option Explicit

'==========================================================
'  CorelDRAW X4 probe #16 - how to make OUR project the
'  active one in the VBE, and what the File menu offers.
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe16.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

Dim app, vbe, p, c, cp, fm, ctl
Dim i, j, k, found

W "probe16 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
W "  create err " & Err.Number & " " & Err.Description
Err.Clear
app.InitializeVBA
W "  initvba err " & Err.Number & " " & Err.Description
Err.Clear
Set vbe = app.VBE
W "  vbe err " & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

If vbe Is Nothing Then
  W "!! no VBE"
  logFile.Close
  WScript.Quit 1
End If

' --- 1. what is active right now ---
On Error Resume Next
Err.Clear
W ""
W "1. ActiveVBProject.Name = [" & vbe.ActiveVBProject.Name & "]  err=" & Err.Number
Err.Clear
On Error GoTo 0

' --- 2. locate our project ---
found = 0
On Error Resume Next
For i = 1 To vbe.VBProjects.Count
  Err.Clear
  Set p = Nothing
  Set p = vbe.VBProjects.Item(i)
  If Err.Number = 0 And Not p Is Nothing Then
    If InStr(1, p.FileName, "CDRX4Toolkit.gms", vbTextCompare) > 0 Then found = i
  End If
  Err.Clear
Next
On Error GoTo 0
W "2. our project index = " & found
If found = 0 Then
  W "!! not loaded"
  logFile.Close
  WScript.Quit 1
End If

Set p = vbe.VBProjects.Item(found)
W "   name=" & p.Name & " comps=" & p.VBComponents.Count & " prot=" & p.Protection

' --- 3. is the VBE main window visible? ---
On Error Resume Next
Err.Clear
W ""
W "3. MainWindow.Visible (read) = " & vbe.MainWindow.Visible & " err=" & Err.Number
Err.Clear
vbe.MainWindow.Visible = True
W "   set Visible=True err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

' --- 4. activation attempt A: VBComponent.Activate ---
On Error Resume Next
Err.Clear
Set c = p.VBComponents.Item(1)
W ""
W "4. comp1 name=[" & c.Name & "] type=" & c.Type & " err=" & Err.Number
Err.Clear
c.Activate
W "   comp.Activate err=" & Err.Number & " " & Err.Description
Err.Clear
WScript.Sleep 1200
W "   ActiveVBProject.Name = [" & vbe.ActiveVBProject.Name & "] err=" & Err.Number
Err.Clear
On Error GoTo 0

' --- 5. activation attempt B: CodePane.Show + Window.SetFocus ---
On Error Resume Next
Err.Clear
Set cp = p.VBComponents.Item(1).CodeModule.CodePane
cp.Show
W ""
W "5. codepane.Show err=" & Err.Number & " " & Err.Description
Err.Clear
cp.Window.SetFocus
W "   Window.SetFocus err=" & Err.Number & " " & Err.Description
Err.Clear
WScript.Sleep 1200
W "   ActiveVBProject.Name = [" & vbe.ActiveVBProject.Name & "] err=" & Err.Number
Err.Clear
On Error GoTo 0

' --- 6. enumerate the VBE menu bar ---
On Error Resume Next
Err.Clear
W ""
W "6. CommandBars.Count = " & vbe.CommandBars.Count & " err=" & Err.Number
Err.Clear
Dim nbars
nbars = 0
nbars = vbe.CommandBars.Count
On Error GoTo 0

For k = 1 To nbars
  On Error Resume Next
  Err.Clear
  Dim bar
  Set bar = Nothing
  Set bar = vbe.CommandBars.Item(k)
  If Err.Number = 0 And Not bar Is Nothing Then
    Dim bn, bc
    bn = ""
    bc = -1
    bn = bar.Name
    bc = bar.Controls.Count
    W "   bar " & k & " name=[" & bn & "] controls=" & bc
    Err.Clear
    If k = 1 Then
      For i = 1 To bc
        Err.Clear
        Set ctl = Nothing
        Set ctl = bar.Controls.Item(i)
        If Err.Number = 0 And Not ctl Is Nothing Then
          W "      top " & i & " cap=[" & ctl.Caption & "] id=" & ctl.ID & _
            " type=" & ctl.Type
        End If
        Err.Clear
      Next
    End If
  End If
  Err.Clear
  On Error GoTo 0
Next

' --- 7. dig into the File menu (top control 1 of bar 1) ---
On Error Resume Next
Err.Clear
Set fm = vbe.CommandBars.Item(1).Controls.Item(1)
W ""
W "7. file menu cap=[" & fm.Caption & "] id=" & fm.ID & " sub=" & fm.Controls.Count
Err.Clear
Dim subn
subn = 0
subn = fm.Controls.Count
Err.Clear
For i = 1 To subn
  Err.Clear
  Set ctl = Nothing
  Set ctl = fm.Controls.Item(i)
  If Err.Number = 0 And Not ctl Is Nothing Then
    W "      sub " & i & " cap=[" & ctl.Caption & "] id=" & ctl.ID & _
      " type=" & ctl.Type & " enabled=" & ctl.Enabled
  Else
    W "      sub " & i & " err=" & Err.Number
  End If
  Err.Clear
Next
Err.Clear
On Error GoTo 0

W ""
W "done"

logFile.Close