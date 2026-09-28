Option Explicit

'==========================================================
'  Run a list of CDRX4Toolkit macros, in order, inside ONE
'  fresh CorelDRAW session, and dump the toolbar state log
'  after each one.
'
'  usage: cscript //nologo tools\run_macros.vbs M_A.Sub M_B.Sub ...
'
'  After every macro the contents of
'  %TEMP%\cdrx4_toolbar_state.log are printed, so a
'  ReportToolbar / RegisterCommands pair shows exactly what
'  the toolbar looks like before and after.
'
'  Log -> _run.log (UTF-16).  ASCII ONLY.
'==========================================================

Dim fso, sh, root, logPath, gLog, tgt, stateLog
Dim app, vbe, p, st, txt, i, macroList

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
logPath = root & "\_run.log"
gLog = ""

If WScript.Arguments.Count < 1 Then
  WScript.Echo "usage: cscript //nologo tools\run_macros.vbs <Macro> [Macro ...]"
  WScript.Quit 2
End If
Set macroList = WScript.Arguments

Sub W(s)
  gLog = gLog & s & vbCrLf
  WScript.Echo s
End Sub

Sub Flush()
  Dim ts
  On Error Resume Next
  Set ts = fso.CreateTextFile(logPath, True, True)
  If Not ts Is Nothing Then
    ts.Write gLog
    ts.Close
  End If
  On Error GoTo 0
End Sub

Sub Die(msg)
  W msg
  Flush
  WScript.Quit 1
End Sub

Function ProjectForTgt(vbe, tgt)
  Dim j, pr, pf
  Set ProjectForTgt = Nothing
  On Error Resume Next
  For j = 1 To vbe.VBProjects.Count
    Err.Clear
    Set pr = Nothing
    Set pr = vbe.VBProjects.Item(j)
    If Err.Number = 0 And Not pr Is Nothing Then
      pf = ""
      pf = pr.FileName
      If Err.Number = 0 And Len(pf) > 0 Then
        If StrComp(fso.GetFileName(pf), fso.GetFileName(tgt), vbTextCompare) = 0 Then
          Set ProjectForTgt = pr
        End If
      End If
    End If
    Err.Clear
  Next
  Err.Clear
  On Error GoTo 0
End Function

Sub DumpState()
  If Not fso.FileExists(stateLog) Then
    W "  (no state log)"
    Exit Sub
  End If
  Set st = fso.OpenTextFile(stateLog, 1, False, -1)
  txt = st.ReadAll
  st.Close
  W txt
End Sub

tgt = sh.ExpandEnvironmentStrings("%APPDATA%") & _
      "\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms"
stateLog = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_toolbar_state.log"

W "==== run_macros " & Now & " ===="
If Not fso.FileExists(tgt) Then Die("FAIL: no GMS at " & tgt)

' --- close anything already running ---------------------------
Set app = Nothing
On Error Resume Next
Set app = GetObject(, "CorelDRAW.Application.14")
On Error GoTo 0
If Not app Is Nothing Then
  W "quitting the running CorelDRAW first"
  On Error Resume Next
  app.Quit
  On Error GoTo 0
  For i = 1 To 60
    WScript.Sleep 500
    Set app = Nothing
    On Error Resume Next
    Set app = GetObject(, "CorelDRAW.Application.14")
    On Error GoTo 0
    If app Is Nothing Then Exit For
  Next
End If
Set app = Nothing

' --- fresh session -------------------------------------------
On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
W "boot   err=" & Err.Number & " " & Err.Description
Err.Clear
app.Visible = True
app.InitializeVBA
Err.Clear
Set vbe = Nothing
For i = 1 To 20
  Set vbe = Nothing
  Err.Clear
  Set vbe = app.VBE
  If Err.Number = 0 And Not vbe Is Nothing Then Exit For
  Err.Clear
  WScript.Sleep 500
Next
On Error GoTo 0
If vbe Is Nothing Then Die("FAIL: no VBE")

Set p = Nothing
For i = 1 To 60
  Set p = ProjectForTgt(vbe, tgt)
  If Not p Is Nothing Then Exit For
  WScript.Sleep 500
Next
If p Is Nothing Then Die("FAIL: CDRX4Toolkit did not load")
W "project loaded comps=" & p.VBComponents.Count

' --- run each macro ------------------------------------------
For i = 0 To macroList.Count - 1
  W ""
  W "--- [" & i & "] " & macroList(i) & " ---"
  If fso.FileExists(stateLog) Then fso.DeleteFile stateLog, True
  On Error Resume Next
  Err.Clear
  app.GMSManager.RunMacro "CDRX4Toolkit", macroList(i)
  W "  err=" & Err.Number & " " & Err.Description
  Err.Clear
  On Error GoTo 0
  Dim waited
  For waited = 1 To 30
    If fso.FileExists(stateLog) Then Exit For
    WScript.Sleep 500
  Next
  DumpState
Next

' --- leave CorelDRAW running so the caller can inspect it -----
W ""
W "==== done (CorelDRAW left running) ===="
Flush
WScript.Quit 0