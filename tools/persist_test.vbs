Option Explicit

'==========================================================
'  Does a VBA-built toolbar survive a CorelDRAW restart?
'
'  Two sessions, back to back:
'    A. start CDR -> report (expect: no toolbar)
'                  -> install -> report (expect: 9 buttons)
'                  -> quit NORMALLY
'    B. start CDR -> report   <-- this is the answer
'
'  M_Install.ReportToolbar only READS CommandBars; it never
'  creates or deletes anything, so it cannot fake a pass.
'
'  Log -> _persist.log (UTF-16)
'
'  ASCII ONLY: WSH parses .vbs as ANSI.
'==========================================================

Dim fso, sh, root, logPath, gLog, tgt, stateLog
Dim app, vbe, p, waited, st, txt

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
logPath = root & "\_persist.log"
gLog = ""

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
  Dim i, pr, pf
  Set ProjectForTgt = Nothing
  On Error Resume Next
  For i = 1 To vbe.VBProjects.Count
    Err.Clear
    Set pr = Nothing
    Set pr = vbe.VBProjects.Item(i)
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

Sub StartSession()
  Dim v
  Set app = Nothing
  On Error Resume Next
  Err.Clear
  Set app = CreateObject("CorelDRAW.Application.14")
  W "  boot   err=" & Err.Number & " " & Err.Description
  Err.Clear
  app.Visible = True
  app.InitializeVBA
  W "  initvba err=" & Err.Number & " " & Err.Description
  Err.Clear
  Set vbe = Nothing
  For v = 1 To 20
    Set vbe = Nothing
    Err.Clear
    Set vbe = app.VBE
    If Err.Number = 0 And Not vbe Is Nothing Then Exit For
    Err.Clear
    WScript.Sleep 500
  Next
  On Error GoTo 0
  If vbe Is Nothing Then Die("  FAIL: no VBE")

  Set p = Nothing
  For v = 1 To 60
    Set p = ProjectForTgt(vbe, tgt)
    If Not p Is Nothing Then Exit For
    WScript.Sleep 500
  Next
  If p Is Nothing Then Die("  FAIL: CDRX4Toolkit did not load")
  W "  project loaded comps=" & p.VBComponents.Count
End Sub

Sub EndSession()
  Dim nw
  On Error Resume Next
  app.Quit
  On Error GoTo 0
  For nw = 1 To 60
    WScript.Sleep 500
    Set app = Nothing
    On Error Resume Next
    Set app = GetObject(, "CorelDRAW.Application.14")
    On Error GoTo 0
    If app Is Nothing Then Exit For
  Next
  W "  session ended"
End Sub

Sub RunMacro(m)
  Dim nw
  If fso.FileExists(stateLog) Then fso.DeleteFile stateLog, True
  On Error Resume Next
  Err.Clear
  app.GMSManager.RunMacro "CDRX4Toolkit", m
  W "  RunMacro " & m & " err=" & Err.Number & " " & Err.Description
  Err.Clear
  On Error GoTo 0
  For nw = 1 To 30
    If fso.FileExists(stateLog) Then Exit For
    WScript.Sleep 500
  Next
  If Not fso.FileExists(stateLog) Then
    W "  !! no state log produced by " & m
    Exit Sub
  End If
  Set st = fso.OpenTextFile(stateLog, 1, False, -1)
  txt = st.ReadAll
  st.Close
  W "  ---- " & m & " ----"
  W txt
  W "  ---- end ----"
End Sub

tgt = sh.ExpandEnvironmentStrings("%APPDATA%") & _
      "\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms"
stateLog = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_toolbar_state.log"

W "==== CDRX4Toolkit toolbar persistence test " & Now & " ===="
If Not fso.FileExists(tgt) Then Die("FAIL: no GMS at " & tgt)
W "gms = " & tgt

W ""
W "--- phase A: fresh session, report, install, report, quit ---"
StartSession
RunMacro "M_Install.ReportToolbar"
RunMacro "M_Install.InstallToolbarSilent"
RunMacro "M_Install.ReportToolbar"
EndSession

W ""
W "--- phase B: restart, report only (does it survive?) ---"
StartSession
RunMacro "M_Install.ReportToolbar"
EndSession

W ""
W "==== done ===="
Flush
WScript.Quit 0