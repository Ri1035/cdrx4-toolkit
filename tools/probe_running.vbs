Option Explicit

'==========================================================
'  Attach to the CorelDRAW that is ALREADY running and run
'  read-only macros in it.
'
'  usage: cscript //nologo tools\probe_running.vbs <Macro> [Macro ...]
'
'  Why not tools\run_macros.vbs: that one quits whatever is
'  running and boots a session of its own.  That is the right
'  thing when the point is to exercise a freshly saved GMS, but
'  it is the wrong thing when the point is to inspect the
'  instance a USER's double-click produced -- a session we
'  create ourselves takes a different boot path (PLAN section
'  10.4), so its toolbar proves nothing about the persisted
'  workspace.
'
'  After each macro the contents of
'  %TEMP%\cdrx4_toolbar_state.log are printed.
'
'  ASCII ONLY: WSH parses .vbs as ANSI.
'==========================================================

Dim fso, sh, app, i, stateLog, st, txt, waited

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

stateLog = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_toolbar_state.log"

If WScript.Arguments.Count < 1 Then
  WScript.Echo "usage: cscript //nologo tools\probe_running.vbs <Macro> [Macro ...]"
  WScript.Quit 2
End If

Set app = Nothing
On Error Resume Next
Err.Clear
Set app = GetObject(, "CorelDRAW.Application.14")
On Error GoTo 0
If app Is Nothing Then
  WScript.Echo "FAIL: no running CorelDRAW to attach to"
  WScript.Quit 1
End If
WScript.Echo "attached to the running CorelDRAW"

For i = 0 To WScript.Arguments.Count - 1
  WScript.Echo ""
  WScript.Echo "--- " & WScript.Arguments(i) & " ---"
  If fso.FileExists(stateLog) Then fso.DeleteFile stateLog, True
  On Error Resume Next
  Err.Clear
  app.GMSManager.RunMacro "CDRX4Toolkit", WScript.Arguments(i)
  WScript.Echo "  err=" & Err.Number & " " & Err.Description
  Err.Clear
  On Error GoTo 0

  For waited = 1 To 30
    If fso.FileExists(stateLog) Then Exit For
    WScript.Sleep 500
  Next

  If fso.FileExists(stateLog) Then
    Set st = fso.OpenTextFile(stateLog, 1, False, -1)
    txt = st.ReadAll
    st.Close
    WScript.Echo txt
  Else
    WScript.Echo "  (no state log - the macro never wrote one)"
  End If
Next

WScript.Echo ""
WScript.Echo "==== done (CorelDRAW left running) ===="