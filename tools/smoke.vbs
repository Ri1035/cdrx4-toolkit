Option Explicit

'==========================================================
'  CDRX4Toolkit smoke test
'
'  Runs against the GMS that is sitting on disk, in a FRESH
'  CorelDRAW session -- a running instance already holds the
'  project in memory, and testing that would not prove the
'  saved file loads:
'    0. restart CorelDRAW so it re-reads User Draw\GMS
'    1. is the project loaded, with all 13 components and a
'       component named exactly M_Install?
'    2. does M_Install.DiagToolbar build 9 Chinese buttons?
'    3. do the nine features actually run (M_Test.SelfTest)?
'
'  Result -> _smoke.log (UTF-16) and stdout.
'
'  ASCII ONLY: WSH parses .vbs as ANSI, so Chinese is built with
'  ChrW() instead of being written literally.
'==========================================================

Dim fso, sh, root, logPath, gLog, st
Dim app, vbe, p, c, i, found, fn, cnt, n, hasInstall
Dim tgt, selLog, toolLog, fails, lines, selTxt, exp, obs, got
Dim waited, docCount, dirty

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
logPath = root & "\_smoke.log"
gLog = ""
fails = 0

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

W "==== CDRX4Toolkit smoke " & Now & " ===="

tgt = sh.ExpandEnvironmentStrings("%APPDATA%") & _
      "\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms"
If Not fso.FileExists(tgt) Then Die("FAIL: GMS not found at " & tgt)
W "gms    = " & tgt
W "size   = " & fso.GetFile(tgt).Size

' --- 0. fresh session -----------------------------------------
' X4 auto-loads User Draw\GMS at startup, so a running instance already
' has the project in memory.  Only a restart proves that the file on disk
' loads, compiles and builds the toolbar.
Set app = Nothing
On Error Resume Next
Set app = GetObject(, "CorelDRAW.Application.14")
On Error GoTo 0

If Not app Is Nothing Then
  docCount = -1
  dirty = -1
  On Error Resume Next
  Err.Clear
  docCount = app.Documents.Count
  dirty = 0
  For i = 1 To docCount
    If app.Documents.Item(i).Dirty Then dirty = dirty + 1
  Next
  Err.Clear
  On Error GoTo 0
  If dirty > 0 Then
    Die "FAIL: CorelDRAW has " & dirty & " document(s) with unsaved changes - save or close them and re-run"
  End If
  W "restarting CorelDRAW (" & docCount & " open doc(s), none dirty)"
  On Error Resume Next
  app.Quit
  On Error GoTo 0
  For waited = 1 To 60
    WScript.Sleep 500
    Set app = Nothing
    On Error Resume Next
    Set app = GetObject(, "CorelDRAW.Application.14")
    On Error GoTo 0
    If app Is Nothing Then Exit For
  Next
  If Not app Is Nothing Then Die("FAIL: CorelDRAW would not exit - close it and re-run")
End If
Set app = Nothing

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
W "boot   err=" & Err.Number & " " & Err.Description
Err.Clear
app.Visible = True
app.InitializeVBA
W "initvba err=" & Err.Number & " " & Err.Description
Err.Clear
Set vbe = Nothing
For waited = 1 To 20
  Set vbe = Nothing
  Err.Clear
  Set vbe = app.VBE
  If Err.Number = 0 And Not vbe Is Nothing Then Exit For
  Err.Clear
  WScript.Sleep 500
Next
On Error GoTo 0

If vbe Is Nothing Then Die("FAIL: no VBE")

' --- 1. project loaded from disk? -----------------------------
W ""
W "--- 1. project ---"
Set p = Nothing
For waited = 1 To 60
  Set p = ProjectForTgt(vbe, tgt)
  If Not p Is Nothing Then Exit For
  WScript.Sleep 500
Next

If p Is Nothing Then
  W "FAIL: CDRX4Toolkit.gms did not load in this CorelDRAW session"
  fails = fails + 1
Else
  W "project = [" & p.Name & "]  comps=" & p.VBComponents.Count & _
    "  prot=" & p.Protection
  cnt = 0
  hasInstall = 0
  For i = 1 To p.VBComponents.Count
    On Error Resume Next
    Err.Clear
    Set c = Nothing
    Set c = p.VBComponents.Item(i)
    If Err.Number = 0 And Not c Is Nothing Then
      W "   comp " & i & " = " & c.Name
      cnt = cnt + 1
      ' the module names are load-bearing: M_Install.<Proc> is what the
      ' startup handler calls, so an import renamed to "M_Install1"
      ' means no toolbar, ever
      If c.Name = "M_Install" Then hasInstall = 1
    End If
    Err.Clear
    On Error GoTo 0
  Next
  If cnt <> 13 Then
    W "FAIL: expected 13 components, saw " & cnt
    fails = fails + 1
  End If
  If hasInstall = 0 Then
    W "FAIL: no component named exactly M_Install - every M_Install.<Proc> macro path is dead"
    fails = fails + 1
  End If
End If

' --- 2. toolbar, verified from inside VBA ---------------------
' X4 does not hand CommandBars to an external automation client
' (app.CommandBars(name) comes back as Nothing with err 13), so the
' toolbar has to be inspected by the macro that built it.
W ""
W "--- 2. toolbar ---"
toolLog = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_toolbar.log"
If fso.FileExists(toolLog) Then fso.DeleteFile toolLog, True

On Error Resume Next
Err.Clear
app.GMSManager.RunMacro "CDRX4Toolkit", "M_Install.DiagToolbar"
W "DiagToolbar err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

For i = 1 To 30
  If fso.FileExists(toolLog) Then Exit For
  WScript.Sleep 500
Next

If Not fso.FileExists(toolLog) Then
  W "FAIL: no toolbar log - M_Install.DiagToolbar never ran (compile error?)"
  fails = fails + 1
Else
  Set st = fso.OpenTextFile(toolLog, 1, False, -1)
  W "---- toolbar log ----"
  W st.ReadAll
  st.Close
  W "---- end ----"

  ' Button order and Chinese captions come from M_Install.CmdList().
  ' DiagToolbar writes both the expected list and the buttons it actually
  ' built into the same log, so this reads one file in one encoding --
  ' no more parsing src\M_Install.bas across encodings.
  Set exp = CreateObject("Scripting.Dictionary")
  Set obs = CreateObject("Scripting.Dictionary")
  FillCaptions toolLog, "exp ", exp
  FillCaptions toolLog, "btn ", obs
  W "expected " & exp.Count & " / built " & obs.Count

  If exp.Count = 0 Then
    W "FAIL: the toolbar log carries no expected captions"
    fails = fails + 1
  ElseIf obs.Count = 0 Then
    W "FAIL: the toolbar log carries no built buttons"
    fails = fails + 1
  Else
    W "observed = " & obs.Count & " button(s), expected = " & exp.Count

    If obs.Count <> exp.Count Then
      W "FAIL: expected " & exp.Count & " buttons, saw " & obs.Count
      fails = fails + 1
    End If

    n = exp.Count - 1
    If obs.Count - 1 < n Then n = obs.Count - 1
    For i = 0 To n
      If obs.Item(i) <> exp.Item(i) Then
        W "FAIL: btn " & (i + 1) & " caption should be [" & exp.Item(i) & "] but is [" & obs.Item(i) & "]"
        fails = fails + 1
      End If
    Next
  End If
End If

' --- 3. the nine features -------------------------------------
W ""
W "--- 3. self test ---"
selLog = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_selftest.log"
If fso.FileExists(selLog) Then fso.DeleteFile selLog, True

On Error Resume Next
Err.Clear
app.GMSManager.RunMacro "CDRX4Toolkit", "M_Test.SelfTest"
W "SelfTest err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

For i = 1 To 30
  If fso.FileExists(selLog) Then Exit For
  WScript.Sleep 500
Next

If Not fso.FileExists(selLog) Then
  W "FAIL: no selftest log - the macro never ran (compile error?)"
  fails = fails + 1
Else
  Set st = fso.OpenTextFile(selLog, 1, False, -1)
  selTxt = st.ReadAll
  st.Close
  W "---- selftest log ----"
  W selTxt
  W "---- end ----"

  ' every feature line must read err=0, and n must be > 0 (0 means it did nothing)
  lines = Split(selTxt, vbLf)
  For i = 0 To UBound(lines)
    If InStr(lines(i), "err=") > 0 Then
      If InStr(lines(i), "err=0") = 0 Then
        W "FAIL: selftest reported an error: " & Trim(lines(i))
        fails = fails + 1
      End If
      If InStr(lines(i), " n=0 ") > 0 Then
        W "FAIL: selftest touched nothing: " & Trim(lines(i))
        fails = fails + 1
      End If
    End If
  Next
End If

W ""
If fails = 0 Then
  W "==== SMOKE PASS ===="
Else
  W "==== SMOKE FAIL (" & fails & " problem(s)) ===="
End If

Flush
WScript.Quit fails


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


' Pull a caption list out of M_Install.DiagToolbar's log, which is UTF-16.
' Both sides of the comparison live in that one file:
'    exp 1 caption=[...] proc=[M_Curves.ConvertAllToCurves]
'    btn 1 caption=[...] tooltip=[...] err=0
' The expected side is the macro's own CmdList(), so this check cannot drift
' from the real button table, and there is no cross-encoding parsing left.
' Fills a caller-supplied Dictionary, keyed 0..n-1 in file order.
Sub FillCaptions(p, prefix, col)
  Dim st2, txt, lines2, i, a, b, s

  On Error Resume Next
  Set st2 = fso.OpenTextFile(p, 1, False, -1)
  txt = st2.ReadAll
  st2.Close
  On Error GoTo 0

  If Len(txt) = 0 Then Exit Sub

  lines2 = Split(txt, vbLf)
  For i = 0 To UBound(lines2)
    s = lines2(i)
    If Left(s, Len(prefix)) = prefix Then
      a = InStr(s, "caption=[")
      If a > 0 Then
        a = a + 9
        b = InStr(a, s, "]")
        If b > a Then col.Add col.Count, Mid(s, a, b - a)
      End If
    End If
  Next
End Sub