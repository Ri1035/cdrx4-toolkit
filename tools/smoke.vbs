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
'    2. does M_Install compile and hand out its 9-button list?
'    3. do the nine features actually run (M_Test.SelfTest)?
'
'  The toolbar itself is NOT built here any more.  It is pure workspace
'  XML written by the installer, so it is checked by
'  tools\verify_toolbar.ps1 reading DRAWUIConfig.xml (PLAN.md section 13).
'
'  Result -> _smoke.log (UTF-16) and stdout.
'
'  ASCII ONLY: WSH parses .vbs as ANSI, so Chinese is built with
'  ChrW() instead of being written literally.
'==========================================================

Dim fso, sh, root, logPath, gLog, st
Dim app, vbe, p, c, i, found, fn, cnt, n, hasInstall
Dim tgt, selLog, fails, lines, selTxt, itemsPath, itemsTxt, nItems, bad, a, proc
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

' --- 2. M_Install compiles, and its button list is intact -----
' M_Install owns the canonical button list (CmdList) and is the module
' the old build reached from a startup hook, so a compile error in it
' would be invisible to the feature test below -- VBA compiles on
' demand.  DumpItems is called explicitly for that reason.  Its output
' is also what tools\verify_toolbar.ps1 compares the workspace markup
' against, so the two halves of the toolbar cannot drift.
'
' The toolbar is NOT built here: it is pure workspace XML written by
' the installer while CorelDRAW is closed (PLAN.md section 13).
W ""
W "--- 2. M_Install ---"
itemsPath = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_toolbar_items.txt"
If fso.FileExists(itemsPath) Then fso.DeleteFile itemsPath, True

On Error Resume Next
Err.Clear
app.GMSManager.RunMacro "CDRX4Toolkit", "M_Install.DumpItems"
W "DumpItems err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

For i = 1 To 30
  If fso.FileExists(itemsPath) Then Exit For
  WScript.Sleep 500
Next

If Not fso.FileExists(itemsPath) Then
  W "FAIL: no items file - M_Install.DumpItems never ran (compile error?)"
  fails = fails + 1
Else
  ' the file is UTF-8 without a BOM.  Decode it as UTF-8 explicitly:
  ' reading it as ANSI on a CJK codepage mis-pairs the trailing byte of a
  ' multi-byte character with the following LF, which merges lines and
  ' undercounts the buttons (seen: 9 -> 6).
  itemsTxt = ReadUtf8(itemsPath)
  W "---- items ----"
  W itemsTxt
  W "---- end ----"

  nItems = 0
  bad = 0
  lines = Split(itemsTxt, vbLf)
  For i = 0 To UBound(lines)
    If Len(Trim(lines(i))) > 0 Then
      nItems = nItems + 1
      a = InStr(lines(i), "|")
      If a = 0 Then
        bad = bad + 1
      Else
        proc = Left(lines(i), a - 1)
        If InStr(proc, ".") = 0 Then bad = bad + 1
      End If
    End If
  Next
  W "items = " & nItems
  If nItems <> 9 Then
    W "FAIL: expected 9 buttons in CmdList, saw " & nItems
    fails = fails + 1
  End If
  If bad > 0 Then
    W "FAIL: " & bad & " malformed item line(s) in the dump"
    fails = fails + 1
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


' read a UTF-8 (no BOM) file.  WSH's OpenTextFile cannot decode UTF-8, and
' the system default (ANSI) corrupts CJK: the trailing byte of a multi-byte
' char pairs with the next LF.  ADODB.Stream handles the encoding properly.
Function ReadUtf8(path)
  Dim stm, s
  s = ""
  On Error Resume Next
  Err.Clear
  Set stm = CreateObject("ADODB.Stream")
  If Err.Number = 0 And Not stm Is Nothing Then
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    s = stm.ReadText
    stm.Close
  End If
  Err.Clear
  On Error GoTo 0
  ReadUtf8 = s
End Function