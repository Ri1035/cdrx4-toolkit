Option Explicit

'==========================================================
'  Build CDRX4Toolkit.gms from the VBA sources in src\
'
'  Requires CorelDRAW X4 (with VBA) installed on this machine.
'
'  Steps:
'    0. get CorelDRAW out of the way (see below)
'    1. copy an unprotected GMS shipped with X4 as a seed
'    2. drop its components and import every module in src\
'    3. write the startup handler into ThisDocument
'    4. save as CDRX4Toolkit.gms in the X4 user GMS folder
'    5. run M_Test.Probe / Diag / SelfTest and M_Install.DiagToolbar
'
'  WHY CORELDRAW IS RESTARTED FIRST
'  X4 auto-loads every GMS under User Draw\GMS at startup, so a running
'  instance already holds CDRX4Toolkit.gms in memory.  Building into that
'  project means inheriting its component names, and the old M_Install --
'  which sits on a compile error -- refuses to be removed.  VBE then
'  quietly renames the fresh import to "M_Install1", every
'  M_Install.<Proc> macro path stops resolving, and the startup handler
'  dies with -2147220224 ("module not found").  So the file is swapped
'  while no instance is running, and X4 is then started fresh.
'
'  src\*.bas are stored as UTF-8 so they read well on GitHub,
'  but VBE imports ANSI text, so every file is re-encoded to a
'  GBK temp copy before import.
'
'  This script is ASCII-only on purpose: Windows Script Host
'  reads .vbs as ANSI, so a UTF-8 file containing Chinese
'  literals would be mis-parsed (and can even break the quotes).
'==========================================================

Dim fso, sh, here, srcDir, tmpDir, seed, tgt, bak
Dim mods, k, i, c, p, app, vbe, docMod, code, fm, saveCtl, cp
Dim logPath, res, probePath, diagPath, toolbarPath
Dim activeName, tries, activated, sizeBefore, sizeAfter, want, hits, saved
Dim attempt, docCount, dirty, waited, fails, bad, cap, m

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

here   = fso.GetParentFolderName(WScript.ScriptFullName)
srcDir = here & "\src"
tmpDir = here & "\_build_tmp"
fails  = 0

mods = Array("M_Util.bas", "M_Curves.bas", "M_Rect.bas", "M_Color.bas", "M_CMYK.bas", _
             "M_FitPath.bas", "M_JPG.bas", "M_PageNo.bas", "M_Calendar.bas", "M_Seal.bas", _
             "M_Install.bas", "M_Test.bas")

seed = FindSeed()
If seed = "" Then
  WScript.Echo "!! no CorelDRAW X4 GMS folder found - install and run X4 once first"
  WScript.Quit 1
End If

tgt = TargetGms()
bak = tgt & ".bak"
WScript.Echo "[0] seed   = " & seed
WScript.Echo "[0] target = " & tgt

If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir, True
fso.CreateFolder tmpDir
For k = 0 To UBound(mods)
  MakeAnsiCopy srcDir & "\" & mods(k), tmpDir & "\" & mods(k)
Next
WScript.Echo "[0] re-encoded " & (UBound(mods) + 1) & " sources to GBK"

' --- 0a. no CorelDRAW while the GMS folder is swapped ---------
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
    WScript.Echo "!! CorelDRAW has " & dirty & " document(s) with unsaved changes."
    WScript.Echo "   Save or close them, then run this build again."
    WScript.Quit 1
  End If

  WScript.Echo "[0a] quitting CorelDRAW (" & docCount & " open doc(s), none dirty)"
  QuitApp
  If Not app Is Nothing Then
    WScript.Echo "!! CorelDRAW is still running - close it by hand and re-run"
    WScript.Quit 1
  End If
  WScript.Echo "[0a] CorelDRAW is gone"
End If
Set app = Nothing

' --- 1. swap in a fresh seed copy -----------------------------
If fso.FileExists(bak) Then fso.DeleteFile bak, True
If fso.FileExists(tgt) Then fso.CopyFile tgt, bak, True
If fso.FileExists(tgt) Then fso.DeleteFile tgt, True
fso.CopyFile seed, tgt, True
WScript.Echo "[1] seed copied (previous build kept as .bak)"

Set app = CreateObject("CorelDRAW.Application.14")
On Error Resume Next
app.Visible = True
app.InitializeVBA
On Error GoTo 0

Set vbe = Nothing
For waited = 1 To 20
  Set vbe = Nothing
  On Error Resume Next
  Err.Clear
  Set vbe = app.VBE
  Err.Clear
  On Error GoTo 0
  If Not vbe Is Nothing Then Exit For
  WScript.Sleep 500
Next
If vbe Is Nothing Then
  WScript.Echo "!! CorelDRAW did not hand out a VBE"
  WScript.Quit 1
End If

' X4 loads User Draw\GMS\*.gms while it starts, which is asynchronous
' relative to this script, so wait for the seed project to show up.
Set p = Nothing
For waited = 1 To 60
  Set p = ProjectForTgt(vbe, tgt)
  If Not p Is Nothing Then Exit For
  WScript.Sleep 500
Next

If p Is Nothing Then
  ' fall back to an explicit load, in case this X4 build does not
  ' auto-load the user GMS folder on its own
  On Error Resume Next
  Err.Clear
  app.GMSManager.LoadGMS tgt
  WScript.Echo "[1a] LoadGMS err=" & Err.Number & " " & Err.Description
  Err.Clear
  On Error GoTo 0
  For waited = 1 To 40
    Set p = ProjectForTgt(vbe, tgt)
    If Not p Is Nothing Then Exit For
    WScript.Sleep 500
  Next
End If

DumpProjects vbe

If p Is Nothing Then
  WScript.Echo "!! " & tgt & " did not load into this CorelDRAW session"
  WScript.Quit 1
End If

p.Name = "CDRX4Toolkit"
WScript.Echo "[2] project = " & p.Name & " prot=" & p.Protection & _
             " comps=" & p.VBComponents.Count

' --- 2. drop every non-document component --------------------
' A module sitting on a compile error can refuse the first Remove, so
' retry, then insist: one stray component is enough to rename the fresh
' import and break every macro path.
For attempt = 1 To 3
  For i = p.VBComponents.Count To 1 Step -1
    Set c = Nothing
    On Error Resume Next
    Err.Clear
    Set c = p.VBComponents.Item(i)
    If Err.Number = 0 And Not c Is Nothing Then
      If c.Type <> 100 Then
        Err.Clear
        If c.CodeModule.CountOfLines > 0 Then c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
        Err.Clear
        p.VBComponents.Remove c
        If Err.Number <> 0 Then WScript.Echo "    remove " & c.Name & " err=" & Err.Number
      End If
    End If
    Err.Clear
    On Error GoTo 0
  Next
  If p.VBComponents.Count <= 1 Then Exit For
Next
WScript.Echo "[3] comps after cleanup = " & p.VBComponents.Count & " (after attempt " & attempt & ")"

If p.VBComponents.Count > 1 Then
  For i = 1 To p.VBComponents.Count
    Set c = Nothing
    On Error Resume Next
    Err.Clear
    Set c = p.VBComponents.Item(i)
    If Err.Number = 0 And Not c Is Nothing Then
      WScript.Echo "    leftover " & i & " name=" & c.Name & " type=" & c.Type
    End If
    Err.Clear
    On Error GoTo 0
  Next
  Fail "the seed project still carries extra components"
End If

' --- 3. import the real modules ------------------------------
For k = 0 To UBound(mods)
  On Error Resume Next
  Err.Clear
  p.VBComponents.Import tmpDir & "\" & mods(k)
  If Err.Number <> 0 Then
    WScript.Echo "    import " & mods(k) & " ERR " & Err.Number & " " & Err.Description
    fails = fails + 1
  End If
  Err.Clear
  On Error GoTo 0
Next
WScript.Echo "[4] comps now = " & p.VBComponents.Count & _
             " (want " & (UBound(mods) + 2) & " = ThisDocument + " & (UBound(mods) + 1) & " modules)"
If fails > 0 Then Fail "one or more modules failed to import"

' Every module must have landed under its OWN name.  VBE renames a
' conflicting import to <name>1 without complaining, and that breaks the
' macro paths -- so this is a hard stop, not a warning.
bad = 0
For k = 0 To UBound(mods)
  want = fso.GetBaseName(mods(k))
  hits = CountComp(p, want)
  If hits <> 1 Then
    WScript.Echo "!! module " & want & " is present " & hits & " time(s)"
    bad = bad + 1
  End If
Next
If bad > 0 Then Fail "module names collided - macro paths would break"

' --- 4. startup handler: this is what auto-installs the toolbar ---
On Error Resume Next
Err.Clear
Set docMod = p.VBComponents.Item(1).CodeModule
code = "Private Sub GlobalMacroStorage_Start()" & vbCrLf & _
       "    On Error Resume Next" & vbCrLf & _
       "    M_Install.InstallToolbarSilent" & vbCrLf & _
       "End Sub" & vbCrLf & _
       "Private Sub GlobalMacroStorage_OnApplicationStart()" & vbCrLf & _
       "    On Error Resume Next" & vbCrLf & _
       "    M_Install.InstallToolbarSilent" & vbCrLf & _
       "End Sub"
If docMod.CountOfLines > 0 Then docMod.DeleteLines 1, docMod.CountOfLines
docMod.AddFromString code
WScript.Echo "[5] startup handler written, err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

' --- 5. save the project back to the target file ---------------
'
' VBProject.SaveAs is dead on X4: err 748 for every destination, including
' the project's own file.  VBProject has no Save method either.  The only
' thing that actually writes a .gms is the VBE File > Save menu item, and
' that saves whatever project the VBE has ACTIVE.
'
' When the active project is a stock CorelDRAW .gms, executing that item
' would rewrite a file under Program Files.  So the item's caption -- which
' names the file that would be written -- is used as the gate: the trigger
' is only pulled once it names our own file, and the item is never executed
' otherwise.
sizeBefore = 0
If fso.FileExists(tgt) Then sizeBefore = fso.GetFile(tgt).Size

saved = 0
On Error Resume Next
Err.Clear
p.SaveAs tgt
If Err.Number = 0 Then saved = 1
WScript.Echo "[7] SaveAs err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

If saved = 0 Then
  On Error Resume Next
  Err.Clear
  activeName = vbe.ActiveVBProject.Name
  WScript.Echo "[7] active project after import = [" & activeName & "]"
  Err.Clear
  vbe.MainWindow.Visible = True
  Err.Clear

  Set saveCtl = Nothing
  Set fm = vbe.CommandBars.Item(1).Controls.Item(1)
  For i = 1 To fm.Controls.Count
    Err.Clear
    Set c = fm.Controls.Item(i)
    If Err.Number = 0 And Not c Is Nothing Then
      If c.ID = 3 Then Set saveCtl = c
    End If
    Err.Clear
  Next
  Err.Clear

  cap = ""
  If saveCtl Is Nothing Then
    WScript.Echo "!! the VBE File menu has no Save item"
  Else
    ' Four different ways of stealing the VBE's active project back.
    ' Showing a code pane alone is not always enough, so they alternate.
    For tries = 1 To 10
      For m = 1 To 4
        Err.Clear
        Select Case m
          Case 1
            p.VBComponents.Item(1).Activate
          Case 2
            Set cp = p.VBComponents.Item(1).CodeModule.CodePane
            cp.Show
            cp.Window.SetFocus
          Case 3
            vbe.MainWindow.SetFocus
            Set cp = p.VBComponents.Item(1).CodeModule.CodePane
            cp.Show
            cp.Window.SetFocus
          Case 4
            vbe.MainWindow.Visible = False
            WScript.Sleep 400
            vbe.MainWindow.Visible = True
            Set cp = p.VBComponents.Item(1).CodeModule.CodePane
            cp.Show
            cp.Window.SetFocus
        End Select
        Err.Clear
        WScript.Sleep 900

        cap = ""
        Err.Clear
        cap = saveCtl.Caption
        Err.Clear
        WScript.Echo "[7a] try " & tries & "." & m & " save caption=[" & cap & "]"
        If InStr(1, cap, "CDRX4Toolkit.gms", vbTextCompare) > 0 Then Exit For
      Next
      If InStr(1, cap, "CDRX4Toolkit.gms", vbTextCompare) > 0 Then Exit For
    Next

    If InStr(1, cap, "CDRX4Toolkit.gms", vbTextCompare) = 0 Then
      WScript.Echo "!! VBE would save [" & cap & "] - refusing to touch another project"
    Else
      saveCtl.Execute
      WScript.Echo "[7b] save err=" & Err.Number & " " & Err.Description
      Err.Clear
      WScript.Sleep 4000
    End If
  End If
  Err.Clear
  On Error GoTo 0
End If

sizeAfter = 0
If fso.FileExists(tgt) Then sizeAfter = fso.GetFile(tgt).Size
WScript.Echo "[8] size before=" & sizeBefore & " after=" & sizeAfter
If sizeAfter <= sizeBefore Then Fail "target did not grow - the save did not land"

If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir, True

' --- 6. run the built-in checks in this very session ---------
' A VBA project with any compile error makes every macro fail, so this
' doubles as the compile check: if the logs come out with err=0 all over,
' the project compiles and the nine features actually run.
'
' M_Install has to be reached explicitly.  VBA compiles on demand and
' SelfTest never touches M_Install, so a compile error in it stays
' invisible until the startup handler fires it -- which is exactly how
' the "Set app = CorelDRAW" bug got out the door and wedged CorelDRAW
' with a modal dialog on every launch.
logPath     = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_selftest.log"
probePath   = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_probe.log"
diagPath    = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_diag.log"
toolbarPath = sh.ExpandEnvironmentStrings("%TEMP%") & "\cdrx4_toolbar.log"
If fso.FileExists(logPath) Then fso.DeleteFile logPath, True
If fso.FileExists(probePath) Then fso.DeleteFile probePath, True
If fso.FileExists(diagPath) Then fso.DeleteFile diagPath, True
If fso.FileExists(toolbarPath) Then fso.DeleteFile toolbarPath, True

WScript.Echo "[9] running probe ..."
On Error Resume Next
Err.Clear
res = app.GMSManager.RunMacro("CDRX4Toolkit", "M_Test.Probe")
WScript.Echo "[9] Probe err=" & Err.Number & " " & Err.Description
Err.Clear
res = app.GMSManager.RunMacro("CDRX4Toolkit", "M_Test.Diag")
WScript.Echo "[9] Diag err=" & Err.Number & " " & Err.Description
Err.Clear
res = app.GMSManager.RunMacro("CDRX4Toolkit", "M_Test.SelfTest")
WScript.Echo "[9] SelfTest err=" & Err.Number & " " & Err.Description
Err.Clear
res = app.GMSManager.RunMacro("CDRX4Toolkit", "M_Install.DiagToolbar")
WScript.Echo "[9] DiagToolbar err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

For i = 1 To 30
  If fso.FileExists(probePath) And fso.FileExists(logPath) And _
     fso.FileExists(diagPath) And fso.FileExists(toolbarPath) Then Exit For
  WScript.Sleep 500
Next

DumpLog fso, diagPath, "diag"
DumpLog fso, logPath, "selftest"
DumpLog fso, toolbarPath, "toolbar"

If fso.FileExists(bak) Then fso.DeleteFile bak, True

WScript.Echo ""
If fails = 0 Then
  WScript.Echo "==== BUILD OK ===="
Else
  WScript.Echo "==== BUILD FAIL (" & fails & " problem(s)) ===="
End If
WScript.Quit fails


' --- helpers -------------------------------------------------

' Hand the previous build back and get out: a GMS that does not compile
' would be auto-loaded on the next X4 start and wedge it with a modal
' dialog, so a failed build must not leave its output behind.
Sub Fail(ByVal msg)
  WScript.Echo "!! " & msg
  QuitApp
  On Error Resume Next
  If fso.FileExists(tgt) Then fso.DeleteFile tgt, True
  If fso.FileExists(bak) Then fso.MoveFile bak, tgt
  If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir, True
  On Error GoTo 0
  WScript.Echo "   restored the previous " & fso.GetFileName(tgt) & " (if there was one)"
  WScript.Quit 1
End Sub


Sub QuitApp()
  Dim w
  On Error Resume Next
  app.Quit
  On Error GoTo 0
  For w = 1 To 60
    WScript.Sleep 500
    Set app = Nothing
    On Error Resume Next
    Set app = GetObject(, "CorelDRAW.Application.14")
    On Error GoTo 0
    If app Is Nothing Then Exit For
  Next
End Sub


Function ProjectForTgt(ByVal vbe, ByVal tgt)
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


Sub DumpProjects(ByVal vbe)
  Dim i, pr
  On Error Resume Next
  For i = 1 To vbe.VBProjects.Count
    Err.Clear
    Set pr = Nothing
    Set pr = vbe.VBProjects.Item(i)
    If Err.Number = 0 And Not pr Is Nothing Then
      WScript.Echo "    proj " & i & " name=" & pr.Name & " file=" & pr.FileName
    End If
    Err.Clear
  Next
  On Error GoTo 0
End Sub


Function CountComp(ByVal p, ByVal nm)
  Dim i, c2, n
  n = 0
  On Error Resume Next
  For i = 1 To p.VBComponents.Count
    Set c2 = Nothing
    Err.Clear
    Set c2 = p.VBComponents.Item(i)
    If Err.Number = 0 And Not c2 Is Nothing Then
      If StrComp(c2.Name, nm, vbTextCompare) = 0 Then n = n + 1
    End If
    Err.Clear
  Next
  On Error GoTo 0
  CountComp = n
End Function


Sub DumpLog(ByVal f, ByVal p, ByVal title)
  Dim st
  If Not f.FileExists(p) Then
    WScript.Echo "!! " & title & " log not produced - macro never ran (compile error?)"
    Exit Sub
  End If
  Set st = f.OpenTextFile(p, 1, False, -1)
  WScript.Echo "---- " & title & " ----"
  WScript.Echo st.ReadAll
  st.Close
  WScript.Echo "---- end " & title & " ----"
End Sub


Function FindSeed()
  Dim cands, i, f
  cands = Array( _
    sh.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\CorelDRAW X4\Draw\GMS", _
    sh.ExpandEnvironmentStrings("%ProgramFiles%") & "\CorelDRAW X4\Draw\GMS", _
    sh.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Corel\CorelDRAW Graphics Suite X4\Draw\GMS", _
    sh.ExpandEnvironmentStrings("%ProgramFiles%") & "\Corel\CorelDRAW Graphics Suite X4\Draw\GMS")

  ' prefer a known unprotected sample
  For i = 0 To UBound(cands)
    If fso.FileExists(cands(i) & "\Emboss.gms") Then
      FindSeed = cands(i) & "\Emboss.gms"
      Exit Function
    End If
  Next
  ' otherwise take any GMS that ships with X4
  For i = 0 To UBound(cands)
    If fso.FolderExists(cands(i)) Then
      For Each f In fso.GetFolder(cands(i)).Files
        If LCase(fso.GetExtensionName(f.Name)) = "gms" Then
          FindSeed = f.Path
          Exit Function
        End If
      Next
    End If
  Next
  FindSeed = ""
End Function


Function TargetGms()
  Dim d
  d = sh.ExpandEnvironmentStrings("%APPDATA%") & _
      "\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS"
  If Not fso.FolderExists(d) Then MkPath d
  TargetGms = d & "\CDRX4Toolkit.gms"
End Function


Sub MkPath(ByVal d)
  Dim parent
  On Error Resume Next
  If fso.FolderExists(d) Then Exit Sub
  parent = fso.GetParentFolderName(d)
  If parent <> "" And Not fso.FolderExists(parent) Then MkPath parent
  fso.CreateFolder d
  On Error GoTo 0
End Sub


' VBE imports ANSI text, so UTF-8 sources are re-encoded to GBK.
' A file that is already ANSI/GBK is copied untouched.
Sub MakeAnsiCopy(ByVal src, ByVal dst)
  Dim st, txt
  On Error Resume Next
  Err.Clear
  Set st = CreateObject("ADODB.Stream")
  st.Type = 2
  st.Charset = "utf-8"
  st.Open
  st.LoadFromFile src
  txt = st.ReadText
  st.Close
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    fso.CopyFile src, dst, True
    Exit Sub
  End If
  Err.Clear
  Set st = CreateObject("ADODB.Stream")
  st.Type = 2
  st.Charset = "gb2312"
  st.Open
  st.WriteText txt
  st.SaveToFile dst, 2
  st.Close
  Err.Clear
  On Error GoTo 0
End Sub