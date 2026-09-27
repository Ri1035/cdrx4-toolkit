Option Explicit

'==========================================================
'  Build CDRX4Toolkit.gms from the VBA sources in src\
'
'  Requires CorelDRAW X4 (with VBA) installed on this machine.
'
'  Steps:
'    1. copy an unprotected GMS shipped with X4 as a seed
'    2. drop its components and import every module in src\
'    3. write the startup handler into ThisDocument
'    4. save as CDRX4Toolkit.gms in the X4 user GMS folder
'
'  src\*.bas are stored as UTF-8 so they read well on GitHub,
'  but VBE imports ANSI text, so every file is re-encoded to a
'  GBK temp copy before import.
'
'  This script is ASCII-only on purpose: Windows Script Host
'  reads .vbs as ANSI, so a UTF-8 file containing Chinese
'  literals would be mis-parsed (and can even break the quotes).
'==========================================================

Dim fso, sh, here, srcDir, tmpDir, seed, tgt
Dim mods, k, i, c, p, app, vbe, docMod, code, fm, saveCtl, found, cp

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

here   = fso.GetParentFolderName(WScript.ScriptFullName)
srcDir = here & "\src"
tmpDir = here & "\_build_tmp"

mods = Array("M_Util.bas", "M_Curves.bas", "M_Rect.bas", "M_Color.bas", "M_CMYK.bas", _
             "M_FitPath.bas", "M_JPG.bas", "M_PageNo.bas", "M_Calendar.bas", "M_Seal.bas", _
             "M_Install.bas")

seed = FindSeed()
If seed = "" Then
  WScript.Echo "!! no CorelDRAW X4 GMS folder found - install and run X4 once first"
  WScript.Quit 1
End If

tgt = TargetGms()
WScript.Echo "[0] seed   = " & seed
WScript.Echo "[0] target = " & tgt

If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir, True
fso.CreateFolder tmpDir
For k = 0 To UBound(mods)
  MakeAnsiCopy srcDir & "\" & mods(k), tmpDir & "\" & mods(k)
Next
WScript.Echo "[0] re-encoded " & (UBound(mods) + 1) & " sources to GBK"

If fso.FileExists(tgt) Then fso.DeleteFile tgt, True
fso.CopyFile seed, tgt, True
WScript.Echo "[1] seed copied"

Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA()
Set vbe = app.VBE

found = 0
For i = 1 To vbe.VBProjects.Count
  Set p = vbe.VBProjects.Item(i)
  If InStr(1, p.FileName, "CDRX4Toolkit.gms", vbTextCompare) > 0 Then found = i
Next
If found = 0 Then
  WScript.Echo "!! seed not loaded"
  WScript.Quit 1
End If
Set p = vbe.VBProjects.Item(found)
p.Name = "CDRX4Toolkit"
WScript.Echo "[2] project = " & p.Name & " prot=" & p.Protection

For i = p.VBComponents.Count To 1 Step -1
  On Error Resume Next
  Err.Clear
  Set c = p.VBComponents.Item(i)
  If Err.Number = 0 Then
    If c.Type <> 100 Then p.VBComponents.Remove c
  End If
  Err.Clear
  On Error GoTo 0
Next
WScript.Echo "[3] comps after cleanup = " & p.VBComponents.Count

For k = 0 To UBound(mods)
  On Error Resume Next
  Err.Clear
  p.VBComponents.Import tmpDir & "\" & mods(k)
  If Err.Number <> 0 Then WScript.Echo "    import " & mods(k) & " ERR " & Err.Number & " " & Err.Description
  Err.Clear
  On Error GoTo 0
Next
WScript.Echo "[4] comps now = " & p.VBComponents.Count

' --- startup handler: this is what auto-installs the toolbar ---
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

' --- activate the project, then save it through the VBE File menu ---
On Error Resume Next
Err.Clear
Set cp = p.VBComponents.Item(1).CodeModule.CodePane
cp.Show
cp.Window.SetFocus
WScript.Sleep 1500
WScript.Echo "[6] active = " & vbe.ActiveVBProject.Name
Err.Clear

Set fm = vbe.CommandBars.Item(1).Controls.Item(1)
Set saveCtl = Nothing
For i = 1 To fm.Controls.Count
  Err.Clear
  Set c = fm.Controls.Item(i)
  If Err.Number = 0 Then
    If c.ID = 3 Then Set saveCtl = c
  End If
  Err.Clear
Next
If saveCtl Is Nothing Then
  WScript.Echo "!! save control not found"
Else
  saveCtl.Execute
  WScript.Echo "[7] saved err=" & Err.Number & " " & Err.Description
End If
Err.Clear
WScript.Sleep 2500
On Error GoTo 0

WScript.Echo "[8] size = " & fso.GetFile(tgt).Size
If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir, True
WScript.Echo "done"


' --- helpers -------------------------------------------------

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