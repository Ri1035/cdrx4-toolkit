Option Explicit

'==========================================================
'  CorelDRAW X4 probe #13  -  find the GMSManager.RunMacro
'  signature so we can run a macro from outside (compile test).
'  ASCII ONLY.
'==========================================================

Dim fso, sh, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe13.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

Sub Chk(label, en, ed)
  If en = 0 Then
    W "  OK    " & label
  Else
    W "  FAIL  " & label & "   [" & en & "] " & ed
  End If
End Sub

Dim app, o, gm
Dim i, gmsPath

W "probe13 start " & Now
On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
W "  boot err " & Err.Number

gmsPath = sh.ExpandEnvironmentStrings("%APPDATA%") & _
          "\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms"
W "  gms exists = " & fso.FileExists(gmsPath)
W "  gms path   = " & gmsPath

On Error Resume Next
Err.Clear
Set gm = Nothing
Set gm = app.GMSManager
W "  GMSManager = " & TypeName(gm) & "  [" & Err.Number & "]"

If Not gm Is Nothing Then
  On Error Resume Next
  Err.Clear
  gm.RunMacro gmsPath, "M_Util.HasDocument"
  Chk "RunMacro(fullpath, M_Util.HasDocument)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  gm.RunMacro "CDRX4Toolkit.gms", "M_Util.HasDocument"
  Chk "RunMacro(CDRX4Toolkit.gms, M_Util.HasDocument)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  gm.RunMacro "", "M_Util.HasDocument"
  Chk "RunMacro("""", M_Util.HasDocument)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  gm.RunMacro "M_Util.HasDocument", gmsPath
  Chk "RunMacro(M_Util.HasDocument, fullpath)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set o = Nothing
  Set o = gm.GMSFile
  W "  gm.GMSFile = " & o & "  [" & Err.Number & "] " & Err.Description

  On Error Resume Next
  Err.Clear
  W "  gm.Macros.Count = " & gm.Macros.Count & "  [" & Err.Number & "] " & Err.Description
End If

' list the VBProjects so we can see whether our GMS is loaded
On Error Resume Next
Err.Clear
Dim p
For i = 1 To app.VBE.VBProjects.Count
  Set p = app.VBE.VBProjects.Item(i)
  If Err.Number = 0 Then
    W "  proj " & i & " = " & p.Name & "   file=" & p.FileName
  End If
  Err.Clear
Next

W ""
W "done"

logFile.Close