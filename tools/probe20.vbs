Option Explicit

'==========================================================
'  CorelDRAW X4 probe #20 - what does GMSManager expose?
'  We want to force a reload of CDRX4Toolkit.gms from disk
'  so the smoke test proves the SAVED file is good, not just
'  the in-memory project.
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe20.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

Dim app, gm, tgt, r, i, p, c, nm, fn

W "probe20 start " & Now

tgt = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%APPDATA%") & _
      "\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms"
W "target = " & tgt
W "exists = " & fso.FileExists(tgt)
If fso.FileExists(tgt) Then W "size   = " & fso.GetFile(tgt).Size

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
W ""
W "create err=" & Err.Number
Err.Clear
app.InitializeVBA
W "initvba err=" & Err.Number
Err.Clear
Set gm = app.GMSManager
W "GMSManager nothing=" & (gm Is Nothing) & " err=" & Err.Number
Err.Clear
On Error GoTo 0

If gm Is Nothing Then
  W "!! no GMSManager"
  logFile.Close
  WScript.Quit 1
End If

W ""
W "--- probing method names ---"

Dim m, args
m = Array("LoadGMS", "UnloadGMS", "IsGMSLoaded", "GMSExists", "GetGMSName", _
          "LoadGMSFile", "UnloadGMSFile", "ReloadGMS", "RunMacro", "ShowGMS")

Dim k, res
For k = 0 To UBound(m)
  On Error Resume Next
  Err.Clear
  res = Empty
  Select Case m(k)
    Case "LoadGMS", "LoadGMSFile", "ReloadGMS", "UnloadGMS", "UnloadGMSFile", "GetGMSName"
      res = Eval("gm." & m(k) & "(tgt)")
    Case "IsGMSLoaded", "GMSExists"
      res = Eval("gm." & m(k) & "(tgt)")
    Case Else
      res = Eval("gm." & m(k))
  End Select
  W "  " & m(k) & " -> err=" & Err.Number & " [" & Err.Description & "] res=" & res
  Err.Clear
  On Error GoTo 0
Next

logFile.Close