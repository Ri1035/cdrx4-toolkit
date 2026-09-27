Option Explicit

'==========================================================
'  CorelDRAW X4 probe #17 - can we save a GMS project with
'  VBProject.SaveAs instead of driving the VBE File menu?
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe17.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

Dim app, vbe, p, i, cnt, matchIdx
Dim nm, fn, prot, comps, saved

W "probe17 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
W "  boot err " & Err.Number & " " & Err.Description
Err.Clear
Set vbe = app.VBE
On Error GoTo 0

If vbe Is Nothing Then
  W "!! no VBE"
  logFile.Close
  WScript.Quit 1
End If

cnt = 0
On Error Resume Next
cnt = vbe.VBProjects.Count
Err.Clear
On Error GoTo 0
W "projects = " & cnt

matchIdx = 0
For i = 1 To cnt
  On Error Resume Next
  Err.Clear
  Set p = Nothing
  Set p = vbe.VBProjects.Item(i)
  If Err.Number = 0 And Not p Is Nothing Then
    nm = ""
    fn = ""
    prot = -1
    comps = -1
    saved = -1
    Err.Clear
    nm = p.Name
    Err.Clear
    fn = p.FileName
    Err.Clear
    prot = p.Protection
    Err.Clear
    comps = p.VBComponents.Count
    Err.Clear
    saved = p.Saved
    Err.Clear
    W "  [" & i & "] name=[" & nm & "] file=[" & fn & "] prot=" & prot & _
      " comps=" & comps & " saved=" & saved & " err=" & Err.Number
    If InStr(1, fn, "CDRX4Toolkit.gms", vbTextCompare) > 0 Then matchIdx = i
    If nm = "CDRX4Toolkit" And matchIdx = 0 Then matchIdx = i
  Else
    W "  [" & i & "] item err " & Err.Number & " " & Err.Description
  End If
  Err.Clear
  On Error GoTo 0
Next

W ""
W "matchIdx = " & matchIdx
If matchIdx = 0 Then
  W "!! our project not loaded"
  logFile.Close
  WScript.Quit 1
End If

On Error Resume Next
Err.Clear
Set p = vbe.VBProjects.Item(matchIdx)
W "target name=[" & p.Name & "] err=" & Err.Number
Err.Clear
On Error GoTo 0

' --- does SaveAs exist / work? ---
Dim tmp
tmp = root & "\_probe17_saveas.gms"
If fso.FileExists(tmp) Then fso.DeleteFile tmp, True

On Error Resume Next
Err.Clear
p.SaveAs tmp
W ""
W "SaveAs err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

If fso.FileExists(tmp) Then
  W "SaveAs produced size = " & fso.GetFile(tmp).Size
Else
  W "SaveAs produced NOTHING"
End If

On Error Resume Next
Err.Clear
W "after: name=[" & p.Name & "] file=[" & p.FileName & "] saved=" & p.Saved & _
  " err=" & Err.Number
Err.Clear
On Error GoTo 0

W ""
W "done"
logFile.Close