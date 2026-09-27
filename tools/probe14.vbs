Option Explicit

'==========================================================
'  CorelDRAW X4 probe #14
'  1) find the working GMSManager.RunMacro form
'  2) dump the VBA source of the reference GMS plugins that
'     ship with this X4 install (API reference only).
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe14.log", True)

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

Dim app, gm, p, c, cm
Dim i, j, outDir, txt

W "probe14 start " & Now
On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
W "  boot err " & Err.Number

'----------------------------------------------------------
W ""
W "=== A. RunMacro forms ==="
On Error Resume Next
Set gm = Nothing
Set gm = app.GMSManager

On Error Resume Next
Err.Clear
gm.RunMacro "CDRX4Toolkit", "M_Util.HasDocument"
Chk "RunMacro(CDRX4Toolkit, M_Util.HasDocument)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
gm.RunMacro "CDRX4Toolkit", "M_Util", "HasDocument"
Chk "RunMacro(CDRX4Toolkit, M_Util, HasDocument)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
gm.RunMacro "CDRX4Toolkit.M_Util", "HasDocument"
Chk "RunMacro(CDRX4Toolkit.M_Util, HasDocument)", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== B. project protection / dump ==="
outDir = root & "\_ref"
On Error Resume Next
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir

Dim want
want = Array("ConverTo", "ColorReplacer", "FitObjectsToPath", "CalendarWizard", _
             "Cachet_公章", "RectangleFixer", "ToJPG", "Node123", "CDRX4Toolkit")

For i = 1 To app.VBE.VBProjects.Count
  On Error Resume Next
  Err.Clear
  Set p = Nothing
  Set p = app.VBE.VBProjects.Item(i)
  If Err.Number = 0 And Not p Is Nothing Then
    Dim nm, wanted
    nm = p.Name
    wanted = False
    For j = 0 To UBound(want)
      If nm = want(j) Then wanted = True
    Next
    If wanted Then
      W ""
      W "  project " & nm & "  comps=" & p.VBComponents.Count & _
        "  protection=" & p.Protection & "  file=" & p.FileName
      For j = 1 To p.VBComponents.Count
        On Error Resume Next
        Err.Clear
        Set c = Nothing
        Set c = p.VBComponents.Item(j)
        If Err.Number = 0 And Not c Is Nothing Then
          On Error Resume Next
          Err.Clear
          Set cm = Nothing
          Set cm = c.CodeModule
          If Err.Number = 0 And Not cm Is Nothing Then
            If cm.CountOfLines > 0 Then
              txt = cm.Lines(1, cm.CountOfLines)
              Dim fname
              fname = outDir & "\" & San(nm) & "__" & San(c.Name) & ".txt"
              On Error Resume Next
              Err.Clear
              Dim st
              Set st = CreateObject("ADODB.Stream")
              st.Type = 2
              st.Charset = "utf-8"
              st.Open
              st.WriteText txt
              st.SaveToFile fname, 2
              st.Close
              W "      wrote " & fname & "  lines=" & cm.CountOfLines & _
                "  err=" & Err.Number & " " & Err.Description
            Else
              W "      comp " & c.Name & " (empty)"
            End If
          End If
        End If
        Err.Clear
        On Error GoTo 0
      Next
    End If
  End If
  Err.Clear
  On Error GoTo 0
Next

W ""
W "done"

logFile.Close

Function San(ByVal s)
  Dim bad, k, ch
  bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
  For k = 0 To UBound(bad)
    s = Replace(s, bad(k), "_")
  Next
  San = s
End Function