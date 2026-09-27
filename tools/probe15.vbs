Option Explicit

'==========================================================
'  CorelDRAW X4 probe #15 - dump every VBA project source.
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe15.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

Dim app, p, c, cm, st
Dim i, j, outDir, txt, nm, fname

W "probe15 start " & Now
On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
W "  boot err " & Err.Number

outDir = root & "\_ref"
On Error Resume Next
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir

Dim total
total = 0
On Error Resume Next
total = app.VBE.VBProjects.Count
W "  projects = " & total

For i = 1 To total
  On Error Resume Next
  Err.Clear
  Set p = Nothing
  Set p = app.VBE.VBProjects.Item(i)
  If Err.Number <> 0 Then
    W "  [" & i & "] item err " & Err.Number & " " & Err.Description
  Else
    nm = ""
    Err.Clear
    nm = p.Name
    W ""
    W "  [" & i & "] name=" & nm & "  file=" & p.FileName & _
      "  prot=" & p.Protection & "  comps=" & p.VBComponents.Count

    For j = 1 To p.VBComponents.Count
      On Error Resume Next
      Err.Clear
      Set c = Nothing
      Set c = p.VBComponents.Item(j)
      If Err.Number <> 0 Then
        W "      comp " & j & " err " & Err.Number
      Else
        Dim cname, ctype, nlines
        cname = ""
        ctype = -1
        nlines = 0
        Err.Clear
        cname = c.Name
        ctype = c.Type
        Err.Clear
        Set cm = Nothing
        Set cm = c.CodeModule
        If Err.Number <> 0 Then
          W "      comp " & j & " (" & cname & ") type=" & ctype & _
            " codemodule err " & Err.Number & " " & Err.Description
        Else
          Err.Clear
          nlines = cm.CountOfLines
          If Err.Number <> 0 Then
            W "      comp " & j & " (" & cname & ") count err " & Err.Number & _
              " " & Err.Description
          Else
            W "      comp " & j & " (" & cname & ") type=" & ctype & _
              " lines=" & nlines
            If nlines > 0 Then
              Err.Clear
              txt = cm.Lines(1, nlines)
              If Err.Number = 0 Then
                fname = outDir & "\" & Pad(i) & "_" & San(nm) & "__" & _
                        San(cname) & ".txt"
                Err.Clear
                Set st = CreateObject("ADODB.Stream")
                st.Type = 2
                st.Charset = "utf-8"
                st.Open
                st.WriteText txt
                st.SaveToFile fname, 2
                st.Close
                If Err.Number <> 0 Then
                  W "        write err " & Err.Number & " " & Err.Description
                End If
              Else
                W "        read lines err " & Err.Number & " " & Err.Description
              End If
            End If
          End If
        End If
      End If
    Next
  End If
  Err.Clear
Next

W ""
W "done"

logFile.Close

Function Pad(ByVal n)
  If n < 10 Then
    Pad = "0" & n
  Else
    Pad = CStr(n)
  End If
End Function

Function San(ByVal s)
  Dim bad, k
  bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
  For k = 0 To UBound(bad)
    s = Replace(s, bad(k), "_")
  Next
  San = s
End Function