Option Explicit

'==========================================================
'  CorelDRAW X4 probe #19 - settle the remaining API doubts:
'    - ShapeRange access: Item(i) vs Shapes(i) vs For Each
'    - all four Rectangle.Radius* writes
'    - OverprintFill / OverprintOutline writes
'    - Outline.Color.CopyAssign
'    - BeginCommandGroup / Optimization
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe19.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

Dim app, doc, pg, lay, sr, sh, r, e, t, c, o, ro, x
Dim i, n, cnt

W "probe19 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
W "  app err=" & Err.Number
Err.Clear
app.InitializeVBA
W "  initvba err=" & Err.Number
Err.Clear
Set doc = app.CreateDocument
W "  createDoc nothing=" & (doc Is Nothing) & " err=" & Err.Number
Err.Clear
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer
W "  page/layer err=" & Err.Number
Err.Clear

Set r = lay.CreateRectangle2(0, 0, 40, 30)
Set e = lay.CreateEllipse2(70, 15, 15, 15)
Set t = lay.CreateArtisticTextWide(0, 60, "AB")
W "  shapes err=" & Err.Number & " r=" & (Not r Is Nothing) & _
  " e=" & (Not e Is Nothing) & " t=" & (Not t Is Nothing)
Err.Clear

Set sr = pg.Shapes.All
W ""
W "1. pg.Shapes.All typename=" & TypeName(sr) & " err=" & Err.Number
Err.Clear
cnt = -1
cnt = sr.Count
W "   Count=" & cnt & " err=" & Err.Number
Err.Clear

W ""
W "2. sr.Item(1)"
Set x = Nothing
Set x = sr.Item(1)
W "   nothing=" & (x Is Nothing) & " err=" & Err.Number & " " & Err.Description
Err.Clear
If Not x Is Nothing Then
  W "   type=" & x.Type & " err=" & Err.Number
  Err.Clear
End If

W ""
W "3. sr.Shapes(1)"
Set x = Nothing
Set x = sr.Shapes(1)
W "   nothing=" & (x Is Nothing) & " err=" & Err.Number & " " & Err.Description
Err.Clear
If Not x Is Nothing Then
  W "   type=" & x.Type & " err=" & Err.Number
  Err.Clear
End If

W ""
W "4. For Each over sr"
n = 0
For Each sh In sr
  n = n + 1
  W "   each " & n & " type=" & sh.Type & " err=" & Err.Number
  Err.Clear
Next
W "   total=" & n & " err=" & Err.Number
Err.Clear

W ""
W "5. For Each over pg.Shapes.All (no temp)"
n = 0
For Each sh In pg.Shapes.All
  n = n + 1
  W "   each " & n & " type=" & sh.Type & " err=" & Err.Number
  Err.Clear
Next
W "   total=" & n & " err=" & Err.Number
Err.Clear

' --- 6. Rectangle radius, all four ---
W ""
W "6. Rectangle.Radius* writes"
Set ro = Nothing
Set ro = r.Rectangle
W "   .Rectangle nothing=" & (ro Is Nothing) & " err=" & Err.Number
Err.Clear
If Not ro Is Nothing Then
  ro.RadiusUpperLeft = 0
  W "   RadiusUpperLeft=0 err=" & Err.Number & " " & Err.Description
  Err.Clear
  ro.RadiusUpperRight = 0
  W "   RadiusUpperRight=0 err=" & Err.Number & " " & Err.Description
  Err.Clear
  ro.RadiusLowerLeft = 0
  W "   RadiusLowerLeft=0 err=" & Err.Number & " " & Err.Description
  Err.Clear
  ro.RadiusLowerRight = 0
  W "   RadiusLowerRight=0 err=" & Err.Number & " " & Err.Description
  Err.Clear
  W "   read back UL=" & ro.RadiusUpperLeft & " UR=" & ro.RadiusUpperRight & _
    " LL=" & ro.RadiusLowerLeft & " LR=" & ro.RadiusLowerRight & " err=" & Err.Number
  Err.Clear
End If

' --- 7. overprint ---
W ""
W "7. overprint on Shape"
Set o = r
o.OverprintFill = True
W "   OverprintFill=True err=" & Err.Number & " " & Err.Description
Err.Clear
o.OverprintOutline = True
W "   OverprintOutline=True err=" & Err.Number & " " & Err.Description
Err.Clear

' --- 8. outline color copy ---
W ""
W "8. outline color"
Set c = app.CreateCMYKColor(0, 0, 0, 100)
W "   makeCMYK nothing=" & (c Is Nothing) & " type=" & c.Type & " err=" & Err.Number
Err.Clear
Set o = r
o.Outline.Color.CopyAssign c
W "   CopyAssign err=" & Err.Number & " " & Err.Description
Err.Clear
o.Outline.Width = 1
W "   Outline.Width=1 err=" & Err.Number
Err.Clear

' --- 9. command group + optimization ---
W ""
W "9. command group / optimization"
doc.BeginCommandGroup "probe19"
W "   BeginCommandGroup err=" & Err.Number & " " & Err.Description
Err.Clear
doc.EndCommandGroup
W "   EndCommandGroup err=" & Err.Number & " " & Err.Description
Err.Clear
app.Optimization = True
W "   Optimization=True err=" & Err.Number
Err.Clear
app.Optimization = False
W "   Optimization=False err=" & Err.Number
Err.Clear
app.Refresh
W "   Refresh err=" & Err.Number
Err.Clear

' --- 10. selection ---
W ""
W "10. selection"
doc.ClearSelection
W "   ClearSelection err=" & Err.Number
Err.Clear
r.AddToSelection
W "   AddToSelection err=" & Err.Number
Err.Clear
Set sr = app.ActiveSelectionRange
W "   ActiveSelectionRange nothing=" & (sr Is Nothing) & " err=" & Err.Number
Err.Clear
If Not sr Is Nothing Then
  W "   Count=" & sr.Count & " err=" & Err.Number
  Err.Clear
  Set x = Nothing
  Set x = sr.Item(1)
  W "   .Item(1) nothing=" & (x Is Nothing) & " err=" & Err.Number & " " & Err.Description
  Err.Clear
  If Not x Is Nothing Then
    W "   type=" & x.Type & " err=" & Err.Number
    Err.Clear
  End If
  n = 0
  For Each sh In sr
    n = n + 1
    W "   each " & n & " type=" & sh.Type & " err=" & Err.Number
    Err.Clear
  Next
  W "   For Each total=" & n & " err=" & Err.Number
  Err.Clear
End If

' --- 11. clean up the test doc ---
W ""
On Error Resume Next
Err.Clear
doc.Dirty = False
doc.Close
W "11. close err=" & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0

W ""
W "done"
logFile.Close