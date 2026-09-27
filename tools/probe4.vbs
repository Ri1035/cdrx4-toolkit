Option Explicit

'==========================================================
'  CorelDRAW X4 probe #4.
'  ASCII ONLY (WSH reads .vbs as ANSI).
'==========================================================

Dim fso, here, root, logFile
Dim nOK, nFail

Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe4.log", True)

nOK = 0
nFail = 0

Sub W(s)
  logFile.WriteLine s
End Sub

Sub T(lbl, n, d)
  If n = 0 Then
    W "  OK    " & lbl
    nOK = nOK + 1
  Else
    W "  FAIL  " & lbl & "   [" & n & "] " & d
    nFail = nFail + 1
  End If
End Sub

Dim app, doc, pg, lay
Dim rect, rect2, ell, txt, poly, s1, sr, opt, crv, o
Dim i, k

W "probe4 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
Set doc = app.CreateDocument
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer
T "new document", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== star: pentagon vs sharpness, node count done right ==="
For k = 0 To 2
  Dim shp, w0, h0, n0
  On Error Resume Next
  Err.Clear
  If k = 0 Then
    Set shp = lay.CreatePolygon2(0, 0, 20, 5)
  ElseIf k = 1 Then
    Set shp = lay.CreatePolygon2(0, 0, 20, 5, 53)
  Else
    Set shp = lay.CreatePolygon2(0, 0, 20, 5)
    shp.Polygon.Sharpness = 53
  End If
  W "  case " & k & ": created, size " & shp.SizeWidth & " x " & shp.SizeHeight & _
    "  sides " & shp.Polygon.Sides & "  sharpness " & shp.Polygon.Sharpness
  On Error Resume Next
  Err.Clear
  shp.ConvertToCurves
  T "    ConvertToCurves (statement)", Err.Number, Err.Description
  On Error Resume Next
  n0 = shp.Curve.Nodes.Count
  W "    after convert: type " & shp.Type & "  nodes " & n0 & _
    "  size " & shp.SizeWidth & " x " & shp.SizeHeight
Next

W ""
W "=== also try CreatePolygon with args ==="
On Error Resume Next
Err.Clear
Set poly = Nothing
Set poly = lay.CreatePolygon(0, 0, 20, 5)
T "CreatePolygon(0,0,20,5)", Err.Number, Err.Description
W "        TypeName = " & TypeName(poly)

On Error Resume Next
Err.Clear
Set poly = Nothing
Set poly = lay.CreatePolygon(0, 0, 20, 5, 53)
T "CreatePolygon(0,0,20,5,53)", Err.Number, Err.Description
W "        TypeName = " & TypeName(poly)

'----------------------------------------------------------
W ""
W "=== building a curve by hand ==="
On Error Resume Next
Err.Clear
Set crv = Nothing
Set crv = lay.CreateLineSegment(0, 0, 10, 10)
T "lay.CreateLineSegment(0,0,10,10)", Err.Number, Err.Description
W "        TypeName = " & TypeName(crv)
If Not crv Is Nothing Then
  On Error Resume Next
  W "        type = " & crv.Type & "  nodes = " & crv.Curve.Nodes.Count & _
    "  length = " & crv.Curve.Length
  On Error Resume Next
  Err.Clear
  Dim sp
  Set sp = Nothing
  Set sp = crv.Curve.CreateSubPath(20, 0)
  T "crv.Curve.CreateSubPath(20,0)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Dim sg
  Set sg = Nothing
  Set sg = crv.Curve.CreateSubPath(20, 20)
  If Err.Number = 0 And Not sg Is Nothing Then
    On Error Resume Next
    Err.Clear
    Set o = Nothing
    Set o = sg.AddNode(30, 30)
    T "subpath.AddNode(30,30)", Err.Number, Err.Description
    On Error Resume Next
    Err.Clear
    sg.Closed = True
    T "subpath.Closed = True", Err.Number, Err.Description
    On Error Resume Next
    W "        nodes now = " & crv.Curve.Nodes.Count
  End If

  On Error Resume Next
  Err.Clear
  Set o = Nothing
  Set o = crv.Curve.Nodes(1)
  T "crv.Curve.Nodes(1)", Err.Number, Err.Description
  If Not o Is Nothing Then
    On Error Resume Next
    Err.Clear
    W "        node1 pos = " & o.PositionX & "," & o.PositionY
    T "node.PositionX/Y (read)", Err.Number, Err.Description
    On Error Resume Next
    Err.Clear
    o.SetPosition 1, 1
    T "node.SetPosition 1,1", Err.Number, Err.Description
    On Error Resume Next
    Err.Clear
    o.PositionX = 2
    T "node.PositionX = 2", Err.Number, Err.Description
  End If

  On Error Resume Next
  Err.Clear
  W "        crv.Curve.SubPaths.Count = " & crv.Curve.SubPaths.Count
  T "crv.Curve.SubPaths.Count", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Dim csh
  Set csh = Nothing
  Set csh = crv.Curve.SubPaths(1)
  T "crv.Curve.SubPaths(1)", Err.Number, Err.Description
  If Not csh Is Nothing Then
    On Error Resume Next
    Err.Clear
    csh.Closed = True
    T "subpaths(1).Closed = True", Err.Number, Err.Description
    On Error Resume Next
    W "        subpath nodes = " & csh.Nodes.Count
  End If
End If

'----------------------------------------------------------
W ""
W "=== curve geometry for fit-to-path ==="
On Error Resume Next
Err.Clear
Set ell = lay.CreateEllipse2(0, 0, 40, 30)
T "CreateEllipse2(0,0,40,30)", Err.Number, Err.Description
On Error Resume Next
Err.Clear
ell.ConvertToCurves
T "ell.ConvertToCurves (statement)", Err.Number, Err.Description
If Not ell Is Nothing Then
  W "        type = " & ell.Type
  On Error Resume Next
  Err.Clear
  W "        ell.Curve.Length = " & ell.Curve.Length
  T "ell.Curve.Length", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set o = Nothing
  Set o = ell.Curve.GetPointPositionAt(ell.Curve.Length / 2)
  T "ell.Curve.GetPointPositionAt(len/2)", Err.Number, Err.Description
  W "        TypeName = " & TypeName(o)
  If Not o Is Nothing Then
    On Error Resume Next
    W "        .x/.y   = " & o.x & "," & o.y
    On Error Resume Next
    W "        .X/.Y   = " & o.X & "," & o.Y
  End If

  On Error Resume Next
  Err.Clear
  W "        ell.Curve.GetPointAngleAt(len/2) = " & ell.Curve.GetPointAngleAt(ell.Curve.Length / 2)
  T "ell.Curve.GetPointAngleAt(len/2)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "        ell.Curve.Nodes.Count = " & ell.Curve.Nodes.Count
  T "ell.Curve.Nodes.Count", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Dim t2
  Set t2 = Nothing
  Set t2 = ell.Curve.SubPaths.Count
  W "        subpaths = " & ell.Curve.SubPaths.Count
End If

'----------------------------------------------------------
W ""
W "=== selection order, distinct sizes ==="
On Error Resume Next
Err.Clear
doc.ClearSelection
Set rect = lay.CreateRectangle2(0, 0, 11, 11)
Set rect2 = lay.CreateRectangle2(50, 0, 22, 22)
T "rect 11x11 and rect2 22x22", Err.Number, Err.Description

On Error Resume Next
Err.Clear
rect.AddToSelection
rect2.AddToSelection
W "        rect then rect2 -> count " & app.ActiveSelection.Shapes.Count & _
  "  shapes(1).w " & app.ActiveSelection.Shapes(1).SizeWidth & _
  "  shapes(2).w " & app.ActiveSelection.Shapes(2).SizeWidth
T "select rect then rect2", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.ClearSelection
rect2.AddToSelection
rect.AddToSelection
W "        rect2 then rect -> count " & app.ActiveSelection.Shapes.Count & _
  "  shapes(1).w " & app.ActiveSelection.Shapes(1).SizeWidth & _
  "  shapes(2).w " & app.ActiveSelection.Shapes(2).SizeWidth
T "select rect2 then rect", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== export signature ==="
Dim tmp
tmp = fso.GetSpecialFolder(2) & "\cdrx4_exp.jpg"
On Error Resume Next
Err.Clear
If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
doc.Export tmp
W "        Export(1 arg)   -> [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
doc.Export tmp, 8
W "        Export(2 args)  -> [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
doc.Export tmp, 8, 1
W "        Export(3 args)  -> [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set opt = app.CreateStructExportOptions
doc.Export tmp, 8, 1, opt
W "        Export(4 args)  -> [" & Err.Number & "] " & Err.Description & _
  "  file=" & fso.FileExists(tmp)

On Error Resume Next
Err.Clear
doc.Export tmp, 8, 1, opt, 1
W "        Export(5 args)  -> [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
doc.ExportEx tmp
W "        ExportEx(1 arg) -> [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
doc.ExportEx tmp, 8, 1, opt
W "        ExportEx(4 args)-> [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Dim pg2
Set pg2 = doc.ActivePage
pg2.Export tmp, 8, 1, opt
W "        Page.Export      -> [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
W "        doc.Name = " & doc.Name & "  FileName=[" & doc.FileName & "]"

'----------------------------------------------------------
W ""
W "=== SaveAs / GetFileName helpers ==="
On Error Resume Next
Err.Clear
W "        TypeName(doc.Pages(1)) = " & TypeName(doc.Pages(1))
T "doc.Pages(1)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.Pages(1).Shapes.All
T "doc.Pages(1).Shapes.All", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.ActivePage.Layers
T "doc.ActivePage.Layers", Err.Number, Err.Description

W ""
W "done: OK=" & nOK & "  FAIL=" & nFail

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close