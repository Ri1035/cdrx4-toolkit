Option Explicit

'==========================================================
'  CorelDRAW X4 probe #6 - node/segment layout (for the star
'  and for walking a path).  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe6.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

Dim app, doc, pg, lay, o, sp, sg, shp
Dim i, k

W "probe6 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
Set doc = app.CreateDocument
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer
W "  new document, err " & Err.Number & " " & Err.Description

'----------------------------------------------------------
W ""
W "=== node layout of a converted polygon ==="
Dim sides
For i = 0 To 1
  sides = 5
  If i = 1 Then sides = 10
  On Error Resume Next
  Err.Clear
  Set shp = Nothing
  Set shp = lay.CreatePolygon2(0, 0, 20, sides)
  W "  polygon sides=" & sides & "  size " & shp.SizeWidth & " x " & shp.SizeHeight
  On Error Resume Next
  Err.Clear
  shp.ConvertToCurves
  W "    converted, err " & Err.Number
  On Error Resume Next
  W "    nodes = " & shp.Curve.Nodes.Count
  For k = 1 To shp.Curve.Nodes.Count
    On Error Resume Next
    W "      node " & k & "  x=" & shp.Curve.Nodes(k).PositionX & _
      "  y=" & shp.Curve.Nodes(k).PositionY & _
      "  type=" & shp.Curve.Nodes(k).Type
  Next
  On Error Resume Next
  W "    subpaths = " & shp.Curve.SubPaths.Count & _
    "  segs = " & shp.Curve.Segments.Count
  On Error Resume Next
  W "    subpath(1).Nodes = " & shp.Curve.SubPaths(1).Nodes.Count & _
    "  subpath(1).Segments = " & shp.Curve.SubPaths(1).Segments.Count & _
    "  subpath(1).Closed = " & shp.Curve.SubPaths(1).Closed
Next

'----------------------------------------------------------
W ""
W "=== Node members ==="
On Error Resume Next
Err.Clear
Set shp = lay.CreatePolygon2(0, 0, 20, 5)
shp.ConvertToCurves
Set o = Nothing
Set o = shp.Curve.Nodes(1)
W "  TypeName(node) = " & TypeName(o)
Dim nm
nm = Array("ControlPoint1X", "ControlPoint1Y", "ControlPoint2X", "ControlPoint2Y", _
           "Segment", "GetSegment", "NextNode", "PreviousNode", "SubPath", "Index", _
           "PositionX", "PositionY", "Type", "GetPointPosition")
For i = 0 To UBound(nm)
  On Error Resume Next
  Err.Clear
  Dim v
  v = Empty
  Execute "v = o." & nm(i)
  If Err.Number = 438 Then
    W "  --    node." & nm(i)
  Else
    W "  ++    node." & nm(i) & " -> [" & Err.Number & "] " & Err.Description & _
      "   value=" & v & "  TypeName=" & TypeName(v)
  End If
Next

On Error Resume Next
Err.Clear
o.ControlPoint1X = 1
W "  node.ControlPoint1X = 1 -> [" & Err.Number & "] " & Err.Description
On Error Resume Next
Err.Clear
o.SetPosition 3, 4
W "  node.SetPosition 3,4 -> [" & Err.Number & "] " & Err.Description & _
  "  now x=" & o.PositionX & " y=" & o.PositionY

'----------------------------------------------------------
W ""
W "=== Segment members ==="
On Error Resume Next
Err.Clear
Set sg = Nothing
Set sg = shp.Curve.Segments(1)
W "  TypeName(seg) = " & TypeName(sg)
nm = Array("ControlPoint1", "ControlPoint2", "StartNode", "EndNode", "Length", _
           "GetLength", "Type", "Index", "GetPointAt", "GetPointPositionAt")
For i = 0 To UBound(nm)
  On Error Resume Next
  Err.Clear
  Dim v2
  v2 = Empty
  Execute "v2 = sg." & nm(i)
  If Err.Number = 438 Then
    W "  --    seg." & nm(i)
  Else
    W "  ++    seg." & nm(i) & " -> [" & Err.Number & "] " & Err.Description & _
      "   value=" & v2 & "  TypeName=" & TypeName(v2)
  End If
Next

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sg.GetPointAt(0.5)
W "  seg.GetPointAt(0.5) -> [" & Err.Number & "] " & Err.Description & _
  "  TypeName " & TypeName(o)
If Not o Is Nothing Then
  On Error Resume Next
  W "        x/y = " & o.x & "," & o.y
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sg.GetPointPositionAt(0.5)
W "  seg.GetPointPositionAt(0.5) -> [" & Err.Number & "] " & Err.Description
If Not o Is Nothing Then
  On Error Resume Next
  W "        x/y = " & o.x & "," & o.y
End If

'----------------------------------------------------------
W ""
W "=== subpath members ==="
On Error Resume Next
Err.Clear
Set sp = Nothing
Set sp = shp.Curve.SubPaths(1)
nm = Array("Nodes", "Segments", "Closed", "Length", "Type", "Index")
For i = 0 To UBound(nm)
  On Error Resume Next
  Err.Clear
  Dim v3
  v3 = Empty
  Execute "v3 = sp." & nm(i)
  If Err.Number = 438 Then
    W "  --    subpath." & nm(i)
  Else
    W "  ++    subpath." & nm(i) & " -> [" & Err.Number & "] " & Err.Description & _
      "   TypeName=" & TypeName(v3)
  End If
Next

'----------------------------------------------------------
W ""
W "=== star by moving alternate nodes of a 10-gon ==="
On Error Resume Next
Err.Clear
Set shp = Nothing
Set shp = lay.CreatePolygon2(0, 0, 20, 10)
W "  10-gon size " & shp.SizeWidth & " x " & shp.SizeHeight
shp.ConvertToCurves
W "  nodes = " & shp.Curve.Nodes.Count
' move every second node inward to radius 8
Dim R, r2, ang, pi
pi = 4 * Atn(1)
R = 20
r2 = 8
Dim moved
moved = 0
For k = 1 To shp.Curve.Nodes.Count
  On Error Resume Next
  Err.Clear
  ang = 2 * pi * (k - 1) / shp.Curve.Nodes.Count
  Dim targetR
  targetR = R
  If (k Mod 2) = 0 Then targetR = r2
  shp.Curve.Nodes(k).SetPosition targetR * Cos(ang), targetR * Sin(ang)
  If Err.Number = 0 Then moved = moved + 1
Next
W "  moved " & moved & " of " & shp.Curve.Nodes.Count & " nodes"
W "  resulting size " & shp.SizeWidth & " x " & shp.SizeHeight
On Error Resume Next
Err.Clear
shp.Fill.ApplyUniformFill app.CreateRGBColor(255, 0, 0)
W "  fill applied, err " & Err.Number & " " & Err.Description

'----------------------------------------------------------
W ""
W "=== empty selection ==="
On Error Resume Next
Err.Clear
doc.ClearSelection
Dim sr
Set sr = Nothing
Set sr = app.ActiveSelectionRange
W "  ActiveSelectionRange with nothing selected -> [" & Err.Number & "] " & Err.Description & _
  "  TypeName " & TypeName(sr)
If Not sr Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "        Count = " & sr.Count & "  err " & Err.Number
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.ActiveSelection
W "  ActiveSelection with nothing selected -> [" & Err.Number & "] " & Err.Description & _
  "  TypeName " & TypeName(o)

'----------------------------------------------------------
W ""
W "=== combining shapes (for welding star parts) ==="
On Error Resume Next
Err.Clear
Set shp = lay.CreatePolygon2(0, 0, 20, 5)
shp.ConvertToCurves
W "  curve.Combine present? " & TypeName(shp.Curve.SubPaths.Count)
On Error Resume Next
Err.Clear
shp.Combine
W "  shp.Combine -> [" & Err.Number & "] " & Err.Description

W ""
W "done"

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close