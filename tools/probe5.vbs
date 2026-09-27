Option Explicit

'==========================================================
'  CorelDRAW X4 probe #5 - star + blend + curve sampling.
'  ASCII ONLY (WSH reads .vbs as ANSI).
'==========================================================

Dim fso, here, root, logFile
Dim nOK, nFail

Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe5.log", True)

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
Dim rect, rect2, ell, poly, s1, sr, crv, o, sp, sg
Dim i, k

W "probe5 start " & Now

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
W "=== ActiveSelectionRange ==="
On Error Resume Next
Err.Clear
Set rect = lay.CreateRectangle2(0, 0, 11, 11)
Set rect2 = lay.CreateRectangle2(50, 0, 22, 22)
rect.AddToSelection
rect2.AddToSelection
T "select two rects", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set sr = Nothing
Set sr = app.ActiveSelectionRange
T "Set sr = app.ActiveSelectionRange", Err.Number, Err.Description
W "        TypeName = " & TypeName(sr)
If Not sr Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "        Count = " & sr.Count
  T "sr.Count", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  Set o = Nothing
  Set o = sr(1)
  T "sr(1)", Err.Number, Err.Description
  On Error Resume Next
  W "        sr(1).SizeWidth = " & sr(1).SizeWidth & "  sr(2).SizeWidth = " & sr(2).SizeWidth

  On Error Resume Next
  Err.Clear
  Set o = Nothing
  Set o = sr.CreateBlend
  T "sr.CreateBlend", Err.Number, Err.Description
  W "        TypeName = " & TypeName(o)
  If Not o Is Nothing Then
    On Error Resume Next
    Err.Clear
    W "        blend.Blend = " & TypeName(o.Blend)
    T "o.Blend", Err.Number, Err.Description
  End If
End If

W ""
W "=== blend on a path ==="
On Error Resume Next
Err.Clear
doc.ClearSelection
Set ell = lay.CreateEllipse2(0, 0, 60, 40)
Set rect = lay.CreateRectangle2(0, 0, 6, 6)
Set rect2 = lay.CreateRectangle2(0, 40, 6, 6)
rect.AddToSelection
rect2.AddToSelection
Set sr = Nothing
Set sr = app.ActiveSelectionRange
Set o = Nothing
Set o = sr.CreateBlend
T "sr.CreateBlend on 2 rects", Err.Number, Err.Description
If Not o Is Nothing Then
  On Error Resume Next
  Err.Clear
  o.Blend.Path = ell
  T "blend.Path = ellipse", Err.Number, Err.Description
End If

'----------------------------------------------------------
W ""
W "=== CreatePolygon variants (looking for a star) ==="
Dim shp
For k = 0 To 3
  On Error Resume Next
  Err.Clear
  Set shp = Nothing
  If k = 0 Then Set shp = lay.CreatePolygon(0, 0, 20, 5, 53)
  If k = 1 Then Set shp = lay.CreatePolygon(0, 0, 5, 20, 53)
  If k = 2 Then Set shp = lay.CreatePolygon(0, 0, 20, 5, 0)
  If k = 3 Then Set shp = lay.CreatePolygon(0, 0, 20, 5)
  If Err.Number <> 0 Then
    W "  case " & k & " -> [" & Err.Number & "] " & Err.Description
  ElseIf shp Is Nothing Then
    W "  case " & k & " -> nothing"
  Else
    W "  case " & k & " -> size " & shp.SizeWidth & " x " & shp.SizeHeight & _
      "  type " & shp.Type
    On Error Resume Next
    Err.Clear
    shp.ConvertToCurves
    If Err.Number = 0 Then
      On Error Resume Next
      W "        after convert: nodes " & shp.Curve.Nodes.Count & _
        "  size " & shp.SizeWidth & " x " & shp.SizeHeight
    End If
  End If
Next

W ""
W "=== node-add API candidates ==="
On Error Resume Next
Err.Clear
Set crv = Nothing
Set crv = lay.CreateLineSegment(0, 0, 10, 10)
T "CreateLineSegment", Err.Number, Err.Description
Dim cands
cands = Array("AddNode", "AddSegment", "AppendNode", "AppendSegment", "CreateNode", _
              "AddNodes", "InsertNode", "AddPoint", "AddLineSegment", "AddCurveSegment")
Dim nm
For i = 0 To UBound(cands)
  On Error Resume Next
  Err.Clear
  Set sp = Nothing
  Set sp = crv.Curve.CreateSubPath(20, 0)
  On Error Resume Next
  Err.Clear
  Set sg = Nothing
  Execute "Set sg = sp." & cands(i) & "(30, 30)"
  If Err.Number = 438 Then
    W "  --    sp." & cands(i) & "  (does not exist)"
  Else
    W "  ++    sp." & cands(i) & "  -> [" & Err.Number & "] " & Err.Description & _
      "   nodes now " & crv.Curve.Nodes.Count
  End If
Next

W ""
W "=== curve sampling candidates ==="
On Error Resume Next
Err.Clear
Set ell = lay.CreateEllipse2(0, 0, 40, 30)
ell.ConvertToCurves
T "ellipse -> curve", Err.Number, Err.Description
Dim cn
cn = Array("GetPointAt", "GetPointPositionAt", "GetPointPosition", "PointAt", _
           "GetPositionAt", "GetTangentAt", "GetPointAngleAt", "GetAngleAt", _
           "GetCurvePointAt", "GetPointAtOffset")
Dim c2
For i = 0 To UBound(cn)
  On Error Resume Next
  Err.Clear
  Set o = Nothing
  Execute "Set o = ell.Curve." & cn(i) & "(10)"
  If Err.Number = 438 Then
    W "  --    Curve." & cn(i) & "  (does not exist)"
  Else
    W "  ++    Curve." & cn(i) & "(10) -> [" & Err.Number & "] " & Err.Description & _
      "  TypeName " & TypeName(o)
    If Err.Number = 0 And Not o Is Nothing Then
      On Error Resume Next
      W "          x/y = " & o.x & "," & o.y
    End If
  End If
Next

W ""
W "=== segments / node geometry ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = ell.Curve.Segments
T "ell.Curve.Segments", Err.Number, Err.Description
W "        TypeName = " & TypeName(o)
If Not o Is Nothing Then
  On Error Resume Next
  W "        Segments.Count = " & ell.Curve.Segments.Count
  On Error Resume Next
  Err.Clear
  Set sg = Nothing
  Set sg = ell.Curve.Segments(1)
  T "ell.Curve.Segments(1)", Err.Number, Err.Description
  W "        TypeName = " & TypeName(sg)
  If Not sg Is Nothing Then
    On Error Resume Next
    Err.Clear
    W "        seg.StartNode = " & TypeName(sg.StartNode) & _
      "  EndNode = " & TypeName(sg.EndNode)
    T "seg.StartNode/EndNode", Err.Number, Err.Description
  End If
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = ell.Curve.Nodes(1)
T "ell.Curve.Nodes(1)", Err.Number, Err.Description
If Not o Is Nothing Then
  On Error Resume Next
  W "        node pos = " & o.PositionX & "," & o.PositionY
  On Error Resume Next
  Err.Clear
  W "        node.GetTangentAt = " & o.GetTangentAt(0.5)
  T "node.GetTangentAt(0.5)", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  W "        node.TangentAngle = " & o.TangentAngle
  T "node.TangentAngle", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  W "        node.GetPointAt = " & o.GetPointAt(0.5)
  T "node.GetPointAt(0.5)", Err.Number, Err.Description
End If

W ""
W "=== shape bounding helpers ==="
On Error Resume Next
Err.Clear
Set rect = lay.CreateRectangle2(0, 0, 10, 10)
T "rect", Err.Number, Err.Description
On Error Resume Next
Err.Clear
rect.GetBoundingBox 0, 0, 0, 0
T "rect.GetBoundingBox 0,0,0,0", Err.Number, Err.Description
On Error Resume Next
Err.Clear
W "        rect.BoundingBox = " & rect.BoundingBox
T "rect.BoundingBox (read)", Err.Number, Err.Description

W ""
W "=== layers / pages for page numbers ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.ActivePage.ActiveLayer
T "doc.ActivePage.ActiveLayer", Err.Number, Err.Description
On Error Resume Next
W "        TypeName = " & TypeName(o)

W ""
W "done: OK=" & nOK & "  FAIL=" & nFail

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close