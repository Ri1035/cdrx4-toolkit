Option Explicit

'==========================================================
'  CorelDRAW X4 probe #2 - the remaining unknowns.
'  ASCII only (WSH reads .vbs as ANSI).  Late-bound, so
'  nothing here can be a compile error.
'==========================================================

Dim fso, here, root, logFile
Dim nOK, nFail

Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe2.log", True)

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

Dim app, doc, pg, lay, rect, ell, txt, poly, s1, sr, crv
Dim i

W "probe2 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
T "attach + InitializeVBA", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set doc = app.CreateDocument
T "app.CreateDocument", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer
T "doc.ActivePage / pg.ActiveLayer", Err.Number, Err.Description

W ""
W "=== layer Create* candidates (does the method exist at all?) ==="
Dim names, args
names = Array("CreateRectangle", "CreateRectangle2", "CreateEllipse", "CreateEllipse2", _
              "CreateCurve", "CreateCurveFromCurve", "CreatePolyline", "CreateLineSegment", _
              "CreateArtisticText", "CreateArtisticTextWide", "CreateParagraphText", _
              "CreateParagraphTextWide", "CreateBitmap", "CreateFreehand", "CreateConnector", _
              "CreateSpiral", "CreateGrid", "CreateStar", "CreatePolygon", "CreatePolygon2", _
              "CreateObject", "CreateOLEObject", "CreateTextRange", "CreateSelection")
For i = 0 To UBound(names)
  On Error Resume Next
  Err.Clear
  Dim dummy
  Set dummy = Nothing
  Execute "Set dummy = lay." & names(i) & "()"
  If Err.Number = 438 Then
    W "  --    " & names(i) & "  (does not exist)"
  Else
    W "  ++    " & names(i) & "  exists, arity-0 call gives [" & Err.Number & "] " & Err.Description
  End If
Next

W ""
W "=== CreatePolygon2 argument meaning ==="
' vary the two numeric args and watch the resulting bounding box
On Error Resume Next
Err.Clear
Set poly = lay.CreatePolygon2(0, 0, 5, 20)
If Err.Number = 0 And Not poly Is Nothing Then
  W "  CreatePolygon2(0,0,5,20)  -> size " & poly.SizeWidth & " x " & poly.SizeHeight & _
    "  pos " & poly.PositionX & "," & poly.PositionY
End If

On Error Resume Next
Err.Clear
Set poly = lay.CreatePolygon2(0, 0, 20, 5)
If Err.Number = 0 And Not poly Is Nothing Then
  W "  CreatePolygon2(0,0,20,5)  -> size " & poly.SizeWidth & " x " & poly.SizeHeight & _
    "  pos " & poly.PositionX & "," & poly.PositionY
End If

On Error Resume Next
Err.Clear
Set poly = lay.CreatePolygon2(0, 0, 5, 20, 10)
If Err.Number = 0 And Not poly Is Nothing Then
  W "  CreatePolygon2(0,0,5,20,10) -> size " & poly.SizeWidth & " x " & poly.SizeHeight
End If

W ""
W "=== polygon / star shape members ==="
On Error Resume Next
Err.Clear
Set poly = lay.CreatePolygon2(0, 0, 5, 20)
If Not poly Is Nothing Then
  Dim pgn
  On Error Resume Next
  Err.Clear
  Set pgn = poly.Polygon
  W "        TypeName(poly.Polygon) = " & TypeName(poly.Polygon)
  T "Set pgn = poly.Polygon", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "        poly.Polygon.Sides = " & poly.Polygon.Sides
  T "poly.Polygon.Sides (read)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  poly.Polygon.Sides = 5
  T "poly.Polygon.Sides = 5", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "        poly.Polygon.Mode = " & poly.Polygon.Mode
  T "poly.Polygon.Mode (read)", Err.Number, Err.Description

  Dim m
  For m = 0 To 4
    On Error Resume Next
    Err.Clear
    poly.Polygon.Mode = m
    If Err.Number = 0 Then
      W "        Polygon.Mode accepts " & m & " (readback " & poly.Polygon.Mode & ")"
    End If
  Next

  On Error Resume Next
  Err.Clear
  poly.Polygon.Sharpness = 53
  T "poly.Polygon.Sharpness = 53", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "        poly.Polygon.Radius = " & poly.Polygon.Radius
  T "poly.Polygon.Radius (read)", Err.Number, Err.Description
End If

W ""
W "=== star via curve nodes ==="
On Error Resume Next
Err.Clear
Set crv = lay.CreateCurve
T "Set crv = lay.CreateCurve", Err.Number, Err.Description
If Not crv Is Nothing Then
  W "        TypeName = " & TypeName(crv)
  On Error Resume Next
  Err.Clear
  Dim sp
  Set sp = crv.CreateSubPath(0, 0)
  T "crv.CreateSubPath(0,0)", Err.Number, Err.Description
  If Not sp Is Nothing Then
    On Error Resume Next
    Err.Clear
    Dim sg
    Set sg = sp.AddNode(10, 10)
    T "sp.AddNode(10,10)", Err.Number, Err.Description
    On Error Resume Next
    Err.Clear
    sp.Closed = True
    T "sp.Closed = True", Err.Number, Err.Description
  End If
  On Error Resume Next
  Err.Clear
  Dim csh
  Set csh = lay.CreateCurveFromCurve(crv)
  T "lay.CreateCurveFromCurve(crv)", Err.Number, Err.Description
End If

W ""
W "=== overprint candidates ==="
On Error Resume Next
Err.Clear
Set rect = lay.CreateRectangle2(0, 0, 20, 20)
T "lay.CreateRectangle2(0,0,20,20)", Err.Number, Err.Description
If Not rect Is Nothing Then
  On Error Resume Next
  Err.Clear
  rect.OverprintFill = True
  T "sh.OverprintFill = True", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.OverprintOutline = True
  T "sh.OverprintOutline = True", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Fill.Overprint = True
  T "sh.Fill.Overprint = True", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Outline.Overprint = True
  T "sh.Outline.Overprint = True", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "        sh.OverprintFill (read) = " & rect.OverprintFill
  T "sh.OverprintFill (read)", Err.Number, Err.Description
End If

W ""
W "=== shaperange on a real ShapeRange ==="
On Error Resume Next
Err.Clear
Set sr = pg.Shapes.All
T "Set sr = pg.Shapes.All", Err.Number, Err.Description
If Not sr Is Nothing Then
  W "        TypeName(sr) = " & TypeName(sr)
  On Error Resume Next
  Err.Clear
  W "        sr.Count = " & sr.Count
  T "sr.Count", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set s1 = Nothing
  Set s1 = sr(1)
  T "sr(1)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set s1 = Nothing
  Set s1 = sr.Item(1)
  T "sr.Item(1)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set s1 = Nothing
  Set s1 = sr.Shapes(1)
  T "sr.Shapes(1)", Err.Number, Err.Description
End If

W ""
W "=== ActiveSelection identity ==="
On Error Resume Next
Err.Clear
rect.AddToSelection
T "rect.AddToSelection", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set sr = app.ActiveSelection
W "        TypeName(app.ActiveSelection) = " & TypeName(app.ActiveSelection)
T "TypeName(app.ActiveSelection)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
W "        TypeName(app.ActiveSelection.Shapes) = " & TypeName(app.ActiveSelection.Shapes)
T "TypeName(app.ActiveSelection.Shapes)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
W "        app.ActiveSelection.Shapes.Count = " & app.ActiveSelection.Shapes.Count
T "app.ActiveSelection.Shapes.Count", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Dim o
Set o = Nothing
Set o = app.ActiveSelection.Shapes(1)
T "app.ActiveSelection.Shapes(1)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.ClearSelection
T "doc.ClearSelection", Err.Number, Err.Description

W ""
W "=== blend (object on path) ==="
On Error Resume Next
Err.Clear
Dim r2, bl
Set r2 = lay.CreateRectangle2(30, 0, 20, 20)
Set bl = Nothing
Set bl = rect.CreateBlend(r2)
T "rect.CreateBlend(r2)", Err.Number, Err.Description
If Not bl Is Nothing Then
  W "        TypeName(blend) = " & TypeName(bl)
  On Error Resume Next
  Err.Clear
  W "        blend.Blend present? " & TypeName(bl.Blend)
  T "bl.Blend", Err.Number, Err.Description
End If

W ""
W "=== document misc ==="
On Error Resume Next
Err.Clear
W "        doc.Unit = " & doc.Unit
T "doc.Unit (read)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.Dirty = False
T "doc.Dirty = False", Err.Number, Err.Description

On Error Resume Next
Err.Clear
W "        doc.Dirty (read) = " & doc.Dirty
T "doc.Dirty (read)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
W "        doc.Name = " & doc.Name
T "doc.Name", Err.Number, Err.Description

On Error Resume Next
Err.Clear
pg.Activate
T "pg.Activate", Err.Number, Err.Description

W ""
W "=== export ==="
On Error Resume Next
Err.Clear
Dim opt
Set opt = app.CreateStructExportOptions
T "app.CreateStructExportOptions", Err.Number, Err.Description
If Not opt Is Nothing Then
  W "        TypeName(opt) = " & TypeName(opt)
  On Error Resume Next
  Err.Clear
  opt.ImageType = 5
  T "opt.ImageType = 5", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  opt.ResolutionX = 96
  opt.ResolutionY = 96
  T "opt.ResolutionX/Y = 96", Err.Number, Err.Description
End If

On Error Resume Next
Err.Clear
Dim tmp
tmp = fso.GetSpecialFolder(2) & "\cdrx4_probe_test.jpg"
If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
doc.ExportEx tmp, 8, 1, opt
T "doc.ExportEx tmp, 8, 1, opt   (8=JPEG?, 1=?)", Err.Number, Err.Description
W "        file created? " & fso.FileExists(tmp)

On Error Resume Next
Err.Clear
If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
doc.ExportEx tmp, 8, 2, opt
T "doc.ExportEx tmp, 8, 2, opt", Err.Number, Err.Description
W "        file created? " & fso.FileExists(tmp)

On Error Resume Next
Err.Clear
If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
doc.ExportEx tmp, 8, 3, opt
T "doc.ExportEx tmp, 8, 3, opt", Err.Number, Err.Description
W "        file created? " & fso.FileExists(tmp)

W ""
W "=== toolbar / plugin commands (read-only checks) ==="
On Error Resume Next
Err.Clear
W "        app.CommandBars.Count = " & app.CommandBars.Count
T "app.CommandBars.Count", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Dim cb
Set cb = Nothing
Set cb = app.CommandBars("增强工具")
T "app.CommandBars(""增强工具"")", Err.Number, Err.Description
If Not cb Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "        controls = " & cb.Controls.Count & "  visible = " & cb.Visible
  T "cb.Controls.Count", Err.Number, Err.Description
  For i = 1 To cb.Controls.Count
    On Error Resume Next
    Err.Clear
    W "          [" & i & "] caption = [" & cb.Controls(i).Caption & "]  type = " & cb.Controls(i).Type
  Next
End If

On Error Resume Next
Err.Clear
app.AddPluginCommand "CDRX4Toolkit.M_Util.HasDocument", "测试按钮", "测试提示"
T "app.AddPluginCommand (application level)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Dim gms
Set gms = app.GMSManager
W "        TypeName(GMSManager) = " & TypeName(gms)
On Error Resume Next
Err.Clear
W "        GMSManager.Count = " & gms.Count
T "GMSManager.Count", Err.Number, Err.Description

W ""
W "done: OK=" & nOK & "  FAIL=" & nFail

On Error Resume Next
doc.Close
logFile.Close