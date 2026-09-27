Option Explicit

'==========================================================
'  CorelDRAW X4 probe #9
'  selection / shaperange item, blend, export, overprint,
'  polygon-star, text-on-path.  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe9.log", True)

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

Dim app, doc, pg, lay
Dim a, b, o, sr, opt, t, p, pc
Dim i, tmpJpg

W "probe9 start " & Now
On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
Set doc = app.CreateDocument
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer
W "  boot err " & Err.Number

On Error Resume Next
Err.Clear
Set a = Nothing
Set a = lay.CreateRectangle2(0, 0, 11, 11)
Set b = Nothing
Set b = lay.CreateRectangle2(0, 40, 22, 22)
W "  two rects err " & Err.Number

'----------------------------------------------------------
W ""
W "=== A. selection ==="
On Error Resume Next
Err.Clear
doc.ClearSelection
Chk "doc.ClearSelection", Err.Number, Err.Description

On Error Resume Next
Err.Clear
a.AddToSelection
Chk "a.AddToSelection", Err.Number, Err.Description

On Error Resume Next
Err.Clear
b.AddToSelection
Chk "b.AddToSelection", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set sr = Nothing
Set sr = app.ActiveSelection
Chk "Set sr = app.ActiveSelection", Err.Number, Err.Description
W "  TypeName(sr) = " & TypeName(sr)

On Error Resume Next
Err.Clear
W "  sr.Count = " & sr.Count & "   [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr.Item(1)
W "  TypeName(sr.Item(1)) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr(1)
W "  TypeName(sr(1)) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr.Shapes(1)
W "  TypeName(sr.Shapes(1)) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr.Shapes.All
W "  TypeName(sr.Shapes.All) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.ActiveSelectionRange
W "  TypeName(app.ActiveSelectionRange) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.ActiveDocument.SelectionRange
W "  TypeName(doc.SelectionRange) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr.Shapes.Item(1)
W "  TypeName(sr.Shapes.Item(1)) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
W "  sr.Count via For Each:"
For Each o In sr
  W "        item TypeName = " & TypeName(o) & "  Type = " & o.Type
Next

'----------------------------------------------------------
W ""
W "=== B. blend ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.CreateBlend(b)
Chk "a.CreateBlend(b)", Err.Number, Err.Description
W "  TypeName = " & TypeName(o)

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.CreateBlend(a, b)
Chk "doc.CreateBlend(a,b)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.CreateBlend(b, 0)
Chk "a.CreateBlend(b,0)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.Effects.CreateBlend(a, b)
Chk "a.Effects.CreateBlend(a,b)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.CreateBlend(a, b)
Chk "app.CreateBlend(a,b)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr.CreateBlend(a, b)
Chk "sr.CreateBlend(a,b)", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== C. overprint ==="
On Error Resume Next
Err.Clear
a.OverprintFill = True
Chk "a.OverprintFill = True", Err.Number, Err.Description
On Error Resume Next
W "        reads " & a.OverprintFill

On Error Resume Next
Err.Clear
a.OverprintOutline = True
Chk "a.OverprintOutline = True", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.Overprint
W "  TypeName(a.Overprint) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.Fill.Overprint
W "  TypeName(a.Fill.Overprint) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.PrintSettings
W "  TypeName(doc.PrintSettings) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

'----------------------------------------------------------
W ""
W "=== D. export ==="
tmpJpg = root & "\_probe9_out.jpg"

On Error Resume Next
Err.Clear
Set opt = Nothing
Set opt = app.CreateStructExportOptions
Chk "CreateStructExportOptions", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.Export tmpJpg, 3
Chk "doc.Export file,3", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
doc.ExportEx tmpJpg, 3, 1
Chk "doc.ExportEx file,3,1", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
doc.ExportEx tmpJpg, 3, 1, opt
Chk "doc.ExportEx file,3,1,opt", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.ExportEx tmpJpg, 3, 1, Nothing
Chk "doc.ExportEx file,3,1,Nothing", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.ExportEx tmpJpg, 3
Chk "doc.ExportEx file,3", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.Export tmpJpg, 3, 1, opt
Chk "doc.Export file,3,1,opt", Err.Number, Err.Description

On Error Resume Next
Err.Clear
opt.ImageType = 4
opt.ResolutionX = 300
opt.ResolutionY = 300
doc.Export tmpJpg, 3, 1, opt
Chk "doc.Export file,3,1,opt(300dpi)", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
W "  opt.ImageType read = " & opt.ImageType
W "  opt.ResolutionX read = " & opt.ResolutionX

'----------------------------------------------------------
W ""
W "=== E. polygon / star ==="
On Error Resume Next
Err.Clear
Set p = Nothing
Set p = lay.CreatePolygon(0, 200, 20, 5, 53)
Chk "CreatePolygon(x,y,r,5,53)", Err.Number, Err.Description
W "  TypeName = " & TypeName(p)
If Not p Is Nothing Then
  On Error Resume Next
  W "  p.Type = " & p.Type & "  [" & Err.Number & "]"
  On Error Resume Next
  Err.Clear
  W "  p.Polygon.Sides = " & p.Polygon.Sides & "  [" & Err.Number & "] " & Err.Description
  On Error Resume Next
  Err.Clear
  W "  p.Polygon.Sharpness = " & p.Polygon.Sharpness & "  [" & Err.Number & "] " & Err.Description
  On Error Resume Next
  Err.Clear
  W "  bbox W = " & p.SizeWidth & "  H = " & p.SizeHeight
  On Error Resume Next
  Err.Clear
  p.SetPolygonProperties 5, 53
  Chk "SetPolygonProperties 5,53", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  p.ConvertToCurves
  W "  after convert: nodes = " & p.Curve.Nodes.Count & "  [" & Err.Number & "] " & Err.Description
End If

On Error Resume Next
Err.Clear
Set p = Nothing
Set p = lay.CreatePolygon(0, 250, 20, 5)
Chk "CreatePolygon(x,y,r,5)", Err.Number, Err.Description
W "  TypeName = " & TypeName(p)

On Error Resume Next
Err.Clear
Set p = Nothing
Set p = lay.CreatePolygon(0, 300, 20, 5, 0)
Chk "CreatePolygon(x,y,r,5,0)", Err.Number, Err.Description
If Not p Is Nothing Then
  On Error Resume Next
  W "  TypeName = " & TypeName(p) & "  Type = " & p.Type
  On Error Resume Next
  Err.Clear
  W "  sharpness = " & p.Polygon.Sharpness & "  [" & Err.Number & "] " & Err.Description
End If

On Error Resume Next
Err.Clear
Set p = Nothing
Set p = lay.CreateStar(0, 350, 20, 5, 1)
Chk "CreateStar(x,y,r,5,1)", Err.Number, Err.Description
W "  TypeName = " & TypeName(p)

'----------------------------------------------------------
W ""
W "=== F. text on path ==="
On Error Resume Next
Err.Clear
Set t = Nothing
Set t = lay.CreateArtisticTextWide(0, 420, "RING TEXT")
Chk "CreateArtisticTextWide", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreateEllipse2(0, 470, 20, 20)
t.Text.FitToPath o
Chk "t.Text.FitToPath(ellipse)", Err.Number, Err.Description
W "  t.Type = " & t.Type
On Error Resume Next
Err.Clear
W "  after fit: t.Type = " & t.Type & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreateArtisticTextWide(0, 520, ChrW(9733))
Chk "CreateArtisticTextWide star glyph", Err.Number, Err.Description
If Not o Is Nothing Then
  On Error Resume Next
  o.Text.Story.Font = "Arial"
  o.Text.Story.Size = 72
  W "  star glyph shape Type = " & o.Type
End If

W ""
W "done"

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close