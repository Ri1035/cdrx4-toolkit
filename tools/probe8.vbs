Option Explicit

'==========================================================
'  CorelDRAW X4 probe #8
'  Color model / rectangle corners / shaperange item /
'  blend / powerclip / export options / text + fitpath
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe8.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

' log one statement result
Sub Chk(label, en, ed)
  If en = 0 Then
    W "  OK    " & label
  Else
    W "  FAIL  " & label & "   [" & en & "] " & ed
  End If
End Sub

Dim app, doc, pg, lay
Dim r, r2, ro, sr, c1, c2, c3, t, p, bl, pc, opt, o
Dim s

W "probe8 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
W "  app -> [" & Err.Number & "] " & Err.Description
app.InitializeVBA
Set doc = app.CreateDocument
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer
W "  doc/pg/lay -> [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
W "  app.Version = " & app.Version
W "  pg.SizeWidth = " & pg.SizeWidth & "  SizeHeight = " & pg.SizeHeight
W "  doc.Pages.Count = " & doc.Pages.Count
W "  pg.Shapes.Count = " & pg.Shapes.Count
W "  TypeName(pg.Shapes.All) = " & TypeName(pg.Shapes.All)

'----------------------------------------------------------
W ""
W "=== A. Color object ==="
On Error Resume Next
Err.Clear
Set r = Nothing
Set r = lay.CreateRectangle2(0, 0, 20, 20)
W "  rect -> [" & Err.Number & "] " & Err.Description & "  TypeName " & TypeName(r)

On Error Resume Next
Err.Clear
r.Fill.ApplyUniformFill app.CreateRGBColor(200, 30, 40)
W "  ApplyUniformFill RGB(200,30,40) -> [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
W "  Fill.Type = " & r.Fill.Type

On Error Resume Next
Err.Clear
Set c1 = Nothing
Set c1 = r.Fill.UniformColor
W "  Fill.UniformColor -> [" & Err.Number & "] " & Err.Description & _
  "  TypeName " & TypeName(c1)

If Not c1 Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "  c1.Type = " & c1.Type & "  [" & Err.Number & "]"
  On Error Resume Next
  Err.Clear
  W "  c1.RGBRed = " & c1.RGBRed & " [" & Err.Number & "] " & Err.Description
  W "  c1.RGBGreen = " & c1.RGBGreen & " [" & Err.Number & "]"
  W "  c1.RGBBlue = " & c1.RGBBlue & " [" & Err.Number & "]"

  On Error Resume Next
  Err.Clear
  W "  c1.CMYKCyan = " & c1.CMYKCyan & " [" & Err.Number & "] " & Err.Description
  W "  c1.CMYKMagenta = " & c1.CMYKMagenta & " [" & Err.Number & "]"
  W "  c1.CMYKYellow = " & c1.CMYKYellow & " [" & Err.Number & "]"
  W "  c1.CMYKBlack = " & c1.CMYKBlack & " [" & Err.Number & "]"

  On Error Resume Next
  Err.Clear
  Set c2 = Nothing
  Set c2 = c1.ConvertToCMYK
  Chk "c1.ConvertToCMYK (as function)", Err.Number, Err.Description
  W "        TypeName = " & TypeName(c2)

  On Error Resume Next
  Err.Clear
  Set c2 = Nothing
  Set c2 = c1.ConvertToRGB
  Chk "c1.ConvertToRGB (as function)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set c2 = Nothing
  Set c2 = c1.ConvertToGray
  Chk "c1.ConvertToGray (as function)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  c1.CMYKAssign 10, 20, 30, 40
  Chk "c1.CMYKAssign 10,20,30,40", Err.Number, Err.Description
  On Error Resume Next
  W "        after: C" & c1.CMYKCyan & " M" & c1.CMYKMagenta & _
    " Y" & c1.CMYKYellow & " K" & c1.CMYKBlack & "  Type=" & c1.Type

  On Error Resume Next
  Err.Clear
  c1.RGBAssign 11, 22, 33
  Chk "c1.RGBAssign 11,22,33", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "  c1.IsSame(c1) = " & c1.IsSame(c1) & " [" & Err.Number & "] " & Err.Description

  On Error Resume Next
  Err.Clear
  Set c3 = Nothing
  Set c3 = app.CreateCMYKColor(0, 0, 0, 100)
  Chk "app.CreateCMYKColor(0,0,0,100)", Err.Number, Err.Description
  W "        TypeName = " & TypeName(c3)
  If Not c3 Is Nothing Then
    On Error Resume Next
    W "        blk CMYK = " & c3.CMYKCyan & "," & c3.CMYKMagenta & "," & _
      c3.CMYKYellow & "," & c3.CMYKBlack & "  Type=" & c3.Type
    On Error Resume Next
    Err.Clear
    W "  c1.IsSame(blk) = " & c1.IsSame(c3) & " [" & Err.Number & "] " & Err.Description
  End If

  On Error Resume Next
  Err.Clear
  Set c2 = Nothing
  Set c2 = app.CreateCMYKColor(0, 0, 0, 100)
  c2.CopyAssign c1
  Chk "c2.CopyAssign c1", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "  TypeName(c1.GetRGB) = " & TypeName(c1.GetRGB)
  Chk "c1.GetRGB", Err.Number, Err.Description
End If

'----------------------------------------------------------
W ""
W "=== B. Fill / Outline write ==="
On Error Resume Next
Err.Clear
r.Fill.OverprintFill = True
Chk "Fill.OverprintFill = True", Err.Number, Err.Description
On Error Resume Next
W "        reads " & r.Fill.OverprintFill

On Error Resume Next
Err.Clear
r.Fill.ApplyNoFill
Chk "Fill.ApplyNoFill", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = r.Outline.Color
Chk "r.Outline.Color (read)", Err.Number, Err.Description
W "        TypeName = " & TypeName(o)

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.CreateRGBColor(255, 0, 0)
r.Outline.Color.CopyAssign o
Chk "r.Outline.Color.CopyAssign red", Err.Number, Err.Description

On Error Resume Next
Err.Clear
r.Outline.Width = 1.2
Chk "r.Outline.Width = 1.2", Err.Number, Err.Description
On Error Resume Next
W "        reads " & r.Outline.Width

'----------------------------------------------------------
W ""
W "=== C. Rectangle corner properties ==="
On Error Resume Next
Err.Clear
Set r2 = Nothing
Set r2 = lay.CreateRectangle2(100, 0, 20, 20)
Set ro = Nothing
Set ro = r2.Rectangle
Chk "Set ro = r.Rectangle", Err.Number, Err.Description
W "  TypeName(ro) = " & TypeName(ro)

If Not ro Is Nothing Then
  Dim props, i
  props = Array("EqualCorners", "Radius", "RelativeCorners", _
                "CornerUpperLeftRadius", "CornerUpperRightRadius", _
                "CornerLowerLeftRadius", "CornerLowerRightRadius", _
                "RadiusUpperLeft", "RadiusUpperRight", _
                "RadiusLowerLeft", "RadiusLowerRight")

  ' read pass
  For i = 0 To UBound(props)
    On Error Resume Next
    Err.Clear
    W "  read  ro." & props(i) & " = " & Eval("ro." & props(i)) & _
      "  [" & Err.Number & "] " & Err.Description
  Next

  ' write pass
  For i = 0 To UBound(props)
    On Error Resume Next
    Err.Clear
    Execute "ro." & props(i) & " = 1"
    W "  write ro." & props(i) & " = 1  -> [" & Err.Number & "] " & Err.Description
  Next
End If

'----------------------------------------------------------
W ""
W "=== D. ShapeRange item semantics ==="
On Error Resume Next
Err.Clear
Set r = Nothing
Set r = lay.CreateRectangle2(0, 100, 11, 11)
Set r2 = Nothing
Set r2 = lay.CreateRectangle2(0, 140, 22, 22)
app.ActiveDocument.ClearSelection
r.AddToSelection
r2.AddToSelection
Set sr = Nothing
Set sr = app.ActiveSelection
Chk "Set sr = app.ActiveSelection", Err.Number, Err.Description
W "  TypeName(sr) = " & TypeName(sr) & "  Count = " & sr.Count

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr(1)
W "  TypeName(sr(1)) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr.Item(1)
W "  TypeName(sr.Item(1)) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr(1).Fill
Chk "sr(1).Fill", Err.Number, Err.Description
W "        TypeName = " & TypeName(o)

On Error Resume Next
Err.Clear
W "  sr(1).Type = " & sr(1).Type & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
W "  sr(1).SizeWidth = " & sr(1).SizeWidth & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
W "  sr(1).Shapes.Count = " & sr(1).Shapes.Count & "  [" & Err.Number & "] " & Err.Description

'----------------------------------------------------------
W ""
W "=== E. Blend ==="
On Error Resume Next
Err.Clear
Set bl = Nothing
Set bl = r.CreateBlend(r2)
Chk "r.CreateBlend(r2)", Err.Number, Err.Description
W "  TypeName(bl) = " & TypeName(bl)

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = r.Effects
W "  TypeName(r.Effects) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = r.Effect
W "  TypeName(r.Effect) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

'----------------------------------------------------------
W ""
W "=== F. PowerClip ==="
On Error Resume Next
Err.Clear
Set pc = Nothing
Set pc = r.PowerClip
Chk "Set pc = r.PowerClip", Err.Number, Err.Description
W "  TypeName(pc) = " & TypeName(pc)

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = r.PowerClip.Shapes.All
Chk "r.PowerClip.Shapes.All", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== G. export options / ExportEx ==="
On Error Resume Next
Err.Clear
Set opt = Nothing
Set opt = app.CreateStructExportOptions
Chk "app.CreateStructExportOptions", Err.Number, Err.Description
W "  TypeName(opt) = " & TypeName(opt)

If Not opt Is Nothing Then
  On Error Resume Next
  Err.Clear
  opt.ImageType = 4
  Chk "opt.ImageType = 4", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  opt.ResolutionX = 300
  Chk "opt.ResolutionX = 300", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  opt.ResolutionY = 300
  Chk "opt.ResolutionY = 300", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  opt.AntiAliasing = 2
  Chk "opt.AntiAliasing = 2", Err.Number, Err.Description
End If

Dim tmpJpg
tmpJpg = root & "\_probe8_out.jpg"
On Error Resume Next
Err.Clear
doc.ExportEx tmpJpg, 3, 1, opt
Chk "doc.ExportEx jpg,3,1,opt", Err.Number, Err.Description
W "        file exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
Set opt = Nothing
Set opt = app.CreateStructExportOptions
opt.ImageType = 4
doc.ExportEx tmpJpg, 3, 1, opt
Chk "ExportEx with ImageType=4", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== H. artistic text + fit to path ==="
On Error Resume Next
Err.Clear
Set t = Nothing
Set t = lay.CreateArtisticTextWide(50, 200, "ABCDEF")
Chk "CreateArtisticTextWide", Err.Number, Err.Description
W "  TypeName(t) = " & TypeName(t)

If Not t Is Nothing Then
  On Error Resume Next
  Err.Clear
  t.Text.Story.Font = "Arial"
  Chk "t.Text.Story.Font = Arial", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  t.Text.Story.Size = 12
  Chk "t.Text.Story.Size = 12", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  t.Text.Story.Alignment = 2
  Chk "t.Text.Story.Alignment = 2", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  Set o = Nothing
  Set o = t.Text.Story
  Chk "t.Text.Story", Err.Number, Err.Description
  W "        TypeName = " & TypeName(o)
  On Error Resume Next
  Err.Clear
  t.FitToPath r
  Chk "t.FitToPath(r)", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  t.Text.FitToPath r
  Chk "t.Text.FitToPath(r)", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  W "  t.Text.Story.Text = " & t.Text.Story.Text
End If

'----------------------------------------------------------
W ""
W "=== I. polygon / star ==="
On Error Resume Next
Err.Clear
Set p = Nothing
Set p = lay.CreatePolygon(0, 250, 20, 5)
Chk "lay.CreatePolygon(0,250,20,5)", Err.Number, Err.Description
W "  TypeName(p) = " & TypeName(p)
If Not p Is Nothing Then
  On Error Resume Next
  Err.Clear
  p.SetPolygonProperties 5, 53
  Chk "p.SetPolygonProperties 5,53", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  W "  p.Polygon.Sharpness = " & p.Polygon.Sharpness & " [" & Err.Number & "]"
  On Error Resume Next
  Err.Clear
  p.ConvertToCurves
  Chk "star ConvertToCurves", Err.Number, Err.Description
  W "  nodes = " & p.Curve.Nodes.Count
End If

'----------------------------------------------------------
W ""
W "=== J. doc misc ==="
On Error Resume Next
Err.Clear
doc.BeginCommandGroup "t"
Chk "doc.BeginCommandGroup", Err.Number, Err.Description
On Error Resume Next
Err.Clear
doc.EndCommandGroup
Chk "doc.EndCommandGroup", Err.Number, Err.Description
On Error Resume Next
Err.Clear
app.Optimization = True
Chk "app.Optimization = True", Err.Number, Err.Description
On Error Resume Next
Err.Clear
app.Refresh
Chk "app.Refresh", Err.Number, Err.Description
On Error Resume Next
Err.Clear
app.ActiveDocument.ClearSelection
Chk "doc.ClearSelection", Err.Number, Err.Description

W ""
W "done"

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close
