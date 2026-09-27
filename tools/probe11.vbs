Option Explicit

'==========================================================
'  CorelDRAW X4 probe #11
'  ellipse size semantics, polygon arg order, duplicate,
'  export filter forms, gms manager, star glyph.
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe11.log", True)

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
Dim a, b, o, opt, tmpJpg, t
Dim i

W "probe11 start " & Now
On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
Set doc = app.CreateDocument
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer
W "  boot err " & Err.Number

'----------------------------------------------------------
W ""
W "=== A. ellipse size semantics ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreateEllipse2(0, 0, 20, 20)
Chk "CreateEllipse2(0,0,20,20)", Err.Number, Err.Description
If Not o Is Nothing Then
  W "  SizeWidth = " & o.SizeWidth & "  SizeHeight = " & o.SizeHeight
  W "  LeftX = " & o.LeftX & "  BottomY = " & o.BottomY
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreateEllipse2(0, 0, 20, 10)
Chk "CreateEllipse2(0,0,20,10)", Err.Number, Err.Description
If Not o Is Nothing Then
  W "  SizeWidth = " & o.SizeWidth & "  SizeHeight = " & o.SizeHeight
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreateRectangle2(0, 0, 20, 10)
Chk "CreateRectangle2(0,0,20,10)", Err.Number, Err.Description
If Not o Is Nothing Then
  W "  SizeWidth = " & o.SizeWidth & "  SizeHeight = " & o.SizeHeight
End If

'----------------------------------------------------------
W ""
W "=== B. polygon arg order ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreatePolygon(0, 200, 20, 53, 5)
Chk "CreatePolygon(0,200,20,53,5)", Err.Number, Err.Description
If Not o Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "  Type = " & o.Type & "  sides = " & o.Polygon.Sides & _
    "  sharpness = " & o.Polygon.Sharpness & "  [" & Err.Number & "] " & Err.Description
  W "  bbox = " & o.SizeWidth & " x " & o.SizeHeight
  On Error Resume Next
  Err.Clear
  o.ConvertToCurves
  W "  nodes after convert = " & o.Curve.Nodes.Count & "  [" & Err.Number & "]"
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreatePolygon(0, 260, 20, 0, 5)
Chk "CreatePolygon(0,260,20,0,5)", Err.Number, Err.Description
If Not o Is Nothing Then
  On Error Resume Next
  W "  sides = " & o.Polygon.Sides & "  sharpness = " & o.Polygon.Sharpness
  W "  bbox = " & o.SizeWidth & " x " & o.SizeHeight
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreatePolygon(0, 320, 20, 100, 5)
Chk "CreatePolygon(0,320,20,100,5)", Err.Number, Err.Description
If Not o Is Nothing Then
  On Error Resume Next
  W "  sides = " & o.Polygon.Sides & "  sharpness = " & o.Polygon.Sharpness
  W "  bbox = " & o.SizeWidth & " x " & o.SizeHeight
End If

'----------------------------------------------------------
W ""
W "=== C. duplicate ==="
On Error Resume Next
Err.Clear
Set a = Nothing
Set a = lay.CreateRectangle2(0, 400, 11, 11)
Set b = Nothing
Set b = a.Duplicate(20, 20)
Chk "a.Duplicate(20,20)", Err.Number, Err.Description
W "  TypeName = " & TypeName(b)
If Not b Is Nothing Then
  W "  dup LeftX = " & b.LeftX & "  BottomY = " & b.BottomY
End If

On Error Resume Next
Err.Clear
Set b = Nothing
Set b = a.Duplicate
Chk "a.Duplicate", Err.Number, Err.Description
W "  TypeName = " & TypeName(b)

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.Clone
Chk "a.Clone", Err.Number, Err.Description
W "  TypeName = " & TypeName(o)

'----------------------------------------------------------
W ""
W "=== D. export filter forms ==="
tmpJpg = root & "\_probe11_out.jpg"
On Error Resume Next
Err.Clear
Set opt = Nothing
Set opt = app.CreateStructExportOptions
opt.ImageType = 4
opt.ResolutionX = 300
opt.ResolutionY = 300

On Error Resume Next
Err.Clear
doc.Export tmpJpg, "JPEG"
Chk "doc.Export file,""JPEG""", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
doc.Export tmpJpg, "JPG - JPEG Bitmaps"
Chk "doc.Export file,JPG name", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
doc.ExportEx tmpJpg, "JPG", 1, opt
Chk "doc.ExportEx file,""JPG"",1,opt", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.CreateStructExportOptions
Chk "doc.CreateStructExportOptions", Err.Number, Err.Description

On Error Resume Next
Err.Clear
W "  doc.Export params probe: " & doc.Export
Chk "read doc.Export", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== E. gms manager ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.GMSManager
If Not o Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "  Macros -> [" & Err.Number & "] " & Err.Description & "  TypeName " & TypeName(o.Macros)
  On Error Resume Next
  Err.Clear
  W "  Macros.Count = " & o.Macros.Count & "  [" & Err.Number & "] " & Err.Description
  On Error Resume Next
  Err.Clear
  o.RunMacro "M_Util", "HasDocument"
  Chk "GMSManager.RunMacro M_Util.HasDocument", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  o.RunMacro "CDRX4Toolkit.M_Util", "HasDocument"
  Chk "GMSManager.RunMacro CDRX4Toolkit.M_Util.HasDocument", Err.Number, Err.Description
End If

'----------------------------------------------------------
W ""
W "=== F. star glyph ==="
On Error Resume Next
Err.Clear
Set t = Nothing
Set t = lay.CreateArtisticTextWide(0, 480, ChrW(9733))
Chk "star glyph create", Err.Number, Err.Description
If Not t Is Nothing Then
  On Error Resume Next
  Err.Clear
  t.Text.Story.Font = "Arial"
  t.Text.Story.Size = 72
  W "  after size 72: bbox " & t.SizeWidth & " x " & t.SizeHeight & "  [" & Err.Number & "]"
  On Error Resume Next
  Err.Clear
  t.Text.Story.Alignment = 2
  W "  align 2 err " & Err.Number
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreateEllipse2(0, 560, 30, 30)
Set t = Nothing
Set t = lay.CreateArtisticTextWide(0, 560, "RING TEXT TEST")
t.Text.Story.Size = 14
t.Text.Story.Alignment = 2
t.Text.FitToPath o
Chk "ring text fit", Err.Number, Err.Description
If Not t Is Nothing Then
  W "  ring text bbox = " & t.SizeWidth & " x " & t.SizeHeight
  On Error Resume Next
  W "  Type = " & t.Type
End If

W ""
W "done"

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close