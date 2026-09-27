Option Explicit

'==========================================================
'  CorelDRAW X4 probe #10
'  ActiveSelectionRange item, blend with shaperange,
'  export signature, overprint readback, macro runner.
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe10.log", True)

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
Dim a, b, o, sr, srB, opt, tmpJpg, flt, rng
Dim i

W "probe10 start " & Now
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
Set a = lay.CreateRectangle2(0, 0, 11, 11)
Set b = lay.CreateRectangle2(0, 40, 22, 22)
W "  two rects err " & Err.Number

'----------------------------------------------------------
W ""
W "=== A. ActiveSelectionRange ==="
On Error Resume Next
Err.Clear
doc.ClearSelection
a.AddToSelection
b.AddToSelection
Set sr = Nothing
Set sr = app.ActiveSelectionRange
Chk "Set sr = app.ActiveSelectionRange", Err.Number, Err.Description
W "  TypeName(sr) = " & TypeName(sr)

On Error Resume Next
Err.Clear
W "  sr.Count = " & sr.Count & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr(1)
W "  TypeName(sr(1)) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description
If Not o Is Nothing Then
  On Error Resume Next
  W "        Type = " & o.Type & "  w = " & o.SizeWidth
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr.Item(1)
W "  TypeName(sr.Item(1)) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sr.Shapes(1)
W "  TypeName(sr.Shapes(1)) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

' single-shape range for the blend test
On Error Resume Next
Err.Clear
doc.ClearSelection
b.AddToSelection
Set srB = Nothing
Set srB = app.ActiveSelectionRange
Chk "single-shape range", Err.Number, Err.Description
W "  TypeName(srB) = " & TypeName(srB) & "  Count = " & srB.Count

'----------------------------------------------------------
W ""
W "=== B. blend with shaperange ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.CreateBlend(srB)
Chk "a.CreateBlend(srB)", Err.Number, Err.Description
W "  TypeName = " & TypeName(o)

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.CreateBlend(b.Shapes.All)
Chk "a.CreateBlend(b.Shapes.All)", Err.Number, Err.Description
W "  TypeName = " & TypeName(o)

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.CreateBlend(a.Shapes.All)
Chk "a.CreateBlend(a.Shapes.All)", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== C. overprint ==="
On Error Resume Next
Err.Clear
a.OverprintFill = True
Chk "a.OverprintFill = True", Err.Number, Err.Description
On Error Resume Next
W "        readback = " & a.OverprintFill

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.Fill
W "  TypeName(a.Fill) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.Fill.UniformColor
W "  TypeName(a.Fill.UniformColor) = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

'----------------------------------------------------------
W ""
W "=== D. export ==="
tmpJpg = root & "\_probe10_out.jpg"

On Error Resume Next
Err.Clear
Set opt = Nothing
Set opt = app.CreateStructExportOptions
opt.ImageType = 4
opt.ResolutionX = 300
opt.ResolutionY = 300
W "  opt ready err " & Err.Number

flt = 3
rng = 1

On Error Resume Next
Err.Clear
doc.Export tmpJpg
Chk "doc.Export file", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.Export tmpJpg, flt
Chk "doc.Export file,flt", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.Export tmpJpg, flt, rng
Chk "doc.Export file,flt,rng", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.Export tmpJpg, flt, rng, opt
Chk "doc.Export file,flt,rng,opt", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
doc.ExportEx tmpJpg, flt, rng, opt
Chk "doc.ExportEx file,flt,rng,opt", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.ExportEx
Chk "read doc.ExportEx", Err.Number, Err.Description

' try the page-level export
On Error Resume Next
Err.Clear
pg.Export tmpJpg, flt
Chk "pg.Export file,flt", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
pg.ExportEx tmpJpg, flt, opt
Chk "pg.ExportEx file,flt,opt", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

'----------------------------------------------------------
W ""
W "=== E. macro runner ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.GMSManager
Chk "app.GMSManager", Err.Number, Err.Description
W "  TypeName = " & TypeName(o)
If Not o Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "  GMSManager.Macros.Count = " & o.Macros.Count & "  [" & Err.Number & "] " & Err.Description
End If

On Error Resume Next
Err.Clear
W "  app.VBE = " & TypeName(app.VBE) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
W "  app.VBProject count = " & app.VBE.VBProjects.Count & "  [" & Err.Number & "] " & Err.Description

W ""
W "done"

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close