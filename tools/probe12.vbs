Option Explicit

'==========================================================
'  CorelDRAW X4 probe #12
'  document unit table, export arg order, macro runner,
'  polygon geometry, overprint on cmyk fill.
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe12.log", True)

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
Dim i, k, oldUnit

W "probe12 start " & Now
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
W "=== A. document unit table ==="
On Error Resume Next
Err.Clear
W "  default doc.Unit = " & doc.Unit & "  [" & Err.Number & "] " & Err.Description
W "  default page W = " & pg.SizeWidth & "  H = " & pg.SizeHeight

oldUnit = 0
On Error Resume Next
oldUnit = doc.Unit
For k = 1 To 12
  On Error Resume Next
  Err.Clear
  doc.Unit = k
  If Err.Number = 0 Then
    W "  unit " & k & " -> page W = " & pg.SizeWidth & "  H = " & pg.SizeHeight
  Else
    W "  unit " & k & " -> [" & Err.Number & "] " & Err.Description
  End If
Next
On Error Resume Next
doc.Unit = oldUnit

'----------------------------------------------------------
W ""
W "=== B. export arg order ==="
tmpJpg = root & "\_probe12_out.jpg"
On Error Resume Next
Err.Clear
Set opt = Nothing
Set opt = app.CreateStructExportOptions
opt.ImageType = 4
opt.ResolutionX = 300
opt.ResolutionY = 300
W "  opt err " & Err.Number

On Error Resume Next
Err.Clear
doc.Export 3, tmpJpg
Chk "doc.Export 3,file", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
doc.Export 3, tmpJpg, 1, opt
Chk "doc.Export 3,file,1,opt", Err.Number, Err.Description
W "        exists = " & fso.FileExists(tmpJpg)

On Error Resume Next
Err.Clear
doc.ExportEx 3, tmpJpg, 1, opt
Chk "doc.ExportEx 3,file,1,opt", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.GetExportFilter
Chk "doc.GetExportFilter", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.ExportFilter
W "  doc.ExportFilter = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.ExportFilter
W "  app.ExportFilter = " & TypeName(o) & "  [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.CreateStructExportOptions
W "  TypeName(opt) = " & TypeName(o)

'----------------------------------------------------------
W ""
W "=== C. macro runner ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = app.GMSManager
If Not o Is Nothing Then
  On Error Resume Next
  Err.Clear
  o.RunMacro "M_Util.HasDocument"
  Chk "RunMacro ""M_Util.HasDocument""", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  o.RunMacro "CDRX4Toolkit.M_Util.HasDocument"
  Chk "RunMacro ""CDRX4Toolkit.M_Util.HasDocument""", Err.Number, Err.Description
  On Error Resume Next
  Err.Clear
  o.RunMacro "M_Install.InstallToolbarSilent"
  Chk "RunMacro ""M_Install.InstallToolbarSilent""", Err.Number, Err.Description
End If

'----------------------------------------------------------
W ""
W "=== D. polygon geometry ==="
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreatePolygon(0, 0, 20, 53, 5)
Chk "CreatePolygon(0,0,20,53,5)", Err.Number, Err.Description
If Not o Is Nothing Then
  W "  bbox before convert = " & o.SizeWidth & " x " & o.SizeHeight
  On Error Resume Next
  Err.Clear
  o.ConvertToCurves
  W "  nodes = " & o.Curve.Nodes.Count & "  [" & Err.Number & "]"
  For i = 1 To o.Curve.Nodes.Count
    W "    node " & i & " = " & o.Curve.Nodes(i).PositionX & "," & _
      o.Curve.Nodes(i).PositionY
  Next
  W "  bbox after convert = " & o.SizeWidth & " x " & o.SizeHeight
End If

'----------------------------------------------------------
W ""
W "=== E. overprint on cmyk fill ==="
On Error Resume Next
Err.Clear
Set a = Nothing
Set a = lay.CreateRectangle2(0, 100, 11, 11)
a.Fill.ApplyUniformFill app.CreateCMYKColor(0, 0, 0, 100)
Chk "cmyk black fill", Err.Number, Err.Description
On Error Resume Next
Err.Clear
a.OverprintFill = True
Chk "OverprintFill = True", Err.Number, Err.Description
On Error Resume Next
W "  readback = " & a.OverprintFill
On Error Resume Next
Err.Clear
Set o = Nothing
Set o = a.Fill.UniformColor
W "  color type = " & o.Type & "  K = " & o.CMYKBlack & "  [" & Err.Number & "] " & Err.Description
On Error Resume Next
Err.Clear
W "  doc.PrintSettings.OverprintBlack? " & doc.PrintSettings
Chk "read doc.PrintSettings", Err.Number, Err.Description

W ""
W "done"

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close