Option Explicit

'==========================================================
'  CorelDRAW X4 probe.
'
'  1. dumps the VBA source of the installed CDRX4Toolkit.gms
'     (so we can see exactly what the user is running)
'  2. exercises every API call the toolkit relies on and logs
'     OK / FAIL per call
'  3. checks whether Application.RunMacro exists
'
'  Late-bound on purpose: nothing here can be a compile error,
'  and no enum constants are used (VBScript has no type
'  library), so every value is derived from a live object.
'
'  ASCII only: Windows Script Host reads .vbs as ANSI.
'==========================================================

Dim fso, here, root, logFile
Dim nOK, nFail

Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe.log", True)

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

'----------------------------------------------------------
' dump one component's source into the log
'----------------------------------------------------------
Sub DumpComp(proj, compName)
  Dim comp, cm
  On Error Resume Next
  Err.Clear
  Set comp = Nothing
  Set comp = proj.VBComponents(compName)
  If Err.Number <> 0 Or comp Is Nothing Then
    W "  -- cannot read component " & compName & " [" & Err.Number & "] " & Err.Description
    Err.Clear
    Exit Sub
  End If
  Set cm = comp.CodeModule
  W "  ===== " & compName & " (" & cm.CountOfLines & " lines) ====="
  If cm.CountOfLines > 0 Then W cm.Lines(1, cm.CountOfLines)
  W "  ===== end " & compName & " ====="
  Err.Clear
End Sub

Dim app, doc, pg, lay
Dim rect, ell, txt, poly, star, s1, sr, opt
Dim col, col2, pgs, pc

W "probe start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
T "CreateObject CorelDRAW.Application.14", Err.Number, Err.Description
If app Is Nothing Then
  W "!! cannot reach CorelDRAW, abort"
  logFile.Close
  WScript.Quit 1
End If

On Error Resume Next
Err.Clear
W "  app.Version = " & app.Version
W "  app.Name    = " & app.Name

On Error Resume Next
Err.Clear
app.InitializeVBA
T "app.InitializeVBA", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== installed VBA projects ==="
Dim vbe, p, i, j, c, cm
Set vbe = Nothing
On Error Resume Next
Err.Clear
Set vbe = app.VBE
T "Set vbe = app.VBE", Err.Number, Err.Description
If Not vbe Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "  VBProjects.Count = " & vbe.VBProjects.Count
  For i = 1 To vbe.VBProjects.Count
    Err.Clear
    Set p = Nothing
    Set p = vbe.VBProjects.Item(i)
    If Err.Number = 0 And Not p Is Nothing Then
      W "  [" & i & "] name=" & p.Name & "  file=" & p.FileName & _
        "  prot=" & p.Protection & "  comps=" & p.VBComponents.Count
      If InStr(1, p.FileName, "CDRX4Toolkit", vbTextCompare) > 0 Then
        W "  --- component list ---"
        For j = 1 To p.VBComponents.Count
          Err.Clear
          Set c = Nothing
          Set c = p.VBComponents.Item(j)
          If Err.Number = 0 And Not c Is Nothing Then
            W "      " & j & ": " & c.Name & "  type=" & c.Type
          End If
        Next
        W ""
        DumpComp p, "M_Rect"
        W ""
        DumpComp p, "M_CMYK"
        W ""
        DumpComp p, "M_Seal"
        W ""
        DumpComp p, "M_Color"
      End If
    Else
      W "  [" & i & "] read failed [" & Err.Number & "] " & Err.Description
    End If
  Next
End If

'----------------------------------------------------------
W ""
W "=== document ==="
On Error Resume Next
Err.Clear
Set doc = app.CreateDocument
T "app.CreateDocument", Err.Number, Err.Description
If doc Is Nothing Then
  On Error Resume Next
  Err.Clear
  Set doc = app.ActiveDocument
  T "app.ActiveDocument (fallback)", Err.Number, Err.Description
End If
If doc Is Nothing Then
  W "!! no document, abort"
  logFile.Close
  WScript.Quit 1
End If

On Error Resume Next
Err.Clear
Set pg = doc.ActivePage
T "doc.ActivePage", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set lay = pg.ActiveLayer
T "pg.ActiveLayer", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== layer Create* methods (values in = left/top/right/bottom) ==="

On Error Resume Next
Err.Clear
Set rect = lay.CreateRectangle2(20, 20, 60, 50)
T "lay.CreateRectangle2(20,20,60,50)", Err.Number, Err.Description
If Not rect Is Nothing Then
  W "        size = " & rect.SizeWidth & " x " & rect.SizeHeight & _
    "   pos = " & rect.PositionX & "," & rect.PositionY & _
    "   l/t/r/b = " & rect.LeftX & "," & rect.TopY & "," & rect.RightX & "," & rect.BottomY & _
    "   type = " & rect.Type
End If

On Error Resume Next
Err.Clear
Set ell = lay.CreateEllipse2(100, 100, 20, 20)
T "lay.CreateEllipse2(100,100,20,20)", Err.Number, Err.Description
If Not ell Is Nothing Then
  W "        size = " & ell.SizeWidth & " x " & ell.SizeHeight & _
    "   pos = " & ell.PositionX & "," & ell.PositionY & _
    "   type = " & ell.Type
End If

On Error Resume Next
Err.Clear
Set txt = lay.CreateArtisticTextWide(50, 200, "hello")
T "lay.CreateArtisticTextWide(50,200,""hello"")", Err.Number, Err.Description
If Not txt Is Nothing Then
  W "        pos = " & txt.PositionX & "," & txt.PositionY & _
    "   type = " & txt.Type
End If

On Error Resume Next
Err.Clear
Set poly = lay.CreatePolygon(150, 150, 5, 20)
T "lay.CreatePolygon(150,150,5,20)", Err.Number, Err.Description
If Not poly Is Nothing Then
  W "        size = " & poly.SizeWidth & " x " & poly.SizeHeight & _
    "   pos = " & poly.PositionX & "," & poly.PositionY & _
    "   type = " & poly.Type
End If

On Error Resume Next
Err.Clear
Set star = lay.CreatePolygon2(150, 150, 5, 20)
T "lay.CreatePolygon2(150,150,5,20)", Err.Number, Err.Description
If Not star Is Nothing Then
  W "        size = " & star.SizeWidth & " x " & star.SizeHeight & _
    "   pos = " & star.PositionX & "," & star.PositionY & _
    "   type = " & star.Type
End If

On Error Resume Next
Err.Clear
Set star = lay.CreateStar(180, 150, 5, 20, 10)
T "lay.CreateStar(180,150,5,20,10)", Err.Number, Err.Description
If Not star Is Nothing Then
  W "        size = " & star.SizeWidth & " x " & star.SizeHeight & _
    "   pos = " & star.PositionX & "," & star.PositionY & _
    "   type = " & star.Type
End If

'----------------------------------------------------------
W ""
W "=== rectangle corner properties (each tested alone) ==="
If Not rect Is Nothing Then
  On Error Resume Next
  Err.Clear
  rect.Rectangle.EqualCorners = True
  T "rect.Rectangle.EqualCorners = True", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Rectangle.Radius = 0
  T "rect.Rectangle.Radius = 0", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Rectangle.RadiusUpperLeft = 0
  T "rect.Rectangle.RadiusUpperLeft = 0", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Rectangle.RadiusUpperRight = 0
  T "rect.Rectangle.RadiusUpperRight = 0", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Rectangle.RadiusLowerLeft = 0
  T "rect.Rectangle.RadiusLowerLeft = 0", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Rectangle.RadiusLowerRight = 0
  T "rect.Rectangle.RadiusLowerRight = 0", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Rectangle.CornerUpperLeft = 0
  T "rect.Rectangle.CornerUpperLeft = 0", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Rectangle.CornerUpperRight = 0
  T "rect.Rectangle.CornerUpperRight = 0", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Rectangle.CornerLowerLeft = 0
  T "rect.Rectangle.CornerLowerLeft = 0", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Rectangle.CornerLowerRight = 0
  T "rect.Rectangle.CornerLowerRight = 0", Err.Number, Err.Description
End If

'----------------------------------------------------------
W ""
W "=== fill / outline / color ==="
If Not rect Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "        rect.Fill.Type = " & rect.Fill.Type
  T "rect.Fill.Type (read)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set col = rect.Fill.UniformColor
  T "Set col = rect.Fill.UniformColor (on fresh rect)", Err.Number, Err.Description
  If Not col Is Nothing Then W "        col.Type = " & col.Type

  On Error Resume Next
  Err.Clear
  Set col = app.CreateRGBColor(255, 0, 0)
  T "app.CreateRGBColor(255,0,0)", Err.Number, Err.Description
  If Not col Is Nothing Then W "        col.Type = " & col.Type

  On Error Resume Next
  Err.Clear
  rect.Fill.ApplyUniformFill col
  T "rect.Fill.ApplyUniformFill col", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "        after apply, rect.Fill.Type = " & rect.Fill.Type
  T "rect.Fill.Type after apply", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Fill.ApplyNoFill
  T "rect.Fill.ApplyNoFill", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set col2 = rect.Outline.Color
  T "Set col2 = rect.Outline.Color", Err.Number, Err.Description
  If Not col2 Is Nothing Then W "        outline col.Type = " & col2.Type

  On Error Resume Next
  Err.Clear
  rect.Outline.Color.CopyAssign col
  T "rect.Outline.Color.CopyAssign col", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Outline.Width = 1.2
  T "rect.Outline.Width = 1.2", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  rect.Fill.OverprintFill = True
  T "rect.Fill.OverprintFill = True", Err.Number, Err.Description
End If

On Error Resume Next
Err.Clear
Set col2 = app.CreateCMYKColor(0, 0, 0, 100)
T "app.CreateCMYKColor(0,0,0,100)", Err.Number, Err.Description
If Not col2 Is Nothing Then W "        cmyk col.Type = " & col2.Type

On Error Resume Next
Err.Clear
Set col = app.CreateRGBColor(255, 0, 0)
Set col2 = Nothing
Set col2 = col.ConvertToCMYK
T "Set col2 = col.ConvertToCMYK", Err.Number, Err.Description
If Not col2 Is Nothing Then W "        converted col.Type = " & col2.Type

On Error Resume Next
Err.Clear
Dim bSame
bSame = col.IsSame(col2)
T "col.IsSame(col2)", Err.Number, Err.Description
W "        IsSame returned " & bSame

On Error Resume Next
Err.Clear
col.ConvertToCMYK
T "col.ConvertToCMYK (statement form)", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== text ==="
If Not txt Is Nothing Then
  On Error Resume Next
  Err.Clear
  txt.Text.Story.Font = "Arial"
  T "txt.Text.Story.Font = ""Arial""", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  txt.Text.Story.Size = 10
  T "txt.Text.Story.Size = 10", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "        Alignment before = " & txt.Text.Story.Alignment
  T "txt.Text.Story.Alignment (read)", Err.Number, Err.Description

  Dim k
  For k = 0 To 7
    On Error Resume Next
    Err.Clear
    txt.Text.Story.Alignment = k
    If Err.Number = 0 Then
      W "        Alignment accepts " & k & " (readback " & txt.Text.Story.Alignment & ")"
    End If
  Next

  On Error Resume Next
  Err.Clear
  txt.SetPosition 50, 200
  T "txt.SetPosition 50,200", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  txt.Text.FitToPath rect
  T "txt.Text.FitToPath rect", Err.Number, Err.Description
End If

'----------------------------------------------------------
W ""
W "=== polygon / star ==="
If Not star Is Nothing Then
  On Error Resume Next
  Err.Clear
  star.SetPolygonProperties 5, 53
  T "star.SetPolygonProperties 5,53", Err.Number, Err.Description
End If

'----------------------------------------------------------
W ""
W "=== selection / shaperange ==="
On Error Resume Next
Err.Clear
rect.AddToSelection
T "rect.AddToSelection", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set sr = app.ActiveSelection
T "Set sr = app.ActiveSelection", Err.Number, Err.Description
If Not sr Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "        sr.Count = " & sr.Count
  T "sr.Count (read)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set s1 = Nothing
  Set s1 = sr(1)
  T "Set s1 = sr(1)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set s1 = Nothing
  Set s1 = sr.Item(1)
  T "Set s1 = sr.Item(1)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set s1 = Nothing
  Set s1 = sr.Shapes(1)
  T "Set s1 = sr.Shapes(1)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set s1 = Nothing
  Set s1 = sr.Shapes.All
  T "Set s1 = sr.Shapes.All", Err.Number, Err.Description
End If

'----------------------------------------------------------
W ""
W "=== misc ==="
On Error Resume Next
Err.Clear
Set pgs = doc.Pages
T "Set pgs = doc.Pages", Err.Number, Err.Description
If Not pgs Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "        doc.Pages.Count = " & doc.Pages.Count
  T "doc.Pages.Count", Err.Number, Err.Description
End If

On Error Resume Next
Err.Clear
Set pc = Nothing
Set pc = rect.PowerClip
T "Set pc = rect.PowerClip (plain rect)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set s1 = Nothing
Set s1 = pg.Shapes.All
T "Set s1 = pg.Shapes.All", Err.Number, Err.Description

On Error Resume Next
Err.Clear
rect.ConvertToCurves
T "rect.ConvertToCurves", Err.Number, Err.Description
If Not rect Is Nothing Then W "        after convert, type = " & rect.Type

On Error Resume Next
Err.Clear
W "        pg.SizeWidth = " & pg.SizeWidth & "  pg.SizeHeight = " & pg.SizeHeight
T "pg.SizeWidth / pg.SizeHeight", Err.Number, Err.Description

On Error Resume Next
Err.Clear
W "        doc.FileName = [" & doc.FileName & "]"
T "doc.FileName (unsaved)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
app.Refresh
T "app.Refresh", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== export options ==="
On Error Resume Next
Err.Clear
Set opt = app.CreateStructExportOptions
T "app.CreateStructExportOptions (no parens)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set opt = app.CreateStructExportOptions()
T "app.CreateStructExportOptions()", Err.Number, Err.Description
If Not opt Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "        opt.ImageType before = " & opt.ImageType
  T "opt.ImageType (read)", Err.Number, Err.Description

  For k = 1 To 8
    On Error Resume Next
    Err.Clear
    opt.ImageType = k
    If Err.Number = 0 Then
      W "        ImageType accepts " & k & " (readback " & opt.ImageType & ")"
    End If
  Next

  On Error Resume Next
  Err.Clear
  opt.ResolutionX = 96
  T "opt.ResolutionX = 96", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  opt.ResolutionY = 96
  T "opt.ResolutionY = 96", Err.Number, Err.Description

  For k = 0 To 3
    On Error Resume Next
    Err.Clear
    opt.AntiAliasing = k
    If Err.Number = 0 Then
      W "        AntiAliasing accepts " & k & " (readback " & opt.AntiAliasing & ")"
    End If
  Next
End If

'----------------------------------------------------------
W ""
W "=== macro invocation ==="
On Error Resume Next
Err.Clear
app.RunMacro "CDRX4Toolkit.M_Util.HasDocument"
T "app.RunMacro ""CDRX4Toolkit.M_Util.HasDocument""", Err.Number, Err.Description

On Error Resume Next
Err.Clear
W "        TypeName(app.GMSManager) = " & TypeName(app.GMSManager)
T "TypeName(app.GMSManager)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
app.GMSManager.RunMacro "CDRX4Toolkit", "M_Util.HasDocument"
T "app.GMSManager.RunMacro ""CDRX4Toolkit"",""M_Util.HasDocument""", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "done: OK=" & nOK & "  FAIL=" & nFail

On Error Resume Next
doc.Close
logFile.Close