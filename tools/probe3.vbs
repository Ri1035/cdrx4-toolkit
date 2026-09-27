Option Explicit

'==========================================================
'  CorelDRAW X4 probe #3 - remaining semantics.
'  ASCII ONLY: WSH reads .vbs as ANSI, so any non-ASCII
'  literal here would be silently mangled.
'==========================================================

Dim fso, here, root, logFile
Dim nOK, nFail

Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe3.log", True)

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

Sub BB(lbl, s)
  If s Is Nothing Then
    W "        " & lbl & " = <nothing>"
  Else
    On Error Resume Next
    W "        " & lbl & " = L" & s.LeftX & " B" & s.BottomY & _
      " R" & s.RightX & " T" & s.TopY & _
      "  size " & s.SizeWidth & " x " & s.SizeHeight & _
      "  pos " & s.PositionX & "," & s.PositionY & _
      "  rot " & s.RotationAngle
  End If
End Sub

Dim app, doc, pg, lay
Dim rect, rect2, ell, txt, poly, s1, sr, opt, crv, o
Dim i, k

W "probe3 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
T "attach", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set doc = app.CreateDocument
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer
T "new document", Err.Number, Err.Description
W "        doc.Unit = " & doc.Unit

'----------------------------------------------------------
W ""
W "=== coordinate semantics ==="
On Error Resume Next
Err.Clear
Set rect = lay.CreateRectangle2(0, 0, 20, 10)
T "CreateRectangle2(0,0,20,10)", Err.Number, Err.Description
BB "rect", rect

On Error Resume Next
Err.Clear
Set rect2 = lay.CreateRectangle2(0, 0, 20, 10)
rect2.SetPosition 100, 100
T "SetPosition 100,100", Err.Number, Err.Description
BB "rect2", rect2

On Error Resume Next
Err.Clear
rect2.SetSize 30, 40
T "SetSize 30,40", Err.Number, Err.Description
BB "rect2", rect2

On Error Resume Next
Err.Clear
rect2.SetSize 30, 40, True
T "SetSize 30,40,True", Err.Number, Err.Description
BB "rect2", rect2

On Error Resume Next
Err.Clear
W "        pg.SizeWidth = " & pg.SizeWidth & "  pg.SizeHeight = " & pg.SizeHeight
W "        doc.WorldScale = " & doc.WorldScale
W "        doc.ReferencePoint = " & doc.ReferencePoint
T "worldscale / referencepoint", Err.Number, Err.Description

On Error Resume Next
Err.Clear
W "        app.Unit = " & app.Unit
T "app.Unit", Err.Number, Err.Description

On Error Resume Next
Err.Clear
W "        app.WorldScale = " & app.WorldScale
T "app.WorldScale", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== text alignment: find cdrCenterAlignment by geometry ==="
On Error Resume Next
Err.Clear
Set txt = lay.CreateArtisticTextWide(0, 0, "ABCDEFGHIJ")
T "CreateArtisticTextWide(0,0,...)", Err.Number, Err.Description
If Not txt Is Nothing Then
  On Error Resume Next
  txt.Text.Story.Size = 20
  BB "text default", txt
  For k = 0 To 6
    On Error Resume Next
    Err.Clear
    txt.Text.Story.Alignment = k
    If Err.Number = 0 Then
      W "        alignment " & k & " -> L" & txt.LeftX & " R" & txt.RightX & _
        "  (anchor x = 0)"
    End If
  Next
End If

'----------------------------------------------------------
W ""
W "=== star: does CreatePolygon2 take a sharpness argument? ==="
Dim nodeCnt
For k = 0 To 3
  On Error Resume Next
  Err.Clear
  Set poly = lay.CreatePolygon2(0, 0, 20, 5, k * 25)
  If Err.Number <> 0 Then
    W "  CreatePolygon2(0,0,20,5," & (k * 25) & ") -> [" & Err.Number & "] " & Err.Description
  Else
    On Error Resume Next
    W "  CreatePolygon2(0,0,20,5," & (k * 25) & ") -> size " & poly.SizeWidth & " x " & poly.SizeHeight & _
      "  sides " & poly.Polygon.Sides & "  sharpness " & poly.Polygon.Sharpness
    On Error Resume Next
    Err.Clear
    Set s1 = poly.ConvertToCurves
    If Err.Number = 0 And Not s1 Is Nothing Then
      On Error Resume Next
      nodeCnt = s1.Curve.Nodes.Count
      W "        after ConvertToCurves: nodes = " & nodeCnt & _
        "  (5 = pentagon, 10 = star)"
    Else
      W "        ConvertToCurves failed [" & Err.Number & "] " & Err.Description
    End If
  End If
Next

W ""
W "=== polygon sharpness on a plain pentagon ==="
On Error Resume Next
Err.Clear
Set poly = lay.CreatePolygon2(0, 0, 20, 5)
T "CreatePolygon2(0,0,20,5)", Err.Number, Err.Description
If Not poly Is Nothing Then
  W "        initial sharpness = " & poly.Polygon.Sharpness
  For k = 1 To 3
    On Error Resume Next
    Err.Clear
    poly.Polygon.Sharpness = 53
    Set s1 = poly.ConvertToCurves
    nodeCnt = s1.Curve.Nodes.Count
    W "        sharpness=53 -> nodes = " & nodeCnt & "  size " & s1.SizeWidth & " x " & s1.SizeHeight
    Exit For
  Next
End If

'----------------------------------------------------------
W ""
W "=== blend candidates ==="
On Error Resume Next
Err.Clear
Set rect = lay.CreateRectangle2(0, 0, 10, 10)
Set rect2 = lay.CreateRectangle2(50, 0, 10, 10)
T "two rects", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set sr = Nothing
Set sr = rect.Range(rect2)
T "rect.Range(rect2)", Err.Number, Err.Description
W "        TypeName = " & TypeName(sr)

If Not sr Is Nothing Then
  On Error Resume Next
  Err.Clear
  Set o = Nothing
  Set o = sr.CreateBlend
  T "sr.CreateBlend", Err.Number, Err.Description
  W "        TypeName = " & TypeName(o)
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = lay.CreateBlend(rect, rect2)
T "lay.CreateBlend(rect,rect2)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = doc.CreateBlend(rect, rect2)
T "doc.CreateBlend(rect,rect2)", Err.Number, Err.Description

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = rect.CreateBlend(rect2)
T "rect.CreateBlend(rect2)", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== curve API for fit-to-path ==="
On Error Resume Next
Err.Clear
Set ell = lay.CreateEllipse2(0, 0, 40, 30)
T "CreateEllipse2(0,0,40,30)", Err.Number, Err.Description
On Error Resume Next
Err.Clear
Set crv = ell.ConvertToCurves
T "ell.ConvertToCurves", Err.Number, Err.Description

If Not crv Is Nothing Then
  On Error Resume Next
  Err.Clear
  W "        crv.Curve.Length = " & crv.Curve.Length
  T "crv.Curve.Length", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  Set o = Nothing
  Set o = crv.Curve.GetPointPositionAt(crv.Curve.Length / 2)
  T "crv.Curve.GetPointPositionAt(len/2)", Err.Number, Err.Description
  If Not o Is Nothing Then
    On Error Resume Next
    W "        point = " & o.x & "," & o.y
    On Error Resume Next
    Err.Clear
    W "        point2 = " & crv.Curve.GetPointPositionAt(crv.Curve.Length / 2).x
  End If

  On Error Resume Next
  Err.Clear
  W "        GetPointAngleAt = " & crv.Curve.GetPointAngleAt(crv.Curve.Length / 2)
  T "crv.Curve.GetPointAngleAt(len/2)", Err.Number, Err.Description

  On Error Resume Next
  Err.Clear
  W "        crv.Curve.Nodes.Count = " & crv.Curve.Nodes.Count
  T "crv.Curve.Nodes.Count", Err.Number, Err.Description
End If

On Error Resume Next
Err.Clear
rect.SetRotationCenter 5, 5
T "rect.SetRotationCenter 5,5", Err.Number, Err.Description

On Error Resume Next
Err.Clear
rect.RotationAngle = 30
T "rect.RotationAngle = 30", Err.Number, Err.Description
BB "rotated rect", rect

On Error Resume Next
Err.Clear
rect.Rotate 15
T "rect.Rotate 15", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== selection order ==="
On Error Resume Next
Err.Clear
doc.ClearSelection
Set rect = lay.CreateRectangle2(0, 0, 10, 10)
Set rect2 = lay.CreateRectangle2(50, 0, 10, 10)
T "fresh rects", Err.Number, Err.Description

On Error Resume Next
Err.Clear
rect.AddToSelection
rect2.AddToSelection
W "        after selecting rect then rect2: count = " & app.ActiveSelection.Shapes.Count & _
  "  shapes(1) size = " & app.ActiveSelection.Shapes(1).SizeWidth & _
  "  shapes(2) = " & app.ActiveSelection.Shapes(2).SizeWidth
T "select rect,rect2", Err.Number, Err.Description

On Error Resume Next
Err.Clear
doc.ClearSelection
rect2.AddToSelection
rect.AddToSelection
W "        after selecting rect2 then rect: count = " & app.ActiveSelection.Shapes.Count
T "select rect2,rect", Err.Number, Err.Description

'----------------------------------------------------------
W ""
W "=== export: find the JPEG filter id ==="
On Error Resume Next
Err.Clear
Set opt = app.CreateStructExportOptions
opt.ImageType = 5
opt.ResolutionX = 96
opt.ResolutionY = 96
T "export options", Err.Number, Err.Description

Dim tmp
For k = 1 To 20
  tmp = fso.GetSpecialFolder(2) & "\cdrx4_f" & k & ".jpg"
  On Error Resume Next
  Err.Clear
  If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
  doc.ExportEx tmp, k, 1, opt
  Dim e1, made1
  e1 = Err.Number
  made1 = fso.FileExists(tmp)
  If made1 Then fso.DeleteFile tmp, True
  If made1 Or e1 <> 13 Then
    W "        ExportEx filter=" & k & " range=1 -> err [" & e1 & "] " & Err.Description & "  file=" & made1
  End If
Next

W ""
W "=== export: plain Export method ==="
For k = 1 To 20
  tmp = fso.GetSpecialFolder(2) & "\cdrx4_g" & k & ".jpg"
  On Error Resume Next
  Err.Clear
  If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
  doc.Export tmp, k
  Dim e2, made2
  e2 = Err.Number
  made2 = fso.FileExists(tmp)
  If made2 Then fso.DeleteFile tmp, True
  If made2 Or e2 = 0 Then
    W "        Export filter=" & k & " -> err [" & e2 & "]  file=" & made2
  End If
Next

On Error Resume Next
Err.Clear
tmp = fso.GetSpecialFolder(2) & "\cdrx4_h.jpg"
If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
doc.Export tmp, 8
W "        Export filter=8 -> err [" & Err.Number & "] " & Err.Description & "  file=" & fso.FileExists(tmp)

On Error Resume Next
Err.Clear
doc.ExportEx tmp, 8, 1
W "        ExportEx 3-arg -> err [" & Err.Number & "] " & Err.Description

On Error Resume Next
Err.Clear
Dim sTmp
sTmp = CStr(tmp)
doc.ExportEx sTmp, 8, 1, opt
W "        ExportEx with CStr -> err [" & Err.Number & "] " & Err.Description & "  file=" & fso.FileExists(tmp)

On Error Resume Next
Err.Clear
Dim vTmp
vTmp = tmp
doc.ExportEx vTmp, 8, 1, opt
W "        ExportEx with variant -> err [" & Err.Number & "] " & Err.Description & "  file=" & fso.FileExists(tmp)

W ""
W "done: OK=" & nOK & "  FAIL=" & nFail

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close