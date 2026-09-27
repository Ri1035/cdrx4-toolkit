Option Explicit

'==========================================================
'  CorelDRAW X4 probe #7 - segment sampling + bbox setters.
'  ASCII ONLY.
'==========================================================

Dim fso, here, root, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(here)
Set logFile = fso.CreateTextFile(root & "\_probe7.log", True)

Sub W(s)
  logFile.WriteLine s
End Sub

Dim app, doc, pg, lay, o, sg, shp
Dim i, k

W "probe7 start " & Now

On Error Resume Next
Err.Clear
Set app = CreateObject("CorelDRAW.Application.14")
app.InitializeVBA
Set doc = app.CreateDocument
Set pg = doc.ActivePage
Set lay = pg.ActiveLayer

'----------------------------------------------------------
W ""
W "=== segment sampling with two arguments ==="
On Error Resume Next
Err.Clear
Set shp = Nothing
Set shp = lay.CreateEllipse2(0, 0, 40, 30)
shp.ConvertToCurves
W "  ellipse curve: length " & shp.Curve.Length & "  segs " & shp.Curve.Segments.Count

For k = 0 To 2
  On Error Resume Next
  Err.Clear
  Set sg = Nothing
  Set sg = shp.Curve.Segments(1)
  Set o = Nothing
  Set o = sg.GetPointPositionAt(0.5, k)
  W "  seg(1).GetPointPositionAt(0.5," & k & ") -> [" & Err.Number & "] " & _
    Err.Description & "  TypeName " & TypeName(o)
  If Err.Number = 0 And Not o Is Nothing Then
    On Error Resume Next
    W "        x/y = " & o.x & "," & o.y
  End If
Next

On Error Resume Next
Err.Clear
Set sg = Nothing
Set sg = shp.Curve.Segments(1)
Set o = Nothing
Set o = sg.GetPointPositionAt(11.18, 0)
W "  seg(1).GetPointPositionAt(11.18,0) -> [" & Err.Number & "] " & Err.Description
If Err.Number = 0 And Not o Is Nothing Then
  On Error Resume Next
  W "        x/y = " & o.x & "," & o.y
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sg.StartNode
W "  seg.StartNode -> TypeName " & TypeName(o)
If Not o Is Nothing Then
  On Error Resume Next
  W "        start = " & o.PositionX & "," & o.PositionY
End If

On Error Resume Next
Err.Clear
Set o = Nothing
Set o = sg.EndNode
W "  seg.EndNode -> TypeName " & TypeName(o)
If Not o Is Nothing Then
  On Error Resume Next
  W "        end = " & o.PositionX & "," & o.PositionY
End If

On Error Resume Next
Err.Clear
W "  seg.Angle = " & sg.Angle
W "  seg.TangentAt(0.5) = " & sg.GetTangentAt(0.5)
W "  seg.GetAngleAt(0.5) = " & sg.GetAngleAt(0.5)

'----------------------------------------------------------
W ""
W "=== bounding box setters ==="
On Error Resume Next
Err.Clear
Set shp = Nothing
Set shp = lay.CreateRectangle2(0, 0, 10, 10)
W "  before: L" & shp.LeftX & " B" & shp.BottomY & " R" & shp.RightX & " T" & shp.TopY

On Error Resume Next
Err.Clear
shp.LeftX = 5
W "  shp.LeftX = 5 -> [" & Err.Number & "] " & Err.Description & _
  "  now L" & shp.LeftX & " R" & shp.RightX

On Error Resume Next
Err.Clear
shp.BottomY = 7
W "  shp.BottomY = 7 -> [" & Err.Number & "] " & Err.Description & _
  "  now B" & shp.BottomY & " T" & shp.TopY

On Error Resume Next
Err.Clear
shp.SetPosition 50, 50
W "  SetPosition 50,50 -> L" & shp.LeftX & " B" & shp.BottomY & _
  " R" & shp.RightX & " T" & shp.TopY & "  ReferencePoint " & doc.ReferencePoint

On Error Resume Next
Err.Clear
shp.SetPosition 50, 50, 1
W "  SetPosition 50,50,1 -> [" & Err.Number & "] " & Err.Description

'----------------------------------------------------------
W ""
W "=== text metrics for centring ==="
On Error Resume Next
Err.Clear
Dim t
Set t = Nothing
Set t = lay.CreateArtisticTextWide(100, 100, "ABC")
W "  text created, err " & Err.Number
If Not t Is Nothing Then
  W "  default: L" & t.LeftX & " B" & t.BottomY & " R" & t.RightX & " T" & t.TopY
  On Error Resume Next
  t.Text.Story.Alignment = 3
  W "  center alignment: L" & t.LeftX & " R" & t.RightX & "  (anchor 100)"
  On Error Resume Next
  t.Text.Story.Size = 12
  W "  size 12: L" & t.LeftX & " R" & t.RightX & " T" & t.TopY & " B" & t.BottomY
  On Error Resume Next
  Err.Clear
  t.Text.Story.Font = "Arial"
  W "  font Arial err " & Err.Number
End If

W ""
W "=== outline width units ==="
On Error Resume Next
Err.Clear
Set shp = lay.CreateEllipse2(0, 0, 20, 20)
shp.Outline.Width = 1
W "  outline width set to 1 -> reads " & shp.Outline.Width
On Error Resume Next
shp.Outline.Width = 0.02
W "  outline width set to 0.02 -> reads " & shp.Outline.Width
On Error Resume Next
Err.Clear
shp.Outline.Style = 0
W "  outline style 0 -> [" & Err.Number & "] " & Err.Description
On Error Resume Next
Err.Clear
W "  outline style read = " & shp.Outline.Style
W "  outline type = " & shp.Outline.Type

W ""
W "done"

On Error Resume Next
doc.Dirty = False
doc.Close
logFile.Close