Option Explicit

'==========================================================
'  Unit test for the caption patcher embedded in
'  dist\安装CDRX4增强工具.vbs (installer_template.txt).
'
'  Why this exists: the first version of the patcher was not
'  idempotent.  It guarded on ` userCaption=` immediately after the
'  command name, but the pair it writes starts with ` userToolTip=`,
'  so the guard never matched and a second install appended a second
'  pair.  DRAWUIConfig.xml is XML, and XML forbids duplicate
'  attributes -- a re-run would have handed CorelDRAW a file it
'  cannot parse.  These cases pin that down.
'
'  ASCII only: the logic is pure string handling, so the samples use
'  ASCII placeholders instead of the real Chinese captions.
'
'  usage: cscript //nologo tools\caption_patch_test.vbs
'==========================================================

Dim fails
fails = 0

Sub Check(ByVal name, ByVal got, ByVal want)
  If StrComp(got, want, vbBinaryCompare) = 0 Then
    WScript.Echo "PASS  " & name
  Else
    WScript.Echo "FAIL  " & name
    WScript.Echo "        got  = [" & got & "]"
    WScript.Echo "        want = [" & want & "]"
    fails = fails + 1
  End If
End Sub

' --- the code under test (kept byte-identical to the installer) ---

Function ReCaption(xml, needle, ttl, cap)
  Dim re, rep

  rep = needle & " userToolTip=""" & ttl & """ userCaption=""" & cap & """"
  rep = Replace(rep, "$", "$$")

  Set re = CreateObject("VBScript.RegExp")
  re.Global = True
  re.IgnoreCase = True
  re.Pattern = ReEsc(needle) & "(?: userToolTip=""[^""]*"" userCaption=""[^""]*"")*"
  ReCaption = re.Replace(xml, rep)
End Function

Function ReEsc(s)
  Dim i, ch, out

  out = ""
  For i = 1 To Len(s)
    ch = Mid(s, i, 1)
    If InStr(1, "\.^$*+?()[]{}|", ch, vbBinaryCompare) > 0 Then out = out & "\"
    out = out & ch
  Next
  ReEsc = out
End Function

' --- fixtures ------------------------------------------------------

Dim needle, other, bare, one, two, oneNew, expect

needle = "dynamicCommand=""CDRX4Toolkit.M_X.Y"""
other  = "dynamicCommand=""Emboss.Main.Emboss"""

bare = "<itemData " & needle & " bmpCol=""7""/>"
one  = "<itemData " & needle & " userToolTip=""T"" userCaption=""C"" bmpCol=""7""/>"
two  = "<itemData " & needle & " userToolTip=""T"" userCaption=""C""" & _
       " userToolTip=""T"" userCaption=""C"" bmpCol=""7""/>"
oneNew = "<itemData " & needle & " userToolTip=""T2"" userCaption=""C2"" bmpCol=""7""/>"

' a neighbour entry that must not be touched
Dim page
page = "<itemData " & other & " bmpCol=""7""/>"

' 1. a bare command gets exactly one pair
Check "adds a pair to a bare command", _
      ReCaption(bare, needle, "T", "C"), one

' 2. running again changes nothing (this is the case that was broken)
Check "is idempotent on an already-captioned command", _
      ReCaption(one, needle, "T", "C"), one

' 3. a doubled entry is repaired down to one pair
Check "collapses a doubled entry", _
      ReCaption(two, needle, "T", "C"), one

' 4. a stale caption is replaced, not appended
Check "replaces a stale caption", _
      ReCaption(one, needle, "T2", "C2"), oneNew

' 5. a neighbour with a different command is left alone
Check "leaves other commands untouched", _
      ReCaption(page, needle, "T", "C"), page

' 6. a real page: bare target + untouched neighbour
Check "patches the target and only the target", _
      ReCaption(bare & page, needle, "T", "C"), one & page

WScript.Echo ""
If fails = 0 Then
  WScript.Echo "==== CAPTION PATCH TEST PASS ===="
Else
  WScript.Echo "==== CAPTION PATCH TEST FAIL (" & fails & ") ===="
End If
WScript.Quit fails