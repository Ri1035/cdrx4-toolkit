# ==========================================================
#  Prototype: put the "增强工具" toolbar into DRAWUIConfig.xml by
#  editing the file, with no CorelDRAW automation at all.
#
#  Why this exists:
#  the automation route -- CreateObject, then
#  CommandBars.Delete / CommandBars.Add / Visible / AddCustomButton --
#  crashes X4.  Reproduced twice on a workspace that has no toolbar
#  yet: CorelDRW.exe dies in CrlFrmWk.dll (0xc0000005 -> 0xc000041d)
#  and leaves an unkillable 0-thread zombie.  That is the same crash
#  signature as the v1.0.2 startup hook, and it means the toolbar must
#  NOT be built through the CommandBars API.
#
#  CorelDRAW persists a toolbar as plain XML anyway: nine <itemData>
#  entries inside <items> (one per button, carrying dynamicCommand +
#  userCaption + userToolTip) and one <commandBarData> inside
#  <commandBars> that lists the button guids.  So the installer can
#  just write those, and X4 picks the toolbar up on the next start.
#
#  This script takes the markup verbatim from a known-good workspace
#  so the prototype cannot be wrong about the schema.  The shipped
#  installer generates the same shape with fresh guids.
#
#  usage:
#     powershell -File tools\workspace_inject.ps1 -From <known-good.xml>
#     powershell -File tools\workspace_inject.ps1 -Report
#
#  ASCII only.  Writes WITHOUT a BOM: CorelDRAW writes this file with
#  no BOM, and an XmlDocument round-trip (which adds one) was already
#  shown to be a bad idea.
# ==========================================================

param(
  [string]$From = "",
  [switch]$Report
)

$ErrorActionPreference = 'Stop'

$dir   = Join-Path $env:APPDATA 'Corel\CorelDRAW Graphics Suite X4\User Workspace\CorelDRAW\_default'
$xml   = Join-Path $dir 'DRAWUIConfig.xml'
$noBom = New-Object System.Text.UTF8Encoding($false)

if ($Report) {
  $c = [System.IO.File]::ReadAllText($xml, [System.Text.Encoding]::UTF8)
  $b = [System.IO.File]::ReadAllBytes($xml)
  Write-Output ("items={0} userCaption={1} bytes={2} bom={3}" -f `
    ([regex]::Matches($c,'dynamicCommand="CDRX4Toolkit')).Count, `
    ([regex]::Matches($c,'userCaption=')).Count, $b.Length, ($b[0] -eq 0xEF))
  exit 0
}

if ($From -eq "" -or -not (Test-Path $From)) { Write-Output "!! need -From <known-good.xml>"; exit 2 }

$src = [System.IO.File]::ReadAllText($From, [System.Text.Encoding]::UTF8)

# --- pull our markup out of the known-good file -------------------
$itemHits = [regex]::Matches($src, '<itemData\b[^>]*dynamicCommand="CDRX4Toolkit\.[^"]*"[^>]*/>')
if ($itemHits.Count -eq 0) { Write-Output "!! source file has no CDRX4Toolkit itemData"; exit 2 }

$guids = @()
foreach ($m in $itemHits) {
  $guids += [regex]::Match($m.Value, 'guid="([^"]+)"').Groups[1].Value
}

$bar = ""
$i = 0
while ($true) {
  $s = $src.IndexOf('<commandBarData', $i)
  if ($s -lt 0) { break }
  $e = $src.IndexOf('</commandBarData>', $s)
  if ($e -lt 0) { break }
  $e = $e + '</commandBarData>'.Length
  $block = $src.Substring($s, $e - $s)
  $hit = $false
  foreach ($g in $guids) { if ($block.Contains('guidRef="' + $g + '"')) { $hit = $true; break } }
  if ($hit) { $bar = $block; break }
  $i = $e
}

if ($bar -eq "") { Write-Output "!! source file has no commandBarData for those buttons"; exit 2 }

Write-Output "extracted $($itemHits.Count) itemData, 1 commandBarData ($($bar.Length) chars)"

# --- splice into the live file -----------------------------------
$t = [System.IO.File]::ReadAllText($xml, [System.Text.Encoding]::UTF8)

# start from a clean slate: drop whatever is already there
$t = [regex]::Replace($t, '<itemData\b[^>]*dynamicCommand="CDRX4Toolkit\.[^"]*"[^>]*/>', '')
# (the old commandBarData was removed by workspace_strip.ps1 in this test)

$eItems = $t.LastIndexOf('</items>')
if ($eItems -lt 0) { Write-Output "!! no </items> in $xml"; exit 2 }
$t = $t.Insert($eItems, ($itemHits | ForEach-Object { $_.Value }) -join '')

$eBars = $t.LastIndexOf('</commandBars>')
if ($eBars -lt 0) { Write-Output "!! no </commandBars> in $xml"; exit 2 }
$t = $t.Insert($eBars, $bar)

[System.IO.File]::WriteAllText($xml, $t, $noBom)

$b = [System.IO.File]::ReadAllBytes($xml)
Write-Output ("written: items={0} userCaption={1} bytes={2} bom={3}" -f `
  ([regex]::Matches($t,'dynamicCommand="CDRX4Toolkit')).Count, `
  ([regex]::Matches($t,'userCaption=')).Count, $b.Length, ($b[0] -eq 0xEF))

try { $d = New-Object System.Xml.XmlDocument; $d.Load($xml); Write-Output "XML parses OK" }
catch { Write-Output "XML ERROR: $($_.Exception.Message)"; exit 3 }