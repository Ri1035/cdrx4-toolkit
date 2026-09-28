# ==========================================================
#  Simulate a fresh user's workspace.
#
#  Why: the installer creates the toolbar through automation and
#  then relies on CorelDRAW having persisted it into
#  DRAWUIConfig.xml.  Whether CorelDRAW actually persists on an
#  automation-driven Quit is exactly the thing that has never been
#  verified -- and getting it wrong is what shipped v1.0.2.
#
#  So: strip every CDRX4Toolkit entry (the 9 itemData command
#  bindings and the commandBarData toolbar skeleton that references
#  them) out of DRAWUIConfig.xml, then run the installer, then start
#  CorelDRAW by hand and see whether the toolbar comes back.
#
#  IMPLEMENTATION NOTE -- do not "improve" this into an XmlDocument
#  round-trip.  XmlDocument.Save writes a UTF-8 BOM, and CorelDRAW
#  writes this file WITHOUT one.  The first version of this script
#  did exactly that and the next CorelDRAW start died in
#  CrlFmWk.dll (0xc0000005 -> 0xc000041d), i.e. the test harness
#  itself poisoned the experiment.  Plain string surgery keeps the
#  file byte-faithful apart from the removed elements.
#
#  usage:
#     powershell -File tools\workspace_strip.ps1            # strip
#     powershell -File tools\workspace_strip.ps1 -Restore   # put back
#     powershell -File tools\workspace_strip.ps1 -Report    # just count
#
#  A backup is written to DRAWUIConfig.xml.pretest before stripping.
#  ASCII only.
# ==========================================================

param(
  [switch]$Restore,
  [switch]$Report
)

$ErrorActionPreference = 'Stop'

$dir  = Join-Path $env:APPDATA 'Corel\CorelDRAW Graphics Suite X4\User Workspace\CorelDRAW\_default'
$xml  = Join-Path $dir 'DRAWUIConfig.xml'
$bak  = Join-Path $dir 'DRAWUIConfig.xml.pretest'
$noBom = New-Object System.Text.UTF8Encoding($false)

if (-not (Test-Path $xml)) { Write-Output "!! no workspace file at $xml"; exit 2 }

if ($Restore) {
  if (-not (Test-Path $bak)) { Write-Output "!! no backup at $bak"; exit 2 }
  Copy-Item $bak $xml -Force
  Write-Output "restored $xml from pretest backup"
  exit 0
}

function Show-Counts([string]$path) {
  $c = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
  $b = [System.IO.File]::ReadAllBytes($path)
  $bom = ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
  return [pscustomobject]@{
    Items = ([regex]::Matches($c, 'dynamicCommand="CDRX4Toolkit')).Count
    Caps  = ([regex]::Matches($c, 'userCaption=')).Count
    Bytes = $b.Length
    Bom   = $bom
  }
}

if ($Report) {
  $s = Show-Counts $xml
  Write-Output "items=$($s.Items) userCaption=$($s.Caps) bytes=$($s.Bytes) bom=$($s.Bom)"
  exit 0
}

Copy-Item $xml $bak -Force

$c = [System.IO.File]::ReadAllText($xml, [System.Text.Encoding]::UTF8)

# 1. the itemData nodes that bind a button to one of our macros
$itemRe = '<itemData\b[^>]*dynamicCommand="CDRX4Toolkit\.[^"]*"[^>]*/>'
$itemHits = [regex]::Matches($c, $itemRe)
if ($itemHits.Count -eq 0) { Write-Output "nothing to strip (already clean)"; exit 0 }

$guids = @()
foreach ($m in $itemHits) {
  $g = [regex]::Match($m.Value, 'guid="([^"]+)"').Groups[1].Value
  if ($g) { $guids += $g }
}

# 2. any toolbar skeleton that references those buttons
$out = $c
$removedBars = 0
$i = 0
while ($true) {
  $s = $out.IndexOf('<commandBarData', $i)
  if ($s -lt 0) { break }
  $e = $out.IndexOf('</commandBarData>', $s)
  if ($e -lt 0) { break }
  $e = $e + '</commandBarData>'.Length
  $block = $out.Substring($s, $e - $s)
  $hit = $false
  foreach ($g in $guids) { if ($block.Contains('guidRef="' + $g + '"')) { $hit = $true; break } }
  if ($hit) {
    $out = $out.Remove($s, $e - $s)
    $removedBars++
    $i = $s
  } else {
    $i = $e
  }
}

# 3. now the itemData nodes themselves
$out = [regex]::Replace($out, $itemRe, '')

[System.IO.File]::WriteAllText($xml, $out, $noBom)

$s = Show-Counts $xml
Write-Output "stripped $($itemHits.Count) itemData and $removedBars commandBarData"
Write-Output "now: items=$($s.Items) userCaption=$($s.Caps) bytes=$($s.Bytes) bom=$($s.Bom)"