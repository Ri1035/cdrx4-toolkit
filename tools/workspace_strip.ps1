# ==========================================================
#  Simulate a fresh user's workspace, and put it back.
#
#  Why: the toolbar is written into DRAWUIConfig.xml by the
#  installer.  To test the install path from a clean slate -- and to
#  test the uninstall path -- there has to be a way to remove every
#  trace of the toolkit from the workspace without running CorelDRAW.
#
#  What counts as "a trace" comes from workspace_markup.txt, which is
#  the same file the installer splices in.  So this script cannot drift
#  from the shipped markup: it reads the guids and names out of it.
#  It also strips older builds by the stable signatures (a
#  dynamicCommand under CDRX4Toolkit, or one of our toolbar names), so
#  a v1.2.0 install is cleaned too.
#
#  IMPLEMENTATION NOTE -- do not "improve" this into an XmlDocument
#  round-trip.  XmlDocument.Save writes a UTF-8 BOM, and CorelDRAW
#  writes this file WITHOUT one.  The first version of this script
#  did exactly that and the next CorelDRAW start died in
#  CrlFrmWk.dll (0xc0000005 -> 0xc000041d), i.e. the test harness
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

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$markup = Join-Path $root 'workspace_markup.txt'

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

# ---- what the toolkit looks like, straight from the shipped markup ----
$mk = [System.IO.File]::ReadAllText($markup, [System.Text.Encoding]::UTF8)

$itemGuids = @()
foreach ($m in [regex]::Matches($mk, '<itemData\b[^>]*\bguid="([^"]+)"')) {
  $itemGuids += $m.Groups[1].Value
}
$barGuids = @()
$barNames = @()
foreach ($m in [regex]::Matches($mk, '<commandBarData\b[^>]*>')) {
  $tag = $m.Value
  $g = [regex]::Match($tag, '\bguid="([^"]+)"').Groups[1].Value
  if ($g) { $barGuids += $g }
  $nm = [regex]::Match($tag, 'nonLocalizableName="([^"]+)"').Groups[1].Value
  if ($nm) { $barNames += $nm }
}

if ($itemGuids.Count -eq 0 -or $barGuids.Count -eq 0) {
  Write-Output "!! could not read guids out of $markup"
  exit 2
}

function Show-Counts([string]$path) {
  $c = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
  $b = [System.IO.File]::ReadAllBytes($path)
  $bom = ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
  $opener = 0
  foreach ($m in [regex]::Matches($c, '<itemData\b[^>]*/>')) {
    if ($m.Value -match 'flyoutBarRef="') { $opener++ }
  }
  return [pscustomobject]@{
    Items  = ([regex]::Matches($c, 'dynamicCommand="CDRX4Toolkit')).Count
    Opener = $opener
    Flyout = ([regex]::Matches($c, 'flyout="true"')).Count
    Bars   = ([regex]::Matches($c, 'nonLocalizableName="(?:' + (($barNames | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')"')).Count
    Caps   = ([regex]::Matches($c, 'userCaption=')).Count
    Bytes  = $b.Length
    Bom    = $bom
  }
}

if ($Report) {
  $s = Show-Counts $xml
  Write-Output "items=$($s.Items) opener=$($s.Opener) flyout=$($s.Flyout) bars=$($s.Bars) userCaption=$($s.Caps) bytes=$($s.Bytes) bom=$($s.Bom)"
  exit 0
}

Copy-Item $xml $bak -Force

$c = [System.IO.File]::ReadAllText($xml, [System.Text.Encoding]::UTF8)
$out = $c

# 1. our commandBarData blocks: by guid, by name, or by the button guids
#    they list (an older build's skeleton has a random guid, so its name
#    or its buttons are what identify it)
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
  foreach ($g in $barGuids) { if ($block -match ('\bguid="' + [regex]::Escape($g) + '"')) { $hit = $true; break } }
  foreach ($nm in $barNames) { if ($block -match ('nonLocalizableName="' + [regex]::Escape($nm) + '"')) { $hit = $true; break } }
  foreach ($g in $itemGuids) { if ($block.Contains('guidRef="' + $g + '"')) { $hit = $true; break } }
  if (-not $hit) { if ($block -match 'dynamicCommand="CDRX4Toolkit\.') { $hit = $true } }

  if ($hit) {
    $out = $out.Remove($s, $e - $s)
    $removedBars++
    $i = $s
  } else {
    $i = $e
  }
}

# 2. our itemData: by guid (covers the flyout opener too), by flyoutBarRef
#    into a bar we just removed, or by the macro path
$removedItems = 0
foreach ($m in [regex]::Matches($c, '<itemData\b[^>]*/>')) {
  $tag = $m.Value
  $hit = $false
  foreach ($g in $itemGuids) { if ($tag -match ('\bguid="' + [regex]::Escape($g) + '"')) { $hit = $true; break } }
  if (-not $hit) { foreach ($g in $barGuids) { if ($tag -match ('flyoutBarRef="' + [regex]::Escape($g) + '"')) { $hit = $true; break } } }
  if (-not $hit) { if ($tag -match 'dynamicCommand="CDRX4Toolkit\.') { $hit = $true } }
  if ($hit) { $out = $out.Replace($tag, ''); $removedItems++ }
}

# 3. the <cmdBarLane> visibility entries for the bars removed above
$removedLanes = 0
foreach ($m in [regex]::Matches($c, '<commandBar\b[^>]*/>')) {
  $tag = $m.Value
  foreach ($g in $barGuids) {
    if ($tag -match ('guidRef="' + [regex]::Escape($g) + '"')) {
      $out = $out.Replace($tag, '')
      $removedLanes++
      break
    }
  }
}

if ($removedItems -eq 0 -and $removedBars -eq 0) {
  Write-Output "nothing to strip (already clean)"
  exit 0
}

[System.IO.File]::WriteAllText($xml, $out, $noBom)

$s = Show-Counts $xml
Write-Output "stripped $removedItems itemData, $removedBars commandBarData, $removedLanes visibility entry(ies)"
Write-Output "now: items=$($s.Items) opener=$($s.Opener) flyout=$($s.Flyout) bars=$($s.Bars) userCaption=$($s.Caps) bytes=$($s.Bytes) bom=$($s.Bom)"