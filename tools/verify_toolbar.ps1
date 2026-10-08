# ==========================================================
#  Verify the toolkit toolbar markup in the CorelDRAW workspace.
#
#  The toolbar is pure XML: the installer splices workspace_markup.txt
#  into DRAWUIConfig.xml while CorelDRAW is closed, and CorelDRAW draws
#  it on the next start.  This script checks the installed file against
#  that markup, so it verifies BOTH directions of the install without
#  ever starting CorelDRAW:
#
#     -Expect present   after the installer: every guid, the flyout
#                       wiring, the button order and the visibility
#                       entry must all be there
#     -Expect absent    after the uninstaller: none of it may be left,
#                       including the flyout opener (which carries no
#                       dynamicCommand, so a name/macro sweep alone
#                       would miss it)
#
#  It also cross-checks %TEMP%\cdrx4_toolbar_items.txt when that exists.
#  That file is written by M_Install.DumpItems inside CorelDRAW, so the
#  check ties the compiled GMS button list to the installed markup.
#
#  usage:
#     powershell -File tools\verify_toolbar.ps1
#     powershell -File tools\verify_toolbar.ps1 -Expect absent
#
#  ASCII only.
# ==========================================================

param(
  [ValidateSet('present','absent')]
  [string]$Expect = 'present',
  [string]$Workspace = ''
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$markup = Join-Path $root 'workspace_markup.txt'

if ($Workspace -eq '') {
  $Workspace = Join-Path $env:APPDATA 'Corel\CorelDRAW Graphics Suite X4\User Workspace\CorelDRAW\_default\DRAWUIConfig.xml'
}
if (-not (Test-Path $markup))   { Write-Output "!! no markup at $markup"; exit 2 }
if (-not (Test-Path $Workspace)) { Write-Output "!! no workspace at $Workspace"; exit 2 }

$fails = 0
function Ok([string]$s)  { Write-Output "  ok   $s" }
function Bad([string]$s) { Write-Output "  FAIL $s"; $script:fails++ }

$mk = [System.IO.File]::ReadAllText($markup, [System.Text.Encoding]::UTF8)
$ws = [System.IO.File]::ReadAllText($Workspace, [System.Text.Encoding]::UTF8)
$wsBytes = [System.IO.File]::ReadAllBytes($Workspace)
$wsBom = ($wsBytes.Length -ge 3 -and $wsBytes[0] -eq 0xEF -and $wsBytes[1] -eq 0xBB -and $wsBytes[2] -eq 0xBF)

# ---- what the shipped markup says the toolbar should be ----
$itemGuids = @()
foreach ($m in [regex]::Matches($mk, '<itemData\b[^>]*\bguid="([^"]+)"')) { $itemGuids += $m.Groups[1].Value }

$barGuids = @()
$barNames = @()
foreach ($m in [regex]::Matches($mk, '<commandBarData\b[^>]*>')) {
  $tag = $m.Value
  $g = [regex]::Match($tag, '\bguid="([^"]+)"').Groups[1].Value
  if ($g) { $barGuids += $g }
  $nm = [regex]::Match($tag, 'nonLocalizableName="([^"]+)"').Groups[1].Value
  if ($nm) { $barNames += $nm }
}

# the opener is the one itemData that carries flyoutBarRef instead of a macro
$openerGuid = ''
$flyoutGuid = ''
foreach ($m in [regex]::Matches($mk, '<itemData\b[^>]*/>')) {
  if ($m.Value -match 'flyoutBarRef="([^"]+)"') {
    $openerGuid = [regex]::Match($m.Value, '\bguid="([^"]+)"').Groups[1].Value
    $flyoutGuid = $Matches[1]
  }
}
$mainBarGuid = ''
foreach ($m in [regex]::Matches($mk, '<commandBarData\b[^>]*>')) {
  if ($m.Value -notmatch 'flyout="true"') {
    $mainBarGuid = [regex]::Match($m.Value, '\bguid="([^"]+)"').Groups[1].Value
  }
}
$markupProcs = @()
foreach ($m in [regex]::Matches($mk, 'dynamicCommand="([^"]+)"')) { $markupProcs += $m.Groups[1].Value }

if ($itemGuids.Count -eq 0 -or $barGuids.Count -eq 0 -or $openerGuid -eq '' -or $flyoutGuid -eq '' -or $mainBarGuid -eq '') {
  Write-Output "!! could not parse the toolbar shape out of $markup"
  exit 2
}

Write-Output "verify_toolbar: expect=$Expect"
Write-Output "  markup    = $markup"
Write-Output "  workspace = $Workspace"
Write-Output "  itemData  = $($itemGuids.Count)  commandBarData = $($barGuids.Count)  buttons = $($markupProcs.Count)"
Write-Output "  opener    = $openerGuid -> flyout $flyoutGuid"
Write-Output "  main bar  = $mainBarGuid"
Write-Output ""

if ($wsBom) { Bad "workspace file has a UTF-8 BOM (CorelDRAW writes none)" } else { Ok "no BOM" }

if ($Expect -eq 'present') {

  # every itemData guid from the markup must be in the file
  $miss = @()
  foreach ($g in $itemGuids) { if ($ws -notmatch ('\bguid="' + [regex]::Escape($g) + '"')) { $miss += $g } }
  if ($miss.Count -gt 0) { Bad "$($miss.Count) itemData guid(s) missing, e.g. $($miss[0])" } else { Ok "all $($itemGuids.Count) itemData guids present" }

  # both commandBarData skeletons
  $miss = @()
  foreach ($g in $barGuids) { if ($ws -notmatch ('<commandBarData\b[^>]*\bguid="' + [regex]::Escape($g) + '"')) { $miss += $g } }
  if ($miss.Count -gt 0) { Bad "$($miss.Count) commandBarData guid(s) missing" } else { Ok "both commandBarData skeletons present" }

  # the flyout flag lives on the flyout bar
  if ($ws -match ('<commandBarData\b[^>]*\bguid="' + [regex]::Escape($flyoutGuid) + '"[^>]*flyout="true"')) { Ok "flyout bar is marked flyout=true" } else { Bad "flyout bar is missing flyout=true" }

  # the opener points at the flyout bar
  if ($ws -match ('<itemData\b[^>]*\bguid="' + [regex]::Escape($openerGuid) + '"[^>]*flyoutBarRef="' + [regex]::Escape($flyoutGuid) + '"')) { Ok "flyout opener present and wired" } else { Bad "flyout opener missing or mis-wired" }

  # the main bar carries the opener and nothing else
  $mb = [regex]::Match($ws, '(?s)<commandBarData\b[^>]*\bguid="' + [regex]::Escape($mainBarGuid) + '".*?</commandBarData>')
  if (-not $mb.Success) {
    Bad "main commandBarData block not found"
  } else {
    $refs = @([regex]::Matches($mb.Value, 'guidRef="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($refs.Count -eq 1 -and $refs[0] -eq $openerGuid) { Ok "main bar references only the opener" } else { Bad "main bar should reference only the opener, found $($refs.Count) item(s)" }
  }

  # the flyout bar carries the nine buttons in CmdList order
  $fb = [regex]::Match($ws, '(?s)<commandBarData\b[^>]*\bguid="' + [regex]::Escape($flyoutGuid) + '".*?</commandBarData>')
  if (-not $fb.Success) {
    Bad "flyout commandBarData block not found"
  } else {
    $refs = @([regex]::Matches($fb.Value, 'guidRef="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    $want = @($itemGuids | Where-Object { $_ -ne $openerGuid })
    if ($refs.Count -ne $want.Count) {
      Bad "flyout lists $($refs.Count) button(s), expected $($want.Count)"
    } elseif (($refs -join ',') -ne ($want -join ',')) {
      Bad "flyout button order differs from the markup"
    } else {
      Ok "flyout lists the $($want.Count) buttons in order"
    }
  }

  # the visibility entry
  if ($ws -match ('<commandBar\b[^>]*guidRef="' + [regex]::Escape($mainBarGuid) + '"[^>]*visible="true"')) { Ok "main bar visibility entry present" } else { Bad "main bar visibility entry missing" }

  # macro bindings, in order
  $got = @()
  foreach ($m in [regex]::Matches($ws, 'dynamicCommand="(CDRX4Toolkit\.[^"]+)"')) { $got += $m.Groups[1].Value }
  if (($got -join ',') -ne ($markupProcs -join ',')) {
    Bad "installed macro bindings differ from the markup (got $($got.Count), want $($markupProcs.Count))"
  } else {
    Ok "all $($got.Count) macro bindings present in order"
  }

  # duplicate attributes would make the XML invalid
  $dups = ([regex]::Matches($ws, '\buserCaption="')).Count
  if ($dups -lt $itemGuids.Count) { Bad "only $dups userCaption attribute(s) for $($itemGuids.Count) buttons" } else { Ok "$dups userCaption attribute(s)" }

} else {

  # nothing of ours may remain
  foreach ($g in $itemGuids) { if ($ws -match ('\bguid="' + [regex]::Escape($g) + '"')) { Bad "itemData guid $g still present" } }
  foreach ($g in $barGuids)  { if ($ws -match ('\bguid="' + [regex]::Escape($g) + '"')) { Bad "commandBarData guid $g still present" } }
  foreach ($g in $barGuids)  { if ($ws -match ('flyoutBarRef="' + [regex]::Escape($g) + '"')) { Bad "flyout opener still points at $g" } }
  foreach ($nm in $barNames) { if ($ws -match ('nonLocalizableName="' + [regex]::Escape($nm) + '"')) { Bad "toolbar name [$nm] still present" } }

  $left = ([regex]::Matches($ws, 'dynamicCommand="CDRX4Toolkit\.')).Count
  if ($left -gt 0) { Bad "$left macro binding(s) still present" }
  # flyoutBarRef is NOT unique to us: the stock workspace already carries
  # other plugins' flyout openers (the built-in "my tools" bar has three),
  # so a blanket count is a false positive.  The per-bar loop above already
  # checked every one of OUR flyout guids, which is the real test.

  if ($fails -eq 0) { Ok "no toolkit markup left in the workspace" }
}

# ---- optional cross-check against the compiled GMS button list ----
$itemsPath = Join-Path $env:TEMP 'cdrx4_toolbar_items.txt'
if (Test-Path $itemsPath) {
  Write-Output ""
  Write-Output "  cross-check $itemsPath"
  $dump = [System.IO.File]::ReadAllText($itemsPath, [System.Text.Encoding]::UTF8)
  $dump = $dump.TrimStart([char]0xFEFF)
  $dumpProcs = @()
  foreach ($ln in ($dump -split "`r?`n")) {
    if ($ln.Trim() -eq '') { continue }
    $dumpProcs += ($ln -split '\|')[0]
  }
  # DumpItems emits the bare proc path ("M_Curves.ConvertAllToCurves");
  # the markup's dynamicCommand carries the project prefix.  Normalise the
  # dump side so the two are actually comparable.
  $dumpFull = @($dumpProcs | ForEach-Object { 'CDRX4Toolkit.' + $_ })
  if (($dumpFull -join ',') -ne ($markupProcs -join ',')) {
    Bad "GMS CmdList differs from the markup ($($dumpProcs.Count) vs $($markupProcs.Count))"
  } else {
    Ok "GMS CmdList matches the markup"
  }
}

Write-Output ""
if ($fails -eq 0) {
  Write-Output "==== VERIFY PASS ($Expect) ===="
  exit 0
} else {
  Write-Output "==== VERIFY FAIL ($Expect, $fails problem(s)) ===="
  exit 1
}