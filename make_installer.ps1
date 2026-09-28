$ErrorActionPreference = "Stop"

# This script is intentionally ASCII-only.
# All Chinese text lives in installer_msgs.txt / uninstaller_msgs.txt,
# which we read explicitly as UTF-8.
#
#   build_gms.vbs          -> writes CDRX4Toolkit.gms into the X4 user GMS folder
#   make_installer.ps1     -> embeds that GMS + the UI messages into
#                             dist\安装CDRX4增强工具.vbs (single self-contained file)
#                             and the UI messages into
#                             dist\卸载CDRX4增强工具.vbs (no payload: it only deletes)

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$u8   = New-Object System.Text.UTF8Encoding($false)
$ascii = New-Object System.Text.ASCIIEncoding

$apGms   = Join-Path $env:APPDATA "Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms"
$distGms = Join-Path $root "dist\CDRX4Toolkit.gms"

if (Test-Path $apGms) { $gms = $apGms }
elseif (Test-Path $distGms) { $gms = $distGms }
else { throw "CDRX4Toolkit.gms not found - run build_gms.vbs first" }

function Chunk([string]$s, [int]$size) {
  $sb = New-Object System.Text.StringBuilder
  for ($i = 0; $i -lt $s.Length; $i += $size) {
    $len = [Math]::Min($size, $s.Length - $i)
    $piece = $s.Substring($i, $len)
    if ($i + $len -ge $s.Length) {
      [void]$sb.AppendLine('"' + $piece + '"')
    } else {
      [void]$sb.AppendLine('"' + $piece + '" & _')
    }
  }
  $sb.ToString().TrimEnd()
}

# a base64 blob for one of the message files
function MsgBlob([string]$name) {
  $t = [System.IO.File]::ReadAllText((Join-Path $root $name), $u8)
  $t = ($t -replace "`r`n", "`n").TrimEnd("`n")
  return [Convert]::ToBase64String($u8.GetBytes($t))
}

# split a template into its "@@OUT:<filename>@@" header and its body
function ReadTpl([string]$name) {
  $tpl = [System.IO.File]::ReadAllText((Join-Path $root $name), $u8)
  $m = [regex]::Match($tpl, '^@@OUT:(.+?)@@\r?\n')
  if (-not $m.Success) { throw "$name is missing its @@OUT:...@@ header" }
  return [pscustomobject]@{ OutName = $m.Groups[1].Value; Body = $tpl.Substring($m.Length) }
}

# the generated .vbs files have two hard requirements:
#   - no leftover placeholder (means a blob was never substituted)
#   - pure ASCII (WSH reads .vbs as ANSI, so UTF-8 Chinese would be
#     mis-decoded and can even be a syntax error)
function Finish([string]$vbs, [string]$label) {
  # the chunker emits CRLF; keep the output LF-only so it diffs cleanly
  $vbs = $vbs -replace "`r`n", "`n"
  if ($vbs -match '@@[A-Z0-9]+@@') {
    throw "unsubstituted placeholder left in ${label}: $($Matches[0])"
  }
  if ($vbs -match '[^\x00-\x7F]') {
    throw "generated ${label} contains non-ASCII characters"
  }
  return $vbs
}

$dist = Join-Path $root "dist"
if (-not (Test-Path $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }

function Emit([string]$outName, [string]$vbs) {
  [System.IO.File]::WriteAllText((Join-Path $dist $outName), $vbs, $ascii)
  return (Join-Path $dist $outName)
}

# ---------------------------------------------------------------
#  installer: plugin payload + UI messages + workspace markup
# ---------------------------------------------------------------
$gmsB64  = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($gms))
$msgsB64 = MsgBlob "installer_msgs.txt"

# workspace markup: the three XML fragments the installer splices into
# DRAWUIConfig.xml, separated by a "%%SECTION%%" line.  Kept in its own
# file so the toolbar markup can be diffed and re-verified on its own.
$ws = [System.IO.File]::ReadAllText((Join-Path $root "workspace_markup.txt"), $u8)
$ws = ($ws -replace "`r`n", "`n").TrimEnd("`n")
$wsB64 = [Convert]::ToBase64String($u8.GetBytes($ws))

$tpl = ReadTpl "installer_template.txt"
$vbs = $tpl.Body.Replace('@@B64CHUNKS@@', (Chunk $gmsB64 500))
$vbs = $vbs.Replace('@@MSGB64CHUNKS@@', (Chunk $msgsB64 500))
$vbs = $vbs.Replace('@@WSB64CHUNKS@@', (Chunk $wsB64 500))
$vbs = Finish $vbs "installer"
$outVbs = Emit $tpl.OutName $vbs

# silent copy (no dialogs) used for smoke testing
[System.IO.File]::WriteAllText((Join-Path $root "_test_silent.vbs"), ($vbs.Replace('MsgBox ', 'WScript.Echo ')), $ascii)

# ---------------------------------------------------------------
#  uninstaller: UI messages only - it deletes, it does not embed
# ---------------------------------------------------------------
$utpl = ReadTpl "uninstaller_template.txt"
$uvbs = $utpl.Body.Replace('@@MSGB64CHUNKS@@', (Chunk (MsgBlob "uninstaller_msgs.txt") 500))
$uvbs = Finish $uvbs "uninstaller"
$uOutVbs = Emit $utpl.OutName $uvbs

[System.IO.File]::WriteAllText((Join-Path $root "_uninstall_silent.vbs"), ($uvbs.Replace('MsgBox ', 'WScript.Echo ')), $ascii)

if ($gms -ne $distGms) { Copy-Item $gms $distGms -Force }

Write-Output "gms         = $gms ($((Get-Item $gms).Length) bytes)"
Write-Output "messages    = $(($msgsB64.Length)) b64 chars"
Write-Output "workspace   = $($ws.Length) chars, $((($ws -split "`n").Count)) lines"
Write-Output "installer   = $outVbs"
Write-Output "size        = $((Get-Item $outVbs).Length) bytes"
Write-Output "uninstaller = $uOutVbs"
Write-Output "size        = $((Get-Item $uOutVbs).Length) bytes"