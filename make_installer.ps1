$ErrorActionPreference = "Stop"

# This script is intentionally ASCII-only.
# All Chinese text lives in installer_template.txt / installer_msgs.txt,
# which we read explicitly as UTF-8.
#
#   build_gms.vbs          -> writes CDRX4Toolkit.gms into the X4 user GMS folder
#   make_installer.ps1     -> embeds that GMS + the UI messages into
#                             dist\安装CDRX4增强工具.vbs (single self-contained file)

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

# plugin payload
$gmsB64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($gms))

# UI messages (keep the file's line structure, drop the trailing newline)
$msgs = [System.IO.File]::ReadAllText((Join-Path $root "installer_msgs.txt"), $u8)
$msgs = ($msgs -replace "`r`n", "`n").TrimEnd("`n")
$msgsB64 = [Convert]::ToBase64String($u8.GetBytes($msgs))

$tpl = [System.IO.File]::ReadAllText((Join-Path $root "installer_template.txt"), $u8)
$m = [regex]::Match($tpl, '^@@OUT:(.+?)@@\r?\n')
if (-not $m.Success) { throw "template missing @@OUT:...@@ header" }
$outName = $m.Groups[1].Value
$body = $tpl.Substring($m.Length)

$vbs = $body.Replace('@@B64CHUNKS@@', (Chunk $gmsB64 500))
$vbs = $vbs.Replace('@@MSGB64CHUNKS@@', (Chunk $msgsB64 500))

# keep the output LF-only: the chunker above emits CRLF, and a file with
# mixed line endings is both ugly in git and annoying to diff
$vbs = $vbs -replace "`r`n", "`n"

# the generated installer must stay pure ASCII, otherwise WSH would
# mis-decode it on a Chinese Windows (ANSI) system
if ($vbs -match '[^\x00-\x7F]') {
  throw "generated installer contains non-ASCII characters"
}

$dist = Join-Path $root "dist"
if (-not (Test-Path $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }

$outVbs = Join-Path $dist $outName
[System.IO.File]::WriteAllText($outVbs, $vbs, $ascii)

# silent copy (no dialogs) used for smoke testing
$testVbs = $vbs.Replace('MsgBox ', 'WScript.Echo ')
[System.IO.File]::WriteAllText((Join-Path $root "_test_silent.vbs"), $testVbs, $ascii)

if ($gms -ne $distGms) { Copy-Item $gms $distGms -Force }

Write-Output "gms        = $gms ($((Get-Item $gms).Length) bytes)"
Write-Output "messages   = $($msgs.Length) chars, $((($msgs -split "`n").Count)) lines"
Write-Output "installer  = $outVbs"
Write-Output "size       = $((Get-Item $outVbs).Length) bytes"