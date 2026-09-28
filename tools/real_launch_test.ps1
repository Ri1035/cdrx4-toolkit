# ==========================================================
#  CDRX4Toolkit -- real double-click launch test
#
#  Why this exists (PLAN.md section 10.4):
#  the startup hook only runs when CorelDRAW is started the way a
#  USER starts it.  CreateObject / InitializeVBA / RunMacro all
#  take a different path through Corel's boot sequence and have
#  been proven NOT to reproduce the crash.  So the one and only
#  acceptable proof for "does launching CDR still die" is to
#  start CorelDRW.exe and watch the process.
#
#  What it does:
#    1. quit any CorelDRAW that is already running
#    2. note the newest Application Error (id 1000) in the event log
#    3. start CorelDRW.exe by hand
#    4. sample threads / handles / working set / Responding for N sec
#    5. look for NEW Application Error events naming CorelDRW
#    6. screenshot (so a human can see whether the toolbar showed up)
#    7. quit CorelDRAW again
#
#  Verdict is printed as PASS / FAIL.  Log -> _launch.log (UTF-16).
#
#  ASCII ONLY on purpose: keep the file portable and avoid
#  PowerShell reading it under the wrong codepage.
# ==========================================================

param(
  [int]$WaitSec = 30,
  [string]$Exe = "C:\Program Files (x86)\CorelDRAW X4\Programs\CorelDRW.exe",
  [string]$Shot = "",
  [switch]$KeepRunning
)

$ErrorActionPreference = 'Continue'

$here   = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$logPath = Join-Path $here '_launch.log'
$script:lines = New-Object System.Collections.ArrayList

function L([string]$s) {
  [void]$script:lines.Add($s)
  Write-Host $s
}

function Flush {
  try { $script:lines -join "`r`n" | Set-Content -Path $logPath -Encoding Unicode } catch {}
}

function NewCrashEvents([datetime]$since) {
  # Application Error, id 1000, CorelDRW -- the "0xc0000005 / 0xc000041d" pair
  $out = @()
  try {
    $out = Get-WinEvent -FilterHashtable @{
      LogName      = 'Application'
      ProviderName = 'Application Error'
      StartTime    = $since
    } -ErrorAction SilentlyContinue |
      Where-Object { $_.Message -match 'CorelDRW' }
  } catch {}
  return $out
}

function SampleProc {
  $p = Get-Process -Name CorelDRW -ErrorAction SilentlyContinue |
       Sort-Object WorkingSet -Descending | Select-Object -First 1
  if ($null -eq $p) { return $null }
  $th = 0
  try { $th = $p.Threads.Count } catch {}
  return [pscustomobject]@{
    Id         = $p.Id
    Threads    = $th
    Handles    = $p.HandleCount
    WorkingSet = $p.WorkingSet64
    Responding = $p.Responding
  }
}

L "==== real_launch_test $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') ===="
L "exe  = $Exe"
L "wait = $WaitSec s"

if (-not (Test-Path $Exe)) { L "!! exe not found"; Flush; exit 2 }

# --- 1. quit anything already running -----------------------------
$app = $null
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application.14') } catch {}
if ($null -ne $app) {
  L "[1] quitting the running CorelDRAW ..."
  try { $app.Quit() } catch {}
  for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 500
    if (-not (Get-Process -Name CorelDRW -ErrorAction SilentlyContinue)) { break }
  }
}
$left = @(Get-Process -Name CorelDRW -ErrorAction SilentlyContinue)
if ($left.Count -gt 0) {
  foreach ($p in $left) {
    $t = 0; try { $t = $p.Threads.Count } catch {}
    L "    still alive: pid=$($p.Id) threads=$t ws=$($p.WorkingSet64)"
  }
  L "    (a 0-thread zombie from an earlier crash cannot be killed; it does not block a new launch)"
}

# --- 2. baseline event log ---------------------------------------
$t0 = (Get-Date).AddSeconds(-2)
$before = NewCrashEvents $t0
L "[2] baseline Application Error (CorelDRW) events in window = $($before.Count)"

# --- 3. start it by hand -----------------------------------------
L "[3] Start-Process $Exe"
try { Start-Process -FilePath $Exe } catch { L "!! start failed: $_"; Flush; exit 3 }

# --- 4. watch it --------------------------------------------------
$bad = 0
$samples = @()
for ($t = 5; $t -le $WaitSec; $t += 5) {
  Start-Sleep -Seconds 5
  $s = SampleProc
  if ($null -eq $s) { L "  +${t}s : process GONE"; $bad = 1; break }
  $samples += $s
  L ("  +{0}s : pid={1} threads={2} handles={3} ws={4:N1}MB responding={5}" -f `
      $t, $s.Id, $s.Threads, $s.Handles, ($s.WorkingSet / 1MB), $s.Responding)
}

# --- 5. new crashes? ---------------------------------------------
$after = NewCrashEvents $t0
$new = @($after | Where-Object { $before -notcontains $_ })
L "[5] NEW Application Error (CorelDRW) events = $($new.Count)"
foreach ($e in $new) {
  $first = ($e.Message -split "`n")[0]
  L "    $($e.TimeCreated) : $first"
}

# --- 6. screenshot -----------------------------------------------
if ($Shot -ne "") {
  try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $b = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $bmp = New-Object System.Drawing.Bitmap($b.Width, $b.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($b.Left, $b.Top, 0, 0, $bmp.Size)
    $bmp.Save($Shot, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    L "[6] screenshot -> $Shot"
  } catch { L "[6] screenshot failed: $_" }
}

# --- 7. quit -----------------------------------------------------
if (-not $KeepRunning) {
  L "[7] quitting CorelDRAW ..."
  $a2 = $null
  try { $a2 = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application.14') } catch {}
  if ($null -ne $a2) { try { $a2.Quit() } catch {} }
  for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 500
    if (-not (Get-Process -Name CorelDRW -ErrorAction SilentlyContinue)) { break }
  }
}

# --- verdict -----------------------------------------------------
$healthy = $false
if ($samples.Count -gt 0) {
  $last = $samples[-1]
  if ($last.Threads -gt 4 -and $last.WorkingSet -gt 40MB -and $last.Responding) { $healthy = $true }
}
L ""
if ($healthy -and $new.Count -eq 0) {
  L "==== LAUNCH PASS ====  CorelDRAW started by hand and stayed healthy"
} else {
  L "==== LAUNCH FAIL ====  healthy=$healthy newCrashEvents=$($new.Count)"
}
Flush
if ($healthy -and $new.Count -eq 0) { exit 0 } else { exit 1 }