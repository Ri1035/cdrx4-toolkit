# ==========================================================
#  capture_cdr.ps1 -- start CorelDRAW by hand, foreground it,
#  and capture BOTH the full desktop and the CorelDRAW window.
#
#  Why not real_launch_test.ps1: that one screenshots whatever
#  happens to be on screen; if TRAE (or anything else) is in
#  front, the shot is useless.  This one forces CorelDRAW to
#  the foreground first and also grabs the window directly via
#  PrintWindow, so an occluded toolbar still lands in the image.
#
#  ASCII only.
# ==========================================================
param(
  [int]$WaitSec = 20,
  [string]$Exe = "C:\Program Files (x86)\CorelDRAW X4\Programs\CorelDRW.exe",
  [string]$Tag = "shot",
  [switch]$KeepRunning
)

$ErrorActionPreference = 'Continue'
$here = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$sig = @'
using System;
using System.Runtime.InteropServices;
public class W {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
  [DllImport("user32.dll")] public static extern int GetWindowTextW(IntPtr h, System.Text.StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int n);
  public delegate bool EnumProc(IntPtr h, IntPtr p);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
}
'@
Add-Type -TypeDefinition $sig -ErrorAction SilentlyContinue

function QuitCdr {
  $a = $null
  try { $a = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application.14') } catch {}
  if ($null -ne $a) { try { $a.Quit() } catch {} }
  for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 500
    if (-not (Get-Process -Name CorelDRW -ErrorAction SilentlyContinue)) { break }
  }
}

function GrabFull([string]$out) {
  $b = [System.Windows.Forms.SystemInformation]::VirtualScreen
  $bmp = New-Object System.Drawing.Bitmap($b.Width, $b.Height)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($b.Left, $b.Top, 0, 0, $bmp.Size)
  $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
}

function GrabWin([IntPtr]$h, [string]$out) {
  $r = New-Object W+RECT
  if (-not [W]::GetWindowRect($h, [ref]$r)) { return $false }
  $w = $r.R - $r.L; $ht = $r.B - $r.T
  if ($w -le 0 -or $ht -le 0) { return $false }
  $bmp = New-Object System.Drawing.Bitmap($w, $ht)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $hdc = $g.GetHdc()
  $ok = [W]::PrintWindow($h, $hdc, 2)   # PW_RENDERFULLCONTENT
  $g.ReleaseHdc($hdc)
  $g.Dispose()
  $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()
  return $ok
}

Write-Host "quitting any running CorelDRAW ..."
QuitCdr
Write-Host "starting $Exe ..."
Start-Process -FilePath $Exe | Out-Null

$h = [IntPtr]::Zero
for ($i = 0; $i -lt ($WaitSec * 2); $i++) {
  Start-Sleep -Milliseconds 500
  $p = Get-Process -Name CorelDRW -ErrorAction SilentlyContinue |
       Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
  if ($null -ne $p) { $h = $p.MainWindowHandle; break }
}
if ($h -eq [IntPtr]::Zero) { Write-Host "!! no main window"; if (-not $KeepRunning) { QuitCdr }; exit 2 }

Start-Sleep -Seconds 6
[W]::ShowWindow($h, 3) | Out-Null     # SW_MAXIMIZE
[W]::SetForegroundWindow($h) | Out-Null
Start-Sleep -Seconds 3

# list every visible top-level window of the process, so floating
# toolbars that live in their own window are not missed
$procIds = @(Get-Process -Name CorelDRW -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
Write-Host "---- top-level windows ----"
$script:found = @()
$cb = [W+EnumProc]{
  param($wh, $lp)
  $pid2 = 0
  [void][W]::GetWindowThreadProcessId($wh, [ref]$pid2)
  if ($procIds -contains [int]$pid2 -and [W]::IsWindowVisible($wh)) {
    $sb = New-Object System.Text.StringBuilder 256
    [void][W]::GetWindowTextW($wh, $sb, 256)
    $cn = New-Object System.Text.StringBuilder 256
    [void][W]::GetClassNameW($wh, $cn, 256)
    $r = New-Object W+RECT
    [void][W]::GetWindowRect($wh, [ref]$r)
    $script:found += [pscustomobject]@{ H = $wh; Title = $sb.ToString(); Class = $cn.ToString(); X = $r.L; Y = $r.T; W = ($r.R - $r.L); Hh = ($r.B - $r.T) }
  }
  return $true
}
[void][W]::EnumWindows($cb, [IntPtr]::Zero)
foreach ($f in $script:found) {
  Write-Host ("  [{0}] {1} x{2} y{3} {4}x{5}  {6}" -f $f.Class, $f.Title, $f.X, $f.Y, $f.W, $f.Hh, $f.H)
}

GrabFull (Join-Path $here ("_" + $Tag + "_desktop.png"))
Write-Host "desktop -> _$Tag`_desktop.png"
if (GrabWin $h (Join-Path $here ("_" + $Tag + "_win.png"))) { Write-Host "window  -> _$Tag`_win.png" } else { Write-Host "PrintWindow failed" }

if (-not $KeepRunning) { Write-Host "quitting ..."; QuitCdr }
Write-Host "done"