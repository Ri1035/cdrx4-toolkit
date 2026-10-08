# click_at.ps1 -- bring CorelDRAW to the foreground, click a screen
# point, and screenshot the result.  Used to prove the flyout opener
# actually opens its menu (the reference plugin uses the same XML
# mechanism, but "same markup" is not "verified").
#
# ASCII only.
param(
  [int]$X = 35,
  [int]$Y = 54,
  [string]$Tag = "click",
  [int]$Clicks = 2
)
$ErrorActionPreference = 'Continue'
$here = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$sig = @'
using System;
using System.Runtime.InteropServices;
public class M {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte sc, uint f, IntPtr e);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern int GetWindowTextW(IntPtr h, System.Text.StringBuilder s, int n);
}
'@
Add-Type -TypeDefinition $sig -ErrorAction SilentlyContinue

$p = Get-Process -Name CorelDRW -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
if ($null -eq $p) { Write-Host "!! CorelDRAW not running"; exit 2 }
$h = $p.MainWindowHandle

# release the foreground lock: a stray ALT keypress makes the next
# SetForegroundWindow succeed even from a background process
[M]::keybd_event(0x12, 0, 0, [IntPtr]::Zero)
[M]::keybd_event(0x12, 0, 2, [IntPtr]::Zero)
[void][M]::ShowWindow($h, 3)
[void][M]::SetForegroundWindow($h)
Start-Sleep -Milliseconds 800
[M]::keybd_event(0x1B, 0, 0, [IntPtr]::Zero)   # ESC closes any open menu
[M]::keybd_event(0x1B, 0, 2, [IntPtr]::Zero)
Start-Sleep -Milliseconds 300

$fg = [M]::GetForegroundWindow()
$sb = New-Object System.Text.StringBuilder 256
[void][M]::GetWindowTextW($fg, $sb, 256)
Write-Host "foreground = $($sb.ToString())"

# click.  The first click on a background window only activates it,
# so default to two.
for ($c = 0; $c -lt $Clicks; $c++) {
  [void][M]::SetCursorPos($X, $Y)
  Start-Sleep -Milliseconds 300
  [M]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)   # LEFTDOWN
  Start-Sleep -Milliseconds 80
  [M]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)   # LEFTUP
  Start-Sleep -Milliseconds 700
}
Start-Sleep -Seconds 1

$b = [System.Windows.Forms.SystemInformation]::VirtualScreen
$bmp = New-Object System.Drawing.Bitmap($b.Width, $b.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($b.Left, $b.Top, 0, 0, $bmp.Size)
$g.Dispose()
$out = Join-Path $here ("_" + $Tag + ".png")
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "shot -> _$Tag.png"