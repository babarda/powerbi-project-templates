# Developer helper: open a template PBIP in a fresh Power BI Desktop instance, refresh it from its
# DataFile, screenshot every page, then close that instance. Other Desktop instances are left alone.
# Usage: .\render.ps1 -Pbip "..\progress-s-curve\Progress S-Curve.pbip" -OutDir ..\_shots -Pages scurve
# Author: Badr Abardazzou. License: MIT.
param([string]$Pbip, [string]$OutDir, [string[]]$Pages, [int]$Scale = 2, [switch]$KeepOpen)
$ErrorActionPreference = "Stop"
$env:PATH = "$env:USERPROFILE\nodejs;" + $env:PATH
$Pbip = (Resolve-Path $Pbip).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$before = @(Get-Process PBIDesktop -ErrorAction SilentlyContinue | ForEach-Object Id)
Start-Process $Pbip
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
public class RWin { public delegate bool P(IntPtr h, IntPtr l);
[DllImport("user32.dll")] public static extern bool EnumWindows(P f, IntPtr l);
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
public static List<string> Titles(){ var r=new List<string>(); EnumWindows((h,l)=>{int pid; GetWindowThreadProcessId(h,out pid); if(IsWindowVisible(h)){var sb=new StringBuilder(256); GetWindowText(h,sb,256); r.Add(pid+"|"+sb.ToString());} return true;}, IntPtr.Zero); return r;} }
"@
$desk = $null
for ($k = 0; $k -lt 60; $k++) {
  Start-Sleep 4
  $new = @(Get-CimInstance Win32_Process -Filter "Name='PBIDesktop.exe'" | Where-Object { $before -notcontains $_.ProcessId -and $_.CommandLine -like "*$Pbip*" } | ForEach-Object ProcessId)
  $bad = [RWin]::Titles() | Where-Object { $t = $_ -split '\|', 2; ($new -contains [int]$t[0]) -and ($t[1] -match 'Unable|went wrong|Error') }
  if ($bad) { $new | ForEach-Object { Stop-Process -Id $_ -Force }; throw "Desktop error dialog: $bad" }
  $job = Start-Job { $env:PATH = "$env:USERPROFILE\nodejs;" + $env:PATH; powerbi-desktop status }
  if (-not (Wait-Job $job -Timeout 25)) { Stop-Job $job; Remove-Job $job -Force; continue }
  try { $st = ((Receive-Job $job -ErrorAction SilentlyContinue 2>$null) | Where-Object { $_ -is [string] }) -join "`n" | ConvertFrom-Json } catch { $st = $null }
  Remove-Job $job -Force
  $desk = $st.instances | Where-Object { $_.currentFilePath -eq $Pbip -and $_.bridgeStatus -eq 'connected' -and ($before -notcontains $_.pid) } | Select-Object -First 1
  if ($desk) { break }
}
if (-not $desk) { throw "Desktop did not open $Pbip" }
$dpid = [int]$desk.pid
"pid=$dpid"
Add-Type @"
using System; using System.Runtime.InteropServices;
public class RW { [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int hh, bool r); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c); }
"@
& (Join-Path $PSScriptRoot "refresh_model.ps1") -DesktopPid $dpid | Out-Null
Start-Sleep 8
$hw = (Get-Process -Id $dpid).MainWindowHandle
[RW]::ShowWindow($hw, 9) | Out-Null
[RW]::MoveWindow($hw, 0, 0, 1900, 1150, $true) | Out-Null
Start-Sleep 6
foreach ($p in $Pages) {
  $out = Join-Path $OutDir "$p.png"
  $ErrorActionPreference = "Continue"
  powerbi-desktop screenshot $p --pid $dpid --output $out --scale 1 --wait-seconds 120 2>$null | Out-Null
  Start-Sleep 40
  for ($t = 0; $t -lt 3; $t++) { powerbi-desktop screenshot $p --pid $dpid --output $out --scale $Scale --wait-seconds 120 2>$null | Out-Null; if ($LASTEXITCODE -eq 0) { break }; Start-Sleep 10 }
  "shot $out"
}
if (-not $KeepOpen) { Stop-Process -Id $dpid -Force }
