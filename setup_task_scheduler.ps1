# =============================================================================
# CSU Advisory Alert - Windows Task Scheduler Setup Script
# =============================================================================
# HOW TO RUN:
#   Right-click this file -> "Run with PowerShell" (as Administrator)
#   OR open PowerShell as Admin and run:
#       cd D:\github-clg-project\pythoncodetest
#       .\setup_task_scheduler.ps1
#
# INTERVAL OPTIONS (edit $IntervalMinutes and $RunMode below):
#   - RunMode "247"     : Every X minutes, 24 hours a day, 7 days a week (default: 30 min)
#   - RunMode "workhrs" : Every X minutes but only 08:00-20:00 (working hours only)
#
# FLOW PER RUN:
#   Task Scheduler calls python advisory_alert_runner.py
#     -> Excel macro RefreshAll runs first (pulls latest SharePoint/Power Query data)
#     -> Waits ~20 seconds for refresh to complete
#     -> Opens Advisory sheet (read-only)
#     -> Finds latest Advisory Preparation Date batch
#     -> Hashes rows, skips already-sent, sends only NEW alerts
#     -> Exits cleanly (next run handled by Task Scheduler)
# =============================================================================

# ── CONFIGURE THESE ──────────────────────────────────────────────────────────
$PythonPath       = "D:\github-clg-project\pythoncodetest\.venv\Scripts\python.exe"
$ScriptPath       = "D:\github-clg-project\pythoncodetest\advisory_alert_runner.py"
$WorkingDir       = "D:\github-clg-project\pythoncodetest"
$TaskName         = "CSU_AdvisoryAlert"
$IntervalMinutes  = 30        # How often to check (every N minutes)
$RunMode          = "247"     # "247" = all day all night | "workhrs" = 08:00-20:00 only
# ─────────────────────────────────────────────────────────────────────────────

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "  CSU Advisory Alert - Task Scheduler Setup" -ForegroundColor Cyan
Write-Host "  Mode     : $RunMode (interval = every $IntervalMinutes min)" -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

if (-not (Test-Path $PythonPath)) { Write-Host "[ERROR] Python not found: $PythonPath" -ForegroundColor Red; exit 1 }
if (-not (Test-Path $ScriptPath)) { Write-Host "[ERROR] Script not found: $ScriptPath" -ForegroundColor Red; exit 1 }
Write-Host "[OK] Python : $PythonPath" -ForegroundColor Green
Write-Host "[OK] Script : $ScriptPath" -ForegroundColor Green

# Remove old task if already registered
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false; Write-Host "[*] Removed old task '$TaskName'." -ForegroundColor Yellow }

$Username    = "$env:USERDOMAIN\$env:USERNAME"
$TempXml     = "$env:TEMP\csu_advisory_task.xml"
$IntervalISO = "PT${IntervalMinutes}M"

# Duration: 24/7 mode = PT24H (full day), working hours = PT12H
if ($RunMode -eq "247") {
    $DurationISO  = "PT24H"
    $StartTime    = (Get-Date).Date.ToString("yyyy-MM-dd") + "T00:00:00"  # midnight = always running
} else {
    $DurationISO  = "PT12H"
    $StartTime    = (Get-Date).Date.ToString("yyyy-MM-dd") + "T08:00:00"  # 08:00-20:00 window
}

Write-Host "[*] Interval : every $IntervalMinutes min | Duration : $DurationISO | Start : $StartTime" -ForegroundColor Cyan

$TaskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>CSU TI Advisory Alert Automation. Checks every $IntervalMinutes min ($RunMode). Runs Excel macro refresh then reads latest Advisory date batch and sends alert emails.</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <Repetition>
        <Interval>$IntervalISO</Interval>
        <Duration>$DurationISO</Duration>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <StartBoundary>$StartTime</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$Username</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT5M</ExecutionTimeLimit>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <Enabled>true</Enabled>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$PythonPath</Command>
      <Arguments>"$ScriptPath"</Arguments>
      <WorkingDirectory>$WorkingDir</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

[System.IO.File]::WriteAllText($TempXml, $TaskXml, [System.Text.Encoding]::Unicode)
$Result = & schtasks.exe /Create /TN $TaskName /XML $TempXml /F 2>&1
Remove-Item $TempXml -ErrorAction SilentlyContinue

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "[OK] Task '$TaskName' registered successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Useful commands:" -ForegroundColor Cyan
    Write-Host "  Run now   : schtasks /Run /TN $TaskName" -ForegroundColor White
    Write-Host "  Check log : schtasks /Query /TN $TaskName /FO LIST /V" -ForegroundColor White
    Write-Host "  Delete    : schtasks /Delete /TN $TaskName /F" -ForegroundColor White
} else {
    Write-Host "[ERROR] Failed to register task. Make sure you are running as Administrator." -ForegroundColor Red
    Write-Host "        Details: $Result" -ForegroundColor Red
    exit 1
}
Write-Host "=================================================================" -ForegroundColor Cyan
