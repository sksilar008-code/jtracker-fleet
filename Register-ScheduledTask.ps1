<#
.SYNOPSIS
  Registers the Search Machine - 2 pipeline to run 4 times a day (every 6 hours)
  in Windows Task Scheduler. Run this once. To remove:
    Unregister-ScheduledTask -TaskName "SearchMachine2-Pipeline" -Confirm:$false
#>
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pipeline = Join-Path $ProjectDir "Run-Pipeline.ps1"

# schtasks.exe is the most reliable way on Windows PowerShell 5.1
# /SC HOURLY /MO 6 = every 6 hours = 4 runs per day, starting at 06:00
# Note: \" escaping is required for paths with spaces on Windows PowerShell 5.1
$cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"' + $pipeline + '\"'
& schtasks /Create /TN "SearchMachine2-Pipeline" /TR $cmd /SC HOURLY /MO 6 /ST 06:00 /F | Out-Null
if ($LASTEXITCODE -ne 0) { Write-Host "schtasks failed (code $LASTEXITCODE)" -ForegroundColor Red; exit 1 }

Write-Host ""
Write-Host "Scheduled task 'SearchMachine2-Pipeline' registered." -ForegroundColor Green
Write-Host "Runs every 6 hours starting 06:00 (4x/day), as user $env:USERNAME."
Write-Host "Test it now with:  Start-ScheduledTask -TaskName 'SearchMachine2-Pipeline'"
Write-Host "Remove it with:    Unregister-ScheduledTask -TaskName 'SearchMachine2-Pipeline' -Confirm:`$false"
