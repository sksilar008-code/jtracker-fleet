<#
.SYNOPSIS
  RUN-PIPELINE : the full chain, designed to be run by Task Scheduler 4x/day.
  1) HuntBot  - find new BIM jobs, add to JobTracker.xlsx
  2) ApplyBot - apply to every job in the tracker with a direct link
#>
param([switch]$DryRun)

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logDir = Join-Path $ProjectDir "logs"
New-Item -ItemType Directory -Force $logDir | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmm"
Start-Transcript -Path (Join-Path $logDir "pipeline-$stamp.log") -Append | Out-Null

Write-Host "======================================"
Write-Host " SEARCH MACHINE - 2  |  pipeline run  $stamp"
Write-Host "======================================"

# Fleet-wide run lock: scheduled tasks (HuntBot-min, SearchMachine2-Pipeline) and manual
# runs must NEVER overlap - concurrent fleets fight over the browser debug port and their
# EPPlus saves clobber tracker rows (triple collision on 10/27 at 12:00). Children
# (HuntBot) see the env var and skip their own lock exit/removal.
$env:JTRACKER_PIPELINE_RUN = $stamp
# The lock lives on LOCAL disk: a lock inside OneDrive proved unreliable - OneDrive holds
# the file open for sync and Remove-Item silently failed at 14:26, locking the fleet out
# for hours. Local disk = no sync engine, no cloud hydrate, no failed deletes.
$runLock = Join-Path $env:LOCALAPPDATA "JTrackerEdge\fleet.lock"
New-Item -ItemType Directory -Force (Split-Path $runLock) | Out-Null
if (Test-Path $runLock) {
    $lockAge = (Get-Date) - (Get-Item $runLock).LastWriteTime
    if ($lockAge.TotalMinutes -lt 120) {
        Write-Host "Another fleet run is active (lock $([int]$lockAge.TotalMinutes) min old) - exiting." -ForegroundColor Yellow
        Stop-Transcript | Out-Null
        exit
    }
}
Set-Content $runLock -Value "pipeline $stamp started $(Get-Date -Format o)" -Encoding UTF8

# One-shot drain mode: if drain-pending.flag exists, this run sends ALL pending
# (EmailBot + FeedBot + LeadsBatch uncapped), then the flag is consumed.
$drainFlag = Join-Path $ProjectDir "drain-pending.flag"
$drain = Test-Path $drainFlag
if ($drain) { Write-Host "*** DRAIN MODE: sending all pending this run ***" -ForegroundColor Yellow }

# Every bot runs in its OWN powershell.exe child process. The 10/27 experiment proved the
# first CDP browser session in a PowerShell process works, but every later session's
# socket dies almost immediately (while the browser itself survives) - the pipeline used
# to run all four bots in one process (6+ sessions), which is why everything after Indeed
# died. Fresh process per step = healthy sockets every time. Child stdout still reaches
# this console; each bot keeps its own log file as the durable record.
function Invoke-BotStep {
    param([string]$Script, [string[]]$BotArgs = @())
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ProjectDir $Script) @BotArgs
    if ($LASTEXITCODE) { Write-Host "  ($Script exited with code $LASTEXITCODE)" -ForegroundColor DarkYellow }
}

Write-Host "`n>>> STEP 1/5 : HuntBot (job search)"
try { Invoke-BotStep "HuntBot.ps1" } catch { Write-Host "HuntBot error: $($_.Exception.Message)" -ForegroundColor Red }

Write-Host "`n>>> STEP 2/5 : ApplyBot (auto apply)"
try {
    if ($DryRun) { Invoke-BotStep "ApplyBot.ps1" @('-DryRun') }
    else         { Invoke-BotStep "ApplyBot.ps1" }
} catch { Write-Host "ApplyBot error: $($_.Exception.Message)" -ForegroundColor Red }

Write-Host "`n>>> STEP 3/4 : TailoredApplyBatch (JD-matched CV + email)"
try {
    if ($DryRun) { Write-Host "(dry run - mail batch skipped)" }
    elseif ($drain) { Invoke-BotStep "TailoredApplyBatch.ps1" @('-MaxSend','30') }
    else { Invoke-BotStep "TailoredApplyBatch.ps1" @('-MaxSend','8') }
} catch { Write-Host "TailoredApply error: $($_.Exception.Message)" -ForegroundColor Red }

Write-Host "`n>>> STEP 4/4 : FeedBot (LinkedIn feed posts)"
try {
    if ($drain) { Invoke-BotStep "FeedBot.ps1" @('-MaxPosts','10') }
    else { Invoke-BotStep "FeedBot.ps1" @('-MaxPosts','8') }
} catch { Write-Host "FeedBot error: $($_.Exception.Message)" -ForegroundColor Red }

if ($drain) { Remove-Item $drainFlag -Force; Write-Host "Drain flag consumed - normal caps resume next run." }
for ($r = 0; $r -lt 5 -and (Test-Path $runLock); $r++) {
    try { Remove-Item $runLock -Force -ErrorAction Stop; break } catch { Start-Sleep -Seconds 2 }
}

Write-Host "`nPipeline finished $(Get-Date -Format 'HH:mm:ss')"
Stop-Transcript | Out-Null
