<#
.SYNOPSIS
  Persists everything that changed during a cloud run back into the PRIVATE state
  repo and pushes it (tracker + Tailored folders + CVS are written in place inside
  the state checkout; only the small volatile files are copied back).
  Also keeps the PUBLIC code repo's schedule alive: GitHub disables cron workflows
  after 60 days with no repository activity, so we push an empty keepalive commit
  when the code repo has been quiet for >6 days.
#>
param(
    [Parameter(Mandatory)][string]$StateDir,
    [string]$ProjectDir = ""
)
$ErrorActionPreference = "Stop"
# NOTE: $PSScriptRoot is empty in param defaults when an earlier param is Mandatory
# (PS5.1 quirk) - resolve in the body instead
if (-not $ProjectDir) { $ProjectDir = $PSScriptRoot }
$StateDir = (Resolve-Path $StateDir).Path
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm"

# ---- 1) volatile working files back into state ----
$stateEp = Join-Path $StateDir "edge-profile"
New-Item -ItemType Directory -Force -Path $stateEp | Out-Null
$ep = Join-Path $ProjectDir "edge-profile"
foreach ($f in @("cookies.json", "feedbot-history.json", "feedbot-rotation.txt")) {
    $p = Join-Path $ep $f
    if (Test-Path $p) { Copy-Item $p (Join-Path $stateEp $f) -Force }
}
$stateLogs = Join-Path $StateDir "logs"
New-Item -ItemType Directory -Force -Path $stateLogs | Out-Null
$logDir = Join-Path $ProjectDir "logs"
Get-ChildItem $logDir -Filter "email-ledger-*.jsonl" -ErrorAction SilentlyContinue | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $stateLogs $_.Name) -Force
}
$ptu = Join-Path $logDir "pending-tracker-updates.json"
if (Test-Path $ptu) { Copy-Item $ptu (Join-Path $stateLogs "pending-tracker-updates.json") -Force }

# ---- 2) commit + push state (remote already carries the token from the clone step) ----
# NOTE: git writes push progress ("To https://...", branch lines) to STDERR. Under
# $ErrorActionPreference="Stop" PowerShell turns that into a fatal NativeCommandError and
# kills the script AFTER a successful push (broke run #2). Go permissive for git calls;
# real failures are caught via $LASTEXITCODE throws below.
$ErrorActionPreference = "Continue"
Push-Location $StateDir
try {
    git add -A 2>&1 | Write-Host
    $dirty = @(git status --porcelain 2>&1)
    $dirty = @($dirty | Where-Object { $_ -and $_ -notmatch '^!!' })
    if ($dirty.Count -gt 0) {
        git -c user.name="jtracker-bot" -c user.email="jtracker-bot@users.noreply.github.com" `
            commit -m "state update $stamp" 2>&1 | Write-Host
        if ($LASTEXITCODE) { throw "state commit failed (exit $LASTEXITCODE)" }
        git push 2>&1 | Write-Host
        if ($LASTEXITCODE) { throw "state push failed (exit $LASTEXITCODE)" }
        Write-Host "STATE PUSHED: $($dirty.Count) changed path(s)"
    } else {
        Write-Host "state: no changes"
    }
} finally { Pop-Location }

# ---- 3) keepalive on the code repo (quiet >6 days -> empty commit) ----
try {
    $lastRaw = git -C $ProjectDir log -1 --format=%ct 2>$null
    if ($lastRaw) {
        $last = [long]("$lastRaw").Trim()
        $age = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - $last
        if ($age -gt (6 * 86400)) {
            git -C $ProjectDir -c user.name="jtracker-bot" -c user.email="jtracker-bot@users.noreply.github.com" `
                commit --allow-empty -m "chore: schedule keepalive $stamp" 2>&1 | Write-Host
            git -C $ProjectDir push 2>&1 | Write-Host
            if ($LASTEXITCODE) { Write-Host "keepalive push failed (non-fatal)" }
            else { Write-Host "code repo keepalive pushed" }
        }
    }
} catch { Write-Host "keepalive skipped: $($_.Exception.Message)" }

Write-Host "Save-CloudState done"
