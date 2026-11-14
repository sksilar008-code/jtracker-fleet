<#
.SYNOPSIS
  Prepares the fleet workspace for a GitHub Actions (cloud) run.
  1. Copies the REAL config from the private state repo over the sanitized code-repo one
  2. Restores session state (cookies, feedbot history/rotation, email ledger)
  3. Rewrites config.json paths to the runner-local state directory
     (regex on raw text - avoids PS5.1 ConvertTo-Json single-item-array corruption)
  4. Sets JTRACKER_CLOUD=1 (headless Edge)
  StateDir = checkout of the PRIVATE jtracker-state repo (cloned by the workflow).
#>
param(
    [Parameter(Mandatory)][string]$StateDir,
    [string]$ProjectDir = ""
)
$ErrorActionPreference = "Stop"
# NOTE: $PSScriptRoot evaluates EMPTY in a param default when an earlier param is
# Mandatory (PS5.1 quirk, broke the first cloud run) - resolve it in the body instead
if (-not $ProjectDir) { $ProjectDir = $PSScriptRoot }
$StateDir = (Resolve-Path $StateDir).Path
Write-Host "Init-CloudEnv: state=$StateDir"
Write-Host "Init-CloudEnv: project=$ProjectDir"

# ---- 1) real config from state (the code repo only ships a sanitized template) ----
$stateCfg = Join-Path $StateDir "config.json"
if (Test-Path $stateCfg) {
    Copy-Item $stateCfg (Join-Path $ProjectDir "config.json") -Force
} else {
    throw "state config.json missing - seed the state repo first (CLOUD-SETUP.md, step 3)"
}
# remember the CV filename the user actually configured before we rewrite the path
$cvLeaf = "Silar cv.pdf"
try { $cvLeaf = Split-Path ((Get-Content (Join-Path $ProjectDir "config.json") -Raw | ConvertFrom-Json).profile.defaultCv) -Leaf } catch {}

# ---- 2) session / history state into their working positions ----
$ep = Join-Path $ProjectDir "edge-profile"
New-Item -ItemType Directory -Force -Path $ep | Out-Null
$stateEp = Join-Path $StateDir "edge-profile"
foreach ($f in @("cookies.json", "feedbot-history.json", "feedbot-rotation.txt")) {
    $s = Join-Path $stateEp $f
    if (Test-Path $s) { Copy-Item $s (Join-Path $ep $f) -Force; Write-Host "  restored $f" }
    else { Write-Host "  (state has no $f yet)" }
}
# email ledger: keeps the shared 8/day cap across cloud runs
$logDir = Join-Path $ProjectDir "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$stateLogs = Join-Path $StateDir "logs"
if (Test-Path $stateLogs) {
    Get-ChildItem $stateLogs -Filter "email-ledger-*.jsonl" -ErrorAction SilentlyContinue | ForEach-Object {
        Copy-Item $_.FullName (Join-Path $logDir $_.Name) -Force
    }
}
# pending tracker updates fallback queue from the previous run
$ptu = Join-Path $stateLogs "pending-tracker-updates.json"
if (Test-Path $ptu) { Copy-Item $ptu (Join-Path $logDir "pending-tracker-updates.json") -Force }

# ---- 3) rewrite config paths to the state dir (raw-text regex, PS5.1-safe) ----
$cfgPath = Join-Path $ProjectDir "config.json"
$raw = Get-Content $cfgPath -Raw
# JSON-escape the backslashes ONCE (raw config text stores paths doubled, e.g. C:\\Users)
$esc = $StateDir -replace '\\', '\\'
$raw = [regex]::Replace($raw, '("tracker"\s*:\s*")[^"]*(")',     ('$1' + $esc + '\\JobTracker.xlsx$2'))
$raw = [regex]::Replace($raw, '("root"\s*:\s*")[^"]*(")',       ('$1' + $esc + '$2'))
$raw = [regex]::Replace($raw, '("defaultCv"\s*:\s*")[^"]*(")',  ('$1' + $esc + '\\CVS\\' + ($cvLeaf -replace '\\','\\\\') + '$2'))
Set-Content -Path $cfgPath -Value $raw -Encoding UTF8

# ---- 4) verify the rewrite actually parsed and the tracker exists ----
$chk = Get-Content $cfgPath -Raw | ConvertFrom-Json
if ($chk.paths.tracker -notlike "$StateDir*") { throw "tracker rewrite failed -> $($chk.paths.tracker)" }
if (-not (Test-Path $chk.paths.tracker))      { throw "state tracker missing -> $($chk.paths.tracker)" }
if (-not (Test-Path $chk.profile.defaultCv))  { throw "state CV missing -> $($chk.profile.defaultCv)" }

# ---- 5) cloud flag for this process tree (workflow env covers the rest) ----
$env:JTRACKER_CLOUD = "1"

Write-Host "Init-CloudEnv OK"
Write-Host "  tracker = $($chk.paths.tracker)"
Write-Host "  root    = $($chk.paths.root)"
Write-Host "  cv      = $($chk.profile.defaultCv)"
