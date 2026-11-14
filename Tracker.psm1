# Tracker.psm1 - JobTracker.xlsx helper module.
# Uses the ImportExcel module (EPPlus, pure .NET) INSTEAD of Excel COM: no Office
# installation required, so the fleet runs identically on a cloud server (2026-11-10:
# pipeline must work without the PC). Function signatures and export list are identical
# to the COM version - callers (HuntBot/ApplyBot/Tailored/FeedBot/EmailBot) unchanged.
$script:TrackerPath = "C:\Users\Silar Shaik\OneDrive - Tiger Group\J-Tracker\JobTracker.xlsx"
# portable default: prefer paths.tracker from a config.json next to this module (the
# server package rewrites config.json for the target machine); every bot also calls
# Set-TrackerPath explicitly, this just keeps the fallback correct anywhere.
$cfgFallback = Join-Path $PSScriptRoot "config.json"
if (Test-Path $cfgFallback) {
    try { $cf = Get-Content $cfgFallback -Raw | ConvertFrom-Json; if ($cf.paths.tracker) { $script:TrackerPath = $cf.paths.tracker } } catch {}
}

# --- load the xlsx engine: bundled copy first (ships in the server package), then installed ---
if (-not (Get-Command Open-ExcelPackage -ErrorAction SilentlyContinue)) {
    $bundled = Join-Path $PSScriptRoot "modules\ImportExcel"
    if (Test-Path $bundled) {
        $psd1 = Get-ChildItem $bundled -Recurse -Filter "ImportExcel.psd1" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($psd1) { try { Import-Module $psd1.FullName -ErrorAction Stop } catch {} }
    }
    if (-not (Get-Command Open-ExcelPackage -ErrorAction SilentlyContinue)) {
        try { Import-Module ImportExcel -ErrorAction Stop } catch {}
    }
}
if (-not (Get-Command Open-ExcelPackage -ErrorAction SilentlyContinue)) {
    throw "ImportExcel module not found - expected bundled at $PSScriptRoot\modules\ImportExcel or installed in PSModulePath"
}

function Set-TrackerPath { param([string]$Path) $script:TrackerPath = $Path }

function Open-Tracker {
    # Retry up to 3 times (OneDrive can briefly lock the file)
    for ($i = 1; $i -le 3; $i++) {
        try {
            $pkg = Open-ExcelPackage -Path $script:TrackerPath
            return @{ Package = $pkg }
        } catch {
            if ($i -eq 3) { throw }
            Start-Sleep -Seconds 2
        }
    }
}

function Close-Tracker {
    param($Session, [bool]$Save = $true)
    # ImportExcel 7.8.10 semantics: Close-ExcelPackage SAVES by default;
    # -NoSave discards. (There is NO -Save switch - passing one prefix-matches
    # -SaveAs and silently writes the workbook to a file named "True"/"False".)
    if ($Session -and $Session.Package) {
        try {
            if ($Save) { Close-ExcelPackage $Session.Package }
            else       { Close-ExcelPackage $Session.Package -NoSave }
        } catch {}
    }
}

function Get-TrackerJobs {
    $s = Open-Tracker
    try {
        $ws = $s.Package.Workbook.Worksheets["Jobs"]
        if (-not $ws) { return @() }
        $last = 0
        if ($ws.Dimension) { $last = $ws.Dimension.End.Row }
        $jobs = @()
        for ($r = 2; $r -le $last; $r++) {
            $title = [string]$ws.Cells[$r, 2].Text
            $url = [string]$ws.Cells[$r, 6].Text
            if (-not $title -and -not $url) { continue }
            $jobs += [PSCustomObject]@{
                Row          = $r
                DateFound    = [string]$ws.Cells[$r, 1].Text
                Title        = $title
                Company      = [string]$ws.Cells[$r, 3].Text
                Location     = [string]$ws.Cells[$r, 4].Text
                Source       = [string]$ws.Cells[$r, 5].Text
                URL          = $url
                Score        = [string]$ws.Cells[$r, 7].Text
                Status       = [string]$ws.Cells[$r, 8].Text
                DatePosted   = [string]$ws.Cells[$r, 9].Text
                DateApplied  = [string]$ws.Cells[$r, 10].Text
                Salary       = [string]$ws.Cells[$r, 11].Text
                TailoredFiles= [string]$ws.Cells[$r, 12].Text
            }
        }
        return $jobs
    } finally { Close-Tracker $s -Save $false }
}

function Add-TrackerJob {
    param([Parameter(Mandatory)]$Job)  # object with Title, Company, Location, Source, URL, Score, Salary, TailoredFiles
    # scalar guard: arrays (from queue nesting bugs) are rejected, never squashed into a cell
    function Get-One($v) {
        if ($null -eq $v) { return "" }
        if ($v -is [System.Array]) {
            if ($v.Count -eq 1) { return [string]$v[0] }
            throw "Refusing to add job with array-valued field (would squash $($v.Count) jobs into one row)"
        }
        return [string]$v
    }
    $s = Open-Tracker
    try {
        $ws = $s.Package.Workbook.Worksheets["Jobs"]
        $next = 1
        if ($ws.Dimension) { $next = $ws.Dimension.End.Row + 1 }
        $ws.Cells[$next, 1].Value  = (Get-Date -Format "yyyy-MM-dd")
        $ws.Cells[$next, 2].Value  = (Get-One $Job.Title)
        $ws.Cells[$next, 3].Value  = (Get-One $Job.Company)
        $ws.Cells[$next, 4].Value  = (Get-One $Job.Location)
        $ws.Cells[$next, 5].Value  = (Get-One $Job.Source)
        $ws.Cells[$next, 6].Value  = (Get-One $Job.URL)
        $ws.Cells[$next, 7].Value  = (Get-One $Job.Score)
        $ws.Cells[$next, 8].Value  = "New"
        $ws.Cells[$next, 11].Value = (Get-One $Job.Salary)
        if ($Job.TailoredFiles) { $ws.Cells[$next, 12].Value = (Get-One $Job.TailoredFiles) }
        Close-Tracker $s -Save $true
        return $next
    } catch { Close-Tracker $s -Save $false; throw }
}

function Set-TrackerJobStatus {
    param([int]$Row, [string]$Status, [string]$DateApplied = "")
    $s = Open-Tracker
    try {
        $ws = $s.Package.Workbook.Worksheets["Jobs"]
        $ws.Cells[$Row, 8].Value = $Status
        if ($DateApplied) { $ws.Cells[$Row, 10].Value = $DateApplied }
        Close-Tracker $s -Save $true
    } catch { Close-Tracker $s -Save $false; throw }
}

function Set-TrackerJobTailored {
    param([int]$Row, [string]$Folder)
    $s = Open-Tracker
    try {
        $ws = $s.Package.Workbook.Worksheets["Jobs"]
        $ws.Cells[$Row, 12].Value = $Folder
        Close-Tracker $s -Save $true
    } catch { Close-Tracker $s -Save $false; throw }
}

function Get-TrackerAccounts {
    $s = Open-Tracker
    try {
        $ws = $s.Package.Workbook.Worksheets["Accounts"]
        if (-not $ws) { return @() }
        $last = 0
        if ($ws.Dimension) { $last = $ws.Dimension.End.Row }
        $accs = @()
        for ($r = 2; $r -le $last; $r++) {
            $platform = [string]$ws.Cells[$r, 1].Text
            if (-not $platform) { continue }
            $accs += [PSCustomObject]@{
                Row      = $r
                Platform = $platform
                Site     = [string]$ws.Cells[$r, 2].Text
                Username = [string]$ws.Cells[$r, 3].Text
                Password = [string]$ws.Cells[$r, 4].Text   # DPAPI-encrypted string
                Verified = [string]$ws.Cells[$r, 5].Text
                Created  = [string]$ws.Cells[$r, 6].Text
                Applied  = [string]$ws.Cells[$r, 7].Text
                Status   = [string]$ws.Cells[$r, 8].Text
            }
        }
        return $accs
    } finally { Close-Tracker $s -Save $false }
}

function Add-TrackerAccount {
    param([string]$Platform, [string]$Site, [string]$Username, [string]$PlainPassword)
    $enc = ConvertTo-SecureString $PlainPassword -AsPlainText -Force | ConvertFrom-SecureString
    $s = Open-Tracker
    try {
        $ws = $s.Package.Workbook.Worksheets["Accounts"]
        $next = 1
        if ($ws.Dimension) { $next = $ws.Dimension.End.Row + 1 }
        $ws.Cells[$next, 1].Value = $Platform
        $ws.Cells[$next, 2].Value = $Site
        $ws.Cells[$next, 3].Value = $Username
        $ws.Cells[$next, 4].Value = $enc
        $ws.Cells[$next, 5].Value = "no"
        $ws.Cells[$next, 6].Value = (Get-Date -Format "yyyy-MM-dd")
        $ws.Cells[$next, 8].Value = "active"
        Close-Tracker $s -Save $true
    } catch { Close-Tracker $s -Save $false; throw }
}

function Unprotect-TrackerPassword {
    param([string]$Encrypted)
    if (-not $Encrypted) { return $null }
    try {
        $sec = ConvertTo-SecureString $Encrypted
        return [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
    } catch { return $null }
}

function Add-TrackerApplicationCount {
    # Increments "Jobs applied" on the Accounts sheet for the platform matching a URL
    param([string]$JobUrl)
    $h = ([uri]$JobUrl).Host -replace "^www\.", ""
    $s = Open-Tracker
    try {
        $ws = $s.Package.Workbook.Worksheets["Accounts"]
        $last = 0
        if ($ws.Dimension) { $last = $ws.Dimension.End.Row }
        for ($r = 2; $r -le $last; $r++) {
            $site = ([string]$ws.Cells[$r, 2].Text -replace "^www\.", "").ToLower()
            $plat = ([string]$ws.Cells[$r, 1].Text).ToLower()
            if (($site -and $h.Contains($site)) -or ($plat -and $h.Contains($plat))) {
                $cur = [string]$ws.Cells[$r, 7].Text
                $n = 1; if ($cur -match "^\d+$") { $n = [int]$cur + 1 }
                $ws.Cells[$r, 7].Value = $n
                Close-Tracker $s -Save $true
                return $n
            }
        }
        Close-Tracker $s -Save $false
        return 0
    } catch { try { Close-Tracker $s -Save $false } catch {}; return 0 }
}

function Add-FeedBotLead {
    param([string]$Email, [string]$Role, [string]$Context, [string]$Status)
    $s = Open-Tracker
    try {
        $wb = $s.Package.Workbook
        $ws = $wb.Worksheets["FeedBot-Link"]
        if (-not $ws) {
            $ws = $wb.Worksheets.Add("FeedBot-Link")
            $ws.Cells[1, 1].Value = "Date Found"; $ws.Cells[1, 2].Value = "Email"
            $ws.Cells[1, 3].Value = "Role"; $ws.Cells[1, 4].Value = "Post Context"
            $ws.Cells[1, 5].Value = "Status"; $ws.Cells[1, 6].Value = "Date Sent"
        }
        $next = 1
        if ($ws.Dimension) { $next = $ws.Dimension.End.Row + 1 }
        # don't re-log an email already in the sheet
        for ($r = 2; $r -lt $next; $r++) {
            if ([string]$ws.Cells[$r, 2].Text -eq $Email) { Close-Tracker $s -Save $false; return -$r }
        }
        $ctx = "" + $Context
        $ws.Cells[$next, 1].Value = (Get-Date -Format "yyyy-MM-dd HH:mm")
        $ws.Cells[$next, 2].Value = $Email
        $ws.Cells[$next, 3].Value = $Role
        $ws.Cells[$next, 4].Value = $(if ($ctx.Length -gt 250) { $ctx.Substring(0, 250) } else { $ctx })
        $ws.Cells[$next, 5].Value = $Status
        Close-Tracker $s -Save $true
        return $next
    } catch { try { Close-Tracker $s -Save $false } catch {}; return 0 }
}

function Set-FeedBotLeadStatus {
    param([string]$Email, [string]$Status)
    $s = Open-Tracker
    try {
        $wb = $s.Package.Workbook
        $ws = $wb.Worksheets["FeedBot-Link"]
        if ($ws) {
            $last = 0
            if ($ws.Dimension) { $last = $ws.Dimension.End.Row }
            for ($r = 2; $r -le $last; $r++) {
                if ([string]$ws.Cells[$r, 2].Text -eq $Email) {
                    $ws.Cells[$r, 5].Value = $Status
                    if ($Status -eq "Sent") { $ws.Cells[$r, 6].Value = (Get-Date -Format "yyyy-MM-dd HH:mm") }
                }
            }
        }
        Close-Tracker $s -Save $true
    } catch { try { Close-Tracker $s -Save $false } catch {} }
}

# ---- shared daily email ledger: hard cap of 8 emails/day across ALL sending agents ----
# One JSONL line per confirmed send, file per calendar day (logs\email-ledger-YYYYMMDD.jsonl).
# Every sender checks Test-EmailBudget before each send and calls Add-SentEmail after a
# confirmed send, so the cap holds across agents and runs (the fleet run-lock serialises
# writers, so no concurrent-append race).
$script:EmailDailyCap = 8
function Get-EmailLedgerFile {
    $dir = Join-Path $PSScriptRoot "logs"
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return (Join-Path $dir ("email-ledger-" + (Get-Date -Format "yyyyMMdd") + ".jsonl"))
}
function Get-EmailsSentToday {
    $f = Get-EmailLedgerFile
    if (-not (Test-Path $f)) { return @() }
    return @((Get-Content $f -Encoding UTF8 | Where-Object { $_ -and $_.Trim() }))
}
function Test-EmailBudget {
    return ((Get-EmailsSentToday).Count -lt $script:EmailDailyCap)
}
function Add-SentEmail {
    param([string]$To, [string]$Agent)
    $line = (@{ to = $To; agent = $Agent; at = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") } | ConvertTo-Json -Compress)
    Add-Content -Path (Get-EmailLedgerFile) -Value $line -Encoding UTF8
}
Export-ModuleMember -Function Set-TrackerPath, Open-Tracker, Close-Tracker, Get-TrackerJobs, Add-TrackerJob, Set-TrackerJobStatus, Set-TrackerJobTailored, Get-TrackerAccounts, Add-TrackerAccount, Unprotect-TrackerPassword, Add-TrackerApplicationCount, Add-FeedBotLead, Set-FeedBotLeadStatus, Get-EmailLedgerFile, Get-EmailsSentToday, Test-EmailBudget, Add-SentEmail
