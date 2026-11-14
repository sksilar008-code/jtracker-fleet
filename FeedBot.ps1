<#
.SYNOPSIS
  FEEDBOT : LinkedIn feed job-post hunter.
  1) Searches LinkedIn posts (last 24h) for BIM jobs in UAE by hashtag/keyword
  2) Extracts recruiter emails from post text (with context)
  3) Sends tailored CV + note from your Gmail
  4) Tracks history - never double-sends. Max posts per run is capped.
#>
param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "config.json"),
    [int]$MaxPosts = 2,
    [switch]$AllSearches   # careful mode: search ALL feed queries this run (skips 3-per-run rotation)
)

$ErrorActionPreference = "Continue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "Tracker.psm1") -Force
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Set-TrackerPath $config.paths.tracker
$P = $config.profile

$logDir = Join-Path $ProjectDir "logs"; New-Item -ItemType Directory -Force $logDir | Out-Null
$logFile = Join-Path $logDir ("feed-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".log")
function Write-Log($msg) { $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"; Write-Host $line; Add-Content $logFile $line }

$historyFile = Join-Path $ProjectDir "edge-profile\feedbot-history.json"
$history = @()
if (Test-Path $historyFile) {
    try {
        $raw = Get-Content $historyFile -Raw | ConvertFrom-Json
        $list = @()
        if ($raw.emails) { $list = @($raw.emails) } else { $list = @($raw) }
        foreach ($item in $list) {
            if ($item -is [string] -and $item) { $history += $item }
            elseif ($item.value) { $history += @($item.value | Where-Object { $_ }) }
        }
        $history = @($history | Sort-Object -Unique)
    } catch { $history = @() }
}

# UAE-ONLY feed queries (user 2026-11-04): every search targets UAE recruitment
$searches = @(
    "%23bimjobs%20uae",
    "%23uaejobs%20bim",
    "%23dubaijobs%20bim",
    "bim%20hiring%20dubai",
    "bim%20jobs%20sharjah",
    "bim%20coordinator%20uae",
    "revit%20jobs%20uae"
)

# excluded countries (user preference): Saudi Arabia + Egypt
$blockedRegions = "saudi|ksa|riyadh|jeddah|dammam|neom|egypt|cairo|masader"
# excluded role types (user preference): Mechanical / MEP (EN+FR+DE variants)
$blockedRoles = "(?i)mechanical|electrical|plumbing|plomberie|hvac|elektro|lectr|cvc|\bmep\b|\btga\b|\bcfo\b|\bcfa\b|\bmech\b"

$badEmail = "noreply|no-reply|donotreply|unsubscribe|example\.|@2x|@3x|\.png|\.jpg|\.gif|\.webp|^person@company\.|^your[._-]?|^tu\.|^name@|^email@|^user@|^test@|^sample@|@dominio\.|@company\.com|@email\.com|@domain\.com|@mail\.com$|@example\.|placeholder|@linkedin\.com|@google\."
$goodRole = "structural|coordinator|civil|infrastructure|roads|bridge|architect|draftsman|modeller|modeler|engineer|lead|manager"

function Restart-FeedBrowser {
    Write-Log "Restarting Edge (connection lost)..."
    try { Stop-EdgeCdp } catch {}
    Start-Sleep -Seconds 2
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -match "JTrackerEdge|edge-profile" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2
    Start-EdgeCdp -Headless:$false
    Enable-Stealth
    if (Test-Path $cookieFile) {
        $sc = Get-Content $cookieFile -Raw | ConvertFrom-Json
        if ($sc) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($sc) } -TimeoutSec 30 | Out-Null }
    }
    Write-Log "Edge restarted."
}

Write-Log "Starting Edge (LinkedIn session)..."
Start-EdgeCdp -Headless:$false
Enable-Stealth
$cookieFile = Join-Path $ProjectDir "edge-profile\cookies.json"
if (Test-Path $cookieFile) {
    $sc = Get-Content $cookieFile -Raw | ConvertFrom-Json
    if ($sc) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($sc) } -TimeoutSec 30 | Out-Null }
}

# ---------- phase 1: collect posts with emails ----------
# rotation: 3 searches per run (cycles through the list) to avoid LinkedIn throttling
# -AllSearches: careful mode - run every query in one go (still respects the back-off)
$rotFile = Join-Path $ProjectDir "edge-profile\feedbot-rotation.txt"
$rotIdx = 0
if (Test-Path $rotFile) { try { $rotIdx = [int](Get-Content $rotFile -Raw) } catch {} }
if ($AllSearches) {
    $runSearches = @($searches)
    Write-Log "CAREFUL MODE: all $($searches.Count) feed queries this run"
} else {
    $runSearches = @()
    for ($i = 0; $i -lt 3; $i++) { $runSearches += $searches[($rotIdx + $i) % $searches.Count] }
    Set-Content $rotFile ((($rotIdx + 3) % $searches.Count)) -NoNewline
}
Write-Log ("This run searches: " + ($runSearches -join " | "))
$found = @()
$consecFail = 0
foreach ($q in $runSearches) {
    $url = "https://www.linkedin.com/search/results/content/?keywords=$q&sortBy=date_posted"
    Write-Log "Searching: $q"
    try {
        try { Invoke-CdpJs "1+1" | Out-Null } catch { Restart-FeedBrowser }
        Navigate-Cdp $url -TimeoutSec 45
        Start-Sleep -Seconds 6
        $body = Get-CdpBodyText
        if ($body.Length -lt 400) { Write-Log "  empty/blocked page"; continue }
        $emails = [regex]::Matches($body, "(?i)[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}") | ForEach-Object { $_.Value.ToLower() } | Sort-Object -Unique
        foreach ($e in $emails) {
            if ($e -match $badEmail) { continue }
            if ($e -match "\.sa$|\.eg$") { Write-Log "  skip (blocked country domain): $e"; continue }
            if ($history -contains $e) { Write-Log "  skip (already emailed): $e"; continue }
            $i = $body.IndexOf($e)
            $start = [Math]::Max(0, $i - 500)
            $ctx = ($body.Substring($start, [Math]::Min(700, $body.Length - $start)) -replace "\s+", " ").Trim()
            if ($ctx -match $blockedRegions) { Write-Log "  skip (blocked country in post): $e"; continue }
            # UAE-ONLY policy (user 2026-11-04): the post itself must be UAE-related
            if ($ctx -notmatch "(?i)dubai|sharjah|abu dhabi|ajman|al ain|ras al khaimah|fujairah|\bu\.?a\.?e\.?\b|united arab") { Write-Log "  skip (not UAE-related): $e"; continue }
            # freshness filter (rule 2026-10-26): UAE posts <= 24h, other countries <= 7 days
            # age token in LinkedIn relative format: "3d" / "2w" / "1mo" near the email
            $window = $body.Substring([Math]::Max(0, $i - 900), [Math]::Min(900, $body.Length - [Math]::Max(0, $i - 900)))
            $ageTok = [regex]::Match($window, "\b(\d+)\s*(d|w|mo)\b")
            if ($ageTok.Success) {
                $an = [int]$ageTok.Groups[1].Value
                $ad = $an
                if ($ageTok.Groups[2].Value -eq "w") { $ad = $an * 7 }
                if ($ageTok.Groups[2].Value -eq "mo") { $ad = 30 }
                $lim = [int]$config.apply.maxAgeDaysOther
                if ($q -match "uae|dubai" -or $ctx -match "(?i)dubai|sharjah|abu dhabi|ajman|\bu\.?a\.?e\.?\b|united arab") { $lim = [int]$config.apply.maxAgeDaysUae }
                if ($ad -gt $lim) { Write-Log "  skip (older than ${lim}d limit: ~${ad}d): $e"; continue }
            }
            $found += [PSCustomObject]@{ Email = $e; Context = $ctx }
            Write-Log "  FOUND: $e"
            # log every fresh lead to the FeedBot-Link tracker tab
            $leadRole = "BIM Role"
            if ($ctx -match "(BIM[^.]{0,45}(Engineer|Coordinator|Modeler|Modeller|Manager|Lead|Architect|Draftsman|Technician))") { $leadRole = $Matches[1].Trim() }
            $leadStatus = "New lead"
            if ($leadRole -match $blockedRoles) { $leadStatus = "Skipped (MEP/Mech)"; Write-Log "  lead role is MEP/Mechanical - will not email" }
            try { [void](Add-FeedBotLead -Email $e -Role $leadRole -Context $ctx -Status $leadStatus) } catch { Write-Log "  lead log failed: $($_.Exception.Message)" }
        }
    } catch {
        Write-Log "  search error: $($_.Exception.Message)"
        $consecFail++
        if ($consecFail -ge 2) { Write-Log "LinkedIn throttling suspected (2 straight failures) - backing off until next run."; break }
        Start-Sleep -Seconds 10
        continue
    }
    $consecFail = 0
    Start-Sleep -Seconds 10
}

Write-Log ("New email targets: " + $found.Count)
if ($found.Count -eq 0) { Stop-EdgeCdp; Write-Log "Nothing new. Done."; exit 0 }

# ---------- phase 2: relevance filter + send ----------
# Judge the EXTRACTED ROLE (not the whole post - multi-role posts mention MEP alongside good roles)
function Get-PostRole($ctx) {
    if ($ctx -match "(BIM[^.]{0,45}(Engineer|Coordinator|Modeler|Modeller|Manager|Lead|Architect|Draftsman|Technician))") { return $Matches[1].Trim() }
    return ""
}
$targets = @($found | ForEach-Object {
    $role = Get-PostRole $_.Context
    if (-not $role -and $_.Context -match $goodRole -and $_.Context -notmatch $blockedRoles) { $role = "BIM Position" }
    $_ | Add-Member -NotePropertyName Role -NotePropertyValue $role -Force -PassThru
} | Where-Object { $_.Role -and $_.Role -notmatch $blockedRoles } | Sort-Object -Property Email -Unique | Select-Object -First $MaxPosts)
if ($targets.Count -eq 0) { Write-Log "No role-relevant targets this run."; Stop-EdgeCdp; exit 0 }

# Gmail check
Navigate-Cdp "https://mail.google.com/mail/u/0/#inbox" -TimeoutSec 60
Start-Sleep -Seconds 6
if ((Get-CdpBodyText) -notmatch "inbox|compose") { Write-Log "GMAIL NOT LOGGED IN"; Stop-EdgeCdp; exit 1 }

$sent = 0
foreach ($t in $targets) {
    $to = $t.Email
    # shared 8/day cap across agents - only defer when the hard cap is hit
    if (-not (Test-EmailBudget)) { Write-Log "daily email cap reached (8/day across agents) - stopping sends; remaining targets go out after midnight"; break }
    Write-Log "Sending to $to ..."
    # subject: try to extract a title from context, else generic
    $role = "BIM Position"
    if ($t.Context -match "(BIM[^.]{0,45}(Engineer|Coordinator|Modeler|Modeller|Manager|Lead|Architect|Draftsman|Technician))") { $role = $Matches[1].Trim() }
    $subject = "BIM Coordinator | 8+ yrs | Revit - Navisworks - ISO 19650"
    $body = @"
Dear Hiring Team,

I came across your post regarding the $role opportunity, and my background maps directly onto what you are looking for.

Over the past 8+ years I have delivered BIM coordination on 25+ projects across six countries - from Singapore MRT station, depot and tunnel packages (WSP) to landmark high-rise towers in the UAE (Tiger Group, my current role). Highlights:

- Federation and clash coordination of architectural, structural and MEP models - Revit, Navisworks, Civil 3D
- ISO 19650 delivery: BEPs, EIR/LOD, CDE administration on Autodesk Construction Cloud (BIM 360)
- Leading clash-resolution workshops and driving coordination actions to closure with design and site teams
- Automation that saves hours: Dynamo scripts, Power BI dashboards, model health-check tools

What I bring from day one:
- Station and rail interface experience - Jurong Region Line J101 depot and CR116 interchange station and tunnels, Singapore
- UAE-based (Sharjah), open to relocation
- GMICE - Graduate Member, Institution of Civil Engineers (UK)

My updated CV is attached. I would welcome a short call to discuss how I can support your team.

Kind regards,
$($P.fullName)
$($P.email) | $($P.phone) | Sharjah, UAE
"@
    try {
        $toEnc = [uri]::EscapeDataString($to)
        $su = [uri]::EscapeDataString($subject)
        $bo = [uri]::EscapeDataString($body)
        Navigate-Cdp "https://mail.google.com/mail/u/0/?view=cm&fs=1&to=$toEnc&su=$su&body=$bo" -TimeoutSec 60
        Start-Sleep -Seconds 8
        $chipOk = Invoke-CdpJs "(function(){ return document.body.innerText.indexOf($($to | ConvertTo-Json -Compress)) >= 0; })()"
        if (-not $chipOk) { Write-Log "  recipient missing - skip"; continue }
        $cvOk = $false
        try { $cvOk = Set-CdpFileInput $P.defaultCv } catch {}
        Write-Log "  CV attached: $cvOk"
        if ($cvOk) { Start-Sleep -Seconds 12 }
        # Gmail send is flaky: a notification/confirm dialog can block the click, or the
        # click lands before compose is ready. Retry up to 3x, dismissing any dialog.
        $confirmed = $false
        for ($try = 1; $try -le 3 -and -not $confirmed; $try++) {
            # dismiss any blocking popup (notification ask, send-without-attachment, etc.)
            Invoke-CdpJs "(function(){ var d = Array.from(document.querySelectorAll('[role=dialog] button, [aria-label*=Close], [aria-label*=Dismiss], button')).find(function(e){ return /^(no thanks|dismiss|close|cancel|send without attachment)$/i.test((e.innerText||'').trim()) || /close/i.test(e.getAttribute('aria-label')||''); }); if(d){ d.click(); } })()" | Out-Null
            Start-Sleep -Seconds 1
            $sentOk = Invoke-CdpJs "(function(){ var b = Array.from(document.querySelectorAll('div[role=button], button')).find(function(e){ return /^send$/i.test((e.innerText||'').trim()); }); if(b){ b.scrollIntoView({block:'center'}); b.click(); return 'ok'; } return null; })()"
            if ($sentOk) {
                # poll for the confirmation toast over ~15s
                for ($w = 0; $w -lt 5 -and -not $confirmed; $w++) {
                    Start-Sleep -Seconds 3
                    $after = Get-CdpBodyText
                    if ($after -match "message sent|your message has been sent|view message") { $confirmed = $true }
                }
            } else {
                Write-Log "  send button not found (try $try)"
            }
            if (-not $confirmed) { Start-Sleep -Seconds 3 }
        }
        if ($confirmed) {
            Write-Log "  SENT (confirmed)"
            $history = @($history + $to | Sort-Object -Unique)
            @{ emails = $history } | ConvertTo-Json -Depth 4 | Set-Content $historyFile -Encoding UTF8
            try { Set-FeedBotLeadStatus -Email $to -Status "Sent" } catch {}
            try { Add-SentEmail -To $to -Agent "FeedBot" } catch {}
            # record in tracker
            try {
                Add-TrackerJob -Job ([PSCustomObject]@{
                    Title = $role; Company = "(via LinkedIn post)"; Location = "UAE"
                    Source = "LinkedInFeed"; URL = "mailto:$to"; Score = 8
                    Salary = $config.search.salaryExpectation; TailoredFiles = ""
                })
                $jobs = Get-TrackerJobs | Sort-Object Row -Descending | Select-Object -First 1
                if ($jobs) { Set-TrackerJobStatus -Row $jobs.Row -Status "Applied (email)" -DateApplied (Get-Date -Format "yyyy-MM-dd") }
            } catch { Write-Log "  tracker note failed: $($_.Exception.Message)" }
            $sent++
        } else {
            Write-Log "  send FAILED after 3 tries - NOT marked sent (will retry next run)"
            try { Set-FeedBotLeadStatus -Email $to -Status "Send failed - retry" } catch {}
        }
    } catch { Write-Log "  send error: $($_.Exception.Message)" }
    Start-Sleep -Seconds 3
}

Stop-EdgeCdp
Write-Log "===== FEEDBOT DONE: $sent sent ====="
Write-Host "`nFeedBot done. $sent LinkedIn-post application(s) sent. Log: $logFile" -ForegroundColor Green
