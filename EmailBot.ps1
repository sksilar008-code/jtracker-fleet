<#
.SYNOPSIS
  EMAILBOT : applies to jobs by EMAIL - the unblockable channel.
  1) Opens each job page, extracts recruiter/agency email addresses
  2) Sends tailored CV + cover note via YOUR Gmail (browser session, no password needed)
  3) Marks "Applied (email)" + date in the tracker
  Use -TestMode to send everything to your own address first (verification).
#>
param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "config.json"),
    [int]$MaxJobs = 5,
    [string]$Rows = "",
    [switch]$TestMode
)
$RowsFilter = @()
if ($Rows) { $RowsFilter = @($Rows -split "," | ForEach-Object { [int]$_.Trim() }) }

$ErrorActionPreference = "Continue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "Tracker.psm1") -Force
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Set-TrackerPath $config.paths.tracker
$P = $config.profile

$logDir = Join-Path $ProjectDir "logs"; New-Item -ItemType Directory -Force $logDir | Out-Null
$logFile = Join-Path $logDir ("email-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".log")
function Write-Log($msg) { $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"; Write-Host $line; Add-Content $logFile $line }

$badDomains = @("adzuna", "indeed.com", "example.com", "sentry", "wixpress", "godaddy", "remoteok", "remotive", "arbeitnow", "googleapis", "schema", "w3.org", "cloudflare", "drjobs", "founditgulf", "naukrigulf", "bayt")

function Get-CvPath($tailoredFolder) {
    if ($tailoredFolder -and (Test-Path $tailoredFolder)) {
        $pdf = Get-ChildItem $tailoredFolder -Filter *.pdf -Recurse | Sort-Object Length -Descending | Select-Object -First 1
        if ($pdf) { return $pdf.FullName }
    }
    return $P.defaultCv
}

function Get-CoverNote($job, $tailoredFolder) {
    if ($tailoredFolder -and (Test-Path (Join-Path $tailoredFolder "cover-letter.md"))) {
        $txt = (Get-Content (Join-Path $tailoredFolder "cover-letter.md") -Raw)
        if ($txt.Length -gt 1200) { $txt = $txt.Substring(0, 1200) }
        return $txt
    }
    return "Dear Hiring Team,`n`nI am writing to apply for the position of $($job.Title) at $($job.Company). I am a BIM Coordinator with 8+ years of experience delivering infrastructure and building projects across six countries, specialising in multidisciplinary model federation, clash coordination and ISO 19650 information management (Revit, Navisworks, ACC/BIM 360).`n`nPlease find my CV attached. I would welcome the opportunity to discuss how I can contribute to your team.`n`nKind regards,`n$($P.fullName)`n$($P.email)`n$($P.phone)"
}

# ---------- pick jobs ----------
$statuses = @($config.apply.statusesToApply)
$allJobs = @(Get-TrackerJobs | Where-Object { $statuses -contains $_.Status -and $_.URL })
# UAE-ONLY policy (user 2026-11-04): email applications only for UAE-related jobs
$jobs = @($allJobs | Where-Object { "$($_.Location) $($_.URL) $($_.Title)" -match "(?i)\bu\.?a\.?e\.?\b|united arab|dubai|sharjah|abu dhabi|ajman|al ain|ae\.indeed\.com" })
if ($jobs.Count -lt $allJobs.Count) { Write-Log "UAE-only: skipped $($allJobs.Count - $jobs.Count) non-UAE job(s)" }
$jobs = @($jobs | Sort-Object @{ Expression = { if ($_.Status -eq "New") { 0 } else { 1 } } }, Row)
if ($RowsFilter.Count -gt 0) { $jobs = @($jobs | Where-Object { $RowsFilter -contains $_.Row }) }
$jobs = @($jobs | Select-Object -First $MaxJobs)
if (-not $jobs) { Write-Log "No jobs to process."; exit 0 }

Write-Log "Starting Edge..."
Start-EdgeCdp -Headless
Enable-Stealth
$cookieFile = Join-Path $ProjectDir "edge-profile\cookies.json"
if (Test-Path $cookieFile) {
    $sc = Get-Content $cookieFile -Raw | ConvertFrom-Json
    if ($sc) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($sc) } -TimeoutSec 30 | Out-Null }
}

# ---------- phase 1: extract emails ----------
$targets = @()
foreach ($job in $jobs) {
    Write-Log "----- [row $($job.Row)] $($job.Title) @ $($job.Company)"
    $emails = @()
    $urlsToScan = @($job.URL)
    # fallback: company website contact pages (recruitment agencies list emails there)
    try {
        $h = ([uri]$job.URL).Host
        $urlsToScan += @("https://$h/contact", "https://$h/contact-us", "https://$h/en/contact")
    } catch {}
    foreach ($scanUrl in $urlsToScan) {
        if ($emails.Count -gt 0) { break }
        try {
            Navigate-Cdp $scanUrl -TimeoutSec 25
            Start-Sleep -Seconds 2
            $html = Invoke-CdpJs "document.documentElement.outerHTML.substring(0, 300000)"
            if (-not $html) { continue }
            $found = [regex]::Matches($html, "(?i)[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}") | ForEach-Object { $_.Value.ToLower() } | Sort-Object -Unique
            foreach ($e in $found) {
                $dom = ($e -split "@")[1]
                $skip = $false
                foreach ($b in $badDomains) { if ($dom -match $b) { $skip = $true; break } }
                if ($e -match "noreply|no-reply|donotreply|unsubscribe|example\.|@2x|@3x|\.png|\.jpg|\.gif|\.webp") { $skip = $true }
                # placeholder/template addresses (from page examples, JSON-LD samples, form placeholders)
                if ($e -match "^person@company\.|^your[._-]?|^tu\.|^name@|^email@|^user@|^test@|^sample@|@dominio\.|@company\.com|@email\.com|@domain\.com|@mail\.com$|@example\.|placeholder") { $skip = $true }
                if (-not $skip -and $emails -notcontains $e) { $emails += $e }
            }
        } catch { Write-Log "  scan error ($scanUrl): $($_.Exception.Message)" }
    }
    if ($emails.Count -gt 0) {
        Write-Log ("  EMAILS FOUND: " + ($emails -join ", "))
        $targets += [PSCustomObject]@{ Job = $job; Emails = $emails }
    } else {
        Write-Log "  no email found on page"
    }
}

if ($targets.Count -eq 0) { Write-Log "No emails found on any job page. Done."; Stop-EdgeCdp; exit 0 }
Write-Log ("Jobs with email targets: " + $targets.Count)

# ---------- phase 2: open Gmail (basic HTML mode - light, automation-friendly) ----------
Write-Log "Opening Gmail (basic HTML mode)..."
Navigate-Cdp "https://mail.google.com/mail/u/0/h/" -TimeoutSec 90
Start-Sleep -Seconds 6
$gBody = Get-CdpBodyText
if ($gBody -notmatch "inbox|compose|search mail|sign out") {
    Write-Log "GMAIL NOT LOGGED IN - cannot send. (Run Login-Setup for Google)"
    Stop-EdgeCdp; exit 1
}
Write-Log "Gmail session OK."

# ---------- phase 3: send applications ----------
$sent = 0
$sentThisRun = @()
foreach ($t in $targets) {
    $job = $t.Job
    $to = $t.Emails[0]
    if ($sentThisRun -contains $to) { Write-Log "  skip (already sent this run): $to"; continue }
    if ($TestMode) { $to = $P.email }
    $cv = Get-CvPath $job.TailoredFiles
    $subject = "Application: $($job.Title) - $($P.fullName)"
    $body = Get-CoverNote $job $job.TailoredFiles
    Write-Log "Sending to $to (job row $($job.Row))..."

    try {
        # Open Gmail compose with recipient + subject + body prefilled via URL
        $toEnc = [uri]::EscapeDataString($to)
        $su = [uri]::EscapeDataString($subject)
        $bo = [uri]::EscapeDataString($body)
        Navigate-Cdp "https://mail.google.com/mail/u/0/?view=cm&fs=1&to=$toEnc&su=$su&body=$bo" -TimeoutSec 60
        Start-Sleep -Seconds 8
        # Verify the recipient chip actually rendered (Gmail blocks send without it)
        $chipOk = Invoke-CdpJs "(function(){ return document.body.innerText.indexOf($($to | ConvertTo-Json -Compress)) >= 0; })()"
        Write-Log "  Recipient present on page: $chipOk ($to)"
        if (-not $chipOk) { Write-Log "  RECIPIENT MISSING - skipping send to avoid empty email."; continue }
        # Attach CV
        $cvOk = $false
        try { $cvOk = Set-CdpFileInput $cv } catch { Write-Log "  attach error: $($_.Exception.Message)" }
        Write-Log "  CV attached: $cvOk"
        if ($cvOk) { Start-Sleep -Seconds 12 }
        # Send
        $sentOk = Invoke-CdpJs @"
(function(){
  var b = Array.from(document.querySelectorAll('div[role=button], button')).find(function(e){ return /^send/i.test((e.innerText||'').trim()); });
  if (b){ b.click(); return 'ok'; }
  return null;
})()
"@
        if ($sentOk) {
            Start-Sleep -Seconds 8
            # verify Gmail confirms the send
            $after = Get-CdpBodyText
            if ($after -match "message sent|your message has been sent|view message") {
                Set-TrackerJobStatus -Row $job.Row -Status "Applied (email)" -DateApplied (Get-Date -Format "yyyy-MM-dd")
                Write-Log "  SENT (confirmed) + tracker updated."
                $sentThisRun += $to
                $sent++
            } else {
                Write-Log "  Send clicked but no confirmation - NOT marking tracker. Check manually."
            }
        } else {
            Write-Log "  Send button not found - draft left open."
        }
        Start-Sleep -Seconds 2
    } catch { Write-Log "  send error: $($_.Exception.Message)" }
}

Stop-EdgeCdp
Write-Log "===== DONE: $sent email application(s) sent ====="
Write-Host "`nEmailBot done. $sent sent. Log: $logFile" -ForegroundColor Green
