<#
.SYNOPSIS
  SEND-LEADSBATCH : sends applications to ALL unsent leads already in the FeedBot-Link tab.
  Uses the upgraded impressive email body. Marks Sent in the sheet + history.
#>
param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "config.json"),
    [int]$MaxSend = 15
)

$ErrorActionPreference = "Continue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "Tracker.psm1") -Force
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Set-TrackerPath $config.paths.tracker
$P = $config.profile

$logDir = Join-Path $ProjectDir "logs"
$logFile = Join-Path $logDir ("leads-batch-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".log")
function Write-Log($msg) { $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"; Write-Host $line; Add-Content $logFile $line }

# ---------- load unsent leads from FeedBot-Link ----------
$historyFile = Join-Path $ProjectDir "edge-profile\feedbot-history.json"
$history = @()
if (Test-Path $historyFile) {
    $raw = Get-Content $historyFile -Raw | ConvertFrom-Json
    if ($raw.emails) { $history = @($raw.emails) }
}

$s = Open-Tracker
$ws = $null
try { $ws = $s.Package.Workbook.Worksheets["FeedBot-Link"] } catch {}
$leads = @()
if ($ws) {
    $last = 0; if ($ws.Dimension) { $last = $ws.Dimension.End.Row }
    for ($r = 2; $r -le $last; $r++) {
        $email = [string]$ws.Cells[$r, 2].Text
        $status = [string]$ws.Cells[$r, 5].Text
        $role = ([string]$ws.Cells[$r, 3].Text -replace "#\S+", "" -replace "\s+", " ").Trim()
        $leadCtx = [string]$ws.Cells[$r, 4].Text
        # UAE-ONLY policy (user 2026-11-04): skip leads whose post isn't UAE-related
        if ($leadCtx -and $leadCtx -notmatch "(?i)dubai|sharjah|abu dhabi|ajman|al ain|ras al khaimah|fujairah|\bu\.?a\.?e\.?\b|united arab") { Write-Log "  skip (not UAE-related): $email"; continue }
        if ($email -and $status -ne "Sent" -and $history -notcontains $email) {
            if ($leads.Email -notcontains $email) {
                $leads += [PSCustomObject]@{ Email = $email; Role = $(if ($role) { $role } else { "BIM Role" }) }
            }
        }
    }
}
Close-Tracker $s -Save $false

$leads = @($leads | Select-Object -First $MaxSend)
Write-Log ("Unique unsent leads to email: " + $leads.Count)
if ($leads.Count -eq 0) { Write-Log "Nothing to send."; exit 0 }

# ---------- start browser + Gmail ----------
Start-EdgeCdp -Headless
$cookieFile = Join-Path $ProjectDir "edge-profile\cookies.json"
if (Test-Path $cookieFile) {
    $sc = Get-Content $cookieFile -Raw | ConvertFrom-Json
    if ($sc) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($sc) } -TimeoutSec 30 | Out-Null }
}
Navigate-Cdp "https://mail.google.com/mail/u/0/#inbox" -TimeoutSec 60
Start-Sleep -Seconds 6
if ((Get-CdpBodyText) -notmatch "inbox|compose") { Write-Log "GMAIL NOT LOGGED IN"; Stop-EdgeCdp; exit 1 }
Write-Log "Gmail session OK."

function Get-ImpressiveBody($role) {
    return @"
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
}

$sent = 0
foreach ($lead in $leads) {
    $to = $lead.Email
    $role = $lead.Role
    Write-Log "Sending to $to ($role)..."
    $subject = "BIM Coordinator | 8+ yrs | Revit - Navisworks - ISO 19650"
    $body = Get-ImpressiveBody $role
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
        $sentOk = Invoke-CdpJs "(function(){ var b = Array.from(document.querySelectorAll('div[role=button], button')).find(function(e){ return /^send/i.test((e.innerText||'').trim()); }); if(b){ b.click(); return 'ok'; } return null; })()"
        if ($sentOk) {
            Start-Sleep -Seconds 8
            $after = Get-CdpBodyText
            if ($after -match "message sent|your message has been sent|view message") {
                Write-Log "  SENT (confirmed)"
                $history = @($history + $to | Sort-Object -Unique)
                @{ emails = $history } | ConvertTo-Json -Depth 4 | Set-Content $historyFile -Encoding UTF8
                try { Set-FeedBotLeadStatus -Email $to -Status "Sent" } catch {}
                $sent++
            } else { Write-Log "  no confirmation" }
        } else { Write-Log "  send button not found" }
    } catch { Write-Log "  error: $($_.Exception.Message)" }
    Start-Sleep -Seconds 4
}

Stop-EdgeCdp
Write-Log "===== BATCH DONE: $sent of $($leads.Count) sent ====="
Write-Host "`nDone. $sent sent." -ForegroundColor Green
