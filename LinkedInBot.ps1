<#
.SYNOPSIS
  LINKEDINBOT : searches LinkedIn people and drafts connection messages (human-in-loop).
  Default is DryRun (search + queue + draft, never clicks Connect).
  Use -AutoSend only for small approved batches - LinkedIn bans aggressive automation.

.EXAMPLE
  .\LinkedInBot.ps1 -Query "BIM Manager UAE" -MaxProfiles 10 -DryRun
  .\LinkedInBot.ps1 -Query "Talent Acquisition Construction Dubai" -MaxProfiles 5 -AutoSend
#>
param(
    [string]$Query = "BIM Manager UAE",
    [int]$MaxProfiles = 10,
    [switch]$DryRun,
    [switch]$AutoSend,
    [switch]$ShowBrowser
)

$ErrorActionPreference = "Continue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking

$logDir = Join-Path $ProjectDir "logs"; New-Item -ItemType Directory -Force $logDir | Out-Null
$shotDir = Join-Path $logDir "shots"; New-Item -ItemType Directory -Force $shotDir | Out-Null
$repDir = Join-Path $ProjectDir "reports"; New-Item -ItemType Directory -Force $repDir | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmm"
$logFile = Join-Path $logDir ("linkedin-" + $stamp + ".log")
$queueFile = Join-Path $logDir "linkedin-queue.json"
function Write-Log($msg) { $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"; Write-Host $line; Add-Content $logFile $line }

if (-not $DryRun -and -not $AutoSend) {
    # Safe default: dry run unless explicitly asked to send
    $DryRun = $true
}

if ($AutoSend -and $MaxProfiles -gt 10) {
    Write-Log "SAFETY: capping AutoSend batch to 10 (you asked $MaxProfiles). LinkedIn restricts >20-25/day."
    $MaxProfiles = 10
}

if ($AutoSend) {
    Write-Host ""
    Write-Host "WARNING: AutoSend clicks real Connect buttons on LinkedIn." -ForegroundColor Red
    Write-Host "This can trigger verification / restriction. Recommended: 5-8/day max." -ForegroundColor Yellow
    $ans = Read-Host "Type YES to continue with $MaxProfiles profile(s) for query '$Query'"
    if ($ans -ne "YES") { Write-Log "Aborted by user."; exit 0 }
}

# Load existing queue to avoid re-processing same profiles
$queue = @()
if (Test-Path $queueFile) {
    try { $queue = @(Get-Content $queueFile -Raw | ConvertFrom-Json) } catch {}
}
$seenUrls = @{}
foreach ($q in $queue) { if ($q.ProfileUrl) { $seenUrls[$q.ProfileUrl] = $true } }

function New-DraftMessage($firstName, $headline) {
    $h = "$headline"
    if ($h.Length -gt 60) { $h = $h.Substring(0, 60) }
    $msg = "Hi $firstName, I'm a BIM Coordinator (8+ yrs MEA/UAE, Revit/Navisworks) exploring roles like $h. Would value connecting. - Silar"
    if ($msg.Length -gt 295) { $msg = $msg.Substring(0, 295) }
    return $msg
}

# ---------- run ----------
$headless = (-not $ShowBrowser)
Write-Log "Starting Edge (headless=$headless)... login via edge-profile (run Login-Setup.ps1 first for LinkedIn)"
try {
    if ($headless) { Start-EdgeCdp -Headless } else { Start-EdgeCdp }
    Enable-Stealth
} catch { Write-Log "FATAL: could not start Edge: $($_.Exception.Message)"; exit 1 }

try {
    $searchUrl = "https://www.linkedin.com/search/results/people/?keywords=" + [uri]::EscapeDataString($Query)
    Write-Log "Search: $Query"
    Write-Log "URL: $searchUrl"
    Navigate-Cdp $searchUrl
    Start-Sleep -Seconds 5
    Save-CdpScreenshot (Join-Path $shotDir ("linkedin-search-" + $stamp + ".png"))
    $body = Get-CdpBodyText

    if ($body -match "_bv|challenge|captcha|verify|unusual|sign up|join now" -and $body.Length -lt 2000) {
        Write-Log "Possible block/login wall. Open with -ShowBrowser and log in manually, then re-run."
        Save-CdpScreenshot (Join-Path $shotDir ("linkedin-block-" + $stamp + ".png"))
        Stop-EdgeCdp
        exit 1
    }
    if ($body -match "log in|sign in" -and $body -notmatch "messaging|notifications|my network") {
        Write-Log "Not logged in. Run: .\Login-Setup.ps1 (LinkedIn step), then re-run with -ShowBrowser to verify."
        Stop-EdgeCdp
        exit 1
    }

    $jsCards = @"
(function(){
  var cards = Array.from(document.querySelectorAll('li.reusable-search__result-container, div.entity-result'));
  var out = [];
  cards.forEach(function(c){
    var a = c.querySelector('a[href*="/in/"]');
    var name = (c.querySelector('span.entity-result__title-text a span[aria-hidden="true"]') || c.querySelector('a[href*="/in/"] span[aria-hidden="true"]') || {}).innerText || '';
    if (!name) { var t = c.innerText.split('\n')[0] || ''; name = t.trim(); }
    var head = (c.querySelector('.entity-result__primary-subtitle') || {}).innerText || '';
    var sub = (c.querySelector('.entity-result__secondary-subtitle') || {}).innerText || '';
    var href = a ? a.getAttribute('href').split('?')[0] : '';
    if (href && href.indexOf('http') !== 0) href = 'https://www.linkedin.com' + href;
    var btn = Array.from(c.querySelectorAll('button')).find(function(b){ return /connect|follow/i.test(b.innerText||''); });
    if (name && href) out.push({name: name.trim(), headline: (head+' '+sub).trim(), url: href, action: btn ? btn.innerText.trim() : ''});
  });
  return JSON.stringify(out);
})()
"@
    $raw = Invoke-CdpJs $jsCards
    $profiles = @()
    if ($raw) { try { $profiles = @(ConvertFrom-Json $raw) } catch { Write-Log "Parse failed" } }
    Write-Log "Found $($profiles.Count) profiles on page 1"

    $fresh = @($profiles | Where-Object { $_.url -and -not $seenUrls.ContainsKey($_.url) } | Select-Object -First $MaxProfiles)
    Write-Log "New (not in queue): $($fresh.Count)"

    foreach ($p in $fresh) {
        $first = ($p.name -split " ")[0]
        $draft = New-DraftMessage $first $p.headline
        $entry = [PSCustomObject]@{
            DateFound = (Get-Date -Format "yyyy-MM-dd HH:mm")
            Query = $Query
            Name = $p.name
            Headline = $p.headline
            ProfileUrl = $p.url
            DraftMessage = $draft
            Status = "Drafted"
        }
        $queue += $entry
        $seenUrls[$p.url] = $true
        Write-Log "- $($p.name) | $($p.headline) | $($p.url)"
        Write-Log "  Draft: $draft"
    }

    $queue | ConvertTo-Json -Depth 5 | Set-Content $queueFile -Encoding UTF8
    Write-Log "Queue saved: $queueFile (total $($queue.Count))"

    $rep = @("# LinkedInBot - $stamp", "", "Query: $Query", "Mode: $(if ($DryRun) { 'DRY RUN (no Connect clicked)' } else { 'AUTOSEND (small batch)' })", "", "## Drafts ($($fresh.Count) new)", "")
    foreach ($e in $fresh) { $rep += "- **$($e.Name)** - $($e.Headline) $($e.ProfileUrl)`n  > $($e.DraftMessage)" }
    $rep += ""
    $rep += "Send manually: open each ProfileUrl, click Connect > Add a note > paste draft (300 char limit)."
    $repPath = Join-Path $repDir ("linkedin-" + $stamp + ".md")
    $rep | Set-Content $repPath -Encoding UTF8
    Write-Log "Report: $repPath"

    if ($AutoSend -and $fresh.Count -gt 0) {
        Write-Log "AUTOSEND starting with human delays (60-120s between invites)..."
        $sent = 0
        foreach ($e in $fresh) {
            try {
                Navigate-Cdp $e.ProfileUrl
                Start-Sleep -Seconds (5 + (Get-Random -Minimum 0 -Maximum 5))
                # Scoped Connect: main profile only, never sidebar. Public profiles show Follow primary,
                # Connect lives in "Connect if you know each other" box or More-actions menu.
                $clicked = Invoke-CdpJs "(function(){
  function inSidebar(el){ return !!(el && el.closest && el.closest('aside, .scaffold-layout__aside, #sidebar, .artdeco-card [class*=pymk]')); }
  function vis(el){ return !!(el && (el.offsetWidth || el.offsetHeight)); }
  var scope = document.querySelector('main') || document;
  var btns = Array.from(scope.querySelectorAll('button')).filter(function(b){ return vis(b) && !inSidebar(b); });
  var box = btns.find(function(b){ var t=(b.innerText||'').trim().toLowerCase(); return t==='connect' && b.closest('section, div') && /connect if you know each other/i.test((b.closest('section, div').innerText||'')); });
  var main = box || btns.find(function(b){ return /^connect$/i.test((b.innerText||'').trim()) && /profile|top-card|pv-top-card/i.test((b.closest('section, div')||{className:''}).className||''); }) || btns.find(function(b){ return /^connect$/i.test((b.innerText||'').trim()); });
  if (main){ main.scrollIntoView({block:'center'}); main.click(); return 'clicked-main'; }
  var more = btns.find(function(b){ return /more actions|show more actions|more/i.test((b.innerText||'')+' '+(b.getAttribute('aria-label')||'')); });
  if (more){ more.click(); return 'clicked-more'; }
  return null;
})()"
                if ($clicked -eq 'clicked-more') {
                    Start-Sleep -Seconds 2
                    $clicked = Invoke-CdpJs "(function(){
  var items = Array.from(document.querySelectorAll('div[role=menu] [role=menuitem], div[role=menu] button, .artdeco-dropdown__item'));
  var c = items.find(function(x){ return /connect$/i.test(((x.innerText||'').trim())) && !/disconnect|remove/i.test(x.innerText||''); });
  if (c){ c.click(); return 'clicked-main'; }
  return null;
})()"
                }
                if ($clicked -ne 'clicked-main') { Write-Log "No Connect button for $($e.Name) (already connected/pending?)"; continue }
                Start-Sleep -Seconds 3
                # Note popup: scope to the invite dialog, not page buttons
                $withNote = Invoke-CdpJs "(function(){
  var dlg = document.querySelector('div[role=dialog][aria-modal=true], div.artdeco-modal, div.send-invite');
  var scope = dlg || document;
  var b = Array.from(scope.querySelectorAll('button')).find(function(x){ return /add a note/i.test(x.innerText||''); });
  if(!b) return null; b.click(); return 'clicked';
})()"
                Start-Sleep -Seconds 2
                if ($withNote) {
                    $msgJson = $e.DraftMessage | ConvertTo-Json -Compress
                    Invoke-CdpJs "(function(){ var t = $msgJson; var dlg = document.querySelector('div[role=dialog][aria-modal=true], div.artdeco-modal, div.send-invite') || document; var ta = dlg.querySelector('textarea[name=message], textarea'); if(!ta) return null; ta.focus(); document.execCommand('insertText', false, t); ta.dispatchEvent(new Event('input',{bubbles:true})); return 'filled'; })()" | Out-Null
                    Start-Sleep -Seconds 2
                    Invoke-CdpJs "(function(){ var dlg = document.querySelector('div[role=dialog][aria-modal=true], div.artdeco-modal, div.send-invite') || document; var b = Array.from(dlg.querySelectorAll('button')).find(function(x){ return /^send$/i.test((x.innerText||'').trim()); }); if(b){ b.click(); return 'sent'; } return null; })()" | Out-Null
                    Write-Log "Sent invite + note to $($e.Name)"
                } else {
                    Write-Log "Connect clicked but no Add-a-note dialog for $($e.Name) - invite sent without note or pending review"
                }
                $sent++
                Save-CdpScreenshot (Join-Path $shotDir ("linkedin-sent-" + $sent + "-" + $stamp + ".png"))
                # update queue status
                foreach ($q in $queue) { if ($q.ProfileUrl -eq $e.ProfileUrl) { $q.Status = "Sent" } }
                $queue | ConvertTo-Json -Depth 5 | Set-Content $queueFile -Encoding UTF8
            } catch {
                Write-Log "Failed on $($e.Name): $($_.Exception.Message)"
            }
            if ($sent -lt $fresh.Count) {
                $wait = Get-Random -Minimum 60 -Maximum 120
                Write-Log "Waiting ${wait}s before next (human pacing)..."
                Start-Sleep -Seconds $wait
            }
        }
        Write-Log "AUTOSEND done: $sent sent."
    } else {
        Write-Log "DRY RUN done. Review $queueFile, send manually. Nothing was clicked."
    }
} finally {
    try { Stop-EdgeCdp } catch {}
}

Write-Host "`nLinkedInBot done. Log: $logFile" -ForegroundColor Green
