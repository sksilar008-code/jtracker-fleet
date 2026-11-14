<#
.SYNOPSIS
  APPLYBOT : reads JobTracker.xlsx rows with status New/Needs manual and applies automatically.
  Drives Microsoft Edge via DevTools Protocol (no external libraries).
  Full auto-submit when config apply.autoSubmit = true. Use -DryRun to fill but never submit.
#>
param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "config.json"),
    [switch]$DryRun,
    [switch]$ShowBrowser,
    [int]$MaxJobs = 0,
    [string]$Rows = ""
)

$ErrorActionPreference = "Continue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "Tracker.psm1") -Force
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $ProjectDir "Tailor.psm1") -Force -DisableNameChecking

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Set-TrackerPath $config.paths.tracker
$P = $config.profile

$logDir = Join-Path $ProjectDir "logs"; New-Item -ItemType Directory -Force $logDir | Out-Null
$shotDir = Join-Path $logDir "shots"; New-Item -ItemType Directory -Force $shotDir | Out-Null
$logFile = Join-Path $logDir ("apply-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".log")
function Write-Log($msg) { $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"; Write-Host $line; Add-Content $logFile $line }

# ---------- JS helpers ----------
$jsClickApply = @"
(function(){
  function inPaywall(e){
    var n = e;
    for (var i=0;i<6 && n;i++){
      var id = ((n.className&&n.className.baseVal!==undefined?n.className.baseVal:n.className)||'') + ' ' + (n.id||'');
      if (/premium|paywall|membership|subscri|upgrade|pricing/i.test(id)) return true;
      n = n.parentElement;
    }
    return false;
  }
  var els = Array.from(document.querySelectorAll('button, a, input[type=submit]'));
  var cand = els.filter(function(e){
    var t = ((e.innerText||'') + ' ' + (e.value||'')).toLowerCase();
    return /apply|bewerben|i'm interested|im interested|ich bin interessiert|jetzt bewerben|schnelle bewerbung|postuler|candidatura|solliciteer|solliciteren/.test(t) && !/applied|save|continue|upgrade|premium/.test(t) && (e.offsetWidth||e.offsetHeight) && !inPaywall(e);
  });
  cand.sort(function(a,b){ return ((a.innerText||a.value||'').length) - ((b.innerText||b.value||'').length); });
  if (cand.length){ cand[0].scrollIntoView({block:'center'}); cand[0].click(); return (cand[0].innerText||cand[0].value||'clicked').trim(); }
  return null;
})()
"@

# Removes paywall/premium overlays entirely (DOM removal), restores scrolling
$jsKillPaywall = @"
(function(){
  var killed = 0;
  Array.from(document.querySelectorAll('div,section,aside')).forEach(function(e){
    var t = (e.innerText||'');
    if (t.length > 1500 || t.length < 10) return;
    if (!/premium member|become a premium|subscribe|upgrade to|membership|unlock/i.test(t)) return;
    var cs = getComputedStyle(e);
    if ((cs.position === 'fixed' || cs.position === 'absolute') && parseInt(cs.zIndex||'0') >= 10){ e.remove(); killed++; }
  });
  Array.from(document.querySelectorAll('div')).forEach(function(d){
    var cs = getComputedStyle(d);
    if (cs.position === 'fixed' && parseInt(cs.zIndex||'0') >= 100 && (d.innerText||'').trim().length === 0){ d.remove(); killed++; }
  });
  document.body.style.overflow = 'auto';
  document.documentElement.style.overflow = 'auto';
  return killed;
})()
"@

$jsHasPassword = "!!document.querySelector('input[type=password]')"

$jsVisibleInputs = @"
Array.from(document.querySelectorAll('input:not([type=hidden]):not([type=submit]), textarea, select')).filter(function(e){ return !!(e.offsetWidth||e.offsetHeight||e.getClientRects().length); }).length
"@

# Closes cookie banners, premium popups, and modal overlays
$jsDismiss = @"
(function(){
  var closed = 0;
  Array.from(document.querySelectorAll('button, a, span[role=button]')).forEach(function(b){
    var t = ((b.innerText||'') + ' ' + (b.getAttribute('aria-label')||'')).trim();
    if (/^(accept all|accept|i agree|agree|got it|allow all|ok|no thanks|not now|maybe later)$/i.test(t) && (b.offsetWidth||b.offsetHeight)){ try{b.click(); closed++}catch(_){}}
  });
  Array.from(document.querySelectorAll('[aria-label="Close" i], .modal .close, .popup-close, .modal-close, .close-button, button.close')).forEach(function(e){
    var p = e.closest('div,section,aside');
    var cs = p ? getComputedStyle(p) : null;
    if (cs && (cs.position === 'fixed' || cs.position === 'absolute') && parseInt(cs.zIndex||'0') >= 10){ try{e.click(); closed++}catch(_){}}
  });
  document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',keyCode:27,bubbles:true}));
  return closed;
})()
"@

function Get-JsFill($profileObj) {
    $json = $profileObj | ConvertTo-Json -Compress
    return @"
(function(){
  var P = $json;
  var map = [
    {pats:['email'], val:P.email},
    {pats:['phone','mobile','tel'], val:P.phone},
    {pats:['firstname','first_name','given-name','givenname'], val:P.firstName},
    {pats:['lastname','last_name','family-name','familyname','surname'], val:P.lastName},
    {pats:['fullname','full_name','name'], val:P.fullName},
    {pats:['city','location','address'], val:P.location},
    {pats:['linkedin'], val:P.linkedin}
  ];
  var inputs = Array.from(document.querySelectorAll('input:not([type=hidden]):not([type=file]):not([type=submit]):not([type=checkbox]):not([type=radio]), textarea'));
  var filled = [];
  inputs.forEach(function(el){
    var id = ((el.name||'')+' '+(el.id||'')+' '+(el.placeholder||'')+' '+(el.autocomplete||'')).toLowerCase();
    for (var m=0; m<map.length; m++){
      if (!map[m].val) continue;
      for (var p=0; p<map[m].pats.length; p++){
        if (id.indexOf(map[m].pats[p]) >= 0 && el.value !== map[m].val){
          el.focus(); el.value = map[m].val;
          el.dispatchEvent(new Event('input',{bubbles:true}));
          el.dispatchEvent(new Event('change',{bubbles:true}));
          filled.push(map[m].pats[p]);
          p = 999; break;
        }
      }
    }
  });
  return filled.join(',');
})()
"@
}

function Get-JsCoverLetter($text) {
    $json = $text | ConvertTo-Json -Compress
    return @"
(function(){
  var t = $json;
  var tas = Array.from(document.querySelectorAll('textarea'));
  for (var i=0;i<tas.length;i++){
    var id = ((tas[i].name||'')+' '+(tas[i].id||'')+' '+(tas[i].placeholder||'')).toLowerCase();
    if (/cover|message|motivation|why/.test(id) || tas[i].rows >= 6){
      tas[i].focus(); tas[i].value = t;
      tas[i].dispatchEvent(new Event('input',{bubbles:true}));
      return 'ok';
    }
  }
  return null;
})()
"@
}

$jsClickSubmit = @"
(function(){
  function inPaywall(e){
    var n = e;
    for (var i=0;i<6 && n;i++){
      var id = ((n.className&&n.className.baseVal!==undefined?n.className.baseVal:n.className)||'') + ' ' + (n.id||'');
      if (/premium|paywall|membership|subscri|upgrade|pricing/i.test(id)) return true;
      n = n.parentElement;
    }
    return false;
  }
  var els = Array.from(document.querySelectorAll('button[type=submit], input[type=submit], button'));
  var cand = els.filter(function(e){
    var t = ((e.innerText||'') + ' ' + (e.value||'')).toLowerCase();
    return /submit|send application|bewerbung absenden|apply now/.test(t) && !/continue|upgrade|pay|subscribe/.test(t) && (e.offsetWidth||e.offsetHeight) && !inPaywall(e);
  });
  if (cand.length){ cand[cand.length-1].scrollIntoView({block:'center'}); cand[cand.length-1].click(); return (cand[cand.length-1].innerText||cand[cand.length-1].value||'submitted').trim(); }
  return null;
})()
"@

function Get-JsLogin($user, $pass) {
    $uj = $user | ConvertTo-Json -Compress
    $pj = $pass | ConvertTo-Json -Compress
    return @"
(function(){
  var u = $uj, p = $pj;
  var inputs = Array.from(document.querySelectorAll('input:not([type=hidden])'));
  var filled = 0;
  inputs.forEach(function(el){
    var id = ((el.name||'')+' '+(el.id||'')+' '+(el.placeholder||'')+' '+(el.type||'')).toLowerCase();
    if (/email|user|login/.test(id) && el.type !== 'password'){ el.value = u; el.dispatchEvent(new Event('input',{bubbles:true})); filled++; }
    if (el.type === 'password'){ el.value = p; el.dispatchEvent(new Event('input',{bubbles:true})); filled++; }
  });
  var btns = Array.from(document.querySelectorAll('button[type=submit], input[type=submit], button'));
  var b = btns.filter(function(e){ return /log ?in|sign ?in|anmelden/i.test((e.innerText||'')+(e.value||'')) && (e.offsetWidth||e.offsetHeight); });
  if (b.length){ b[0].click(); } else if (btns.length){ btns[0].click(); }
  return filled;
})()
"@
}

# ---------- pick jobs ----------
$statuses = @($config.apply.statusesToApply)
$jobs = @(Get-TrackerJobs | Where-Object { $statuses -contains $_.Status -and $_.URL })
# Easy-apply platforms first (Indeed Easy Apply / NaukriGulf) - user priority; then New; then oldest first
$easyRx = "(?i)indeed-easyapply|naukrigulf|linkedin-easyapply"
$jobs = @($jobs | Sort-Object @{ Expression = { if ($_.Source -match $easyRx) { 0 } else { 1 } } }, @{ Expression = { if ($_.Status -eq "New") { 0 } else { 1 } } }, Row)
$RowsFilter = @()
if ($Rows) { $RowsFilter = @($Rows -split "," | ForEach-Object { [int]$_.Trim() }) }
if ($RowsFilter.Count -gt 0) { $jobs = @($jobs | Where-Object { $RowsFilter -contains $_.Row }) }
if ($MaxJobs -le 0) { $MaxJobs = $config.apply.maxJobsPerRun }
$jobs = @($jobs | Select-Object -First $MaxJobs)
if (-not $jobs) { Write-Log "No jobs to apply to. Done."; exit 0 }

$searchPagePattern = "(/search|/q-|-jobs\.html$|/jobs\?|/job-search|/vacancies)"
$accounts = @(Get-TrackerAccounts)

function Get-CvPath($tailoredFolder) {
    if ($tailoredFolder -and (Test-Path $tailoredFolder)) {
        $pdf = Get-ChildItem $tailoredFolder -Filter *.pdf -Recurse | Sort-Object Length -Descending | Select-Object -First 1
        if ($pdf) { return $pdf.FullName }
    }
    return $P.defaultCv
}

function Get-CoverLetter($tailoredFolder) {
    if ($tailoredFolder -and (Test-Path $tailoredFolder)) {
        $cl = Get-ChildItem $tailoredFolder -Filter "cover-letter.md" -Recurse | Select-Object -First 1
        if ($cl) { return (Get-Content $cl.FullName -Raw) }
    }
    return "Dear Hiring Manager,`n`nPlease find attached my CV. I am a BIM Coordinator with 8+ years of MEA/UAE experience and a strong fit for this role.`n`nKind regards,`n$($P.fullName)"
}

# ---------- freshness (user rule 2026-10-26): UAE <= 24h, other countries <= 7 days ----------
function Get-ListedAgeDays($job, $pageText) {
    # 1) explicit age phrase on the posting (EN/DE/FR/ES/NL) - most reliable
    $t = $pageText
    if ($t.Length -gt 6000) { $t = $t.Substring(0, 6000) }
    $pats = @(
        "(?i)(\d+)\s*\+?\s*days?\s+ago",
        "(?i)posted\s+(\d+)\s*\+?\s*days?",
        "(?i)vor\s+(\d+)\s*Tagen",
        "(?i)vor\s+(\d+)\s*Tag\b",
        "(?i)il y a\s+(\d+)\s*jours?",
        "(?i)hace\s+(\d+)\s*d[ií]as?",
        "(?i)(\d+)\s*dagen\s+geleden"
    )
    foreach ($rx in $pats) {
        $m = [regex]::Match($t, $rx)
        if ($m.Success) { return [int]$m.Groups[1].Value }
    }
    $m = [regex]::Match($t, "(?i)vor\s+(\d+)\s*Wochen")
    if ($m.Success) { return [int]$m.Groups[1].Value * 7 }
    if ($t -match "(?i)\byesterday\b|\bgestern\b") { return 1 }
    # 2) fallback: DateFound - posting is at least that old (tracker date is dd/MM/yyyy or yyyy-MM-dd)
    $df = $job.DateFound
    $dt = $null
    if ($df -match "^(\d{1,2})/(\d{1,2})/(\d{4})") {
        $a = [int]$Matches[1]; $b = [int]$Matches[2]
        if ($a -gt 12) { $dt = Get-Date -Day $a -Month $b -Year ([int]$Matches[3]) -Hour 0 -Minute 0 -Second 0 }
        elseif ($b -gt 12) { $dt = Get-Date -Day $b -Month $a -Year ([int]$Matches[3]) -Hour 0 -Minute 0 -Second 0 }
        else { $dt = Get-Date -Day $a -Month $b -Year ([int]$Matches[3]) -Hour 0 -Minute 0 -Second 0 }  # dd/MM (system locale)
    } elseif ($df -match "^(\d{4})-(\d{2})-(\d{2})") {
        $dt = Get-Date -Day ([int]$Matches[3]) -Month ([int]$Matches[2]) -Year ([int]$Matches[1]) -Hour 0 -Minute 0 -Second 0
    }
    if ($dt) { return [Math]::Floor(((Get-Date) - $dt).TotalDays) }
    return -1  # age unknown - do not block
}

function Test-IsUaeJob($job) {
    $locAll = "$($job.Location) $($job.URL) $($job.Title)"
    return [bool]($locAll -match "(?i)\bu\.?a\.?e\.?\b|united arab|dubai|sharjah|abu dhabi|ajman|ae\.indeed\.com")
}

function Find-Account($url) {
    $h = ([uri]$url).Host -replace "^www\.", ""
    foreach ($a in $accounts) {
        $site = ($a.Site -replace "^www\.", "").ToLower()
        $plat = $a.Platform.ToLower()
        if ($site -and $h.Contains($site)) { return $a }
        if ($plat -and $h.Contains($plat)) { return $a }
    }
    return $null
}

# ---------- run ----------
$headless = ($config.apply.headless -and -not $ShowBrowser)
Write-Log "Starting Edge (headless=$headless)..."
try { if ($headless) { Start-EdgeCdp -Headless } else { Start-EdgeCdp }; Enable-Stealth; Enable-AssetBlocking }
catch { Write-Log "FATAL: could not start Edge: $($_.Exception.Message)"; exit 1 }

# Restore saved session cookies (survives browser restarts)
$cookieFile = Join-Path $ProjectDir "edge-profile\cookies.json"
if (Test-Path $cookieFile) {
    try {
        $savedCookies = Get-Content $cookieFile -Raw | ConvertFrom-Json
        if ($savedCookies) {
            Send-Cdp "Network.enable" | Out-Null
            Send-Cdp "Network.setCookies" @{ cookies = @($savedCookies) } -TimeoutSec 30 | Out-Null
            Write-Log "Restored $($savedCookies.Count) saved session cookies."
        }
    } catch { Write-Log "Cookie restore failed: $($_.Exception.Message)" }
}

function Restart-Browser {
    Write-Log "Restarting Edge (connection lost)..."
    try { Stop-EdgeCdp } catch {}
    Start-Sleep -Seconds 2
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -match "JTrackerEdge|edge-profile" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2
    if ($headless) { Start-EdgeCdp -Headless } else { Start-EdgeCdp }
    Enable-Stealth
    Enable-AssetBlocking
    if (Test-Path $cookieFile) {
        $sc = Get-Content $cookieFile -Raw | ConvertFrom-Json
        if ($sc) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($sc) } -TimeoutSec 30 | Out-Null }
    }
    Write-Log "Edge restarted."
}

$results = @()
foreach ($job in $jobs) {
    Write-Log "----- [row $($job.Row)] $($job.Title) @ $($job.Company)"
    Write-Log "URL: $($job.URL)"
    try { Invoke-CdpJs "1+1" | Out-Null } catch { try { Restart-Browser } catch { Write-Log "Browser restart failed: $($_.Exception.Message)"; continue } }

    try {
    if ($job.URL -match $searchPagePattern) {
        Write-Log "SKIP: search-page link, not a direct job post."
        Set-TrackerJobStatus -Row $job.Row -Status "Needs direct link"
        $results += "$($job.Title) @ $($job.Company): Needs direct link"
        continue
    }

    # Adzuna: ApplyIQ never exposes a submit button to automation (verified repeatedly) -
    # stop retrying these forever, they clog the queue ahead of appliable jobs
    if ($job.URL -match "adzuna\.") {
        Write-Log "Adzuna listing - ApplyIQ has no bot-accessible submit. Marking 'External apply' (terminal)."
        Set-TrackerJobStatus -Row $job.Row -Status "External apply"
        $results += "$($job.Title) @ $($job.Company): EXTERNAL apply (Adzuna - manual)"
        continue
    }

    # UAE-ONLY policy (user 2026-11-04): guards rows that entered the tracker before the
    # rule existed - anything without a UAE signal in Location/URL/Title closes unapplied.
    # (Subsumes the old Saudi/Egypt-only exclusion.)
    if (-not (Test-IsUaeJob $job)) {
        Write-Log "SKIP: not a UAE job - closing row (UAE-only policy)."
        Set-TrackerJobStatus -Row $job.Row -Status "Closed (non-UAE)"
        $results += "$($job.Title) @ $($job.Company): Closed (non-UAE)"
        continue
    }

    # Navigate with slow-network recovery. A hung media-heavy page kills the RENDERER
    # (Chromium hang watchdog) which closes the CDP socket - restarting the browser just
    # re-navigates the same heavy page and dies again (10/27 13:14 run: 11/11 jobs).
    # First failure: wait it out + reload, NO restart. Second failure: probe; restart only
    # if the socket is truly dead. Asset blocking prevents most of this in the first place.
    $navOk = $false
    for ($navTry = 0; $navTry -lt 3 -and -not $navOk; $navTry++) {
        try { Navigate-Cdp $job.URL; $navOk = $true }
        catch {
            Write-Log "Navigation issue (try $($navTry+1)): $($_.Exception.Message)"
            if ($navTry -eq 0) {
                Write-Log "  slow/hung page - waiting 10s, then reloading (no restart)."
                Start-Sleep -Seconds 10
                try { Send-Cdp "Page.reload" @{} -TimeoutSec 15 | Out-Null } catch {}
            } else {
                $alive = $false
                for ($pb = 0; $pb -lt 2 -and -not $alive; $pb++) {
                    try { Invoke-CdpJs "1+1" | Out-Null; $alive = $true } catch { Start-Sleep -Seconds 4 }
                }
                if (-not $alive) {
                    Write-Log "CDP dead - one restart, then final try."
                    try { Restart-Browser } catch { Write-Log "Restart failed: $($_.Exception.Message)"; Start-Sleep -Seconds 5 }
                } else {
                    Start-Sleep -Seconds 10
                }
            }
        }
    }
    if (-not $navOk) { Write-Log "Navigation failed 3x - skipping job."; continue }
    $dismissed = Invoke-CdpJs $jsDismiss
    if ($dismissed -gt 0) { Write-Log "Dismissed $dismissed popup(s)/cookie banner(s)."; Start-Sleep -Seconds 2 }
    Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-landing.png")
    $body = Get-CdpBodyText
    for ($retry = 0; $retry -lt 3 -and $body.Length -lt 100; $retry++) { Start-Sleep -Seconds 3; $body = Get-CdpBodyText }
    if ($body.Length -lt 100) {
        Write-Log "Page appears empty/blocked - status 'Needs manual'"
        Set-TrackerJobStatus -Row $job.Row -Status "Needs manual"
        $results += "$($job.Title) @ $($job.Company): Needs manual (page blocked)"
        continue
    }

    # --- FRESHNESS GATE (rule 2026-10-26): UAE posts <= 24h, other countries <= 7 days ---
    # covers ALL channels: Indeed, NaukriGulf Easy Apply, StepStone, TotalJobs, generic forms
    $isUae = Test-IsUaeJob $job
    $maxAge = [int]$config.apply.maxAgeDaysOther
    if ($isUae) { $maxAge = [int]$config.apply.maxAgeDaysUae }
    $age = Get-ListedAgeDays $job $body
    if ($age -gt $maxAge) {
        $zone = "other"; if ($isUae) { $zone = "UAE" }
        Write-Log "TOO OLD - listed ~$age day(s) ago (limit $maxAge day(s) for $zone). No apply - marking expired."
        Set-TrackerJobStatus -Row $job.Row -Status "Expired (age)"
        $results += "$($job.Title) @ $($job.Company): Expired ($age days > ${maxAge}d limit)"
        continue
    }
    if ($age -ge 0) { Write-Log "Freshness OK: ~$age day(s) old (limit $maxAge)." }

    # 1) Click Apply (kill paywalls first, retry once if no form appears)
    $pw = Invoke-CdpJs $jsKillPaywall
    if ($pw -gt 0) { Write-Log "Killed $pw paywall/premium overlay(s)."; Start-Sleep -Seconds 1 }
    $clicked = Invoke-CdpJs $jsClickApply
    if ($clicked) {
        Write-Log "Clicked Apply button: '$clicked'"; Start-Sleep -Seconds 3
        if (Switch-EdgeTab -PreviousUrl $job.URL) {
            Write-Log "Switched to new tab: $(Get-CdpUrl)"
            # Adzuna-style tracking pages auto-redirect onward - follow the chain
            for ($w = 0; $w -lt 10; $w++) {
                Start-Sleep -Seconds 2
                $nowUrl = Get-CdpUrl
                if ($nowUrl -notmatch "adzuna\.(co\.uk|com|co\.nz|nl|de|fr|es|pl|com\.au)/land") { break }
            }
            Write-Log "Final destination: $(Get-CdpUrl)"
            Start-Sleep -Seconds 3
        }
        # confirmation page reached (e.g. StepStone /application/confirmation)? Stop here - do NOT re-click
        $landedUrl = Get-CdpUrl
        if ($landedUrl -match "confirmation|success|thank|merci|danke|bedankt|gracias") {
            Write-Log "Landed on confirmation page: $landedUrl"
            Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-submitted.png")
            $confText = Get-CdpBodyText
            if ($confText -match "thank you|successfully|bestätigt|confirmée|submitted|bewerbung.*erhalten|application.*received") {
                Set-TrackerJobStatus -Row $job.Row -Status "Applied" -DateApplied (Get-Date -Format "yyyy-MM-dd")
                $cnt = Add-TrackerApplicationCount $job.URL
                Write-Log "CONFIRMED applied (confirmation page). (platform total: $cnt)"
                $results += "$($job.Title) @ $($job.Company): APPLIED"
            } else {
                Set-TrackerJobStatus -Row $job.Row -Status "Attempted - check"
                Write-Log "Confirmation page but no success text - 'Attempted - check'"
                $results += "$($job.Title) @ $($job.Company): Attempted - check (confirmation page)"
            }
            continue
        }
    } else { Write-Log "No Apply button found." }
    $pw2 = Invoke-CdpJs $jsKillPaywall
    if ($pw2 -gt 0) { Write-Log "Killed $pw2 paywall overlay(s) after apply click."; Start-Sleep -Seconds 1 }

    # Adzuna email-alert teaser modal? close it and click the REAL apply button
    $alertClosed = Invoke-CdpJs @"
(function(){
  var mods = Array.from(document.querySelectorAll('div, section'));
  for (var i=0;i<mods.length;i++){
    var m = mods[i];
    var t = (m.innerText||'');
    if (t.length < 600 && /create email alert|similar new jobs|job alert/i.test(t) && (m.offsetWidth||m.offsetHeight)){
      var x = m.querySelector('button[class*=close], .close, [aria-label="Close" i], [aria-label="Dismiss" i]');
      if (x){ x.click(); } else { m.remove(); }
      return 'closed';
    }
  }
  return null;
})()
"@
    if ($alertClosed) {
        Write-Log "Closed an email-alert teaser modal - re-clicking real Apply."
        Start-Sleep -Seconds 2
        $clicked3 = Invoke-CdpJs $jsClickApply
        if ($clicked3) { Write-Log "Clicked real Apply: '$clicked3'"; Start-Sleep -Seconds 3; [void](Switch-EdgeTab -PreviousUrl $job.URL) }
    }

    # Adzuna "No thanks, take me to the job" bypass link -> the actual application
    $bypass = Invoke-CdpJs @"
(function(){
  var a = Array.from(document.querySelectorAll('a, button')).find(function(e){ return /no thanks.*take me to the job|take me to the job/i.test(e.innerText||'') && (e.offsetWidth||e.offsetHeight); });
  if (a){ a.click(); return 'clicked'; }
  return null;
})()
"@
    if ($bypass) {
        Write-Log "Clicked 'No thanks, take me to the job' - heading to real application."
        Start-Sleep -Seconds 5
        [void](Switch-EdgeTab -PreviousUrl $job.URL)
        # follow redirect chains (appcast/click trackers) until we settle on a real site
        for ($w = 0; $w -lt 8; $w++) {
            $u = Get-CdpUrl
            if ($u -match "appcast|/click/|click\.|/track/|redirect") { Start-Sleep -Milliseconds 2500 } else { break }
        }
        Write-Log "Now at: $(Get-CdpUrl)"
        [void](Invoke-CdpJs $jsDismiss)
        Start-Sleep -Seconds 2
    }
    $visInputs = Invoke-CdpJs $jsVisibleInputs
    if ($clicked -and $visInputs -lt 2) {
        $d2 = Invoke-CdpJs $jsDismiss
        if ($d2 -gt 0) { Write-Log "Dismissed $d2 more popup(s) after apply click."; Start-Sleep -Seconds 2 }
        $clicked2 = Invoke-CdpJs $jsClickApply
        if ($clicked2) { Write-Log "Re-clicked Apply button: '$clicked2'"; Start-Sleep -Seconds 3 }
    }

    # 1b) DrJobs-style payment wall? (premium required to apply - never pay automatically)    $paymentWall = Invoke-CdpJs "(function(){var m=document.getElementById('subscriptionModal'); if(m && m.offsetWidth && /pay|premium|\\$/i.test(m.innerText)){ var c=m.querySelector('[aria-label=Close], a.pull-right, button.close'); if(c){c.click();} return true; } return false; })()"
    if ($paymentWall -eq $true) {
        Write-Log "PAYMENT WALL: site requires paid premium to apply - closed popup, skipping (never pays automatically)."
        Set-TrackerJobStatus -Row $job.Row -Status "Paywall - manual decision"
        Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-paywall.png")
        $results += "$($job.Title) @ $($job.Company): PAYWALL - needs your decision"
        continue
    }

    # 1c) Ineligible verdict? (platform auto-screened the candidate: location/right-to-work)
    $inelig = Invoke-CdpJs "(function(){ var b = document.querySelector('[data-testid=ineligible-apply-button]'); if (b && b.offsetWidth) return true; var t = (document.body.innerText||''); if (/you are not eligible|not eligible to apply/i.test(t)) return true; return false; })()"
    if ($inelig -eq $true) {
        Write-Log "INELIGIBLE: platform auto-screened this profile (usually location/right-to-work) - marking, will not retry."
        Set-TrackerJobStatus -Row $job.Row -Status "Ineligible (location)"
        Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-ineligible.png")
        $results += "$($job.Title) @ $($job.Company): INELIGIBLE (location screening)"
        continue
    }

    # 2) Login wall?
    $hasPwd = Invoke-CdpJs $jsHasPassword
    if ($hasPwd -eq $true) {
        $acc = Find-Account (Get-CdpUrl)
        if ($acc -and $acc.Password -eq "GOOGLE-LOGIN") {
            # Session-based account: trust logged-in page markers over password-field detection
            $bodyNow = Get-CdpBodyText
            $looksLoggedIn = ($bodyNow -match "logout|log out|sign out|my account|my profile|dashboard|silar")
            Start-Sleep -Seconds 3
            if ((Invoke-CdpJs $jsHasPassword) -eq $true -and -not $looksLoggedIn) {
                Write-Log "Session expired - run Login-Setup.ps1 to log in again."
                Set-TrackerJobStatus -Row $job.Row -Status "Needs login session"
                Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-login.png")
                $results += "$($job.Title) @ $($job.Company): Needs login session"
                continue
            }
            Write-Log "Session alive (logged-in markers found) - proceeding."
        }
        elseif ($acc) {
            $plain = Unprotect-TrackerPassword $acc.Password
            if ($plain) {
                Write-Log "Login wall - trying account $($acc.Username)"
                $n = Invoke-CdpJs (Get-JsLogin $acc.Username $plain)
                Write-Log "Login fill count: $n"
                Start-Sleep -Seconds 5
                Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-after-login.png")
            } else { Write-Log "Account found but password could not be decrypted (created by different Windows user?)." }
        }
        # still a password field after login attempt?
        if ((Invoke-CdpJs $jsHasPassword) -eq $true) {
            Write-Log "Still at login - status 'Needs account'"
            Set-TrackerJobStatus -Row $job.Row -Status "Needs account"
            Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-login.png")
            $results += "$($job.Title) @ $($job.Company): Needs account"
            continue
        }
    }

    # 2b) Indeed Easy Apply multi-step flow (contact -> resume -> questions -> review -> submit)
    if ((Get-CdpUrl) -match "indeed\.com|smartapply\.indeed") {
        # If still on the job post, click the real 'Apply now' link first (generic clicker misses it)
        if ((Get-CdpUrl) -match "indeed\.com/viewjob") {
            $c = $null
            for ($try = 1; $try -le 5 -and -not $c; $try++) {
                $c = Invoke-CdpJs "(function(){ var a = Array.from(document.querySelectorAll('a, button')).find(function(e){ return /apply now/i.test(e.innerText||'') && (e.offsetWidth||e.offsetHeight); }); if (!a) return null; a.scrollIntoView({block:'center'}); a.click(); return 'clicked'; })()"
                if (-not $c) { Start-Sleep -Seconds 3 }
            }
            if ($c) { Write-Log "Clicked Indeed 'Apply now' link."; Start-Sleep -Seconds 6; [void](Switch-EdgeTab -PreviousUrl $job.URL) }
        }
        # Wait for the smartapply SPA to route into a real form module
        for ($w = 0; $w -lt 8; $w++) {
            $u = Get-CdpUrl
            if ($u -match "smartapply\.indeed\.com" -and $u -notmatch "/form/") { Start-Sleep -Milliseconds 2500 } else { break }
        }
        Write-Log "Indeed flow page: $(Get-CdpUrl)"
        # --- prepare JD-matched tailored CV (Easy Apply must get a tailored CV, not the generic one) ---
        $tailoredRoot = Join-Path $config.paths.root "Tailored"
        if (-not $job.TailoredFiles) {
            try {
                $jdText = Get-CdpBodyText
                $tPdf = New-TailoredCVForJob -Job $job -JdText $jdText -TailoredRoot $tailoredRoot -Profile $P
                if ($tPdf) {
                    $job.TailoredFiles = Split-Path $tPdf -Parent
                    try { Set-TrackerJobTailored -Row $job.Row -Folder $job.TailoredFiles } catch {}
                    Write-Log "Tailored CV prepared: $tPdf"
                } else { Write-Log "Tailored CV generation failed - using default CV" }
            } catch { Write-Log "Tailoring error: $($_.Exception.Message)" }
        }
        $cv2 = Get-CvPath $job.TailoredFiles
        $isTailored = ($cv2 -ne $P.defaultCv)
        $done = $false
        $stuckCount = 0
        $noneCount = 0
        $prevModule = ""
        for ($step = 1; $step -le 12; $step++) {
            [void](Invoke-CdpJs (Get-JsFill $P))
            # answer dropdowns (first option), radio groups (prefer 'Yes'), required checkboxes
            $ans = Invoke-CdpJs @"
(function(){
  var n = 0;
  Array.from(document.querySelectorAll('select')).forEach(function(s){
    if (!s.value && s.options.length > 1){ s.selectedIndex = 1; s.dispatchEvent(new Event('change',{bubbles:true})); n++; }
  });
  var groups = {};
  Array.from(document.querySelectorAll('input[type=radio]')).forEach(function(r){ var g = r.name || r.id; (groups[g] = groups[g] || []).push(r); });
  Object.keys(groups).forEach(function(g){
    if (!groups[g].some(function(r){ return r.checked; })){
      var yes = groups[g].find(function(r){ return /yes/i.test((r.value||'')); });
      (yes || groups[g][0]).click(); n++;
    }
  });
  Array.from(document.querySelectorAll('input[type=checkbox][required]')).forEach(function(c){ if (!c.checked){ c.click(); n++; } });
  return n;
})()
"@
            if ($ans -gt 0) { Write-Log "Answered $ans question(s) (step $step)." }
            # Tailored CV available -> upload it (click "Upload resume" first if no file input yet).
            # Otherwise fall back to the user's existing Indeed Resume (less error-prone than upload).
            $uploadedTailored = $false
            if ($isTailored) {
                $cvOk2 = $false
                try { $cvOk2 = Set-CdpFileInput $cv2 } catch {}
                if (-not $cvOk2) {
                    Invoke-CdpJs "(function(){ var t = Array.from(document.querySelectorAll('label, div[role=radio], div[role=button], button')).find(function(e){ return /upload (your )?resume/i.test(e.innerText||''); }); if(t){ t.click(); return 'clicked'; } return null; })()" | Out-Null
                    Start-Sleep -Seconds 2
                    try { $cvOk2 = Set-CdpFileInput $cv2 } catch {}
                }
                if ($cvOk2) { $uploadedTailored = $true; Write-Log "Tailored CV uploaded (step $step)." }
                else { Write-Log "Tailored CV upload failed (step $step) - falling back to Indeed Resume" }
            }
            if (-not $uploadedTailored) {
                $useIa = Invoke-CdpJs @"
(function(){
  if (!/add a resume|resume/i.test(document.body.innerText.substring(0,2000))) return 'not-resume-page';
  var cand = Array.from(document.querySelectorAll('label, div[role=radio], div[role=button], input[type=radio]'));
  var t = cand.find(function(e){ return /use your indeed resume/i.test(e.innerText||''); });
  if (t){ t.click(); return 'clicked'; }
  return null;
})()
"@
                if ($useIa -eq 'clicked') {
                    [void](Invoke-CdpRealClick "Array.from(document.querySelectorAll('label, div[role=radio], div[role=button]')).find(function(e){ return /use your indeed resume/i.test(e.innerText||''); })")
                    Write-Log "Selected 'Use your Indeed Resume' (real click)."
                } else {
                    $cvOk2 = $false
                    try { $cvOk2 = Set-CdpFileInput $cv2 } catch {}
                    if ($cvOk2) { Write-Log "CV uploaded (step $step)." }
                }
            }
            $curModule = Get-CdpUrl
            if ($curModule -eq $prevModule) { $stuckCount++ } else { $stuckCount = 0 }
            $prevModule = $curModule
            if ($stuckCount -ge 3) { Write-Log "Stuck on same module - likely an unsupported question type."; break }
            $btn = Invoke-CdpJs @"
(function(){
  var els = Array.from(document.querySelectorAll('button, input[type=submit]'));
  function vis(e){ return !!(e.offsetWidth||e.offsetHeight); }
  var sub = els.filter(function(e){ return /submit application|submit/i.test((e.innerText||'')+(e.value||'')) && vis(e) && !/continue/.test(e.innerText||''); });
  if (sub.length){ window.__indeedBtn = sub[sub.length-1]; return 'SUBMIT'; }
  var cont = els.filter(function(e){ return /continue|review|next/i.test((e.innerText||'')+(e.value||'')) && vis(e); });
  if (cont.length){ window.__indeedBtn = cont[0]; return 'CONTINUE'; }
  return 'NONE';
})()
"@
            if ($btn -ne 'NONE') { [void](Invoke-CdpRealClick "window.__indeedBtn") }
            Write-Log "Indeed step $step -> $btn (real click)"
            if ($btn -eq "SUBMIT") {
                Start-Sleep -Seconds 6
                Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-submitted.png")
                $after2 = Get-CdpBodyText
                if ($after2 -match "application (was |has been )?(submitted|sent)|thank you|successfully") {
                    Set-TrackerJobStatus -Row $job.Row -Status "Applied" -DateApplied (Get-Date -Format "yyyy-MM-dd")
                    $cnt = Add-TrackerApplicationCount $job.URL
                    Write-Log "CONFIRMED applied via Indeed Easy Apply. (platform total: $cnt)"
                    $results += "$($job.Title) @ $($job.Company): APPLIED (Indeed Easy Apply)"
                } else {
                    Set-TrackerJobStatus -Row $job.Row -Status "Attempted - check"
                    $results += "$($job.Title) @ $($job.Company): Attempted - check"
                }
                $done = $true
                break
            }
            elseif ($btn -eq "CONTINUE") { $noneCount = 0; Start-Sleep -Seconds (4 + (Get-Random -Minimum 0 -Maximum 3)); continue }
            else { $noneCount++; if ($noneCount -ge 3) { break }; Start-Sleep -Seconds 3 }
        }
        if (-not $done) {
            Set-TrackerJobStatus -Row $job.Row -Status "New"
            Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-indeed-stuck.png")
            Write-Log "Indeed flow incomplete - leaving as New for retry"
            $results += "$($job.Title) @ $($job.Company): retry next run"
        }
        continue
    }

    # 2c) NaukriGulf Easy Apply flow (Easy Apply -> Submit & Apply, CV from profile)
    if ($job.URL -match "naukrigulf\.com" -or (Get-CdpUrl) -match "naukrigulf\.com") {
        # error-page recovery (NaukriGulf throttles rapid loads): reload with backoff
        $recovered = $false
        for ($rl = 1; $rl -le 3; $rl++) {
            $curUrl = Get-CdpUrl
            $bodyLen = (Get-CdpBodyText).Length
            if ($curUrl -match "^chrome-error" -or $bodyLen -lt 200) {
                Write-Log "Error/empty page detected (attempt $rl) - waiting and reloading..."
                Start-Sleep -Seconds (10 + $rl * 5)
                try { Navigate-Cdp $job.URL } catch {}
                Start-Sleep -Seconds 5
            } else { $recovered = $true; break }
        }
        if (-not $recovered) {
            Write-Log "Page still blocked after retries - restarting browser once..."
            try { Restart-Browser } catch {}
            try { Navigate-Cdp $job.URL } catch {}
            Start-Sleep -Seconds 6
        }
        $ngDone = $false
        $already = Invoke-CdpJs "(function(){ return /already applied/i.test(document.body.innerText); })()"
        if ($already -eq $true) {
            Write-Log "Already applied (page shows confirmation)."
            $ngDone = $true
        } else {
            $c = $null
            for ($try = 1; $try -le 4 -and -not $c; $try++) {
                $c = Invoke-CdpJs "(function(){ var b = Array.from(document.querySelectorAll('button, a')).find(function(e){ return /^easy apply$/i.test((e.innerText||'').trim()); }); if (!b) return null; b.scrollIntoView({block:'center'}); b.click(); return 'clicked'; })()"
                if (-not $c) { Start-Sleep -Seconds 4 }
            }
            if ($c) { Write-Log "Clicked NaukriGulf 'Easy Apply'."; Start-Sleep -Seconds 6 }
            else {
                # no Easy Apply: check for plain external Apply (jobSource=WEB) - not appliable on-site
                $extApply = Invoke-CdpJs "(function(){ var b = Array.from(document.querySelectorAll('button, a')).find(function(e){ return /^apply$/i.test((e.innerText||'').trim()); }); return !!b; })()"
                if ($extApply -eq $true) {
                    Write-Log "EXTERNAL apply only (no on-site application) - marking terminal."
                    Set-TrackerJobStatus -Row $job.Row -Status "External apply"
                    $results += "$($job.Title) @ $($job.Company): EXTERNAL apply (manual)"
                    Start-Sleep -Seconds 8
                    continue
                }
            }
            $s = $null
            for ($try2 = 1; $try2 -le 3 -and -not $s; $try2++) {
                $s = Invoke-CdpJs "(function(){ var b = Array.from(document.querySelectorAll('button')).find(function(e){ return /submit & apply/i.test((e.innerText||'')); }); if (!b) return null; b.click(); return 'clicked'; })()"
                if (-not $s) { Start-Sleep -Seconds 4 }
            }
            if ($s) {
                Write-Log "Clicked 'Submit & Apply'."
                Start-Sleep -Seconds 8
                Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-submitted.png")
                $afterNg = Get-CdpBodyText
                if ($afterNg -match "already applied|successfully applied|application submitted|thank you") {
                    $ngDone = $true
                } else {
                    Write-Log "No confirmation text - rechecking page state..."
                    Start-Sleep -Seconds 5
                    $recheck = Invoke-CdpJs "(function(){ return /already applied|successfully applied/i.test(document.body.innerText); })()"
                    if ($recheck -eq $true) { $ngDone = $true }
                }
            } else {
                $landedUrl = Get-CdpUrl
                if ($landedUrl -match "simJobs") {
                    Write-Log "Landed on similar-jobs page (external apply) - marking terminal."
                    Set-TrackerJobStatus -Row $job.Row -Status "External apply"
                    $results += "$($job.Title) @ $($job.Company): EXTERNAL apply (manual)"
                    Start-Sleep -Seconds 8
                    continue
                }
                Write-Log "Submit & Apply button not found."
            }
        }
        if ($ngDone) {
            Set-TrackerJobStatus -Row $job.Row -Status "Applied" -DateApplied (Get-Date -Format "yyyy-MM-dd")
            $cnt = Add-TrackerApplicationCount $job.URL
            Write-Log "CONFIRMED applied via NaukriGulf Easy Apply. (platform total: $cnt)"
            $results += "$($job.Title) @ $($job.Company): APPLIED (NaukriGulf)"
        } else {
            Set-TrackerJobStatus -Row $job.Row -Status "New"
            Write-Log "NaukriGulf flow incomplete - leaving as New for retry"
            $results += "$($job.Title) @ $($job.Company): retry next run"
        }
        Start-Sleep -Seconds 8  # pacing: NaukriGulf throttles rapid-fire loads
        continue
    }

    # 3) Fill form
    $cv = Get-CvPath $job.TailoredFiles
    $filled = Invoke-CdpJs (Get-JsFill $P)
    Write-Log "Fields filled: $filled"
    $cover = Get-CoverLetter $job.TailoredFiles
    $cl = Invoke-CdpJs (Get-JsCoverLetter $cover)
    if ($cl) { Write-Log "Cover letter pasted." }
    $cvOk = $false
    try { $cvOk = Set-CdpFileInput $cv } catch { Write-Log "CV upload error: $($_.Exception.Message)" }
    Write-Log "CV upload: $cvOk ($cv)"
    Start-Sleep -Seconds 2
    Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-filled.png")

    # 4) Submit
    if ($DryRun) {
        Write-Log "DRY RUN - not submitting."
        Set-TrackerJobStatus -Row $job.Row -Status "Ready (dry run)"
        $results += "$($job.Title) @ $($job.Company): DRY RUN filled"
        continue
    }
    if ($config.apply.autoSubmit) {
        [void](Invoke-CdpJs $jsKillPaywall)
        $sub = Invoke-CdpJs $jsClickSubmit
        if ($sub) {
            Write-Log "Clicked submit: '$sub'"
            Start-Sleep -Seconds 6
            Save-CdpScreenshot (Join-Path $shotDir "job-$($job.Row)-submitted.png")
            $after = Get-CdpBodyText
            if ($after -match "thank you|application (has been|was) (submitted|received|sent)|successfully (applied|submitted)|application complete|application sent") {
                Set-TrackerJobStatus -Row $job.Row -Status "Applied" -DateApplied (Get-Date -Format "yyyy-MM-dd")
                $cnt = Add-TrackerApplicationCount $job.URL
                Write-Log "CONFIRMED applied. (platform total: $cnt)"
                $results += "$($job.Title) @ $($job.Company): APPLIED"
            } else {
                Set-TrackerJobStatus -Row $job.Row -Status "Attempted - check"
                Write-Log "No confirmation detected - 'Attempted - check'"
                $results += "$($job.Title) @ $($job.Company): Attempted - check"
            }
        } else {
            Set-TrackerJobStatus -Row $job.Row -Status "Needs manual"
            Write-Log "No submit button found - 'Needs manual'"
            $results += "$($job.Title) @ $($job.Company): Needs manual"
        }
    }
    } catch {
        Write-Log "Job crashed: $($_.Exception.Message) - recovering browser and moving on"
        try { Restart-Browser } catch { Write-Log "Recovery failed: $($_.Exception.Message)" }
    }
}

Stop-EdgeCdp
Write-Log "===== RUN SUMMARY ====="
foreach ($r in $results) { Write-Log $r }
Write-Host "`nApplyBot done. Log: $logFile" -ForegroundColor Green
