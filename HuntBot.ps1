<#
.SYNOPSIS
  HUNTBOT : searches job boards for BIM/Revit roles and adds new matches to JobTracker.xlsx.
  Sources with free public APIs: RemoteOK, Remotive, Arbeitnow.
  Writes a report to reports\hunt-report-<date>.md
#>
param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "config.json")
)

$ErrorActionPreference = "Continue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "Tracker.psm1") -Force

# Deep JSON parser: PS 5.1 ConvertFrom-Json collapses arrays of objects
# (returns 1 object with array props) and has no -Depth parameter.
# JavaScriptSerializer handles them correctly.
Add-Type -AssemblyName System.Web.Extensions
function ConvertTo-PSObjectDeep($o) {
    if ($o -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $o.Keys) { $h[[string]$k] = ConvertTo-PSObjectDeep $o[$k] }
        return [PSCustomObject]$h
    } elseif ($o -is [System.Collections.IList] -and $o -isnot [string]) {
        $arr = @()
        foreach ($e in $o) { $arr += ConvertTo-PSObjectDeep $e }
        return $arr
    } else { return $o }
}
function ConvertFrom-JsonDeep {
    param([string]$Json)
    $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $ser.MaxJsonLength = 10485760
    $ser.RecursionLimit = 100
    return (ConvertTo-PSObjectDeep ($ser.DeserializeObject($Json)))
}

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Set-TrackerPath $config.paths.tracker

# Guard against CONCURRENT runs: two HuntBots racing EPPlus load-modify-save on
# JobTracker.xlsx lose rows (last save wins) - happened 10/27 12:00 when the scheduled
# task overlapped a manual pipeline run and 334 drained rows vanished. Lock file on
# LOCAL disk (OneDrive-hosted locks proved undeletable when the sync engine holds them
# open - that locked the fleet out for hours on 10/27). Pipeline locks get a 120-min
# staleness window, standalone HuntBot locks 30-min; crashes self-heal after that.
$runLock = Join-Path $env:LOCALAPPDATA "JTrackerEdge\fleet.lock"
New-Item -ItemType Directory -Force (Split-Path $runLock) | Out-Null
if (-not $env:JTRACKER_PIPELINE_RUN -and (Test-Path $runLock)) {
    $lockAge = (Get-Date) - (Get-Item $runLock).LastWriteTime
    $lockTxt = ""; try { $lockTxt = Get-Content $runLock -Raw -ErrorAction Stop } catch {}
    $limit = if ($lockTxt -match "^\s*pipeline") { 120 } else { 30 }
    if ($lockAge.TotalMinutes -lt $limit) {
        Write-Host "Another fleet run is active (lock $([int]$lockAge.TotalMinutes) min old, limit $limit) - exiting."
        exit
    }
}
if (-not $env:JTRACKER_PIPELINE_RUN) { Set-Content $runLock -Value "huntbot standalone $(Get-Date -Format o)" -Encoding UTF8 }

$logDir = Join-Path $ProjectDir "logs"; New-Item -ItemType Directory -Force $logDir | Out-Null
$repDir = Join-Path $ProjectDir "reports"; New-Item -ItemType Directory -Force $repDir | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmm"
$report = Join-Path $repDir "hunt-report-$stamp.md"

function Write-Log($msg) { $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"; Write-Host $line; Add-Content (Join-Path $logDir "hunt.log") $line }

function Get-Score($title, $desc) {
    # Title must carry the signal. Description/location can only add small bonuses.
    $t = "$title".ToLower()
    $d = "$desc".ToLower()
    $score = 0
    foreach ($q in $config.search.queries) { if ($t -match [regex]::Escape($q.ToLower())) { $score += 5 } }
    foreach ($k in $config.search.keywords) { if ($t -match [regex]::Escape($k.ToLower())) { $score += 3 } }
    $descHits = 0
    foreach ($k in $config.search.keywords) { if ($d -match [regex]::Escape($k.ToLower())) { $descHits++ } }
    $score += [Math]::Min($descHits, 2)   # max +2 from description
    foreach ($l in $config.search.locations) { if ($t -match [regex]::Escape($l.ToLower()) -or $d.Substring(0, [Math]::Min(300, $d.Length)) -match [regex]::Escape($l.ToLower())) { $score += 1; break } }
    return $score
}

$found = New-Object System.Collections.Generic.List[object]
$ua = @{ "User-Agent" = "Mozilla/5.0 (JobHunt research)" }

# --- Source 1: RemoteOK ---
try {
    $raw = Invoke-RestMethod "https://remoteok.com/api" -Headers $ua -TimeoutSec 30
    foreach ($j in ($raw | Where-Object { $_.position })) {
        $found.Add([PSCustomObject]@{
            Title = $j.position; Company = $j.company; Location = $j.location
            Source = "RemoteOK"; URL = $j.url; Desc = ($j.description -replace "<[^>]+>", " ")
            Posted = $j.date
        })
    }
    Write-Log "RemoteOK: $($raw.Count - 1) listings fetched"
} catch { Write-Log "RemoteOK FAILED: $($_.Exception.Message)" }

# --- Source 2: Remotive ---
foreach ($q in $config.search.queries) {
    try {
        $r = Invoke-RestMethod ("https://remotive.com/api/remote-jobs?search=" + [uri]::EscapeDataString($q)) -Headers $ua -TimeoutSec 30
        foreach ($j in $r.jobs) {
            $found.Add([PSCustomObject]@{
                Title = $j.title; Company = $j.company_name; Location = $j.candidate_required_location
                Source = "Remotive"; URL = $j.url; Desc = ($j.description -replace "<[^>]+>", " ")
                Posted = $j.publication_date
            })
        }
        Write-Log "Remotive '$q': $($r.jobs.Count) listings fetched"
    } catch { Write-Log "Remotive '$q' FAILED: $($_.Exception.Message)" }
}

# --- Source 3: Arbeitnow ---
try {
    $r = Invoke-RestMethod "https://www.arbeitnow.com/api/job-board-api" -Headers $ua -TimeoutSec 30
    foreach ($j in $r.data) {
        $found.Add([PSCustomObject]@{
            Title = $j.title; Company = $j.company_name; Location = $j.location
            Source = "Arbeitnow"; URL = $j.url; Desc = ($j.description -replace "<[^>]+>", " ")
            Posted = ""
        })
    }
    Write-Log "Arbeitnow: $($r.data.Count) listings fetched"
} catch { Write-Log "Arbeitnow FAILED: $($_.Exception.Message)" }

# --- Source 4: DrJobs UAE (disabled - paywall blocks applying, user chose to skip) ---
if ($config.sources.drjobs -ne $false) {
try {
    $html = (Invoke-WebRequest "https://drjobs.ae/bim-jobs" -Headers $ua -TimeoutSec 30 -UseBasicParsing).Content
    $cities = @("abu-dhabi","ras-al-khaimah","umm-al-quwain","dubai","sharjah","ajman","fujairah","al-ain")
    $cardRx = [regex]'(?si)href="(https://drjobs\.ae/jobs/([a-z0-9-]+)-([A-Z0-9]{10,}))".{0,4000}?(\d+\s+hours?\s+ago|Yesterday|\d+\s+days?\s+ago)'
    $seen = @{}
    foreach ($m in $cardRx.Matches($html)) {
        $url = $m.Groups[1].Value
        if ($seen.ContainsKey($url)) { continue }
        $seen[$url] = $true
        $slug = $m.Groups[2].Value
        $when = $m.Groups[4].Value
        # only fresh posts (<= 3 days)
        $posted = Get-Date
        if ($when -match "^(\d+)\s+day") { if ([int]$Matches[1] -gt 3) { continue }; $posted = (Get-Date).AddDays(-[int]$Matches[1]) }
        elseif ($when -match "^Yesterday") { $posted = (Get-Date).AddDays(-1) }
        # parse slug: {title}-{city}-{company}
        $parts = $slug -split "-"
        $cityIdx = -1; $cityName = ""
        for ($i = 0; $i -lt $parts.Count; $i++) {
            foreach ($c in $cities) {
                $cp = $c -split "-"
                if ($i + $cp.Count -le $parts.Count -and (($parts[$i..($i + $cp.Count - 1)] -join "-") -eq $c)) { $cityIdx = $i; $cityName = ($cp | ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_) }) -join " "; break }
            }
            if ($cityIdx -ge 0) { break }
        }
        if ($cityIdx -lt 1) { continue }
        $title = (Get-Culture).TextInfo.ToTitleCase(($parts[0..($cityIdx - 1)] -join " "))
        $company = ""
        if ($cityIdx + 1 -lt $parts.Count) { $company = (Get-Culture).TextInfo.ToTitleCase(($parts[($cityIdx + 1)..($parts.Count - 1)] -join " ")) }
        $found.Add([PSCustomObject]@{
            Title = $title; Company = $company; Location = "$cityName, UAE"
            Source = "DrJobs"; URL = $url; Desc = $title; Posted = $posted.ToString("yyyy-MM-dd")
        })
    }
    Write-Log "DrJobs UAE: $($seen.Count) fresh BIM listings parsed"
} catch { Write-Log "DrJobs FAILED: $($_.Exception.Message)" }
} else { Write-Log "DrJobs source: disabled in config (paywall)" }

# --- Source 5: Adzuna API (UK, USA, AU, NL, DE, IE, FR, PL) ---
if ($config.sources.adzuna -ne $false -and $config.adzuna.appId) {
    foreach ($cc in $config.adzuna.countries) {
        $done = $false
        for ($attempt = 1; $attempt -le 3 -and -not $done; $attempt++) {
        try {
            $u = "https://api.adzuna.com/v1/api/jobs/$cc/search/1?app_id=$($config.adzuna.appId)&app_key=$($config.adzuna.appKey)&what=bim&results_per_page=50&sort_by=date"
            $r = Invoke-RestMethod $u -Headers $ua -TimeoutSec 30
            $n = 0
            foreach ($j in $r.results) {
                $found.Add([PSCustomObject]@{
                    Title = ($j.title -replace "<[^>]+>", ""); Company = $j.company.display_name
                    Location = $j.location.display_name; Source = "Adzuna-$($cc.ToUpper())"
                    URL = $j.redirect_url; Desc = ($j.description -replace "<[^>]+>", " ")
                    Posted = $j.created
                })
                $n++
            }
            Write-Log "Adzuna $($cc.ToUpper()): $n listings fetched (total $($r.count))"
            $done = $true
        } catch {
            if ($attempt -lt 3) { Write-Log "Adzuna $($cc.ToUpper()) attempt $attempt failed - retrying..."; Start-Sleep -Seconds 5 }
            else { Write-Log "Adzuna $($cc.ToUpper()) FAILED after 3 tries: $($_.Exception.Message)" }
        }
        }
    }
}

# --- Source 6: Indeed UAE (browser + session + stealth; direct Easy Apply links) ---
if ($config.sources.indeed -ne $false) {
    try {
        Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking
        Start-EdgeCdp -Headless
        Enable-Stealth
        Enable-AssetBlocking
        $cookieFile = Join-Path $ProjectDir "edge-profile\cookies.json"
        if (Test-Path $cookieFile) {
            $sc = Get-Content $cookieFile -Raw | ConvertFrom-Json
            if ($sc) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($sc) } -TimeoutSec 30 | Out-Null }
        }
        $jsCards = @"
(function(){
  var seen = {};
  var cards = Array.from(document.querySelectorAll('div.job_seen_beacon, td.resultContent'));
  var jobs = [];
  cards.forEach(function(c){
    var a = c.querySelector('h2.jobTitle a, a.jcs-JobTitle');
    if (!a) return;
    var href = a.getAttribute('href') || '';
    var m = href.match(/jk=([a-f0-9]+)/i);
    if (!m) return;
    var jk = m[1];
    if (seen[jk]) return;
    seen[jk] = true;
    var comp = c.querySelector('[data-testid=company-name], .companyName');
    var loc = c.querySelector('[data-testid=text-location], .companyLocation');
    var date = c.querySelector('[data-testid=myJobsStateDate], .date');
    jobs.push({ title: a.innerText.trim(), company: comp ? comp.innerText.trim() : '',
      loc: loc ? loc.innerText.trim() : '', date: date ? date.innerText.trim() : '',
      jk: jk, easy: /easily apply/i.test(c.innerText) });
  });
  return JSON.stringify(jobs);
})()
"@
        $total = 0
        $indeedQueries = @($config.search.queries) + @("BIM", "Revit", "VDC")
        $indeedQueries = @($indeedQueries | Sort-Object -Unique)
        # Multi-location sweep: every query on the primary location, top queries on the rest
        $indeedLocs = @($config.search.indeedLocations)
        if ($indeedLocs.Count -eq 0) { $indeedLocs = @("Dubai") }
        $locPlan = @()
        foreach ($iq in $indeedQueries) { $locPlan += , @($iq, $indeedLocs[0]) }
        foreach ($iq in @($indeedQueries | Select-Object -First 2)) {
            foreach ($il in @($indeedLocs | Select-Object -Skip 1)) { $locPlan += , @($iq, $il) }
        }
        foreach ($pair in $locPlan) {
            $q = $pair[0]; $loc = $pair[1]
            $url = "https://ae.indeed.com/jobs?q=" + [uri]::EscapeDataString($q) + "&l=" + [uri]::EscapeDataString($loc) + "&sort=date"
            Navigate-Cdp $url
            Start-Sleep -Seconds 4
            $body = Get-CdpBodyText
            if ($body -match "Request Blocked|unusual traffic|Just a moment|Verifying you are|verify you are human|captcha") { Write-Log "Indeed BLOCKED on '$q' - skipping"; continue }
            $raw = Invoke-CdpJs $jsCards
            $ijobs = @()
            if ($raw) {
                $parsed = ConvertFrom-JsonDeep $raw
                if ($parsed -isnot [System.Array]) { $parsed = @($parsed) }
                $ijobs = $parsed
            }
            if ($ijobs.Count -eq 0) {
                $pgTitle = Invoke-CdpJs "document.title"
                $snippet = Invoke-CdpJs "document.body ? document.body.innerText.substring(0, 200) : ''"
                Write-Log "Indeed '$q': 0 cards (title='$pgTitle', page='$snippet') - retrying once..."
                Start-Sleep -Seconds 8
                $raw = Invoke-CdpJs $jsCards
                if ($raw) {
                    $parsed = ConvertFrom-JsonDeep $raw
                    if ($parsed -isnot [System.Array]) { $parsed = @($parsed) }
                    $ijobs = $parsed
                }
                if ($ijobs.Count -eq 0) { Write-Log "Indeed '$q': still 0 after retry" }
            }
            foreach ($j in $ijobs) {
                $found.Add([PSCustomObject]@{
                    Title = $j.title; Company = $j.company; Location = ($j.loc -replace "\s+", " ")
                    Source = $(if ($j.easy) { "Indeed-EasyApply" } else { "Indeed" })
                    URL = "https://ae.indeed.com/viewjob?jk=$($j.jk)"
                    Desc = "$($j.title) $($j.date)"; Posted = $j.date
                })
                $total++
            }
            Write-Log "Indeed '$q' @ $loc : $($ijobs.Count) direct posts ($( @($ijobs | Where-Object { $_.easy }).Count ) Easy Apply)"
            Start-Sleep -Seconds 3
        }
        Stop-EdgeCdp
        Write-Log "Indeed total: $total direct job posts"
    } catch { Write-Log "Indeed FAILED: $($_.Exception.Message)"; try { Stop-EdgeCdp } catch {} }
}

# --- Source 7: StepStone Germany (plain HTTP, 985 BIM jobs, date-sorted) ---
if ($config.sources.stepstone -ne $false) {
    foreach ($stq in @("bim", "revit")) {
    try {
        $stHtml = (Invoke-WebRequest "https://www.stepstone.de/jobs/$stq/" -Headers $ua -TimeoutSec 45 -UseBasicParsing).Content
        $cards = $stHtml -split 'data-testid="job-item"'
        $n = 0
        for ($i = 1; $i -lt $cards.Count; $i++) {
            $c = $cards[$i]
            $um = [regex]::Match($c, 'href="(https://www\.stepstone\.de/stellenangebote--[^"]+?)"')
            if (-not $um.Success) { continue }
            $tm = [regex]::Match($c, 'data-testid="job-item-title"[^>]*>.*?>([^<>]{3,120})<', [Text.RegularExpressions.RegexOptions]::Singleline)
            $title = if ($tm.Success) { $tm.Groups[1].Value.Trim() } else { "" }
            if (-not $title) { continue }
            $cm = [regex]::Match($c, 'alt="([^"]{2,80})"')
            $company = if ($cm.Success) { $cm.Groups[1].Value.Trim() } else { "" }
            $locs = [regex]::Matches($c, 'data-genesis-element="TEXT">([^<>]{2,60})<') | ForEach-Object { $_.Groups[1].Value.Trim() } | Where-Object { $_ -notmatch "^(Teilweise Home-Office|Vollzeit|Teilzeit|Schnelle Bewerbung|Anschreiben nicht erforderlich)$" }
            $loc = if ($locs.Count -gt 0) { $locs[0] } else { "Germany" }
            $dm = [regex]::Match($c, "(vor \d+ (Stunden|Stunde|Tagen|Tag|Wochen|Woche)|vor \d+ Minuten|Heute|Gestern)")
            $found.Add([PSCustomObject]@{
                Title = $title; Company = $company; Location = ($loc + ", Germany")
                Source = "StepStone"; URL = $um.Groups[1].Value
                Desc = "$title $company"; Posted = $(if ($dm.Success) { $dm.Groups[1].Value } else { "" })
            })
            $n++
        }
        Write-Log "StepStone '$stq': $n listings parsed"
    } catch { Write-Log "StepStone '$stq' FAILED: $($_.Exception.Message)" }
    }
}

# --- Sources 8+9: TotalJobs UK + NaukriGulf (stealth browser, one session) ---
if ($config.sources.totaljobs -ne $false -or $config.sources.naukrigulf -ne $false) {
    try {
        Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking
        Start-EdgeCdp -Headless
        Enable-Stealth
        Enable-AssetBlocking

        if ($config.sources.totaljobs -ne $false) {
            foreach ($tjq in @("bim", "bim-coordinator")) {
            try {
                Navigate-Cdp "https://www.totaljobs.com/jobs/${tjq}?postedwithin=7" -TimeoutSec 45
                Start-Sleep -Seconds 4
                $jsTJ = '(function(){ var seen = {}; var out = []; var all = document.querySelectorAll("a"); for (var i=0;i<all.length;i++){ var h = all[i].href.split("?")[0]; var m = h.match(/\/job\/([a-z0-9\-]+\/[^\/]+\-job\d+)/); if (!m || seen[m[1]]) continue; seen[m[1]] = true; var card = all[i].closest("article, li, div"); var txt = card ? card.innerText.replace(/\s+/g," ").substring(0,300) : ""; out.push({ title: (all[i].innerText||"").trim().substring(0,100), url: h, ctx: txt }); } return JSON.stringify(out.slice(0,30)); })()'
                $tjRaw = Invoke-CdpJs $jsTJ
                $tjJobs = @()
                if ($tjRaw) { $tjJobs = @(ConvertFrom-JsonDeep $tjRaw) }
                foreach ($j in $tjJobs) {
                    $cm = [regex]::Match($j.ctx, "(?i)([A-Z][\w&'\-\. ]{1,50}(Ltd|Limited|Group|Recruitment|Associates|Partners|Consult|Services|UK|plc))")
                    $found.Add([PSCustomObject]@{
                        Title = $j.title; Company = $(if ($cm.Success) { $cm.Groups[1].Value.Trim() } else { "" })
                        Location = "UK"; Source = "TotalJobs"; URL = $j.url; Desc = $j.ctx; Posted = ""
                    })
                }
                Write-Log "TotalJobs '$tjq': $($tjJobs.Count) posts (last 7d)"
            } catch { Write-Log "TotalJobs '$tjq' FAILED: $($_.Exception.Message)" }
            }
        }

        if ($config.sources.naukrigulf -ne $false) {
            foreach ($ngSlug in @("bim-coordinator-jobs", "bim-engineer-jobs", "bim-modeler-jobs")) {
            try {
                Navigate-Cdp "https://www.naukrigulf.com/$ngSlug" -TimeoutSec 45
                for ($s = 0; $s -lt 4; $s++) {
                    Invoke-CdpJs '(function(){ window.scrollBy(0, 1200); return window.scrollY; })()' | Out-Null
                    Start-Sleep -Seconds 2
                }
                $jsNG = '(function(){ var seen = {}; var out = []; var cards = document.querySelectorAll("div.srp-tuple"); for (var i=0;i<cards.length;i++){ var a = cards[i].querySelector("a[href]"); if (!a) continue; var h = a.href.split("?")[0]; if (seen[h] || h.indexOf("-jid-") < 0) continue; seen[h] = true; var txt = cards[i].innerText.replace(/\s+/g," ").substring(0,400); out.push({ title: (a.innerText||"").trim().substring(0,100), url: h, ctx: txt }); } return JSON.stringify(out.slice(0,30)); })()'
                $ngRaw = Invoke-CdpJs $jsNG
                $ngJobs = @()
                if ($ngRaw) { $ngJobs = @(ConvertFrom-JsonDeep $ngRaw) }
                foreach ($j in $ngJobs) {
                    $lines = @($j.title -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    $ngTitle = if ($lines.Count -gt 0) { $lines[0] } else { $j.title }
                    $ngCompany = if ($lines.Count -gt 1) { $lines[1] } else { "" }
                    $lm = [regex]::Match($j.ctx, "(?i)(Dubai|Abu Dhabi|Sharjah|Ajman|Al Ain|United Arab Emirates|UAE)[^,]{0,20}(UAE|Emirates)?")
                    $found.Add([PSCustomObject]@{
                        Title = $ngTitle; Company = $ngCompany; Location = $(if ($lm.Success) { $lm.Value.Trim() } else { "Gulf" })
                        Source = "NaukriGulf"; URL = $j.url; Desc = $j.ctx; Posted = ""
                    })
                }
                Write-Log "NaukriGulf '$ngSlug': $($ngJobs.Count) posts"
            } catch { Write-Log "NaukriGulf '$ngSlug' FAILED: $($_.Exception.Message)" }
            }
        }

        # --- Source 11: Bayt.com (UAE/MENA board; SEO keyword pages, real browser only - HTTP gets 403) ---
        if ($config.sources.bayt -ne $false) {
            try {
                $jsBayt = '(function(){ var seen={}, out=[]; var as=document.querySelectorAll("a[href]"); for (var i=0;i<as.length;i++){ var h=(as[i].href||"").split("?")[0]; if(!/bayt\.com\/en\/uae\/jobs\/[a-z0-9\-]+-\d+\/?$/i.test(h)) continue; var t=(as[i].innerText||"").replace(/\s+/g," ").trim(); if(!t||t.length<4) continue; if(seen[h]) continue; seen[h]=1; var card=as[i].closest("li, article, div"); out.push({title:t.substring(0,110), url:h, ctx:card?(card.innerText||"").substring(0,300):""}); if(out.length>=25) break; } return JSON.stringify(out); })()'
                foreach ($bq in @("bim-coordinator-jobs", "bim-engineer-jobs", "bim-jobs")) {
                    Navigate-Cdp "https://www.bayt.com/en/uae/jobs/$bq/" -TimeoutSec 45
                    Start-Sleep -Seconds 4
                    $bRaw = Invoke-CdpJs $jsBayt
                    $bJobs = @()
                    if ($bRaw) { $bJobs = @(ConvertFrom-JsonDeep $bRaw) }
                    foreach ($j in $bJobs) {
                        $bl = @($j.ctx -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                        $found.Add([PSCustomObject]@{
                            Title = $j.title; Company = $(if ($bl.Count -gt 1) { $bl[1] } else { "" })
                            Location = "UAE"; Source = "Bayt"; URL = $j.url; Desc = ($j.ctx -replace "\s+", " ")
                            Posted = $j.ctx
                        })
                    }
                    Write-Log "Bayt '$bq': $($bJobs.Count) posts"
                    Start-Sleep -Seconds 3
                }
            } catch { Write-Log "Bayt FAILED: $($_.Exception.Message)" }
        }

        # LinkedIn runs in its OWN clean-profile browser (port 9446): the bot profile's persisted
        # state makes LinkedIn serve a degraded view (only ~7 SSR cards, never hydrates).
        # A fresh profile + fleet session cookies renders the full list (~40-60) - verified by probe.
        try { Stop-EdgeCdp } catch {}

        # --- Source 12: LinkedIn Jobs (full public sweep: every post type, all ages, UAE only) ---
        # Freshness rule stays intact via parsed age text (UAE<=1d, other<=7d) at the add-time gate.
        if ($config.sources.linkedin -ne $false) {
            try {
                $edgeExe = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                if (-not (Test-Path $edgeExe)) { $edgeExe = "C:\Program Files\Microsoft\Edge\Application\msedge.exe" }
                $liUd = Join-Path $env:TEMP "opencode\huntbot-linkedin-profile"
                if (Test-Path $liUd) { try { Remove-Item $liUd -Recurse -Force -ErrorAction Stop } catch {} }
                New-Item -ItemType Directory -Force -Path $liUd | Out-Null
                Start-Process $edgeExe -ArgumentList "--remote-debugging-port=9446", "--user-data-dir=`"$liUd`"", "--no-first-run", "--no-default-browser-check", "--start-maximized", "about:blank"
                $liUp = $false
                for ($lwi = 0; $lwi -lt 40; $lwi++) { try { $null = Invoke-RestMethod "http://127.0.0.1:9446/json/version" -TimeoutSec 1; $liUp = $true; break } catch { Start-Sleep -Milliseconds 500 } }
                if (-not $liUp) { throw "LinkedIn browser (port 9446) did not start" }
                Connect-EdgeCdp -Port 9446
                $lCookieFile = Join-Path $ProjectDir "edge-profile\cookies.json"
                if (Test-Path $lCookieFile) {
                    $lsc = Get-Content $lCookieFile -Raw | ConvertFrom-Json
                    if ($lsc) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($lsc) } -TimeoutSec 30 | Out-Null }
                }
                $jsLi = '(function(){ var seen={}, out=[]; var as=document.querySelectorAll("a[href*=\"/jobs/view/\"]"); for (var i=0;i<as.length;i++){ var h=(as[i].href||"").split("?")[0]; var t=(as[i].innerText||"").replace(/\s+/g," ").trim(); if(!t||t.length<4) continue; if(seen[h]) continue; seen[h]=1; var card=as[i].closest("li, article, div"); out.push({title:t.substring(0,110), url:h, ctx:card?(card.innerText||"").substring(0,300):""}); if(out.length>=40) break; } return JSON.stringify(out); })()'
                $liQueries = @(@($config.search.queries) + @("BIM", "Revit", "VDC") | Sort-Object -Unique)
                $liLocs = @($config.search.linkedinLocations)
                if ($liLocs.Count -eq 0) { $liLocs = @("United Arab Emirates") }
                # full sweep (2 pages, newest first) on the primary location; top queries, 1 page on the rest
                $liPlan = @()
                foreach ($lq in $liQueries) { $liPlan += , @($lq, $liLocs[0], 0); $liPlan += , @($lq, $liLocs[0], 25) }
                foreach ($lq in @($liQueries | Select-Object -First 2)) {
                    foreach ($ll2 in @($liLocs | Select-Object -Skip 1)) { $liPlan += , @($lq, $ll2, 0) }
                }
                foreach ($p3 in $liPlan) {
                    $lq = $p3[0]; $lLoc = $p3[1]; $lStart = $p3[2]
                    $lurl = "https://www.linkedin.com/jobs/search?keywords=" + [uri]::EscapeDataString($lq) + "&location=" + [uri]::EscapeDataString($lLoc) + "&sortBy=DD&start=$lStart"
                    Navigate-Cdp $lurl -TimeoutSec 45
                    Start-Sleep -Seconds 3
                    # LinkedIn lazy-hydrates results (headless keeps only ~7 SSR placeholder cards)
                    # - scroll + poll until the real list has loaded
                    $prevN = -1
                    for ($lw = 0; $lw -lt 5; $lw++) {
                        Invoke-CdpJs '(function(){ window.scrollBy(0, 900); return document.querySelectorAll("a[href*=\"/jobs/view/\"]").length; })()' | Out-Null
                        Start-Sleep -Seconds 2
                        $nNow = Invoke-CdpJs '(function(){ return document.querySelectorAll("a[href*=\"/jobs/view/\"]").length; })()'
                        if ("$nNow" -match "^\d+$") {
                            if ([int]$nNow -ge 25 -or ([int]$nNow -eq $prevN -and $lw -ge 2)) { break }
                            $prevN = [int]$nNow
                        }
                    }
                    $lRaw = Invoke-CdpJs $jsLi
                    $lJobs = @()
                    if ($lRaw) { $lJobs = @(ConvertFrom-JsonDeep $lRaw) }
                    $nEasy = 0
                    foreach ($j in $lJobs) {
                        $ll = @($j.ctx -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                        $liComp = ""; $liLoc = $lLoc
                        foreach ($ln in $ll) {
                            if (-not $liComp -and $ln -ne $j.title -and $ln -notmatch "(?i)applicant|actively hiring|easy apply|be the first|promoted|follow") { $liComp = $ln }
                            if ($ln -match "(?i)(Dubai|Sharjah|Abu Dhabi|Ajman|Al Ain|United Arab Emirates|UAE)") { $liLoc = $ln }
                        }
                        $liAge = ""
                        if ($j.ctx -match "(?i)\d+\s*(?:minutes?|hours?|days?|weeks?)\s+ago") { $liAge = $Matches[0] }
                        $lSrc = "LinkedIn"
                        if ($j.ctx -match "(?i)easy apply") { $lSrc = "LinkedIn-EasyApply"; $nEasy++ }
                        $found.Add([PSCustomObject]@{
                            Title = $j.title; Company = $liComp; Location = $liLoc
                            Source = $lSrc; URL = $j.url; Desc = ($j.ctx -replace "\s+", " ")
                            Posted = $liAge
                        })
                    }
                    Write-Log "LinkedIn '$lq' @ $lLoc (start=$lStart): $($lJobs.Count) posts ($nEasy Easy Apply)"
                    Start-Sleep -Seconds 4
                }
            } catch { Write-Log "LinkedIn FAILED: $($_.Exception.Message)" }
            finally {
                Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -match "huntbot-linkedin-profile" } | ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
            }
        }

        try { Stop-EdgeCdp } catch {}
    } catch { Write-Log "JobBoards browser FAILED: $($_.Exception.Message)"; try { Stop-EdgeCdp } catch {} }
}

Write-Log "Total raw listings: $($found.Count)"

# --- Score + filter ---
$exSubs = @($config.search.excludeTitles | ForEach-Object { [regex]::Escape($_) })
$exWords = @($config.search.excludeWords | ForEach-Object { "\b" + [regex]::Escape($_) + "\b" })
$excludeRx = "(?i)" + (($exSubs + $exWords) -join "|")
$blockedLocRx = "(?i)saudi|riyadh|jeddah|dammam|neom|egypt|cairo|alexandria"
# UAE-ONLY policy (user 2026-11-04): a listing must show a UAE signal in Location/URL/Title
$uaeLocRx = "(?i)\bu\.?a\.?e\.?\b|united arab|dubai|sharjah|abu dhabi|ajman|al ain|ras al khaimah|fujairah|ae\.indeed\.com|bayt\.com/en/uae/"
Write-Log "UAE-only gate: $(@($found | Where-Object { "$($_.Location) $($_.URL) $($_.Title)" -match $uaeLocRx }).Count) of $($found.Count) raw listings are UAE-related"
$scored = @($found | Where-Object { $_.Title -notmatch $excludeRx -and $_.Location -notmatch $blockedLocRx -and $_.URL -notmatch $blockedLocRx -and ("$($_.Location) $($_.URL) $($_.Title)" -match $uaeLocRx) } | ForEach-Object {
    $_ | Add-Member -NotePropertyName Score -NotePropertyValue (Get-Score $_.Title $_.Desc) -Force -PassThru
} | Where-Object { $_.Score -ge $config.search.minScore })

Write-Log "Excluded $(($found.Count) - (@($found | Where-Object { $_.Title -notmatch $excludeRx }).Count)) MEP/Mechanical-titled listings"

Write-Log "After scoring (>= $($config.search.minScore)): $($scored.Count)"

# --- FRESHNESS at add time (user rule 2026-10-26): UAE <= 24h, other countries <= 7 days ---
# Only skips listings whose AGE TEXT already exceeds the limit (unknown age -> keep, ApplyBot gate re-checks).
function Get-PostedAgeDays([string]$p) {
    if (-not $p) { return -1 }
    # LinkedIn-style relative times
    if ($p -match "(?i)\b(\d+)\s*(?:minutes?|hours?|mins?|hrs?)\b") { return 0 }
    $m = [regex]::Match($p, "(?i)(\d+)\s*\+?\s*weeks?\s+ago")
    if ($m.Success) { return [int]$m.Groups[1].Value * 7 }
    foreach ($rx in @("(?i)(\d+)\s*\+?\s*days?\s+ago", "(?i)posted\s+(\d+)\s*\+?\s*days?", "(?i)vor\s+(\d+)\s*Tagen", "(?i)vor\s+(\d+)\s*Tag\b", "(?i)il y a\s+(\d+)\s*jours?", "(?i)(\d+)\s*dagen\s+geleden")) {
        $m = [regex]::Match($p, $rx)
        if ($m.Success) { return [int]$m.Groups[1].Value }
    }
    $m = [regex]::Match($p, "(?i)vor\s+(\d+)\s*Wochen")
    if ($m.Success) { return [int]$m.Groups[1].Value * 7 }
    if ($p -match "(?i)\byesterday\b|\bgestern\b") { return 1 }
    if ($p -match "(?i)(\d+)\s*month|(\d+)\s*monat") { return [int]$Matches[1] * 30 }
    return -1
}
$ageCut = $scored.Count
$scored = @($scored | Where-Object {
    $pd = Get-PostedAgeDays $_.Posted
    if ($pd -lt 0) { return $true }
    $isU = ("$($_.Location) $($_.URL) $($_.Title)") -match "(?i)\bu\.?a\.?e\.?\b|united arab|dubai|sharjah|abu dhabi|ajman|ae\.indeed\.com"
    $lim = [int]$config.apply.maxAgeDaysOther
    if ($isU) { $lim = [int]$config.apply.maxAgeDaysUae }
    return ($pd -le $lim)
})
if ($ageCut -ne $scored.Count) { Write-Log "Freshness filter: dropped $(($ageCut) - ($scored.Count)) too-old listing(s) (UAE>1d / other>7d)" }

# --- Queue file: if the tracker is locked (Excel open), matches wait here ---
$queueFile = Join-Path $logDir "pending-jobs.json"

# --- Dedupe against tracker ---
$existing = $null
try { $existing = Get-TrackerJobs } catch { Write-Log "Tracker unavailable (Excel open?) - queueing matches for next run" }
if ($null -eq $existing) {
    $queued = @()
    if (Test-Path $queueFile) { $queued = @(Get-Content $queueFile -Raw | ConvertFrom-JsonDeep) }
    $queued += $scored
    $queued | ConvertTo-Json -Depth 5 | Set-Content $queueFile -Encoding UTF8
    Write-Log "Queued $($scored.Count) match(es) -> $queueFile (total queued: $($queued.Count))"
    Write-Host "Excel is open - matches queued for next run." -ForegroundColor Yellow
    exit 0
}

# Drain the queue from previous locked runs
$pending = @()
if (Test-Path $queueFile) {
    $pending = @(Get-Content $queueFile -Raw | ConvertFrom-Json)
    if ($pending.Count -gt 0) { Write-Log "Draining $($pending.Count) queued match(es) from previous run" }
}
$allCandidates = @($pending) + @($scored)
$existingUrls = @{}; $existingKeys = @{}
foreach ($e in $existing) {
    if ($e.URL) { $existingUrls[$e.URL.ToLower().TrimEnd('/')] = $true }
    $existingKeys[("$($e.Title)|$($e.Company)").ToLower()] = $true
}

$newOnes = @($allCandidates | Where-Object {
    $u = $_.URL.ToLower().TrimEnd('/')
    $k = ("$($_.Title)|$($_.Company)").ToLower()
    if ($existingUrls.ContainsKey($u) -or $existingKeys.ContainsKey($k)) { return $false }
    # in-run dedupe: first occurrence wins, later duplicates (same job via other queries) are dropped here
    $existingUrls[$u] = $true
    $existingKeys[$k] = $true
    return $true
} | Sort-Object Score -Descending | Select-Object -First 50)

Write-Log "New unique matches: $($newOnes.Count)"

# --- Write to tracker ---
$added = 0
$addFailed = @()
foreach ($j in $newOnes) {
    try {
        Add-TrackerJob -Job ([PSCustomObject]@{
            Title = $j.Title; Company = $j.Company; Location = $j.Location
            Source = $j.Source; URL = $j.URL; Score = $j.Score
            Salary = $config.search.salaryExpectation; TailoredFiles = ""
        })
        $added++
        # in-run dedupe: mark this URL/title as seen so the same job from
        # another query in THIS run cannot be added a second time
        if ($j.URL) { $existingUrls[$j.URL.ToLower().TrimEnd('/')] = $true }
        $existingKeys[("$($j.Title)|$($j.Company)").ToLower()] = $true
    } catch {
        if ($j.Title -is [System.Array] -or $j.URL -is [System.Array]) {
            Write-Log "DROPPED squashed job (array fields) - will be re-found cleanly next run"
        } else {
            Write-Log "Failed to add '$($j.Title)' (Excel locked?) - queued for next run"
            $addFailed += $j
        }
    }
}
if ($addFailed.Count -gt 0) {
    $queued = @()
    if (Test-Path $queueFile) { $queued = @(Get-Content $queueFile -Raw | ConvertFrom-JsonDeep) }
    $queued += $addFailed
    $queued | ConvertTo-Json -Depth 5 | Set-Content $queueFile -Encoding UTF8
    Write-Log "Queued $($addFailed.Count) failed add(s) for next run"
}
Write-Log "Added to tracker: $added"
if ($addFailed.Count -eq 0 -and $added -gt 0 -and (Test-Path $queueFile)) { Remove-Item $queueFile -Force }

# --- Report ---
$rep = @("# HuntBot report - $(Get-Date -Format 'yyyy-MM-dd HH:mm')", "", "Queries: $($config.search.queries -join ', ')", "New matches added: **$added**", "", "## Top matches", "")
foreach ($j in ($newOnes | Select-Object -First 15)) {
    $rep += "- **$($j.Title)** - $($j.Company) ($($j.Location)) [score $($j.Score)] $($j.URL)"
}
$rep | Set-Content $report -Encoding UTF8
Write-Log "Report: $report"
if (-not $env:JTRACKER_PIPELINE_RUN) { Remove-Item $runLock -Force -ErrorAction SilentlyContinue }
Write-Host "`nHuntBot done. $added new job(s) added to tracker." -ForegroundColor Green
