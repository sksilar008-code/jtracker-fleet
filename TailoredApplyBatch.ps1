<#
.SYNOPSIS
  TAILOREDAPPLYBATCH : for every pending job/lead -
    1) reads the job description, 2) builds a JD-matched tailored CV (HTML+PDF),
    3) writes a tailored cover email, 4) sends it with the CV attached,
    5) updates the tracker. Recipients required: no email = honest skip.
#>
param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "config.json"),
    [int]$MaxSend = 30
)

$ErrorActionPreference = "Continue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "Tracker.psm1") -Force
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking
Add-Type -AssemblyName System.IO.Compression.FileSystem

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Set-TrackerPath $config.paths.tracker
$P = $config.profile
$tailoredRoot = Join-Path $config.paths.root "Tailored"

$logDir = Join-Path $ProjectDir "logs"
$logFile = Join-Path $logDir ("tailored-batch-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".log")
function Write-Log($msg) { $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"; Write-Host $line; Add-Content $logFile $line }

$historyFile = Join-Path $ProjectDir "edge-profile\feedbot-history.json"
$history = @()
if (Test-Path $historyFile) { $raw = Get-Content $historyFile -Raw | ConvertFrom-Json; if ($raw.emails) { $history = @($raw.emails) } }

# ---------- skill map: description pattern -> display label + achievement bullet ----------
$skillMap = @(
    @{ pat = "acc\b|autodesk construction cloud|bim 360|common data environment|\bcde\b"; label = "Autodesk Construction Cloud (ACC/BIM 360) & CDE administration"; bullet = "Administered ACC / BIM 360 CDE environments: permissions, folder governance, design collaboration and information workflows to ISO 19650." },
    @{ pat = "navisworks|clash"; label = "Navisworks clash detection & coordination"; bullet = "Ran Navisworks clash detection and resolution workshops across architectural, structural, civil and MEP models to submission-ready deliverables." },
    @{ pat = "\brevit\b"; label = "Autodesk Revit modelling & coordination"; bullet = "Produced and coordinated Revit models to LOD 350+, enforcing family standards and model health across disciplines." },
    @{ pat = "autocad"; label = "AutoCAD drafting & detailing"; bullet = "Delivered precise AutoCAD shop drawings, GA drawings and as-builts coordinated back to the federated model." },
    @{ pat = "civil 3d|12d"; label = "Civil 3D / civil modelling"; bullet = "Coordinated civil models, alignments and shared coordinates with building disciplines on infrastructure packages." },
    @{ pat = "iso 19650|\bbep\b|\beir\b|\blod\b|midp|tidp"; label = "ISO 19650 (BEP / EIR / LOD)"; bullet = "Developed and enforced BEPs (EIR, LOD, MIDP/TIDP) and ran QA/QC audits to client and authority standards." },
    @{ pat = "federat"; label = "Multidisciplinary model federation"; bullet = "Federated architectural, structural, civil and MEP models across 25+ projects, checking quality and submission readiness." },
    @{ pat = "dynamo|power bi|python|automat"; label = "Dynamo / Power BI / Python automation"; bullet = "Built Dynamo and Power BI automations for clash reporting, model health checks and progress dashboards." },
    @{ pat = "station|rail|metro|depot|tunnel"; label = "Station & rail interface experience"; bullet = "Delivered station, depot and tunnel BIM packages on Singapore's Jurong Region Line (J101 Tengah Depot) and CR116 Ang Mo Kio Interchange." },
    @{ pat = "4d|quantit|take-?off|boq"; label = "4D planning & quantity support"; bullet = "Supported 4D planning inputs, quantity extraction and model-based visualisation for design and construction reviews." },
    @{ pat = "asset|handover|as-?built|fm"; label = "Asset information & handover"; bullet = "Prepared asset information and as-built documentation for handover, aligned to client information requirements." },
    @{ pat = "shop drawing|qa/?qc|quality"; label = "Shop drawings & QA/QC"; bullet = "QA/QC'd design packages, shop drawings and subcontractor deliverables against project, client and authority standards." },
    @{ pat = "workshop|stakeholder|coordination meeting|interface"; label = "Coordination workshops & stakeholder management"; bullet = "Led coordination and resolution workshops, tracking agreed actions and interface risks to closure with designers and site teams." },
    @{ pat = "shared coordinat|survey|scan-?to-?bim|point cloud"; label = "Shared coordinates & reality capture"; bullet = "Managed shared coordinates and delivered scan-to-BIM packages (Frankfurt station) for the Netherlands team." },
    @{ pat = "high-?rise|tower|residential"; label = "High-rise residential delivery"; bullet = "Coordinated UAE high-rise residential towers (Tiger Sky Tower, AZHA, Jade Tower) from design through construction." },
    @{ pat = "hospital|healthcare"; label = "Healthcare projects"; bullet = "Supported hospital BIM delivery (Omraniyoun, Al-Moosa, King Abdullah Medical City) with strict QA/QC regimes." },
    @{ pat = "data ?centre|mission critical"; label = "Data centres"; bullet = "Delivered BIM packages for Microsoft Data Center Singapore and mission-critical facilities." },
    @{ pat = "precast|prefab|modular"; label = "Precast & prefab"; bullet = "Coordinated precast/prefab modelling and shop-drawing QA for modular construction (MBS Hybrid Casco, Netherlands)." }
)

function Get-Matched($text) {
    $m = @()
    foreach ($s in $skillMap) { if ($text -match $s.pat) { $m += $s } }
    return $m
}

function New-SafeName($s) { return (($s -replace "[^\w\- ]", "") -replace "\s+", "_").Trim("_") }

function New-TailoredCV($folder, $role, $company, $location, $matched, $descSnippet) {
    $matchLabels = @($matched | ForEach-Object { $_.label } | Select-Object -Unique | Select-Object -First 8)
    $bullets = @($matched | ForEach-Object { $_.bullet } | Select-Object -Unique | Select-Object -First 5)
    $generic = @(
        "Federated architectural, structural and MEP models across 25+ projects; ran clash detection and resolution workshops to submission-ready, clash-free deliverables.",
        "Administered ACC / BIM 360 CDE environments to ISO 19650, tracking coordination actions to closure with engineers and construction teams.",
        "8+ years across six countries (UAE, Singapore, India, Saudi Arabia, Netherlands, USA) on high-rise, infrastructure and transport programs."
    )
    $i = 0
    while ($bullets.Count -lt 4 -and $i -lt $generic.Count) { if ($bullets -notcontains $generic[$i]) { $bullets += $generic[$i] }; $i++ }
    $sector = "Transport & Infrastructure"
    if ($descSnippet -match "high-?rise|tower|residential") { $sector = "High-Rise & Infrastructure" }
    elseif ($descSnippet -match "hospital|healthcare") { $sector = "Healthcare & Infrastructure" }
    elseif ($descSnippet -match "data ?centre") { $sector = "Data Centres & Infrastructure" }
    $tags = @($matchLabels | ForEach-Object { "<span class=`"tag`">$_</span>" }) -join ""
    $blist = ($bullets | ForEach-Object { "<li>$_</li>" }) -join ""
    $matchLine = ($matchLabels -join " ● ")
    $html = @"
<!doctype html><html><head><meta charset="utf-8"><style>
@page{size:A4;margin:0}
body{font-family:'Segoe UI',Georgia,serif;color:#1b2733;margin:0;padding:0;font-size:11.5px;line-height:1.42}
.header{background:linear-gradient(90deg,#0f2a43 0%,#134a7c 70%,#1b6ca8 100%);color:#fff;padding:26px 34px 20px}
.header h1{margin:0;font-size:30px;letter-spacing:2.5px;color:#ffd166}
.header .subtitle{margin:4px 0 8px;font-size:12.5px;letter-spacing:1.2px;color:#cfe3f5;text-transform:uppercase}
.header .contact{font-size:11px;color:#dbe7f3}
.body{padding:18px 34px 28px}
.cols{display:flex;gap:26px}
.main{flex:2.1}
.side{flex:1}
h3{color:#0f2a43;font-size:12.5px;text-transform:uppercase;letter-spacing:2px;border-bottom:2.5px solid #1b6ca8;padding-bottom:3px;margin:16px 0 8px}
ul{margin:5px 0;padding-left:18px} li{margin-bottom:4px}
.exp b{color:#134a7c;font-size:12.5px}
.period{float:right;color:#7f8c9b;font-size:10.5px;font-weight:normal}
.tag{display:inline-block;background:#eaf3fb;color:#134a7c;border:1px solid #bfd9ec;border-radius:4px;padding:2px 8px;margin:2px 4px 2px 0;font-size:10.5px}
.ach li::marker{color:#1b6ca8}
p{margin:5px 0}
.footerline{height:5px;background:linear-gradient(90deg,#ffd166,#1b6ca8,#0f2a43);margin-top:18px}
</style></head><body>
<div class="header">
<h1>SILAR SHAIK</h1>
<div class="subtitle">$role &nbsp;|&nbsp; $sector &nbsp;|&nbsp; ISO 19650</div>
<div class="contact">Phone: $($P.phone) &nbsp;&nbsp;|&nbsp;&nbsp; Email: $($P.email) &nbsp;&nbsp;|&nbsp;&nbsp; Sharjah, UAE (open to $location) &nbsp;&nbsp;|&nbsp;&nbsp; English (Fluent)</div>
</div>
<div class="body">
<h3>Profile</h3>
<p>BIM Coordinator with <b>8+ years</b> delivering complex building and infrastructure programs across six countries, including Singapore MRT station, depot and tunnel packages with WSP. Specialist in multidisciplinary model federation, clash coordination and submission-ready information management under <b>ISO 19650</b> (ACC, Navisworks, Revit, AutoCAD, Civil 3D), with Dynamo/Power BI automation. GMICE graduate in Civil Engineering (JNTU Kakinada).</p>
<h3>Job Match — $role @ $company</h3>
<p><b>Directly matching your requirements:</b> $matchLine.</p>
<h3>Why This Role — Matched Achievements</h3>
<ul>$blist</ul>
<div class="cols">
<div class="main">
<h3>Professional Experience</h3>
<p class="exp"><b>BIM Coordinator — Tiger Group, UAE</b> <span class="period">2025 – Present</span></p>
<ul><li>Coordinate and federate architectural, structural and MEP models; run clash detection cycles and resolution workshops to submission-ready deliverables.</li><li>Administer ACC CDE to ISO 19650; QA/QC design packages; track coordination actions to closure with engineers and leadership.</li></ul>
<p class="exp"><b>BIM Technician — Transportation &amp; Infrastructure — WSP</b> <span class="period">2023 – 2025</span></p>
<ul><li>Delivered BIM packages for <b>Singapore MRT J101 (Tengah Depot) and CR116 (Ang Mo Kio Interchange Station &amp; Tunnels)</b> and Mumbai Coastal Road interchanges.</li><li>Ran interdisciplinary clash reviews; managed BIM 360/ACC digital delivery aligned to BIM codes.</li></ul>
<p class="exp"><b>Team Lead — BIM — The BIM Engineers</b> <span class="period">2021 – 2023</span></p>
<ul><li>Led model production and audits; enforced BEPs; scan-to-BIM for Frankfurt station; mentored junior modellers.</li></ul>
<p class="exp"><b>Team Lead — BIM (Architecture) — Innovspatial</b> <span class="period">2018 – 2021</span></p>
<ul><li>Led CAD/BIM production via shared CDE; 30+ authority permit sets; QA/QC audits to client requirements.</li></ul>
<h3>Key Projects</h3>
<ul><li>J101 Tengah Depot &amp; MRT Facilities, Jurong Region Line, Singapore</li><li>CR116 Ang Mo Kio Interchange Station &amp; Tunnels, Singapore</li><li>Tiger Sky Tower — AZHA — Jade Tower (High-Rise UAE)</li><li>Frankfurt Station Scan-to-BIM — Microsoft Data Center Singapore</li></ul>
</div>
<div class="side">
<h3>Expertise</h3>
<p>Model federation &amp; QA<br>Clash coordination &amp; workshops<br>ISO 19650 / BEP / EIR / LOD<br>CDE administration (ACC)<br>Submission readiness<br>Action tracking to closure<br>Automation &amp; reporting</p>
<h3>Core Skills</h3>
<p>$tags</p>
<h3>Achievements</h3>
<ul class="ach"><li>25+ projects, 6 countries</li><li>Singapore MRT station, depot &amp; tunnel packages</li><li>ACC/BIM 360 across 25+ projects</li><li>ISO 19650-compliant delivery</li></ul>
<h3>Education</h3>
<p>B.Tech Civil Engineering<br>JNTU Kakinada, India<br>2014–2018</p>
<h3>Professional Body</h3>
<p>Graduate Member (GMICE)<br>Institution of Civil Engineers, UK</p>
</div>
</div>
<div class="footerline"></div>
</div>
</body></html>
"@
    $html | Set-Content (Join-Path $folder "cv.html") -Encoding UTF8
    # PDF via Edge print engine (no CDP - stable)
    $pdf = Join-Path $folder "SILAR - BIM Coordinator.pdf"
    $edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    if (-not (Test-Path $edge)) { $edge = "C:\Program Files\Microsoft\Edge\Application\msedge.exe" }
    & $edge --headless --disable-gpu --no-pdf-header-footer --print-to-pdf="$pdf" ("file:///" + ((Join-Path $folder "cv.html") -replace "\\","/")) 2>$null
    Start-Sleep -Seconds 5
    if (Test-Path $pdf) { return $pdf } else { return $null }
}

function Get-TailoredBody($to, $role, $company, $matched) {
    $labels = @($matched | ForEach-Object { $_.label } | Select-Object -Unique | Select-Object -First 4)
    $reqLine = "multidisciplinary BIM coordination, clash detection and ISO 19650 information management"
    if ($labels.Count -gt 0) { $reqLine = ($labels -join ", ") }
    $greet = "Dear Hiring Team,"
    if ($to -match "^([a-z]+)[\.\-_]([a-z]+)@") { $greet = "Dear " + (Get-Culture).TextInfo.ToTitleCase($Matches[1]) + "," }
    return @"
$greet

I am applying for the $role position$(if ($company -and $company -notmatch "via LinkedIn") { " at $company" }), and my background maps directly onto the role's requirements in $reqLine.

Over 8+ years across six countries I have delivered BIM coordination on 25+ projects - from Singapore MRT station, depot and tunnel packages (WSP) to landmark high-rise towers in the UAE (Tiger Group, my current role). I federate architectural, structural and MEP models, lead clash-resolution workshops, and keep project information submission-ready under ISO 19650 - working daily in Revit, Navisworks and ACC/BIM 360, with Dynamo and Power BI automation.

I am based in Sharjah, UAE, and open to relocation. My tailored CV is attached - I would welcome a short call to discuss the role.

Kind regards,
$($P.fullName)
$($P.email) | $($P.phone) | Sharjah, UAE | GMICE
"@
}

# ---------- load pending (XML snapshot - no Excel lock; verified fresh) ----------
$tmpSnap = Join-Path $env:TEMP "opencode\tailored-snap.xlsx"
$snapOk = $false
for ($try = 1; $try -le 3 -and -not $snapOk; $try++) {
    try {
        if (Test-Path $tmpSnap) { Remove-Item $tmpSnap -Force -ErrorAction Stop }
        Copy-Item $config.paths.tracker $tmpSnap -Force -ErrorAction Stop
        $fi = Get-Item $tmpSnap -ErrorAction Stop
        if ($fi.Length -gt 10000) { $snapOk = $true; Write-Log ("Snapshot verified fresh (" + $fi.Length + " bytes).") }
    } catch { Write-Log ("Snapshot copy failed (try $try): " + $_.Exception.Message); Start-Sleep -Seconds 3 }
}
if (-not $snapOk) { Write-Log "FATAL: snapshot copy failed - aborting"; exit 1 }
$zip = [System.IO.Compression.ZipFile]::Open($tmpSnap, 'Read')
$sr = New-Object IO.StreamReader($zip.GetEntry("xl/sharedStrings.xml").Open())
[xml]$ssx = $sr.ReadToEnd(); $sr.Close()
$strings = @($ssx.sst.si | ForEach-Object { if ($_.t) { $_.t } else { ($_.r | ForEach-Object { $_.t }) -join "" } })
function Get-Cell($c) { if ($c.t -eq 's') { return $strings[[int]$c.v] } else { return $c.v } }

$queue = @()
$s1 = New-Object IO.StreamReader($zip.GetEntry("xl/worksheets/sheet1.xml").Open())
[xml]$s1x = $s1.ReadToEnd(); $s1.Close()
foreach ($row in $s1x.worksheet.sheetData.row) {
    if ($row.r -eq 1) { continue }
    $v = @{}; foreach ($c in $row.c) { $v[($c.r -replace "\d","")] = Get-Cell $c }
    if ($v["H"] -eq "New" -and ($v["B"] -or $v["F"])) {
        $queue += [PSCustomObject]@{ Kind = "Job"; Row = [int]$row.r; Title = $v["B"]; Company = $v["C"]; Location = $v["D"]; URL = $v["F"]; Folder = $v["L"]; Desc = "" }
    }
}
$s3 = New-Object IO.StreamReader($zip.GetEntry("xl/worksheets/sheet3.xml").Open())
[xml]$s3x = $s3.ReadToEnd(); $s3.Close()
$ua = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }
$badDomains = @("adzuna", "indeed.com", "example.com", "sentry", "wixpress", "godaddy", "remoteok", "remotive", "arbeitnow", "googleapis", "schema", "w3.org", "cloudflare", "drjobs", "founditgulf", "naukrigulf", "bayt", "linkedin.com")
$blockedRoles = "(?i)mechanical|electrical|plumbing|plomberie|hvac|elektro|lectr|cvc|\bmep\b|\btga\b|\bcfo\b|\bcfa\b|\bmech\b"

foreach ($row in $s3x.worksheet.sheetData.row) {
    if ($row.r -eq 1) { continue }
    $v = @{}; foreach ($c in $row.c) { $v[($c.r -replace "\d","")] = Get-Cell $c }
    if ($v["B"] -and $v["E"] -ne "Sent" -and $v["E"] -notmatch "Skipped" -and $history -notcontains $v["B"]) {
        if (($v["C"] + " " + $v["D"]) -match $blockedRoles) { continue }  # MEP/Mechanical - never email
        $queue += [PSCustomObject]@{ Kind = "Lead"; Row = [int]$row.r; Title = $v["C"]; Company = ""; Location = ""; URL = ""; Folder = ""; Desc = $v["D"]; Email = $v["B"] }
    }
}
$zip.Dispose()
Write-Log ("Snapshot rows - Jobs: " + (@($s1x.worksheet.sheetData.row).Count) + ", FeedBot-Link: " + (@($s3x.worksheet.sheetData.row).Count))
$queue = @($queue | Sort-Object @{ Expression = { if ($_.Kind -eq "Lead") { 0 } else { 1 } } }, Row)
# UAE-ONLY policy (user 2026-11-04): never mail about non-UAE jobs/leads
$uaeQ = "(?i)\bu\.?a\.?e\.?\b|united arab|dubai|sharjah|abu dhabi|ajman|al ain|ras al khaimah|fujairah|ae\.indeed\.com|bayt\.com/en/uae/"
$qPre = $queue.Count
$queue = @($queue | Where-Object { "$($_.Location) $($_.URL) $($_.Title) $($_.Desc)" -match $uaeQ })
if ($queue.Count -lt $qPre) { Write-Log ("UAE-only: dropped " + ($qPre - $queue.Count) + " non-UAE item(s) from queue") }
Write-Log ("Pending queue: " + $queue.Count + " item(s), leads first")

$ua = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }

# ---------- phase 1: descriptions + recipients + tailored CVs (no browser) ----------
function Select-RecipientEmail([string]$txt) {
    if (-not $txt) { return "" }
    $found = [regex]::Matches($txt, "(?i)[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}") | ForEach-Object { $_.Value.ToLower() } | Sort-Object -Unique
    foreach ($e in $found) {
        $dom = ($e -split "@")[1]
        $skip = $false
        foreach ($b in $badDomains) { if ($dom -match $b) { $skip = $true; break } }
        if ($e -match "noreply|no-reply|donotreply|unsubscribe|example\.|@2x|@3x|\.png|\.jpg|\.gif|\.webp|^person@company\.|^your[._-]?|^tu\.|^name@|^email@|^user@|^test@|^sample@|@dominio\.|@company\.com|@email\.com|@domain\.com|@mail\.com$|@example\.|placeholder") { $skip = $true }
        if (-not $skip -and $history -notcontains $e) { return $e }
    }
    return ""
}
$ready = @()
$fetchFails = 0
foreach ($item in $queue) {
    Write-Log "----- [$($item.Kind) row $($item.Row)] $($item.Title)"
    $desc = ""; $email = ""
    if ($item.Kind -eq "Lead") {
        $desc = $item.Desc
        $email = $item.Email
        $role = $(if ($item.Title -and $item.Title -ne "BIM Role") { $item.Title } else { "BIM Coordinator" })
        $company = ""
    } else {
        $role = $(if ($item.Title) { $item.Title } else { "BIM Coordinator" })
        $company = $item.Company
        try {
            $pg = Invoke-WebRequest $item.URL -Headers $ua -TimeoutSec 40 -UseBasicParsing
            $txt = (($pg.Content -replace "<script[^>]*>.*?</script>", " " -replace "<style[^>]*>.*?</style>", " " -replace "<[^>]+>", " ") -replace "\s+", " ").Trim()
            if ($txt.Length -gt 6000) { $txt = $txt.Substring(0, 6000) }
            $desc = $txt
            $email = Select-RecipientEmail $txt
            $fetchFails = 0
        } catch {
            Write-Log "  page fetch failed: $($_.Exception.Message)"
            # Browser fallback: boards that block plain HTTP (LinkedIn etc.) - fetch via CDP session instead
            if ($item.URL -match "(?i)linkedin\.com|naukrigulf\.com|xing\.com|bayt\.com|gulftalent\.com|foundit\.ae") {
                try {
                    if (-not $script:TFBrowserUp) {
                        Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking
                        Start-EdgeCdp -Headless
                        Enable-Stealth
                        Enable-AssetBlocking
                        $cbFile = Join-Path $ProjectDir "edge-profile\cookies.json"
                        if (Test-Path $cbFile) {
                            $scb = Get-Content $cbFile -Raw | ConvertFrom-Json
                            if ($scb) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($scb) } -TimeoutSec 30 | Out-Null }
                        }
                        $script:TFBrowserUp = $true
                        Write-Log "  (browser fallback started for HTTP-blocked boards)"
                    }
                    Navigate-Cdp $item.URL -TimeoutSec 45
                    Start-Sleep -Seconds 4
                    $btxt = Invoke-CdpJs "(function(){ return (document.body ? document.body.innerText : '').replace(/\s+/g,' '); })()"
                    if ($btxt) {
                        if ($btxt.Length -gt 6000) { $btxt = $btxt.Substring(0, 6000) }
                        $desc = $btxt
                        $email = Select-RecipientEmail $btxt
                        $fetchFails = 0
                    }
                } catch {
                    Write-Log "  browser fetch failed: $($_.Exception.Message)"
                    $fetchFails++
                }
            } else {
                $fetchFails++
            }
            if ($fetchFails -ge 5) { Write-Log "Network throttling suspected (5 straight fetch failures) - stopping fetches, sending what is ready."; break }
        }
        Start-Sleep -Seconds 2
        if (-not $desc) { $desc = "$role $company $($item.Location)" }
    }
    if (-not $email) { Write-Log "  SKIP: no recipient email found"; continue }
    $matched = Get-Matched ($role + " " + $desc)
    $folder = Join-Path $tailoredRoot ((New-SafeName $company) + "_" + (New-SafeName $role))
    if (-not $company) { $folder = Join-Path $tailoredRoot ("LinkedIn_" + (New-SafeName $email.Split("@")[0]) + "_" + (New-SafeName $role)) }
    New-Item -ItemType Directory -Force $folder | Out-Null
    $pdf = New-TailoredCV $folder $role $company $(if ($item.Location) { $item.Location } else { "UAE" }) $matched $desc
    if (-not $pdf) { Write-Log "  SKIP: PDF generation failed"; continue }
    $body = Get-TailoredBody $email $role $company $matched
    ($body | Set-Content (Join-Path $folder "cover-letter.md") -Encoding UTF8)
    Write-Log ("  CV ready: $pdf | recipient: $email | matched: " + (($matched | ForEach-Object { $_.label } | Select-Object -First 4) -join " / "))
    $ready += [PSCustomObject]@{ Item = $item; Email = $email; Role = $role; Company = $company; Pdf = $pdf; Body = $body; Folder = $folder }
}

if ($ready.Count -eq 0) { Write-Log "Nothing sendable (no recipients). Done."; exit 0 }
$ready = @($ready | Select-Object -First $MaxSend)
Write-Log ("Sending " + $ready.Count + " tailored application(s)...")

# ---------- phase 2: send via Gmail ----------
Start-EdgeCdp -Headless
Enable-AssetBlocking
$cookieFile = Join-Path $ProjectDir "edge-profile\cookies.json"
if (Test-Path $cookieFile) {
    $sc = Get-Content $cookieFile -Raw | ConvertFrom-Json
    if ($sc) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($sc) } -TimeoutSec 30 | Out-Null }
}
Navigate-Cdp "https://mail.google.com/mail/u/0/#inbox" -TimeoutSec 60
Start-Sleep -Seconds 6
if ((Get-CdpBodyText) -notmatch "inbox|compose") { Write-Log "GMAIL NOT LOGGED IN"; Stop-EdgeCdp; exit 1 }

$sent = 0
$results = @()
$sentThisRun = @()
foreach ($r in $ready) {
    # shared 8/day cap across agents - only defer when the hard cap is hit
    if (-not (Test-EmailBudget)) { Write-Log "daily email cap reached (8/day across agents) - stopping sends; rest go out after midnight"; break }
    if ($sentThisRun -contains $r.Email) { Write-Log ("  skip (already sent this run): " + $r.Email); continue }
    Write-Log ("Sending to " + $r.Email + " [" + $r.Role + "]...")
    $subject = "BIM Coordinator | 8+ yrs | Revit - Navisworks - ISO 19650"
    try {
        $toEnc = [uri]::EscapeDataString($r.Email)
        $su = [uri]::EscapeDataString($subject)
        $bo = [uri]::EscapeDataString($r.Body)
        Navigate-Cdp "https://mail.google.com/mail/u/0/?view=cm&fs=1&to=$toEnc&su=$su&body=$bo" -TimeoutSec 60
        Start-Sleep -Seconds 8
        $chipOk = Invoke-CdpJs "(function(){ return document.body.innerText.indexOf($($r.Email | ConvertTo-Json -Compress)) >= 0; })()"
        if (-not $chipOk) { Write-Log "  recipient missing - skip"; continue }
        $cvOk = $false
        try { $cvOk = Set-CdpFileInput $r.Pdf } catch { Write-Log "  attach error" }
        if ($cvOk) { Start-Sleep -Seconds 12 }
        $sentOk = Invoke-CdpJs "(function(){ var b = Array.from(document.querySelectorAll('div[role=button], button')).find(function(e){ return /^send/i.test((e.innerText||'').trim()); }); if(b){ b.click(); return 'ok'; } return null; })()"
        if ($sentOk) {
            Start-Sleep -Seconds 8
            if ((Get-CdpBodyText) -match "message sent|your message has been sent|view message") {
                Write-Log "  SENT (confirmed)"
                $history = @($history + $r.Email | Sort-Object -Unique)
                $sentThisRun += $r.Email
                $results += [PSCustomObject]@{ Item = $r.Item; Email = $r.Email; Folder = $r.Folder }
                $sent++
                try { Add-SentEmail -To $r.Email -Agent "Tailored" } catch {}
            } else { Write-Log "  no confirmation" }
        } else { Write-Log "  send button missing" }
    } catch { Write-Log "  error: $($_.Exception.Message)" }
    Start-Sleep -Seconds 3
}
@{ emails = $history } | ConvertTo-Json -Depth 4 | Set-Content $historyFile -Encoding UTF8
Stop-EdgeCdp

# ---------- phase 3: tracker update (single xlsx session - ImportExcel, no Office needed) ----------
if ($results.Count -gt 0) {
    $s = $null
    try {
        $s = Open-Tracker
        $wsJ = $s.Package.Workbook.Worksheets["Jobs"]
        $wsF = $s.Package.Workbook.Worksheets["FeedBot-Link"]
        foreach ($res in $results) {
            if ($res.Item.Kind -eq "Job") {
                $wsJ.Cells[$res.Item.Row, 8].Value = "Applied (email)"
                $wsJ.Cells[$res.Item.Row, 10].Value = (Get-Date -Format "yyyy-MM-dd")
                $wsJ.Cells[$res.Item.Row, 12].Value = $res.Folder
            } else {
                $lastF = 0; if ($wsF -and $wsF.Dimension) { $lastF = $wsF.Dimension.End.Row }
                for ($rr = 2; $rr -le $lastF; $rr++) {
                    if ([string]$wsF.Cells[$rr, 2].Text -eq $res.Email) {
                        $wsF.Cells[$rr, 5].Value = "Sent"
                        $wsF.Cells[$rr, 6].Value = (Get-Date -Format "yyyy-MM-dd HH:mm")
                    }
                }
            }
        }
        Close-Tracker $s -Save $true
        $s = $null
        Write-Log "Tracker updated for $sent send(s)."
    } catch {
        if ($s) { try { Close-Tracker $s -Save $false } catch {} }
        $q = @($results | ForEach-Object { [PSCustomObject]@{ Kind = $_.Item.Kind; Row = $_.Item.Row; Email = $_.Email; Folder = $_.Folder; Date = (Get-Date -Format "yyyy-MM-dd") } })
        $q | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $logDir "pending-tracker-updates.json") -Encoding UTF8
        Write-Log "Tracker LOCKED - $($results.Count) update(s) queued to pending-tracker-updates.json"
    }
}

Write-Log ("===== TAILORED BATCH DONE: $sent sent =====")
Write-Host "`nDone. $sent tailored application(s) sent. Log: $logFile" -ForegroundColor Green
