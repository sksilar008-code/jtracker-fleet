<#
  Tailor.psm1 : JD-matched CV tailoring engine (shared).
  Extracted from TailoredApplyBatch.ps1 so ApplyBot can prepare a tailored CV
  for form/Easy-Apply submissions the same way email applications get one.
  All facts in the CV are true (Silar's real experience) - only emphasis changes per JD.
#>

$script:SkillMap = @(
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

function Get-MatchedSkills($text) {
    $m = @()
    foreach ($s in $script:SkillMap) { if ($text -match $s.pat) { $m += $s } }
    return $m
}

function New-SafeName($s) { return (($s -replace "[^\w\- ]", "") -replace "\s+", "_").Trim("_") }

function New-TailoredCV {
    param([string]$Folder, [string]$Role, [string]$Company, [string]$Location, $Matched, [string]$DescSnippet, $Profile)
    New-Item -ItemType Directory -Force -Path $Folder | Out-Null
    $matchLabels = @($Matched | ForEach-Object { $_.label } | Select-Object -Unique | Select-Object -First 8)
    $bullets = @($Matched | ForEach-Object { $_.bullet } | Select-Object -Unique | Select-Object -First 5)
    $generic = @(
        "Federated architectural, structural and MEP models across 25+ projects; ran clash detection and resolution workshops to submission-ready, clash-free deliverables.",
        "Administered ACC / BIM 360 CDE environments to ISO 19650, tracking coordination actions to closure with engineers and construction teams.",
        "8+ years across six countries (UAE, Singapore, India, Saudi Arabia, Netherlands, USA) on high-rise, infrastructure and transport programs."
    )
    $i = 0
    while ($bullets.Count -lt 4 -and $i -lt $generic.Count) { if ($bullets -notcontains $generic[$i]) { $bullets += $generic[$i] }; $i++ }
    $sector = "Transport & Infrastructure"
    if ($DescSnippet -match "high-?rise|tower|residential") { $sector = "High-Rise & Infrastructure" }
    elseif ($DescSnippet -match "hospital|healthcare") { $sector = "Healthcare & Infrastructure" }
    elseif ($DescSnippet -match "data ?centre") { $sector = "Data Centres & Infrastructure" }
    $tags = @($matchLabels | ForEach-Object { "<span class=`"tag`">$_</span>" }) -join ""
    $blist = ($bullets | ForEach-Object { "<li>$_</li>" }) -join ""
    $matchLine = ($matchLabels -join " &#9679; ")
    if (-not $matchLine) { $matchLine = "multidisciplinary BIM coordination, clash detection and ISO 19650 information management" }
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
<div class="subtitle">$Role &nbsp;|&nbsp; $sector &nbsp;|&nbsp; ISO 19650</div>
<div class="contact">Phone: $($Profile.phone) &nbsp;&nbsp;|&nbsp;&nbsp; Email: $($Profile.email) &nbsp;&nbsp;|&nbsp;&nbsp; Sharjah, UAE (open to $Location) &nbsp;&nbsp;|&nbsp;&nbsp; English (Fluent)</div>
</div>
<div class="body">
<h3>Profile</h3>
<p>BIM Coordinator with <b>8+ years</b> delivering complex building and infrastructure programs across six countries, including Singapore MRT station, depot and tunnel packages with WSP. Specialist in multidisciplinary model federation, clash coordination and submission-ready information management under <b>ISO 19650</b> (ACC, Navisworks, Revit, AutoCAD, Civil 3D), with Dynamo/Power BI automation. GMICE graduate in Civil Engineering (JNTU Kakinada).</p>
<h3>Job Match - $Role @ $Company</h3>
<p><b>Directly matching your requirements:</b> $matchLine.</p>
<h3>Why This Role - Matched Achievements</h3>
<ul>$blist</ul>
<div class="cols">
<div class="main">
<h3>Professional Experience</h3>
<p class="exp"><b>BIM Coordinator - Tiger Group, UAE</b> <span class="period">2025 - Present</span></p>
<ul><li>Coordinate and federate architectural, structural and MEP models; run clash detection cycles and resolution workshops to submission-ready deliverables.</li><li>Administer ACC CDE to ISO 19650; QA/QC design packages; track coordination actions to closure with engineers and leadership.</li></ul>
<p class="exp"><b>BIM Technician - Transportation &amp; Infrastructure - WSP</b> <span class="period">2023 - 2025</span></p>
<ul><li>Delivered BIM packages for <b>Singapore MRT J101 (Tengah Depot) and CR116 (Ang Mo Kio Interchange Station &amp; Tunnels)</b> and Mumbai Coastal Road interchanges.</li><li>Ran interdisciplinary clash reviews; managed BIM 360/ACC digital delivery aligned to BIM codes.</li></ul>
<p class="exp"><b>Team Lead - BIM - The BIM Engineers</b> <span class="period">2021 - 2023</span></p>
<ul><li>Led model production and audits; enforced BEPs; scan-to-BIM for Frankfurt station; mentored junior modellers.</li></ul>
<p class="exp"><b>Team Lead - BIM (Architecture) - Innovspatial</b> <span class="period">2018 - 2021</span></p>
<ul><li>Led CAD/BIM production via shared CDE; 30+ authority permit sets; QA/QC audits to client requirements.</li></ul>
<h3>Key Projects</h3>
<ul><li>J101 Tengah Depot &amp; MRT Facilities, Jurong Region Line, Singapore</li><li>CR116 Ang Mo Kio Interchange Station &amp; Tunnels, Singapore</li><li>Tiger Sky Tower - AZHA - Jade Tower (High-Rise UAE)</li><li>Frankfurt Station Scan-to-BIM - Microsoft Data Center Singapore</li></ul>
</div>
<div class="side">
<h3>Expertise</h3>
<p>Model federation &amp; QA<br>Clash coordination &amp; workshops<br>ISO 19650 / BEP / EIR / LOD<br>CDE administration (ACC)<br>Submission readiness<br>Action tracking to closure<br>Automation &amp; reporting</p>
<h3>Core Skills</h3>
<p>$tags</p>
<h3>Achievements</h3>
<ul class="ach"><li>25+ projects, 6 countries</li><li>Singapore MRT station, depot &amp; tunnel packages</li><li>ACC/BIM 360 across 25+ projects</li><li>ISO 19650-compliant delivery</li></ul>
<h3>Education</h3>
<p>B.Tech Civil Engineering<br>JNTU Kakinada, India<br>2014-2018</p>
<h3>Professional Body</h3>
<p>Graduate Member (GMICE)<br>Institution of Civil Engineers, UK</p>
</div>
</div>
<div class="footerline"></div>
</div>
</body></html>
"@
    $html | Set-Content (Join-Path $Folder "cv.html") -Encoding UTF8
    # PDF via Edge print engine (no CDP - stable)
    $pdf = Join-Path $Folder "SILAR - BIM Coordinator.pdf"
    $edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    if (-not (Test-Path $edge)) { $edge = "C:\Program Files\Microsoft\Edge\Application\msedge.exe" }
    & $edge --headless --disable-gpu --no-pdf-header-footer --print-to-pdf="$pdf" ("file:///" + ((Join-Path $Folder "cv.html") -replace "\\", "/")) 2>$null
    Start-Sleep -Seconds 5
    if (Test-Path $pdf) { return $pdf } else { return $null }
}

# High-level helper: tailor a CV for one tracker job (idempotent - reuses existing PDF)
function New-TailoredCVForJob {
    param($Job, [string]$JdText, [string]$TailoredRoot, $Profile)
    $role = $Job.Title; if ($role.Length -gt 70) { $role = $role.Substring(0, 70) }
    $folderName = "row$($Job.Row)_" + (New-SafeName ("$($Job.Title) $($Job.Company)"))
    $folder = Join-Path $TailoredRoot $folderName
    $existing = Get-ChildItem $folder -Filter *.pdf -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($existing) { return $existing.FullName }
    $matched = Get-MatchedSkills $JdText
    $loc = $Job.Location; if (-not $loc) { $loc = "UAE" }
    $pdf = New-TailoredCV -Folder $folder -Role $role -Company $Job.Company -Location $loc -Matched $matched -DescSnippet $JdText -Profile $Profile
    return $pdf
}

Export-ModuleMember -Function Get-MatchedSkills, New-SafeName, New-TailoredCV, New-TailoredCVForJob
