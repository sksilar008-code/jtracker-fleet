<#
.SYNOPSIS
  SIGNUPBOT : creates accounts on job platforms.
  For each platform: opens the signup page in the bot's Edge, autofills your data,
  then YOU complete the CAPTCHA and click submit. Bot captures the session and
  saves the account (encrypted password) to the tracker.
  Run once; re-run anytime to add more platforms.
#>
param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "config.json")
)
$ErrorActionPreference = "Continue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "Tracker.psm1") -Force
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Set-TrackerPath $config.paths.tracker
$P = $config.profile
$sec = ConvertTo-SecureString $config.signup.passwordEnc
$plainPwd = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))

$platforms = @(
    @{ Platform = "NaukriGulf";     Site = "naukrigulf.com";     Url = "https://www.naukrigulf.com/register" },
    @{ Platform = "FounditGulf";    Site = "founditgulf.com";    Url = "https://www.founditgulf.com/seeker/registration" },
    @{ Platform = "Bayt";           Site = "bayt.com";           Url = "https://www.bayt.com/en/register/" },
    @{ Platform = "GulfTalent";     Site = "gulftalent.com";     Url = "https://www.gulftalent.com/uae/registration" },
    @{ Platform = "Reed";           Site = "reed.co.uk";         Url = "https://www.reed.co.uk/register" },
    @{ Platform = "TotalJobs";      Site = "totaljobs.com";      Url = "https://www.totaljobs.com/account/register" },
    @{ Platform = "CV-Library";     Site = "cv-library.co.uk";   Url = "https://www.cv-library.co.uk/register" },
    @{ Platform = "CareerStructure";Site = "careerstructure.com";Url = "https://www.careerstructure.com/account/register" },
    @{ Platform = "IrishJobs";      Site = "irishjobs.ie";       Url = "https://www.irishjobs.ie/register" },
    @{ Platform = "SEEK";           Site = "seek.com.au";        Url = "https://www.seek.com.au/sign-up" },
    @{ Platform = "ZipRecruiter";   Site = "ziprecruiter.com";   Url = "https://www.ziprecruiter.com/register" },
    @{ Platform = "StepStone";      Site = "stepstone.de";       Url = "https://www.stepstone.de/en/register" },
    @{ Platform = "WorkInBIM";      Site = "workinbim.com";      Url = "https://www.workinbim.com/register" }
)

$existing = @(Get-TrackerAccounts)
$todo = $platforms | Where-Object { $p = $_; -not ($existing | Where-Object { $_.Platform -eq $p.Platform }) }
if (-not $todo) { Write-Host "All platforms already have accounts. Nothing to do." -ForegroundColor Green; exit 0 }

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  SIGNUPBOT - creating accounts on $($todo.Count) platforms" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "For each site: bot fills the form -> YOU solve CAPTCHA + click Register/Sign up"
Write-Host "-> come back here and press ENTER. Press S+ENTER to skip a site."
Write-Host ""

Start-EdgeCdp

function Save-CookiesNow {
    try {
        Send-Cdp "Network.enable" | Out-Null
        $r = Send-Cdp "Network.getAllCookies" -TimeoutSec 30
        $out = Join-Path $ProjectDir "edge-profile\cookies.json"
        $r.cookies | ConvertTo-Json -Depth 6 | Set-Content $out -Encoding UTF8
        return $r.cookies.Count
    } catch { return 0 }
}

function Save-Account($platform, $site) {
    $encPwd = ConvertTo-SecureString $plainPwd -AsPlainText -Force | ConvertFrom-SecureString
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($config.paths.tracker)
    $ws = $wb.Worksheets.Item("Accounts")
    $next = $ws.UsedRange.Rows.Count + 1
    $ws.Cells.Item($next, 1) = $platform
    $ws.Cells.Item($next, 2) = $site
    $ws.Cells.Item($next, 3) = $P.email
    $ws.Cells.Item($next, 4) = $encPwd
    $ws.Cells.Item($next, 5) = "no"
    $ws.Cells.Item($next, 6) = (Get-Date -Format "yyyy-MM-dd")
    $ws.Cells.Item($next, 8) = "active"
    $wb.Save(); $wb.Close($false); $excel.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

function Get-JsSignupFill($profileObj, $pwd) {
    $json = $profileObj | ConvertTo-Json -Compress
    $pj = $pwd | ConvertTo-Json -Compress
    return @"
(function(){
  var P = $json, PW = $pj;
  var filled = 0;
  Array.from(document.querySelectorAll('input:not([type=hidden]):not([type=submit]):not([type=checkbox]):not([type=radio])')).forEach(function(el){
    var id = ((el.name||'')+' '+(el.id||'')+' '+(el.placeholder||'')+' '+(el.autocomplete||'')).toLowerCase();
    var val = null;
    if (el.type === 'password') val = PW;
    else if (/email/.test(id)) val = P.email;
    else if (/first.?name|given/.test(id)) val = P.firstName;
    else if (/last.?name|family|surname/.test(id)) val = P.lastName;
    else if (/full.?name|your.?name|\bname\b/.test(id)) val = P.fullName;
    else if (/phone|mobile|tel/.test(id)) val = P.phone;
    else if (/city|location/.test(id)) val = P.location;
    if (val){ el.focus(); el.value = val; el.dispatchEvent(new Event('input',{bubbles:true})); el.dispatchEvent(new Event('change',{bubbles:true})); filled++; }
  });
  return filled;
})()
"@
}

$done = 0
foreach ($site in $todo) {
    Write-Host ""
    Write-Host "====> $($site.Platform) ($($site.Site))" -ForegroundColor Yellow
    try { Navigate-Cdp $site.Url } catch { Write-Host "    Navigation issue - page may still load." }
    Start-Sleep 3
    $n = Invoke-CdpJs (Get-JsSignupFill $P $plainPwd)
    Write-Host "    Auto-filled $n field(s). Complete CAPTCHA + submit on the Edge window."
    $ans = Read-Host "    Press ENTER when registered (S = skip)"
    if ($ans -eq "S" -or $ans -eq "s") { Write-Host "    Skipped."; continue }
    $ck = Save-CookiesNow
    Save-Account $site.Platform $site.Site
    $done++
    Write-Host "    SAVED: $($site.Platform) account recorded + $ck cookies captured." -ForegroundColor Green
}

Stop-EdgeCdp
Write-Host ""
Write-Host "SignupBot complete: $done new account(s) created and saved." -ForegroundColor Green
Write-Host "IMPORTANT: check your Gmail and click the verification links from each site."
Read-Host "Press ENTER to close"
