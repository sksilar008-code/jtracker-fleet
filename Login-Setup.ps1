<#
.SYNOPSIS
  LOGIN SETUP - one-time login for all your job platforms.
  Opens the bot's own Edge window. On each site, click "Sign in with Google"
  (or log in however you normally do). Your session is saved in the bot's
  browser profile, so ApplyBot stays logged in on all future runs.
  Run again anytime a session expires.
#>
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $ProjectDir "Tracker.psm1") -Force

$sites = @(
    @{ Platform = "DrJobs";      Site = "drjobs.ae";       LoginUrl = "https://drjobs.ae/login" },
    @{ Platform = "Indeed";      Site = "indeed.com";      LoginUrl = "https://secure.indeed.com/auth" },
    @{ Platform = "FounditGulf"; Site = "founditgulf.com"; LoginUrl = "https://www.founditgulf.com/login" },
    @{ Platform = "NaukriGulf";  Site = "naukrigulf.com";  LoginUrl = "https://www.naukrigulf.com/login" },
    @{ Platform = "Bayt";        Site = "bayt.com";        LoginUrl = "https://www.bayt.com/en/login/" },
    @{ Platform = "LinkedIn";    Site = "linkedin.com";    LoginUrl = "https://www.linkedin.com/login" }
)

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  SEARCH MACHINE - 2  |  one-time login setup" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "An Edge window will open. For each site:"
Write-Host "  1. Click 'Sign in with Google' and log in as sksilar008@gmail.com"
Write-Host "  2. Come back here and press ENTER"
Write-Host "Press Q + ENTER at any site to skip it."
Write-Host ""

Start-EdgeCdp

foreach ($s in $sites) {
    Write-Host "----> $($s.Platform) ($($s.Site))" -ForegroundColor Yellow
    Navigate-Cdp $s.LoginUrl
    $ans = Read-Host "    Log in on the Edge window, then press ENTER (or Q to skip)"
    if ($ans -eq "Q" -or $ans -eq "q") { Write-Host "    Skipped."; continue }

    # Record the account in the tracker (no password - session-based login)
    $existing = @(Get-TrackerAccounts | Where-Object { $_.Platform -eq $s.Platform })
    if (-not $existing) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open("C:\Users\Silar Shaik\OneDrive - Tiger Group\J-Tracker\JobTracker.xlsx")
        $ws = $wb.Worksheets.Item("Accounts")
        $next = $ws.UsedRange.Rows.Count + 1
        $ws.Cells.Item($next, 1) = $s.Platform
        $ws.Cells.Item($next, 2) = $s.Site
        $ws.Cells.Item($next, 3) = "sksilar008@gmail.com"
        $ws.Cells.Item($next, 4) = "GOOGLE-LOGIN"
        $ws.Cells.Item($next, 5) = "yes"
        $ws.Cells.Item($next, 6) = (Get-Date -Format "yyyy-MM-dd")
        $ws.Cells.Item($next, 8) = "active"
        $wb.Save(); $wb.Close($false); $excel.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Write-Host "    Saved to tracker (session login)." -ForegroundColor Green
    } else {
        Write-Host "    Already in tracker." -ForegroundColor Green
    }
}

Stop-EdgeCdp
Write-Host ""
Write-Host "Login setup complete. ApplyBot will reuse these sessions automatically." -ForegroundColor Green
Read-Host "Press ENTER to close"
