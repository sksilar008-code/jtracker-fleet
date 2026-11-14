<#
.SYNOPSIS
  Gmail re-login handoff for the CLOUD fleet (GitHub runners get new IPs each run,
  so Google occasionally invalidates the shared session cookie).
  Opens the bot's own Edge window at mail.google.com. YOU sign in there - the
  assistant never sees your credentials. Then this captures ALL session cookies
  into edge-profile\cookies.json, which gets pushed to the private state repo.
.USAGE
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Gmail-Login.ps1
  (run from the project folder; if the bot Edge is already open, close it first)
#>
$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Gmail re-login  (cloud fleet session refresh)" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

Start-EdgeCdp
Start-Sleep -Seconds 2
Navigate-Cdp "https://mail.google.com/mail/u/0/" -TimeoutSec 60

Write-Host "In the Edge window that just opened:" -ForegroundColor Yellow
Write-Host "  1. Sign in to Gmail as sksilar008@gmail.com"
Write-Host "  2. Wait until your INBOX is fully visible"
Write-Host "  3. Come back here and press ENTER" -ForegroundColor Yellow
Read-Host "Press ENTER when the inbox is visible"

Send-Cdp "Network.enable" | Out-Null
$r = Send-Cdp "Network.getAllCookies" -TimeoutSec 30
$out = Join-Path $ProjectDir "edge-profile\cookies.json"
New-Item -ItemType Directory -Force -Path (Split-Path $out) | Out-Null
$r.cookies | ConvertTo-Json -Depth 6 | Set-Content $out -Encoding UTF8

$gm = @($r.cookies | Where-Object { $_.domain -match "google|gmail" })
Write-Host ""
Write-Host ("Saved {0} cookies ({1} google/gmail) -> {2}" -f $r.cookies.Count, $gm.Count, $out) -ForegroundColor Green

# verify the session actually opens the inbox before we push it anywhere
Navigate-Cdp "https://mail.google.com/mail/u/0/#inbox" -TimeoutSec 45 | Out-Null
$body = Get-CdpBodyText
if ($body -match "inbox|compose") {
    Write-Host "VERIFIED: Gmail inbox loads - session is good for cloud runs." -ForegroundColor Green
} else {
    Write-Host "WARNING: inbox not detected - Google may still be challenging this session." -ForegroundColor Yellow
}
Stop-EdgeCdp
Write-Host ""
Write-Host "Done. Tell the assistant: 'gmail cookies captured' - the state repo will be pushed." -ForegroundColor Cyan
