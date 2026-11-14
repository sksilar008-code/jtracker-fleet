<#
  SEND-DIRECTEMAIL : one-off targeted email application via YOUR Gmail (browser session).
  Sends a specific CV + cover note to a specific address. Based on EmailBot's proven flow.
  Use -TestMode to send to your own address first (verifies attachment works).
#>
param(
    [Parameter(Mandatory = $true)][string]$To,
    [Parameter(Mandatory = $true)][string]$Subject,
    [Parameter(Mandatory = $true)][string]$BodyFile,
    [Parameter(Mandatory = $true)][string]$CvPath,
    [switch]$Headless,
    [switch]$TestMode
)
$ErrorActionPreference = "Continue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking
$config = Get-Content (Join-Path $ProjectDir "config.json") -Raw | ConvertFrom-Json
$P = $config.profile

$logDir = Join-Path $ProjectDir "logs"; New-Item -ItemType Directory -Force $logDir | Out-Null
$logFile = Join-Path $logDir ("direct-email-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".log")
function Write-Log($msg) { $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"; Write-Host $line; Add-Content $logFile $line }

if (-not (Test-Path $CvPath)) { Write-Log "CV NOT FOUND: $CvPath"; exit 1 }
if (-not (Test-Path $BodyFile)) { Write-Log "BODY FILE NOT FOUND: $BodyFile"; exit 1 }
$body = (Get-Content $BodyFile -Raw).Trim()
$sendTo = $To
if ($TestMode) { $sendTo = $P.email; Write-Log "TEST MODE -> sending to self: $sendTo" }

Write-Log "Starting Edge..."
Start-EdgeCdp -Headless:$Headless
Enable-Stealth
$cookieFile = Join-Path $ProjectDir "edge-profile\cookies.json"
if (Test-Path $cookieFile) {
    $sc = Get-Content $cookieFile -Raw | ConvertFrom-Json
    if ($sc) { Send-Cdp "Network.enable" | Out-Null; Send-Cdp "Network.setCookies" @{ cookies = @($sc) } -TimeoutSec 30 | Out-Null }
}

Write-Log "Opening Gmail (basic HTML mode)..."
Navigate-Cdp "https://mail.google.com/mail/u/0/h/" -TimeoutSec 90
Start-Sleep -Seconds 6
$gBody = Get-CdpBodyText
if ($gBody -notmatch "inbox|compose|search mail|sign out") {
    Write-Log "GMAIL NOT LOGGED IN - cannot send."
    Stop-EdgeCdp; exit 1
}
Write-Log "Gmail session OK."

try {
    $toEnc = [uri]::EscapeDataString($sendTo)
    $su = [uri]::EscapeDataString($Subject)
    $bo = [uri]::EscapeDataString($body)
    Navigate-Cdp "https://mail.google.com/mail/u/0/?view=cm&fs=1&to=$toEnc&su=$su&body=$bo" -TimeoutSec 60
    Start-Sleep -Seconds 8
    $chipOk = Invoke-CdpJs "(function(){ return document.body.innerText.indexOf($($sendTo | ConvertTo-Json -Compress)) >= 0; })()"
    Write-Log "Recipient present on page: $chipOk ($sendTo)"
    if (-not $chipOk) { Write-Log "RECIPIENT MISSING - aborting to avoid empty email."; Stop-EdgeCdp; exit 1 }

    $cvOk = $false
    try { $cvOk = Set-CdpFileInput $CvPath } catch { Write-Log "attach error: $($_.Exception.Message)" }
    Write-Log "CV attached: $cvOk"
    if ($cvOk) { Start-Sleep -Seconds 12 }
    else { Write-Log "ATTACHMENT FAILED - aborting (application needs the CV)."; Stop-EdgeCdp; exit 1 }

    # screenshot before send for the record
    Save-CdpScreenshot (Join-Path $logDir ("direct-email-before-send-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".png"))

    $sentOk = Invoke-CdpJs @"
(function(){
  var b = Array.from(document.querySelectorAll('div[role=button], button')).find(function(e){ return /^send/i.test((e.innerText||'').trim()); });
  if (b){ b.click(); return 'ok'; }
  return null;
})()
"@
    if ($sentOk) {
        Start-Sleep -Seconds 8
        $after = Get-CdpBodyText
        if ($after -match "message sent|your message has been sent|view message") {
            Write-Log "SENT (confirmed)."
        } else {
            Write-Log "Send clicked but NO CONFIRMATION - check Sent folder manually."
        }
    } else {
        Write-Log "Send button not found - draft left open."
    }
} catch { Write-Log "send error: $($_.Exception.Message)" }

Stop-EdgeCdp
Write-Log "===== DONE ====="
