<#
.SYNOPSIS
  Connects to the CURRENTLY OPEN bot Edge window (while you are still logged in)
  and saves all session cookies so ApplyBot can restore them on every run.
#>
$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $ProjectDir "EdgeCdp.psm1") -Force -DisableNameChecking

Connect-EdgeCdp
Send-Cdp "Network.enable" | Out-Null
$r = Send-Cdp "Network.getAllCookies" -TimeoutSec 30
$outDir = Join-Path $ProjectDir "edge-profile"
New-Item -ItemType Directory -Force $outDir | Out-Null
$out = Join-Path $outDir "cookies.json"
$r.cookies | ConvertTo-Json -Depth 6 | Set-Content $out -Encoding UTF8
Write-Host ""
Write-Host ("Saved " + $r.cookies.Count + " cookies to " + $out) -ForegroundColor Green
$drjobs = @($r.cookies | Where-Object { $_.domain -match "drjob" })
Write-Host ("DrJobs cookies captured: " + $drjobs.Count)
foreach ($c in $drjobs) { Write-Host ("  " + $c.name + " (" + $c.domain + ")") }
if ($drjobs.Count -eq 0) { Write-Host "WARNING: no DrJobs cookies found - are you logged in on DrJobs in this window?" -ForegroundColor Yellow }
