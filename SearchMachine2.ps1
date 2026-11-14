<#
.SYNOPSIS
  SEARCH MACHINE - 2 : searches your own files (PDF, Excel, Word, text, CSV, JSON, HTML)
  for a name or keyword and shows where it was found.

.EXAMPLE
  .\SearchMachine2.ps1 -Query "AECOM"
  .\SearchMachine2.ps1 -Query "Navisworks" -Report
#>
param(
    [string]$Query,
    [string]$Root = "C:\Users\Silar Shaik\OneDrive - Tiger Group\J-Tracker",
    [switch]$Report
)

$ErrorActionPreference = "SilentlyContinue"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Lib = Join-Path $ProjectDir "lib"

if (-not $Query) { $Query = Read-Host "Search for" }
if (-not $Query) { Write-Host "No query given."; exit 1 }

# Load PdfPig for PDF text extraction
Get-ChildItem $Lib -Filter "UglyToad*.dll" | ForEach-Object { [Reflection.Assembly]::LoadFrom($_.FullName) | Out-Null }
Add-Type -AssemblyName System.IO.Compression.FileSystem

$textExt = @(".txt", ".md", ".csv", ".json", ".html", ".htm", ".log", ".xml")
$results = New-Object System.Collections.Generic.List[object]
$q = [regex]::Escape($Query)

function Add-Hit($file, $where, $line) {
    $snippet = ($line -replace "\s+", " ").Trim()
    if ($snippet.Length -gt 120) {
        $i = $snippet.IndexOf($Query, [StringComparison]::OrdinalIgnoreCase)
        if ($i -lt 0) { $i = 0 }
        $start = [Math]::Max(0, $i - 50)
        $snippet = "..." + $snippet.Substring($start, [Math]::Min(120, $snippet.Length - $start)) + "..."
    }
    $script:results.Add([PSCustomObject]@{ File = $file; Where = $where; Match = $snippet })
}

Write-Host ""
Write-Host "SEARCH MACHINE - 2" -ForegroundColor Cyan
Write-Host "Searching '$Query' under $Root ..." -ForegroundColor Cyan
Write-Host ""

$files = Get-ChildItem $Root -Recurse -File -Force
foreach ($f in $files) {
    # 1) File name match
    if ($f.Name -match $q) { Add-Hit $f.FullName "file name" $f.Name }

    $ext = $f.Extension.ToLower()
    try {
        if ($textExt -contains $ext) {
            $ln = 0
            foreach ($line in [IO.File]::ReadLines($f.FullName)) {
                $ln++
                if ($line -match $q) { Add-Hit $f.FullName "line $ln" $line }
            }
        }
        elseif ($ext -eq ".pdf") {
            $doc = [UglyToad.PdfPig.PdfDocument]::Open($f.FullName)
            for ($p = 1; $p -le $doc.NumberOfPages; $p++) {
                $t = $doc.GetPage($p).Text
                if ($t -match $q) {
                    foreach ($line in ($t -split "`n")) { if ($line -match $q) { Add-Hit $f.FullName "page $p" $line; break } }
                }
            }
            $doc.Dispose()
        }
        elseif ($ext -eq ".xlsx" -or $ext -eq ".docx") {
            $zip = [IO.Compression.ZipFile]::OpenRead($f.FullName)
            foreach ($entry in $zip.Entries) {
                if ($entry.FullName -notmatch "\.xml$") { continue }
                if ($entry.FullName -notmatch "sharedStrings|document\.xml|sheet\d+\.xml") { continue }
                $sr = New-Object IO.StreamReader($entry.Open())
                $xml = $sr.ReadToEnd(); $sr.Close()
                $plain = $xml -replace "<[^>]+>", " "
                if ($plain -match $q) {
                    foreach ($m in [regex]::Matches($plain, ".{0,60}$q.{0,60}", "IgnoreCase")) {
                        Add-Hit $f.FullName $entry.FullName $m.Value
                        break
                    }
                }
            }
            $zip.Dispose()
        }
    } catch { }
}

if ($results.Count -eq 0) {
    Write-Host "No matches found for '$Query'." -ForegroundColor Yellow
} else {
    Write-Host "$($results.Count) match(es):" -ForegroundColor Green
    $results | Format-Table -AutoSize -Wrap | Out-String -Width 200 | Write-Host
}

if ($Report) {
    $repDir = Join-Path $ProjectDir "reports"
    New-Item -ItemType Directory -Force $repDir | Out-Null
    $rep = Join-Path $repDir ("search-" + ($Query -replace '[^\w-]', '_') + "-" + (Get-Date -Format "yyyyMMdd-HHmm") + ".md")
    $out = @("# Search Machine - 2 report", "", "Query: **$Query**", "Date: $(Get-Date)", "", "| File | Where | Match |", "|---|---|---|")
    foreach ($r in $results) { $out += "| $($r.File) | $($r.Where) | $($r.Match -replace '\|','\|') |" }
    $out | Set-Content $rep -Encoding UTF8
    Write-Host "Report saved: $rep" -ForegroundColor Cyan
}
