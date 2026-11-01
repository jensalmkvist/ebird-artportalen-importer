<#
.SYNOPSIS
    Convert eBird checklists into rows for Artportalen's Excel bulk import.

.DESCRIPTION
    Without date parameters: converts the CSVs in "Import files\" as before.

    With -StartDate, -StartDate/-EndDate or -SinceLastImport: first downloads
    your eBird checklists for those dates with ebird-checklist-scraper (into a
    new folder under "Import files\"), then converts them. Checklists that were
    imported before are skipped, so overlapping dates don't give duplicates.

    import_state.json remembers the last import date and the imported
    checklist IDs. On the first dated run, every S*.csv already in
    "Import files\" counts as imported.

.PARAMETER StartDate
    Get all checklists from this date (yyyy-MM-dd) up to today, or up to -EndDate.

.PARAMETER EndDate
    Last date to include (default: today).

.PARAMETER SinceLastImport
    Get all checklists from the last import date up to today.

.PARAMETER ScraperPath
    Folder of ebird-checklist-scraper (default: next to this repo).

.EXAMPLE
    .\ebird_artportalen_importer.ps1 -StartDate 2026-10-01
    .\ebird_artportalen_importer.ps1 -StartDate 2026-10-01 -EndDate 2026-10-15
    .\ebird_artportalen_importer.ps1 -SinceLastImport
#>

param(
  [DateTime]$StartDate,
  [DateTime]$EndDate,
  [switch]$SinceLastImport,
  [string]$ScraperPath = (Join-Path $PSScriptRoot '..\ebird-checklist-scraper')
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

# set the console to UTF-8 before running the lua script (for å/ä/ö)
chcp 65001 >$null
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

$importDir = Join-Path $PSScriptRoot 'Import files'
$statePath = Join-Path $PSScriptRoot 'import_state.json'

# --- No dates: convert what is in "Import files\" ----------------------------
if (-not $PSBoundParameters.ContainsKey('StartDate') -and
    -not $PSBoundParameters.ContainsKey('EndDate') -and -not $SinceLastImport)
{
  lua .\ebird_to_artportalen.lua "$importDir/" .\out_for_excel.csv
  exit $LASTEXITCODE
}

# --- State -------------------------------------------------------------------
function Get-ChecklistIds($dir)
{
  Get-ChildItem $dir -Filter 'S*.csv' -Recurse -File |
    Where-Object { $_.Name -match '^(S\d+)' } |
    ForEach-Object { $Matches[1] }
}

if (Test-Path $statePath)
{
  $state = Get-Content $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
  $lastImport = $state.last_import
  $importedIds = @($state.imported_ids)
} else
{
  $lastImport = $null
  $importedIds = @(Get-ChecklistIds $importDir | Sort-Object -Unique)
  Write-Host "First dated import: $($importedIds.Count) checklists already in 'Import files' count as imported."
}

# --- Date range --------------------------------------------------------------
if ($SinceLastImport)
{
  if ($PSBoundParameters.ContainsKey('StartDate'))
  { throw 'Use either -SinceLastImport or -StartDate, not both.'
  }
  if (-not $lastImport)
  { throw 'No previous import recorded yet. Run once with -StartDate.'
  }
  $StartDate = [DateTime]::ParseExact($lastImport, 'yyyy-MM-dd', $null)
} elseif (-not $PSBoundParameters.ContainsKey('StartDate'))
{
  throw '-EndDate needs -StartDate.'
}
if (-not $PSBoundParameters.ContainsKey('EndDate'))
{ $EndDate = Get-Date
}
$start = $StartDate.ToString('yyyy-MM-dd')
$end = $EndDate.ToString('yyyy-MM-dd')
if ($end -lt $start)
{ throw "End date $end is before start date $start."
}

# --- Download ----------------------------------------------------------------
$python = Join-Path $ScraperPath '.venv\Scripts\python.exe'
$scraper = Join-Path $ScraperPath 'ebird_month_download.py'
foreach ($p in $python, $scraper)
{
  if (-not (Test-Path $p))
  { throw "Not found: $p (set -ScraperPath to your ebird-checklist-scraper folder)"
  }
}

$outDir = Join-Path $importDir ("{0} eBird {1} to {2}" -f (Get-Date -Format 'yyyy-MM-dd HHmmss'), $start, $end)
$skipFile = Join-Path ([IO.Path]::GetTempPath()) 'ebird_imported_ids.txt'
Set-Content -Path $skipFile -Value $importedIds -Encoding UTF8

Write-Host "Downloading eBird checklists $start to $end ..."
& $python $scraper --start $start --end $end --out $outDir --skip-ids $skipFile
$downloadOk = ($LASTEXITCODE -eq 0)
Remove-Item $skipFile -ErrorAction SilentlyContinue
if (-not $downloadOk)
{
  Write-Warning 'Some checklists failed to download. Converting the rest; the last import date is not moved, so a later -SinceLastImport retries them.'
}

$newIds = @()
if (Test-Path $outDir)
{ $newIds = @(Get-ChecklistIds $outDir)
}

# --- Convert -----------------------------------------------------------------
if ($newIds.Count -eq 0)
{
  Write-Host 'No new checklists to import.'
  if (Test-Path $outDir)
  { Remove-Item $outDir -Recurse -Force
  }
} else
{
  lua .\ebird_to_artportalen.lua "$outDir/" .\out_for_excel.csv
  if ($LASTEXITCODE -ne 0)
  { throw 'Conversion failed; import state not updated.'
  }
}

# --- Save state --------------------------------------------------------------
if ($downloadOk -and ((-not $lastImport) -or $end -gt $lastImport))
{ $lastImport = $end
}
$state = [ordered]@{
  last_import  = $lastImport
  imported_ids = @(@($importedIds) + $newIds | Sort-Object -Unique)
}
$state | ConvertTo-Json | Set-Content -Path $statePath -Encoding UTF8
Write-Host "Imported $($newIds.Count) checklists. Last import date: $lastImport"
