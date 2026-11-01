<#
.SYNOPSIS
    Convert eBird checklists into rows for Artportalen's Excel bulk import.

.DESCRIPTION
    Without date parameters: converts the CSVs in "Import files\" as before.

    With -StartDate, -StartDate/-EndDate or -SinceLastImport: first downloads
    your eBird checklists for those dates with ebird-checklist-scraper (into a
    new folder under "Import files\"), then converts them. Checklists that were
    imported before are skipped, so overlapping dates don't give duplicates.

    Type p at any question to pause. Your answers so far are kept, and
    -Resume continues with the already downloaded checklists later.

    import_state.json remembers the last import date, the imported checklist
    IDs and a paused import. On the first dated run, every S*.csv already in
    "Import files\" counts as imported.

.PARAMETER StartDate
    Get all checklists from this date (yyyy-MM-dd) up to today, or up to -EndDate.

.PARAMETER EndDate
    Last date to include (default: today).

.PARAMETER SinceLastImport
    Get all checklists from the last import date up to today.

.PARAMETER Resume
    Continue a paused import (no new download).

.PARAMETER ScraperPath
    Folder of ebird-checklist-scraper (default: next to this repo).

.EXAMPLE
    .\ebird_artportalen_importer.ps1 -StartDate 2026-10-01
    .\ebird_artportalen_importer.ps1 -StartDate 2026-10-01 -EndDate 2026-10-15
    .\ebird_artportalen_importer.ps1 -SinceLastImport
    .\ebird_artportalen_importer.ps1 -Resume
#>

param(
  [DateTime]$StartDate,
  [DateTime]$EndDate,
  [switch]$SinceLastImport,
  [switch]$Resume,
  [string]$ScraperPath = (Join-Path $PSScriptRoot '..\ebird-checklist-scraper')
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

# set the console to UTF-8 before running the lua script (for å/ä/ö)
chcp 65001 >$null
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

$importDir = Join-Path $PSScriptRoot 'Import files'
$statePath = Join-Path $PSScriptRoot 'import_state.json'
$datesGiven = $PSBoundParameters.ContainsKey('StartDate') -or
  $PSBoundParameters.ContainsKey('EndDate') -or $SinceLastImport

# Exit code of ebird_to_artportalen.lua when you type p at a question
$EXIT_PAUSED = 4

# --- No dates: convert what is in "Import files\" ----------------------------
if (-not $datesGiven -and -not $Resume)
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

$pending = $null
if (Test-Path $statePath)
{
  $state = Get-Content $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
  $lastImport = $state.last_import
  $importedIds = @($state.imported_ids)
  $pending = $state.pending
} else
{
  $lastImport = $null
  $importedIds = @(Get-ChecklistIds $importDir | Sort-Object -Unique)
  Write-Host "First dated import: $($importedIds.Count) checklists already in 'Import files' count as imported."
}

function Save-State
{
  [ordered]@{
    last_import  = $script:lastImport
    imported_ids = @($script:importedIds | Sort-Object -Unique)
    pending      = $script:pending
  } | ConvertTo-Json | Set-Content -Path $statePath -Encoding UTF8
}

# Convert a downloaded folder. On success the checklists count as imported;
# when paused (or on an error) the folder stays pending for -Resume.
function Convert-Download($folder, $start, $end, $downloadOk)
{
  lua .\ebird_to_artportalen.lua "$folder/" .\out_for_excel.csv
  $luaExit = $LASTEXITCODE
  if ($luaExit -ne 0)
  {
    $script:pending = [ordered]@{ folder = $folder; start = $start; end = $end; download_ok = $downloadOk }
    Save-State
    if ($luaExit -eq $EXIT_PAUSED)
    { Write-Host "Import $start to $end paused. Continue with: .\ebird_artportalen_importer.ps1 -Resume"
    } else
    { Write-Warning "Conversion stopped with an error (exit code $luaExit). Fix it, then continue with -Resume."
    }
    exit $luaExit
  }

  $newIds = @(Get-ChecklistIds $folder)
  $script:importedIds = @($script:importedIds) + $newIds
  if ($downloadOk -and ((-not $script:lastImport) -or $end -gt $script:lastImport))
  { $script:lastImport = $end
  }
  $script:pending = $null
  Save-State
  Write-Host "Imported $($newIds.Count) checklists. Last import date: $script:lastImport"
}

# --- Resume a paused import --------------------------------------------------
if ($Resume)
{
  if ($datesGiven)
  { throw 'Use -Resume on its own; it continues the paused import.'
  }
  if (-not $pending)
  { throw 'There is no paused import to resume.'
  }
  if (-not (Test-Path $pending.folder))
  { throw "The paused import's folder is gone: $($pending.folder)"
  }
  Write-Host "Resuming import $($pending.start) to $($pending.end) ..."
  Convert-Download $pending.folder $pending.start $pending.end $pending.download_ok
  exit 0
}

if ($pending)
{
  throw "Import $($pending.start) to $($pending.end) is paused. Finish it first with: .\ebird_artportalen_importer.ps1 -Resume"
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

# The scraper reads the login from EBIRD_USERNAME / EBIRD_PASSWORD. A shell
# started before they were set doesn't have them, so fall back to the saved
# user environment variables.
foreach ($name in 'EBIRD_USERNAME', 'EBIRD_PASSWORD')
{
  if (-not [Environment]::GetEnvironmentVariable($name))
  {
    [Environment]::SetEnvironmentVariable($name, [Environment]::GetEnvironmentVariable($name, 'User'))
  }
}

Write-Host "Downloading eBird checklists $start to $end ..."
& $python $scraper --start $start --end $end --out $outDir --skip-ids $skipFile
$scraperExit = $LASTEXITCODE
Remove-Item $skipFile -ErrorAction SilentlyContinue
# 0 = all downloaded, 3 = some checklists failed, anything else = no usable run
if ($scraperExit -ne 0 -and $scraperExit -ne 3)
{
  if ((Test-Path $outDir) -and -not (Get-ChildItem $outDir -Filter *.csv -File))
  { Remove-Item $outDir -Recurse -Force
  }
  throw "Download failed (exit code $scraperExit); nothing imported, import state not changed."
}
$downloadOk = ($scraperExit -eq 0)
if (-not $downloadOk)
{
  Write-Warning 'Some checklists failed to download. Converting the rest; the last import date is not moved, so a later -SinceLastImport retries them.'
}

# --- Convert -----------------------------------------------------------------
if (-not ((Test-Path $outDir) -and (Get-ChecklistIds $outDir)))
{
  Write-Host 'No new checklists to import.'
  if (Test-Path $outDir)
  { Remove-Item $outDir -Recurse -Force
  }
  if ($downloadOk -and ((-not $lastImport) -or $end -gt $lastImport))
  { $lastImport = $end
  }
  Save-State
  Write-Host "Last import date: $lastImport"
  exit 0
}

Convert-Download $outDir $start $end $downloadOk
