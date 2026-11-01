# ebird-artportalen-importer

Converts **eBird checklist CSV exports** into rows for the
**Artportalen bulk-import Excel template**, so you don't have to re-enter
observations by hand.

For each observation it:

- maps the eBird species name (e.g. `gråsiska (Acanthis flammea)`) to the
  Artportalen species name (`Gråsiska`)
- maps the eBird location to an Artportalen *Lokalnamn* + *Huvudlokal*
- converts dates (`Jun 18, 2024` → `2024-06-18`) and times (`1:00 PM` → `13:00`)
- computes end time and end date from the checklist duration (a checklist that
  runs past midnight gets the next day as *Slutdatum*)

The result is written to a CSV **and** copied to the clipboard as
tab-separated text (Windows), ready to paste straight into Excel.

## Output columns

| Artnamn | Antal | Lokalnamn | Huvudlokal | Startdatum | Starttid | Slutdatum | Sluttid |
|---------|-------|-----------|------------|------------|----------|-----------|---------|

## Requirements

- [Lua](https://www.lua.org/) 5.1+ (5.4 tested)
- [dkjson](http://dkolf.de/dkjson-lua/), either via LuaRocks
  (`luarocks install dkjson`) or by putting `dkjson.lua` in a `vendor/` folder
  next to the scripts
- Windows + PowerShell for the clipboard copy (the CSV is written on any OS)

## Usage

1. Export your eBird checklists as CSV and put them in a folder, e.g.
   `Import files/`. The CSVs must have these columns:
   `Species, Count, Location, Observation Date, Start Time, Duration`.
2. Run (needs internet access for site suggestions):

   ```powershell
   lua .\ebird_to_artportalen.lua '.\Import files\' .\out_for_excel.csv
   ```

   or use `.\ebird_artportalen_importer.ps1`, which converts everything in
   `Import files\` and also switches the console to UTF-8 so å/ä/ö display and
   input work.
3. Paste the clipboard into the Artportalen import template in Excel.

### Download and convert by date

With [ebird-checklist-scraper](https://github.com/jensalmkvist/ebird-checklist-scraper)
cloned next to this repo (set up with its `.venv`, and `EBIRD_USERNAME` /
`EBIRD_PASSWORD` set), the script can fetch your checklists from eBird first:

```powershell
.\ebird_artportalen_importer.ps1 -StartDate 2026-10-01                       # 1 Oct up to today
.\ebird_artportalen_importer.ps1 -StartDate 2026-10-01 -EndDate 2026-10-15   # a date span
.\ebird_artportalen_importer.ps1 -SinceLastImport                            # last import date up to today
```

The checklists are downloaded to a new folder in `Import files\` and only
those are converted to `out_for_excel.csv` / the clipboard. `import_state.json`
(not in git) remembers the last import date and every checklist ID imported,
so overlapping dates never import a checklist twice; `-SinceLastImport` starts
*on* the last import date to catch checklists entered later that day. On the
first dated run, the `S…_observations.csv` files already in `Import files\`
count as imported. If a download fails, the last import date isn't moved, so
the next `-SinceLastImport` tries again. Use `-ScraperPath` if the scraper is
somewhere else.

### Unknown species

When a species isn't in `species_map.json`, the script asks in the terminal:

```
No mapping for species 'stare (Sturnus vulgaris)'. Enter Artnamn (blank to skip):
```

Your answer is saved immediately; blank writes the row with Artnamn empty.

### Unknown locations: Artportalen site suggestions

When a location isn't in `mappings.json`, the script looks up public
Artportalen sites and lets you pick one by number:

```
No mapping for location 'Länsväg T 771, Grythyttan SE-Örebro län 59.63403, 14.76846'.
Nearest public Artportalen sites to 59.63403, 14.76846:
  1) Gungbromossen  [huvudlokal: Grythyttans sn]  (Hällefors, 346 m)
  2) Bjurtjärnen (Moshyttan)  [huvudlokal: Järnboås sn]  (Nora, 765 m)
  ...
  m) enter Lokalnamn/Huvudlokal manually
  s) skip (leave empty, ask again next run)
  Or type coordinates "lat, lon" or a site name to search again.
Choice:
```

- **With coordinates** the nearest sites are shown with their distance.
  Coordinates come from optional `Latitude`/`Longitude` columns in the CSV, or
  from the `lat, lon` / `(lat,lon)` that eBird puts at the end of personal
  location names.
- **Without coordinates** (eBird hotspots) the site names are searched using
  the parts of the location name, e.g. `Hornborgasjön--Utloppet` searches for
  `Utloppet` and `Hornborgasjön`. Check the region shown, as names repeat
  across Sweden.
- **Your own sites** are listed first. Private Artportalen sites can't be
  looked up without logging in, so these come from the sites you've already
  used in `mappings.json`: those whose Lokalnamn matches the location name
  (`Hemma Bråviksvägen` finds `Bråviksvägen`), then those within 5 km. Their
  coordinates come from the eBird location name or the `lat`/`lon` saved in
  the mapping; the script adds `lat`/`lon` to a mapping whenever it sees the
  location with coordinates.
- At the prompt you can paste coordinates (e.g. from the eBird map) or type a
  site name to search again.
- `s` (or blank) skips the location for the rest of the run; it is asked again
  next time.

Your choice is saved to `mappings.json` immediately, so each location is only
asked once, ever. The lookup needs `curl` (built into Windows 10+) and
internet access; if it fails you can still use `m` to type the names.

It uses the same unauthenticated endpoints as Artportalen's public "Sök fynd"
page (`/ViewSighting/GetSelectableSitesAsMapGeoJsonSites` and
`/ViewSighting/FindSitesByName`). They are not a documented API and may change.

## Mapping files

Both files are plain JSON. You can edit them by hand; keys are kept sorted.

- `species_map.json`: eBird species → Artportalen *Artnamn*

  ```json
  { "stare (Sturnus vulgaris)": "Stare" }
  ```

- `mappings.json`: eBird location → Artportalen site

  ```json
  {
    "Trollskogen": { "lokalnamn": "Trollskogen", "huvudlokal": "Ölands norra udde" }
  }
  ```

If a mapping file contains invalid JSON it is renamed to `<file>.corrupt` and a
fresh empty mapping is started.

## Files

| File | Purpose |
|------|---------|
| `ebird_to_artportalen.lua` | Main conversion script |
| `KeyValueMapper.lua` | Small JSON-backed key → value store used for the mappings |
| `ArtportalenSites.lua` | Looks up public Artportalen sites (nearest / by name) |
| `ebird_artportalen_importer.ps1` | Windows launcher (UTF-8 console + run) |
| `species_map.json`, `mappings.json` | Your accumulated species/location mappings |

## License

See [LICENSE](LICENSE).
