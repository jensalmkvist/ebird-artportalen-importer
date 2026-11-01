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
2. Run:

   ```powershell
   lua .\ebird_to_artportalen.lua '.\Import files\' .\out_for_excel.csv
   ```

   or use `ebird_artportalen_importer.ps1`, which also switches the console to
   UTF-8 so å/ä/ö display and input work. (Edit the input path in the script to
   match your machine.)
3. Paste the clipboard into the Artportalen import template in Excel.

### Unknown species and locations

When a species or location isn't in the mapping files, the script asks you in
the terminal:

```
No mapping for species 'stare (Sturnus vulgaris)'. Enter Artnamn (blank to skip):
No mapping for location 'Trollskogen'. Enter Lokalnamn (blank to skip):
Enter Huvudlokal for 'Trollskogen' (blank to skip):
```

Your answers are saved immediately, so each species/location is only asked
once, ever. Leaving the answer blank writes the row with that field empty.

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
| `ebird_artportalen_importer.ps1` | Windows launcher (UTF-8 console + run) |
| `species_map.json`, `mappings.json` | Your accumulated species/location mappings |

## License

See [LICENSE](LICENSE).
