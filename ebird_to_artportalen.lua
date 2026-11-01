-- ebird_to_artportalen.lua
-- Transform eBird checklist CSVs into columns expected by the Artportalen
-- bulk-import Excel template:
-- Artnamn, Antal, Lokalnamn, Huvudlokal, Startdatum, Starttid, Slutdatum, Sluttid
--
-- Uses two JSON-backed mappers (KeyValueMapper.lua) for:
--   Species  -> Artnamn                  (species_map.json)
--   Location -> {lokalnamn, huvudlokal}  (mappings.json)
--
-- Assumes:
--   - CSV columns: Species, Count, Location, Observation Date, Start Time, Duration
--   - Duration as minutes ("90"), "HH:MM", or text ("1 hour(s), 26 minute(s)").
--   - If the checklist runs past midnight, Slutdatum is moved forward accordingly.
--   - Unknown locations get numbered suggestions of public Artportalen sites
--     (ArtportalenSites.lua), nearest first when coordinates are known.
--
-- Usage:
--   lua ebird_to_artportalen.lua <input_folder> <output.csv>
--
-- Tip (Windows/PowerShell, for å/ä/ö): run once per session before lua:
--   chcp 65001 > $null
--   [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

-- If you vendored dkjson.lua next to KeyValueMapper.lua, ensure package.path for vendor (uncomment if needed):
-- package.path = package.path .. ";./vendor/?.lua;./vendor/?/init.lua"

-- Toggle this to enable/disable debug logs
local DEBUG = false

local function dprintf(fmt, ...)
  if DEBUG then
    io.stderr:write('[debug] ' .. string.format(fmt, ...) .. '\n')
  end
end

local Mapper = require 'KeyValueMapper' -- generic key->value JSON mapper
local Sites = require 'ArtportalenSites' -- public Artportalen site lookup

-- -------- Utilities --------

-- Detect platform
local function is_windows()
  return package.config:sub(1, 1) == '\\'
end

-- Return true if path looks like a directory
local function path_is_dir(path)
  -- Cheap heuristic: if a trailing slash exists, treat as dir
  if path:match '[\\/]$' then
    return true
  end
  -- Try to open as file; if it fails, might be dir (best-effort without lfs)
  local f = io.open(path, 'rb')
  if f then
    f:close()
    return false
  end
  -- On Windows, we can query "dir" (no output parsing here)
  return true
end

-- List *.csv files in a directory (non-recursive). Works on Windows/macOS/Linux.
local function list_csv_files(dir)
  local files = {}
  local cmd
  if is_windows() then
    -- /b: bare format, /a:-d exclude directories
    cmd = string.format('cmd /C dir /b /a:-d "%s\\*.csv"', dir)
  else
    cmd = string.format('sh -lc \'ls -1 "%s"/*.csv 2>/dev/null\'', dir)
  end
  local p = io.popen(cmd)
  if not p then
    return files
  end
  for line in p:lines() do
    line = tostring(line or ''):gsub('%s+$', '')
    if line ~= '' then
      -- Windows returns just the filename; prepend dir
      if is_windows() and not line:match '^%a:' and not line:match '^[/\\]' then
        files[#files + 1] = dir .. '\\' .. line
      else
        files[#files + 1] = line
      end
    end
  end
  p:close()
  return files
end

local function trim(s)
  return (tostring(s or ''):gsub('^%s+', ''):gsub('%s+$', ''))
end

-- Replace no-break spaces with plain spaces. Newer eBird exports write times
-- as "7:55<U+202F>AM" (narrow no-break space), which %s does not match.
local function normalize_spaces(s)
  return (tostring(s or ''):gsub('\226\128\175', ' '):gsub('\194\160', ' '))
end

-- Normalize "H:MM AM/PM" (also supports "HH:MM:SS AM/PM") to 24-hour "HH:MM"
local function to_24h_time(s)
  s = trim(normalize_spaces(s))
  if s == '' then
    return ''
  end

  -- capture optional seconds and AM/PM (case/space tolerant)
  local h, m, sec, ap = s:match '^(%d%d?):(%d%d):(%d%d)%s*([AaPp][Mm])$'
  if h then
    h, m, sec = tonumber(h), tonumber(m), tonumber(sec)
    ap = ap:upper()
    if ap == 'AM' then
      if h == 12 then
        h = 0
      end
    else -- PM
      if h ~= 12 then
        h = h + 12
      end
    end
    return string.format('%02d:%02d', h, m) -- drop seconds to align with your format
  end

  h, m, ap = s:match '^(%d%d?):(%d%d)%s*([AaPp][Mm])$'
  if h then
    h, m = tonumber(h), tonumber(m)
    ap = ap:upper()
    if ap == 'AM' then
      if h == 12 then
        h = 0
      end
    else -- PM
      if h ~= 12 then
        h = h + 12
      end
    end
    return string.format('%02d:%02d', h, m)
  end

  -- If already 24h ("HH:MM" or "HH:MM:SS") keep your old path
  local hh, mm, ss = s:match '^(%d%d?):(%d%d):(%d%d)$'
  if hh then
    return string.format('%02d:%02d', tonumber(hh), tonumber(mm))
  end
  hh, mm = s:match '^(%d%d?):(%d%d)$'
  if hh and mm then
    return string.format('%02d:%02d', tonumber(hh), tonumber(mm))
  end

  -- Fallback: return as-is (your later logic will handle/log)
  return s
end

-- Convert "Jun 18, 2024" -> "2024-06-18"
local function normalize_date_mmddyyyy(s)
  local mon, day, year = s:match '^(%a+)%s+(%d%d?),%s*(%d%d%d%d)$'
  if not (mon and day and year) then
    return s -- return unchanged if pattern doesn't match
  end

  local months = {
    Jan = '01',
    Feb = '02',
    Mar = '03',
    Apr = '04',
    May = '05',
    Jun = '06',
    Jul = '07',
    Aug = '08',
    Sep = '09',
    Oct = '10',
    Nov = '11',
    Dec = '12',
  }

  local mm = months[mon:sub(1, 3)]
  return string.format('%s-%s-%02d', year, mm, tonumber(day))
end

-- Robust CSV parser for one line: handles quotes, commas, escaped quotes ("")
local function parse_csv_line(line)
  local res = {}
  local i, len = 1, #line
  while i <= len do
    if line:sub(i, i) == '"' then
      i = i + 1
      local buf = {}
      while i <= len do
        local c = line:sub(i, i)
        if c == '"' then
          if i + 1 <= len and line:sub(i + 1, i + 1) == '"' then
            buf[#buf + 1] = '"'
            i = i + 2
          else
            i = i + 1
            break
          end
        else
          buf[#buf + 1] = c
          i = i + 1
        end
      end
      if i <= len and line:sub(i, i) == ',' then
        i = i + 1
      end
      res[#res + 1] = table.concat(buf)
    else
      local j = line:find(',', i, true)
      if j then
        res[#res + 1] = trim(line:sub(i, j - 1))
        i = j + 1
      else
        res[#res + 1] = trim(line:sub(i))
        break
      end
    end
  end
  return res
end

local function csv_escape(field)
  if field == nil then
    return ''
  end
  field = tostring(field)
  if field:find '[,"\r\n]' then
    field = '"' .. field:gsub('"', '""') .. '"'
  end
  return field
end

local function write_csv_row(fh, fields)
  local out = {}
  for i = 1, #fields do
    out[i] = csv_escape(fields[i])
  end
  fh:write(table.concat(out, ',') .. '\n')
end

-- Parse "HH:MM" or "HH:MM:SS"
local function parse_time_hms(s)
  s = trim(s)
  dprintf("parse_time_hms trimmed string: '%s'", s)

  local h, m, sec = s:match '^(%d%d?):(%d%d):(%d%d)$'
  dprintf("parse_time_hms matched h='%s' m='%s', sec='%s'", h, m, sec)

  if h then
    return tonumber(h), tonumber(m), tonumber(sec)
  end
  h, m = s:match '^(%d%d?):(%d%d)$'
  if h then
    return tonumber(h), tonumber(m), 0
  end
  return nil, nil, nil
end

-- Add minutes to a time string "HH:MM" (or "HH:MM:SS").
-- Returns "HH:MM" (wrapped after 24h) and the number of days crossed,
-- e.g. ("23:30", 60) -> "00:30", 1
local function add_minutes_to_time(start_str, minutes)
  local h, m = parse_time_hms(start_str)
  if not h then
    dprintf("add_minutes_to_string start_str='%s'", start_str)
    return start_str, 0
  end
  dprintf("add_minutes_to_string h='%d', m='%d'", h, m)
  minutes = tonumber(minutes) or 0
  local total = h * 60 + m + minutes
  local days = math.floor(total / (24 * 60))
  total = total % (24 * 60)
  local out_h = math.floor(total / 60)
  local out_m = total % 60
  return string.format('%02d:%02d', out_h, out_m), days
end

-- Add whole days to a "YYYY-MM-DD" date. Returns input unchanged if it
-- doesn't look like an ISO date. os.time normalizes day overflow
-- (e.g. Dec 32 -> Jan 1); noon avoids DST edge cases.
local function add_days_to_date(date_str, days)
  if not days or days == 0 then
    return date_str
  end
  local y, mo, d = tostring(date_str or ''):match '^(%d%d%d%d)-(%d%d)-(%d%d)$'
  if not y then
    return date_str
  end
  local t = os.time {
    year = tonumber(y),
    month = tonumber(mo),
    day = tonumber(d) + days,
    hour = 12,
  }
  return os.date('%Y-%m-%d', t)
end

-- Accept:
--  - plain minutes: "15", "90"
--  - HH:MM: "1:30"
--  - text: "7 minute(s)", "1 hour(s), 26 minute(s)", "2 hours, 5 minutes"
local function parse_duration_to_minutes(s)
  local raw = normalize_spaces(s)
  dprintf("raw='%s'", raw)

  s = trim(raw)
  if s == '' then
    dprintf '-> empty string, minutes=0'
    return 0
  end

  -- numeric minutes
  if s:match '^%d+$' or s:match '^%d+%.%d+$' then
    return math.floor(tonumber(s) or 0)
  end

  -- HH:MM
  local h, m = s:match '^(%d%d?):(%d%d)$'
  if h and m then
    return tonumber(h) * 60 + tonumber(m)
  end

  -- text formats
  local t = s:lower()
  -- normalize: remove "(s)", commas → space
  t = t:gsub('%(s%)', '')
  t = t:gsub(',', ' ')

  dprintf("normalized text='%s'", t)

  local hours = 0
  local minutes = 0
  local total = 0

  -- capture "N hour" or "N hours"
  local hnum = t:match '(%d+)%s*hour%a*'
  if hnum then
    hours = tonumber(hnum) or 0
  end

  -- capture "N minute" or "N minutes"
  local mnum = t:match '(%d+)%s*min%a*'
  if mnum then
    minutes = tonumber(mnum) or 0
  end

  total = hours * 60 + minutes

  dprintf(
    '-> text matched: hours=%d, minutes=%d, total=%d',
    hours,
    minutes,
    total
  )

  return total
end

-- Copy UTF‑8 text to Windows clipboard using PowerShell (works with å/ä/ö)
local function copy_to_clipboard_utf8(text)
  local tmp = os.tmpname() .. '.txt'
  local f = assert(io.open(tmp, 'wb'))
  f:write(text) -- write UTF-8 bytes
  f:close()
  local cmd = ([[powershell -NoProfile -Command "Get-Content -LiteralPath '%s' -Raw -Encoding UTF8 | Set-Clipboard"]]):format(
    tmp:gsub("'", "''")
  )
  os.execute(cmd)
  os.remove(tmp)
end

-- Read a CSV file and return a single TSV blob (header + rows), using parse_csv_line()
local function csv_file_to_tsv_blob(csv_path)
  local fin = assert(
    io.open(csv_path, 'r'),
    'Cannot open output CSV to convert: ' .. csv_path
  )
  local lines = {}
  for line in fin:lines() do
    -- Parse CSV line (handles quoted fields / commas / escaped quotes)
    local fields = parse_csv_line(line or '')
    -- Join with tabs
    lines[#lines + 1] = table.concat(fields, '\t')
  end
  fin:close()
  return table.concat(lines, '\n')
end

-- -------- Unknown locations: suggest Artportalen sites --------

-- Number of site suggestions to show for an unknown location
local MAX_SUGGESTIONS = 5
-- Your own sites (already used in mappings.json) shown above the public ones
local MAX_OWN_SUGGESTIONS = 3
local OWN_SITE_RADIUS_M = 5000

-- string.lower only handles ASCII; also lowercase Å/Ä/Ö
local function lower_sv(s)
  return (tostring(s or ''):lower():gsub('Å', 'å'):gsub('Ä', 'ä'):gsub('Ö', 'ö'))
end

-- Coordinates at the end of eBird personal location names:
-- "Lantmannavägen 4, Växjö SE-G (56.8681,14.8354)" or "... 58.32, 13.55"
local COORDS_SUFFIX = '%(?(%-?%d+%.%d+%s*,%s*%-?%d+%.%d+)%)?%s*$'

local function coords_from_name(name)
  return Sites.parse_coords(tostring(name or ''):match(COORDS_SUFFIX))
end

local function strip_coords(name)
  return (tostring(name or ''):gsub('%s*' .. COORDS_SUFFIX, ''))
end

-- Great-circle distance in metres
local function distance_m(lat1, lon1, lat2, lon2)
  local r = math.rad
  local dlat, dlon = r(lat2 - lat1), r(lon2 - lon1)
  local a = math.sin(dlat / 2) ^ 2
    + math.cos(r(lat1)) * math.cos(r(lat2)) * math.sin(dlon / 2) ^ 2
  return 6371000 * 2 * math.asin(math.sqrt(a))
end

local function format_distance(m)
  if not m then
    return ''
  end
  if m < 1000 then
    return ('%d m'):format(math.floor(m + 0.5))
  end
  return ('%.1f km'):format(m / 1000)
end

-- Search terms for a location name without coordinates.
-- eBird hotspots look like "Hornborgasjön--Utloppet" or "Kvarnberget, Hjälstaviken".
local function name_search_terms(location)
  local terms, seen = {}, {}
  local function add(t)
    t = trim(t)
    if #t >= 3 and not seen[t:lower()] then
      seen[t:lower()] = true
      terms[#terms + 1] = t
    end
  end
  local parts = {}
  for part in (strip_coords(location) .. '--'):gmatch '(.-)%-%-' do
    for sub in (part .. ','):gmatch '([^,]*),' do
      parts[#parts + 1] = sub
    end
  end
  -- Most specific part (last) first
  for i = #parts, 1, -1 do
    add(parts[i])
  end
  return terms
end

-- Name search across all parts of the location; sites matching more parts
-- (e.g. both "Utloppet" and "Hornborgasjön") rank first.
local function suggest_by_name(location)
  local terms = name_search_terms(location)
  local found, order = {}, {}
  for _, term in ipairs(terms) do
    local hits = Sites.search_by_name(term)
    for _, s in ipairs(hits or {}) do
      if not found[s.id] then
        found[s.id] = s
        order[#order + 1] = s
        s.rank = #order
      end
    end
  end
  for _, s in ipairs(order) do
    local text = (s.lokalnamn .. ' ' .. s.huvudlokal):lower()
    s.score = 0
    for _, term in ipairs(terms) do
      if text:find(term:lower(), 1, true) then
        s.score = s.score + 1
      end
    end
  end
  table.sort(order, function(a, b)
    if a.score ~= b.score then
      return a.score > b.score
    end
    return a.rank < b.rank
  end)
  local out = {}
  for i = 1, math.min(MAX_SUGGESTIONS, #order) do
    out[i] = order[i]
  end
  return out
end

-- Sites you have already mapped to, one per Lokalnamn/Huvudlokal. These
-- include your private Artportalen sites, which the public lookup can't see.
-- Coordinates come from the saved lat/lon or the eBird location name.
local function own_sites(location_mapper)
  local by_key, list = {}, {}
  for ebird_name, v in location_mapper:each() do
    if type(v) == 'table' and trim(v.lokalnamn) ~= '' then
      local key = v.lokalnamn .. '\0' .. (v.huvudlokal or '')
      local s = by_key[key]
      if not s then
        s = { lokalnamn = v.lokalnamn, huvudlokal = v.huvudlokal or '' }
        by_key[key] = s
        list[#list + 1] = s
      end
      if not s.lat then
        local lat, lon = tonumber(v.lat), tonumber(v.lon)
        if not (lat and lon) then
          lat, lon = coords_from_name(ebird_name)
        end
        s.lat, s.lon = lat, lon
      end
    end
  end
  return list
end

-- True if a search term and the site's name contain one another
-- ("Hemma Bråviksvägen" matches your site "Bråviksvägen" and vice versa).
-- Huvudlokal is only searched when asked: a hotspot like "Tåkern--Glänås"
-- would otherwise list every site you have under Tåkern.
local function name_matches(site, terms, with_huvudlokal)
  local lokal = lower_sv(site.lokalnamn)
  local text = with_huvudlokal and (lokal .. ' ' .. lower_sv(site.huvudlokal)) or lokal
  for _, t in ipairs(terms) do
    t = lower_sv(t)
    if text:find(t, 1, true) or (#lokal >= 3 and t:find(lokal, 1, true)) then
      return true
    end
  end
  return false
end

-- Own sites whose name matches first, then own sites near lat/lon.
local function suggest_own(sites, terms, lat, lon, with_huvudlokal)
  local named, near = {}, {}
  for _, s in ipairs(sites) do
    local d = lat and lon and s.lat and s.lon and distance_m(lat, lon, s.lat, s.lon)
    if name_matches(s, terms, with_huvudlokal) then
      named[#named + 1] = { site = s, dist = d }
    elseif d and d <= OWN_SITE_RADIUS_M then
      near[#near + 1] = { site = s, dist = d }
    end
  end
  table.sort(named, function(a, b)
    return a.site.lokalnamn < b.site.lokalnamn
  end)
  table.sort(near, function(a, b)
    return a.dist < b.dist
  end)
  local out = {}
  for _, group in ipairs { named, near } do
    for _, e in ipairs(group) do
      if #out < MAX_OWN_SUGGESTIONS then
        out[#out + 1] = {
          lokalnamn = e.site.lokalnamn,
          huvudlokal = e.site.huvudlokal,
          dist_m = e.dist,
        }
      end
    end
  end
  return out
end

-- Exit code when you pause at a question; ebird_artportalen_importer.ps1
-- uses it to keep the import open for -Resume.
local EXIT_PAUSED = 4

-- Stop at a question. Answers given so far are already saved in the mapping
-- files, and the output CSV is only written at the end, so nothing is lost.
local function pause()
  print ''
  print 'Paused. Your answers so far are saved; the output CSV was not changed.'
  print 'Run the same command again to continue (with the .ps1: -Resume).'
  os.exit(EXIT_PAUSED)
end

local function print_site(i, s)
  local extra = {}
  if s.region and s.region ~= '' then
    extra[#extra + 1] = s.region
  end
  if s.dist_m then
    extra[#extra + 1] = format_distance(s.dist_m)
  end
  print(
    ('  %d) %s%s%s'):format(
      i,
      s.lokalnamn,
      s.huvudlokal ~= '' and ('  [huvudlokal: ' .. s.huvudlokal .. ']') or '',
      #extra > 0 and ('  (' .. table.concat(extra, ', ') .. ')') or ''
    )
  )
end

-- Ask the user how to map an unknown eBird location.
-- Returns {lokalnamn, huvudlokal, lat, lon} to save, or nil to skip.
local function choose_location(location, lat, lon, location_mapper)
  print ''
  print(("No mapping for location '%s'."):format(location))

  local own = own_sites(location_mapper)
  local search_text = nil -- free-text name search typed by the user
  while true do
    local sites, err, heading
    if search_text then
      sites, err = Sites.search_by_name(search_text)
      heading = ("Public Artportalen sites matching '%s':"):format(search_text)
      if sites and #sites > MAX_SUGGESTIONS * 2 then
        for i = #sites, MAX_SUGGESTIONS * 2 + 1, -1 do
          sites[i] = nil
        end
      end
    elseif lat and lon then
      sites, err = Sites.nearest(lat, lon, MAX_SUGGESTIONS)
      heading = ('Nearest public Artportalen sites to %.5f, %.5f:'):format(lat, lon)
    else
      sites, err = suggest_by_name(location)
      heading = 'Public Artportalen sites with a similar name (no coordinates known):'
    end

    local mine = suggest_own(
      own,
      search_text and { search_text } or name_search_terms(location),
      lat,
      lon,
      search_text ~= nil
    )
    local choices = {}
    if #mine > 0 then
      print 'Your own sites (already used in mappings.json):'
      for _, s in ipairs(mine) do
        choices[#choices + 1] = s
        print_site(#choices, s)
      end
    end

    if not sites then
      print('  Could not look up Artportalen sites: ' .. tostring(err))
    elseif #sites == 0 then
      print '  No matching public Artportalen sites found.'
    else
      print(heading)
      for _, s in ipairs(sites) do
        local dup = false
        for _, m in ipairs(mine) do
          dup = dup or (m.lokalnamn == s.lokalnamn and m.huvudlokal == s.huvudlokal)
        end
        if not dup then
          choices[#choices + 1] = s
          print_site(#choices, s)
        end
      end
    end
    print '  m) enter Lokalnamn/Huvudlokal manually'
    print '  s) skip (leave empty, ask again next run)'
    print '  p) pause (stop here, continue later)'
    print '  Or type coordinates "lat, lon" or a site name to search again.'
    io.write 'Choice: '
    local answer = trim(io.read '*l' or '')
    local lower = answer:lower()

    local n = tonumber(answer)
    if n and choices[n] then
      return {
        lokalnamn = choices[n].lokalnamn,
        huvudlokal = choices[n].huvudlokal,
        lat = lat,
        lon = lon,
      }
    elseif lower == '' or lower == 's' then
      return nil
    elseif lower == 'p' then
      pause()
    elseif lower == 'm' then
      io.write(("Enter Lokalnamn for '%s' (blank to skip): "):format(location))
      local ln = trim(io.read '*l' or '')
      io.write(("Enter Huvudlokal for '%s' (blank for none): "):format(location))
      local hl = trim(io.read '*l' or '')
      if ln == '' and hl == '' then
        return nil
      end
      return { lokalnamn = ln, huvudlokal = hl, lat = lat, lon = lon }
    elseif n then
      print '  Not one of the listed numbers.'
    else
      local clat, clon = Sites.parse_coords(answer)
      if clat then
        lat, lon, search_text = clat, clon, nil
      else
        search_text = answer
      end
    end
  end
end

-- -------- Main --------

local function main()
  local input_dir = arg[1]
  local output = arg[2]
  if not (input_dir and output) then
    io.stderr:write 'Usage: lua ebird_to_artportalen.lua <input_folder> <output.csv>\n'
    os.exit(2)
  end

  -- Enumerate CSV files
  local inputs = list_csv_files(input_dir)
  if #inputs == 0 then
    io.stderr:write(('No CSV files found in directory: %s\n'):format(input_dir))
    os.exit(1)
  end

  -- Mappers
  local species_mapper = Mapper.new 'species_map.json' -- Species -> Artnamn (string)
  local location_mapper =
    Mapper.new('mappings.json', { value_key_for_reverse = 'lokalnamn' }) -- Location -> {lokalnamn, huvudlokal}

  -- Stats
  local used_s, added_s = 0, 0
  local used_l, added_l = 0, 0

  -- Locations skipped this run (not asked again for every row)
  local skipped_locations = {}

  -- Output rows, written at the end so a pause leaves the old output alone
  local out_rows = {}

  -- Process each CSV file in the directory
  for _, in_csv in ipairs(inputs) do
    dprintf('Processing: %s', in_csv)

    local fin =
      assert(io.open(in_csv, 'r'), 'Cannot open input CSV: ' .. in_csv)
    local header_line = fin:read '*l'
    if not header_line then
      io.stderr:write('Empty CSV: ' .. in_csv .. '\n')
      fin:close()
      goto continue
    end
    local header = parse_csv_line(header_line)

    -- Build header index map
    local idx = {}
    for i, name in ipairs(header) do
      idx[trim(name)] = i
    end

    -- Required columns
    local function need(col)
      local i = idx[col]
      if not i then
        io.stderr:write(
          ("[%s] Missing column '%s'. Found: %s\n"):format(
            in_csv,
            col,
            table.concat(header, ', ')
          )
        )
        os.exit(1)
      end
      return i
    end

    local iSpecies = need 'Species'
    local iCount = need 'Count'
    local iLocation = need 'Location'
    local iDate = need 'Observation Date'
    local iStart = need 'Start Time'
    local iDur = need 'Duration'
    -- Optional coordinate columns (not in eBird's own export)
    local iLat = idx['Latitude']
    local iLon = idx['Longitude']

    -- Rows
    for line in fin:lines() do
      if line ~= '' then
        local f = parse_csv_line(line)

        local species = trim(f[iSpecies] or '')
        local count = trim(f[iCount] or '')
        local location = trim(f[iLocation] or '')
        local startdate = trim(f[iDate] or '')
        startdate = normalize_date_mmddyyyy(startdate)
        local starttime = trim(f[iStart] or '')
        local duration = trim(f[iDur] or '')

        -- If you have AM/PM normalization, keep it here
        dprintf("starttime from csv (raw)='%s'", starttime)
        starttime = to_24h_time(starttime)
        dprintf("starttime normalized to 24h='%s'", starttime)

        -- Species -> Artnamn
        local art = species ~= '' and species_mapper:get(species) or nil
        if not art and species ~= '' then
          io.write(
            ("No mapping for species '%s'. Enter Artnamn (blank to skip, p to pause): "):format(
              species
            )
          )
          local val = trim(io.read '*l' or '')
          if val:lower() == 'p' then
            pause()
          elseif val ~= '' then
            local ok, err = species_mapper:add(species, val)
            if ok then
              art = val
              added_s = added_s + 1
            else
              io.stderr:write('Species add failed: ' .. err .. '\n')
            end
          end
        else
          if art then
            used_s = used_s + 1
          end
        end

        -- Location -> {lokalnamn, huvudlokal}
        local loc_obj = location ~= '' and location_mapper:get(location) or nil

        -- Coordinates: optional Latitude/Longitude columns, else the
        -- "lat, lon" at the end of eBird personal location names
        local lat = iLat and tonumber(f[iLat])
        local lon = iLon and tonumber(f[iLon])
        if not (lat and lon) then
          lat, lon = coords_from_name(location)
        end

        if not loc_obj and location ~= '' and not skipped_locations[location] then
          local choice = choose_location(location, lat, lon, location_mapper)
          if choice then
            local ok, err = location_mapper:add(location, choice)
            if ok then
              loc_obj = choice
              added_l = added_l + 1
              print(
                ("Saved '%s' -> %s / %s"):format(
                  location,
                  choice.lokalnamn,
                  choice.huvudlokal
                )
              )
            else
              io.stderr:write('Location add failed: ' .. tostring(err) .. '\n')
            end
          else
            skipped_locations[location] = true
          end
        elseif loc_obj then
          -- Backward-compat: if old mappings.json had a plain string, migrate it.
          if type(loc_obj) == 'string' then
            loc_obj = { lokalnamn = loc_obj, huvudlokal = '' }
            local ok, err = location_mapper:set(location, loc_obj)
            if not ok then
              io.stderr:write('Location migrate failed: ' .. tostring(err) .. '\n')
            end
          end
          -- Remember coordinates so this site can be suggested near new locations
          if lat and lon and loc_obj.lat == nil then
            loc_obj.lat, loc_obj.lon = lat, lon
            local ok, err = location_mapper:set(location, loc_obj)
            if not ok then
              io.stderr:write('Location update failed: ' .. tostring(err) .. '\n')
            end
          end
          used_l = used_l + 1
        end

        local lokal = (loc_obj and loc_obj.lokalnamn) or ''
        local huvudlokal = (loc_obj and loc_obj.huvudlokal) or ''

        -- Time math
        local mins = parse_duration_to_minutes(duration)
        local sluttid, days_crossed = '', 0
        if starttime ~= '' then
          sluttid, days_crossed = add_minutes_to_time(starttime, mins)
        end
        local slutdatum = add_days_to_date(startdate, days_crossed)

        -- Output row
        out_rows[#out_rows + 1] = {
          art or '',
          count or '',
          lokal or '',
          huvudlokal or '',
          startdate or '',
          starttime or '',
          slutdatum or '',
          sluttid or '',
        }
      end
    end

    fin:close()
    ::continue::
  end

  local fout = assert(io.open(output, 'w'), 'Cannot open output CSV')
  write_csv_row(fout, {
    'Artnamn',
    'Antal',
    'Lokalnamn',
    'Huvudlokal',
    'Startdatum',
    'Starttid',
    'Slutdatum',
    'Sluttid',
  })
  for _, row in ipairs(out_rows) do
    write_csv_row(fout, row)
  end
  fout:close()

  -- Convert the just-written output.csv to TSV and copy to clipboard (includes header)
  local tsv_blob = csv_file_to_tsv_blob(output)
  copy_to_clipboard_utf8(tsv_blob)
  print 'Copied tab-separated output (including header) to clipboard.'

  print(
    ('Species mappings used: %d, added: %d (species_map.json)'):format(
      used_s,
      added_s
    )
  )
  print(
    ('Location mappings used: %d, added: %d (mappings.json)'):format(
      used_l,
      added_l
    )
  )
  print(('Wrote output CSV: %s'):format(output))
  print 'Columns: Artnamn, Antal, Lokalnamn, Huvudlokal, Startdatum, Starttid, Slutdatum, Sluttid'
end

main()
