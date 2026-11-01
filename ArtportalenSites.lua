-- ArtportalenSites.lua
-- Look up public Artportalen sites (lokaler), without logging in, using the
-- same endpoints as the public "Sök fynd" page on artportalen.se:
--
--   POST /ViewSighting/GetSelectableSitesAsMapGeoJsonSites
--        public sites inside a map box (EPSG:3857 metres)
--   POST /ViewSighting/FindSitesByName
--        public sites whose name matches a search term
--
-- Both are undocumented and may change. Requests go through `curl`, which is
-- included in Windows 10+ and in macOS/Linux.
--
-- Each site is returned as:
--   { id, lokalnamn, huvudlokal, region, dist_m }
-- where dist_m is only set for coordinate searches.

-- Prefer LuaRocks if available; fall back to project-local vendor
local ok_loader = pcall(require, 'luarocks.loader')
if not ok_loader then
  package.path = package.path .. ';./vendor/?.lua;./vendor/?/init.lua'
end

local dkjson = require 'dkjson'

local M = {}

local BASE_URL = 'https://www.artportalen.se'
local EARTH_RADIUS = 6378137 -- Web Mercator sphere radius (m)

-- Search radii (real metres) tried in order until enough sites are found.
-- The server caps a box at ~300 sites, so start small.
local RADII = { 1000, 3000, 10000, 25000 }

local function is_windows()
  return package.config:sub(1, 1) == '\\'
end

local function trim(s)
  return (tostring(s or ''):gsub('^%s+', ''):gsub('%s+$', ''))
end

-- -------- HTTP --------

local function temp_path()
  local dir = os.getenv 'TEMP' or os.getenv 'TMP' or os.getenv 'TMPDIR'
  if not dir or dir == '' then
    dir = is_windows() and '.' or '/tmp'
  end
  local sep = is_windows() and '\\' or '/'
  return ('%s%sap_request_%d_%d.json'):format(
    dir,
    sep,
    os.time(),
    math.random(100000, 999999)
  )
end

local function shell_quote(s)
  if is_windows() then
    return '"' .. s .. '"'
  end
  return "'" .. s:gsub("'", "'\\''") .. "'"
end

-- POST a JSON body and decode the JSON response.
-- Returns the decoded value, or nil + error message.
local function post_json(path, body)
  local tmp = temp_path()
  local f, ferr = io.open(tmp, 'wb')
  if not f then
    return nil, 'cannot write temp file: ' .. tostring(ferr)
  end
  f:write(dkjson.encode(body))
  f:close()

  local cmd = table.concat({
    'curl -sS -f -m 30 -X POST',
    '-H "X-Requested-With: XMLHttpRequest"',
    '-H "Content-Type: application/json; charset=UTF-8"',
    '--data-binary @' .. shell_quote(tmp),
    BASE_URL .. path,
  }, ' ')

  local p = io.popen(cmd)
  if not p then
    os.remove(tmp)
    return nil, 'could not run curl'
  end
  local out = p:read '*a'
  p:close()
  os.remove(tmp)

  if not out or trim(out) == '' then
    return nil, 'empty response from ' .. path
  end
  local data, _, err = dkjson.decode(out)
  if err or data == nil then
    return nil, 'invalid JSON from ' .. path .. ': ' .. tostring(err)
  end
  return data
end

-- -------- Geometry --------

local function to_mercator(lat, lon)
  local x = math.rad(lon) * EARTH_RADIUS
  local y = math.log(math.tan(math.pi / 4 + math.rad(lat) / 2)) * EARTH_RADIUS
  return x, y
end

-- Squared distance from point p to segment a-b (all in mercator metres)
local function seg_dist2(px, py, ax, ay, bx, by)
  local dx, dy = bx - ax, by - ay
  local len2 = dx * dx + dy * dy
  local t = 0
  if len2 > 0 then
    t = math.max(0, math.min(1, ((px - ax) * dx + (py - ay) * dy) / len2))
  end
  local cx, cy = ax + t * dx - px, ay + t * dy - py
  return cx * cx + cy * cy
end

-- Mercator distance from a point to a polygon (0 if inside).
-- coords is GeoJSON Polygon coordinates: { outer_ring, hole1, ... }
local function polygon_dist(px, py, coords)
  local ring = coords and coords[1]
  if not ring or #ring == 0 then
    return math.huge
  end
  local inside = false
  local best = math.huge
  local j = #ring
  for i = 1, #ring do
    local xi, yi = ring[i][1], ring[i][2]
    local xj, yj = ring[j][1], ring[j][2]
    if (yi > py) ~= (yj > py) and px < (xj - xi) * (py - yi) / (yj - yi) + xi then
      inside = not inside
    end
    best = math.min(best, seg_dist2(px, py, xj, yj, xi, yi))
    j = i
  end
  if inside then
    return 0
  end
  return math.sqrt(best)
end

local function feature_dist(px, py, feature)
  local g = feature.geometry or {}
  if g.type == 'Point' then
    local dx, dy = g.coordinates[1] - px, g.coordinates[2] - py
    return math.sqrt(dx * dx + dy * dy)
  elseif g.type == 'Polygon' then
    return polygon_dist(px, py, g.coordinates)
  elseif g.type == 'MultiPolygon' then
    local best = math.huge
    for _, poly in ipairs(g.coordinates or {}) do
      best = math.min(best, polygon_dist(px, py, poly))
    end
    return best
  end
  return math.huge
end

-- -------- Site name handling --------

local function ends_with(s, suffix)
  return suffix ~= '' and s:sub(-#suffix) == suffix
end

-- Site name search. Returns list of { id, lokalnamn, huvudlokal, region }.
function M.search_by_name(term)
  term = trim(term)
  if term == '' then
    return {}
  end
  local data, err = post_json('/ViewSighting/FindSitesByName', {
    Term = term,
    IncludePublicBirdSites = true,
    IncludeOthersPrivateSites = false,
  })
  if not data then
    return nil, err
  end
  local out = {}
  for _, s in ipairs(data) do
    if s.Id and s.Id ~= 0 and s.IsPublicSite ~= false then
      out[#out + 1] = {
        id = s.Id,
        lokalnamn = trim(s.Name),
        huvudlokal = trim((tostring(s.ParentName or ''):gsub('^%s*,%s*', ''))),
        region = trim(s.Region),
      }
    end
  end
  return out
end

-- The map endpoint gives the full name ("Utloppet, Hornborgasjön") plus a
-- parentId. Work out the parent name so we can split it into
-- lokalnamn / huvudlokal.
local function split_site_name(props, by_id)
  local name = trim(props.siteName)
  local parent_id = tonumber(props.parentId) or 0
  if parent_id == 0 then
    return name, ''
  end

  -- 1) Parent is in the same result set
  local parent = by_id[parent_id]
  if parent then
    local pname = trim(parent.properties.siteName)
    if ends_with(name, ', ' .. pname) then
      return name:sub(1, #name - #pname - 2), pname
    end
  end

  -- 2) Ask the name search, which returns name and parent separately
  local first = name:match '^([^,]+)' or name
  local hits = M.search_by_name(first)
  for _, h in ipairs(hits or {}) do
    if h.id == props.siteId then
      return h.lokalnamn, h.huvudlokal
    end
  end

  -- 3) Best guess: last comma-separated part is the parent
  local head, tail = name:match '^(.*),%s*([^,]+)$'
  if head then
    return trim(head), trim(tail)
  end
  return name, ''
end

local function fetch_box(px, py, half_side)
  local bbox = ('%d,%d,%d,%d'):format(
    math.floor(px - half_side),
    math.floor(py - half_side),
    math.ceil(px + half_side),
    math.ceil(py + half_side)
  )
  local data, err = post_json('/ViewSighting/GetSelectableSitesAsMapGeoJsonSites', {
    zoomLevel = 14,
    bbox = bbox,
    userId = 0,
    speciesGroupId = 8, -- the value the public page uses; others return nothing
    projectIds = '',
    showOnlyProjects = false,
    showCluster = false,
    breakYear = 0,
    showValidated = false,
  })
  if not data then
    return nil, err
  end
  local features = {}
  for _, layer in ipairs { 'points', 'polygons' } do
    local fc = data[layer]
    for _, feat in ipairs(fc and fc.features or {}) do
      if feat.properties and feat.properties.siteId then
        features[#features + 1] = feat
      end
    end
  end
  return features
end

-- Nearest public sites to lat/lon.
-- Returns up to `limit` sites sorted by distance, or nil + error message.
function M.nearest(lat, lon, limit)
  limit = limit or 5
  local px, py = to_mercator(lat, lon)
  -- Mercator stretches distances by 1/cos(lat)
  local scale = math.cos(math.rad(lat))

  local features, err
  local max_dist
  for _, r in ipairs(RADII) do
    features, err = fetch_box(px, py, r / scale)
    if not features then
      return nil, err
    end
    max_dist = r
    -- Count sites actually within the radius (box corners are farther away)
    local n = 0
    for _, feat in ipairs(features) do
      if feature_dist(px, py, feat) * scale <= r then
        n = n + 1
      end
    end
    if n >= limit then
      break
    end
  end

  local by_id = {}
  for _, feat in ipairs(features) do
    by_id[feat.properties.siteId] = feat
  end

  local ranked = {}
  for _, feat in ipairs(features) do
    local d = feature_dist(px, py, feat) * scale
    if d <= max_dist * 1.5 then
      ranked[#ranked + 1] = { feat = feat, dist = d }
    end
  end
  table.sort(ranked, function(a, b)
    return a.dist < b.dist
  end)

  local out = {}
  for i = 1, math.min(limit, #ranked) do
    local props = ranked[i].feat.properties
    local lokal, huvud = split_site_name(props, by_id)
    out[#out + 1] = {
      id = props.siteId,
      lokalnamn = lokal,
      huvudlokal = huvud,
      region = trim(props.siteAreaName),
      dist_m = ranked[i].dist,
    }
  end
  return out
end

-- Parse "58.32, 13.55" / "58.32 13.55" into lat, lon (decimal degrees).
function M.parse_coords(s)
  local lat, lon = tostring(s or ''):match '(%-?%d+%.%d+)%s*[,;%s]%s*(%-?%d+%.%d+)'
  lat, lon = tonumber(lat), tonumber(lon)
  if lat and lon and math.abs(lat) <= 90 and math.abs(lon) <= 180 then
    return lat, lon
  end
  return nil
end

return M
