-- KeyValueMapper.lua
-- Standalone JSON-backed mapper: csv_key -> excel_value
-- Requires: dkjson (pure Lua JSON). Install via: luarocks install dkjson

-- Prefer LuaRocks if available; fall back to project-local vendor
local ok_loader = pcall(require, 'luarocks.loader')
if not ok_loader then
  package.path = package.path .. ';./vendor/?.lua;./vendor/?/init.lua'
end

local dkjson = require 'dkjson'

local KeyValueMapper = {}
KeyValueMapper.__index = KeyValueMapper

-- --- helpers -----------------------------------------------------

local function norm_str(s)
  -- Normalize inputs to trimmed strings
  s = tostring(s or '')
  s = s:gsub('^%s+', ''):gsub('%s+$', '')
  return s
end

local function file_exists(path)
  local f = io.open(path, 'r')
  if f then
    f:close()
    return true
  end
  return false
end

local function read_all(path)
  local f, err = io.open(path, 'r')
  if not f then
    return nil, err
  end
  local data = f:read '*a'
  f:close()
  return data
end

local function write_file(path, content)
  local f, err = io.open(path, 'w')
  if not f then
    return nil, 'open failed: ' .. tostring(err)
  end
  f:write(content)
  f:flush()
  f:close()
  return true
end

local function sorted_copy(map)
  local keys = {}
  for k in pairs(map) do
    table.insert(keys, k)
  end
  table.sort(keys, function(a, b)
    return a:lower() < b:lower()
  end)
  local out = {}
  for _, k in ipairs(keys) do
    out[k] = map[k]
  end
  return out
end

-- --- constructor -------------------------------------------------

--- Create a new mapper backed by a JSON file.
-- @param mapping_path string: path to JSON file (object with string keys/values)
-- @return KeyValueMapper
function KeyValueMapper.new(mapping_path)
  local self = setmetatable({}, KeyValueMapper)
  self.mapping_path = assert(mapping_path, 'mapping_path is required')
  self.map = {} -- csv_location -> excel_location
  self.reverse = {} -- excel_location -> csv_location (for uniqueness check)

  if file_exists(mapping_path) then
    local content, rerr = read_all(mapping_path)
    if not content then
      error('Failed to read mapping file: ' .. tostring(rerr))
    end
    local decoded, _, derr = dkjson.decode(content)
    if not decoded or type(decoded) ~= 'table' then
      -- Backup corrupt file and start fresh
      local backup = mapping_path .. '.corrupt'
      os.rename(mapping_path, backup)
      io.stderr:write('Invalid JSON. Backed up to: ' .. backup .. '\n')
      decoded = {}
    end
    -- normalize & build both maps
    for k, v in pairs(decoded) do
      local ck = norm_str(k)
      local ev = norm_str(v)
      if ck ~= '' and ev ~= '' then
        self.map[ck] = ev
        self.reverse[ev] = ck
      end
    end
  else
    -- Create an empty file on first save
    self:save()
  end

  return self
end

-- --- API ---------------------------------------------------------

--- Get mapped Excel location by CSV location.
-- @param csv_location string
-- @return string|nil
function KeyValueMapper:get(csv_location)
  csv_location = norm_str(csv_location)
  if csv_location == '' then
    return nil
  end
  return self.map[csv_location]
end

--- Check if a CSV location key exists.
function KeyValueMapper:exists_key(csv_location)
  csv_location = norm_str(csv_location)
  return self.map[csv_location] ~= nil
end

--- Check if an Excel location value exists (used by some other CSV key).
function KeyValueMapper:exists_value(excel_location)
  excel_location = norm_str(excel_location)
  return self.reverse[excel_location] ~= nil
end

--- Add a new mapping. Enforces uniqueness for both key and value.
-- @param csv_location string
-- @param excel_location string
-- @return true | nil, err
function KeyValueMapper:add(csv_location, excel_location)
  csv_location = norm_str(csv_location)
  excel_location = norm_str(excel_location)

  if csv_location == '' then
    return nil, 'csv_location is empty'
  end
  if excel_location == '' then
    return nil, 'excel_location is empty'
  end

  if self.map[csv_location] ~= nil then
    return nil, ("csv_location already exists: '%s'"):format(csv_location)
  end
  if self.reverse[excel_location] ~= nil then
    return nil,
      ("excel_location already in use by csv '%s'"):format(
        self.reverse[excel_location]
      )
  end

  self.map[csv_location] = excel_location
  self.reverse[excel_location] = csv_location

  local ok, err = self:save()
  if not ok then
    -- Roll back in-memory if save fails
    self.reverse[excel_location] = nil
    self.map[csv_location] = nil
    return nil, err
  end

  return true
end

--- Persist to JSON file (atomic write).
-- @return true | nil, err
function KeyValueMapper:save()
  local ordered = sorted_copy(self.map)
  local content = dkjson.encode(ordered, { indent = true })
  local ok, err = write_file(self.mapping_path, content)
  if not ok then
    return nil, 'save failed: ' .. tostring(err)
  end
  return true
end

return KeyValueMapper
