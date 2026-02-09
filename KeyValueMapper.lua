-- KeyValueMapper.lua
-- JSON-backed mapper: csv_key -> value (string or table)
-- Requires: dkjson

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
-- @param mapping_path string
-- @param opts table|nil  (e.g., { value_key_for_reverse = 'lokalnamn' })
function KeyValueMapper.new(mapping_path, opts)
  local self = setmetatable({}, KeyValueMapper)
  self.mapping_path = assert(mapping_path, 'mapping_path is required')
  self.value_key_for_reverse = opts and opts.value_key_for_reverse or nil
  self.map = {} -- csv_key -> value (string or table)
  self.reverse = {} -- reverse_key -> { [csv_key] = true } (set of csv_keys)
  self._migrated = false

  local function extract_reverse_key(v)
    if type(v) == 'table' then
      if not self.value_key_for_reverse then
        return nil
      end
      local kk = norm_str(v[self.value_key_for_reverse] or '')
      return kk ~= '' and kk or nil
    elseif type(v) == 'string' then
      local kk = norm_str(v)
      return kk ~= '' and kk or nil
    end
    return nil
  end
  self._extract_reverse_key = extract_reverse_key

  if file_exists(mapping_path) then
    local content, rerr = read_all(mapping_path)
    if not content then
      error('Failed to read mapping file: ' .. tostring(rerr))
    end
    local decoded, _, derr = dkjson.decode(content)
    if not decoded or type(decoded) ~= 'table' then
      local backup = mapping_path .. '.corrupt'
      os.rename(mapping_path, backup)
      io.stderr:write('Invalid JSON. Backed up to: ' .. backup .. '\n')
      decoded = {}
    end

    for k, v in pairs(decoded) do
      local ck = norm_str(k)
      local val = v
      -- auto-migrate string -> table if we enforce a field for reverse
      if type(v) == 'string' and self.value_key_for_reverse then
        val = { [self.value_key_for_reverse] = norm_str(v) }
        self._migrated = true
      end
      if ck ~= '' and val ~= nil then
        self.map[ck] = val
        local rk = extract_reverse_key(val)
        if rk then
          self.reverse[rk] = self.reverse[rk] or {}
          self.reverse[rk][ck] = true
        end
      end
    end
  else
    -- Create an empty file on first save
    self:save()
  end

  if self._migrated then
    -- Persist migration of existing entries to objects
    self:save()
  end

  return self
end

-- --- API ---------------------------------------------------------

--- Get mapped value by key (string or table).
function KeyValueMapper:get(csv_key)
  csv_key = norm_str(csv_key)
  if csv_key == '' then
    return nil
  end
  return self.map[csv_key]
end

--- Check if a key exists.
function KeyValueMapper:exists_key(csv_key)
  csv_key = norm_str(csv_key)
  return self.map[csv_key] ~= nil
end

--- Check if a reverse value exists (string key used for reverse lookups).
-- For string values: pass the value.
-- For table values: pass the configured field value (e.g., lokalnamn).
function KeyValueMapper:exists_value(value_key)
  local k = norm_str(value_key)
  local rev = self.reverse[k]
  return rev ~= nil and next(rev) ~= nil
end

--- Add a new mapping. Enforces:
--   - unique key (csv_key)
--   - value may be shared across many keys (no uniqueness check)
function KeyValueMapper:add(csv_key, value)
  csv_key = norm_str(csv_key)
  if csv_key == '' then
    return nil, 'csv_key is empty'
  end
  if value == nil then
    return nil, 'value is nil'
  end

  if self.map[csv_key] ~= nil then
    return nil, ("csv_key already exists: '%s'"):format(csv_key)
  end

  local rk = self._extract_reverse_key(value)

  self.map[csv_key] = value
  if rk then
    self.reverse[rk] = self.reverse[rk] or {}
    self.reverse[rk][csv_key] = true
  end

  local ok, err = self:save()
  if not ok then
    -- rollback
    if rk and self.reverse[rk] then
      self.reverse[rk][csv_key] = nil
      if next(self.reverse[rk]) == nil then
        self.reverse[rk] = nil
      end
    end
    self.map[csv_key] = nil
    return nil, err
  end

  return true
end

--- Set/replace a mapping (upsert). Updates reverse safely.
function KeyValueMapper:set(csv_key, value)
  csv_key = norm_str(csv_key)
  if csv_key == '' then
    return nil, 'csv_key is empty'
  end
  if value == nil then
    return nil, 'value is nil'
  end

  local old = self.map[csv_key]
  local old_rk = old and self._extract_reverse_key(old) or nil
  local new_rk = self._extract_reverse_key(value)

  -- remove old reverse membership
  if old_rk and self.reverse[old_rk] then
    self.reverse[old_rk][csv_key] = nil
    if next(self.reverse[old_rk]) == nil then
      self.reverse[old_rk] = nil
    end
  end

  -- update map and add new reverse membership
  self.map[csv_key] = value
  if new_rk then
    self.reverse[new_rk] = self.reverse[new_rk] or {}
    self.reverse[new_rk][csv_key] = true
  end

  -- persist
  local ok, err = self:save()
  if not ok then
    -- best-effort rollback
    -- remove new membership
    if new_rk and self.reverse[new_rk] then
      self.reverse[new_rk][csv_key] = nil
      if next(self.reverse[new_rk]) == nil then
        self.reverse[new_rk] = nil
      end
    end
    -- restore old map
    self.map[csv_key] = old
    -- restore old membership
    if old and old_rk then
      self.reverse[old_rk] = self.reverse[old_rk] or {}
      self.reverse[old_rk][csv_key] = true
    end
    return nil, err
  end

  return true
end

--- Persist to JSON file.
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
