-- csv_prompt_map.lua
-- Read a CSV file, resolve each location via LocationMapper, and prompt to add missing mappings.

-- If you vendored dkjson and don't use LuaRocks loader, uncomment the next line:
-- package.path = package.path .. ";./vendor/?.lua;./vendor/?/init.lua"

local KeyValueMapper = require 'KeyValueMapper'

-- ----------- small CSV helpers (safe enough for common CSV) ------------

local function trim(s)
  return (s:gsub('^%s+', ''):gsub('%s+$', ''))
end

-- Split a CSV line into fields, supporting quoted fields with commas and escaped quotes ("")
-- Note: This is a minimal parser for RFC4180-ish CSV; good enough for common cases.
local function parse_csv_line(line)
  local res = {}
  local i, len = 1, #line
  while i <= len do
    if line:sub(i, i) == '"' then
      -- Quoted field
      i = i + 1
      local buf = {}
      local closed = false
      while i <= len do
        local c = line:sub(i, i)
        if c == '"' then
          -- Check for escaped quote ("")
          if i + 1 <= len and line:sub(i + 1, i + 1) == '"' then
            table.insert(buf, '"')
            i = i + 2
          else
            closed = true
            i = i + 1
            break
          end
        else
          table.insert(buf, c)
          i = i + 1
        end
      end
      local field = table.concat(buf)
      -- Skip trailing comma (if present)
      if i <= len and line:sub(i, i) == ',' then
        i = i + 1
      end
      table.insert(res, field)
    else
      -- Unquoted field: read until comma or end
      local j = line:find(',', i, true)
      if j then
        table.insert(res, trim(line:sub(i, j - 1)))
        i = j + 1
      else
        table.insert(res, trim(line:sub(i)))
        break
      end
    end
  end
  return res
end

local function read_header(fp)
  local header_line = fp:read '*l'
  if not header_line then
    return nil, 'CSV is empty or unreadable'
  end
  local cols = parse_csv_line(header_line)
  return cols
end

-- ----------- main logic ------------

local function main()
  -- Args: csv_path, location_col_name, mappings_json_path
  local csv_path = arg[1]
  local csv_col = arg[2]
  local mappings_json = arg[3]

  if not csv_path or not csv_col then
    io.stderr:write 'Usage: lua csv_prompt_map.lua <csv_path> <location_col_name> [mappings.json]\n'
    os.exit(2)
  end

  -- Open CSV
  local fp, ferr = io.open(csv_path, 'r')
  if not fp then
    io.stderr:write('Failed to open CSV: ' .. tostring(ferr) .. '\n')
    os.exit(1)
  end

  -- Initialize mapper (loads JSON or creates empty)
  local mapper = KeyValueMapper.new(mappings_json)

  -- Read header and find column index
  local header = read_header(fp)
  if not header then
    io.stderr:write 'Failed to read header row.\n'
    fp:close()
    os.exit(1)
  end

  -- Build column name -> index
  local col_index = {}
  for i, name in ipairs(header) do
    col_index[trim(name)] = i
  end

  local idx = col_index[csv_col]
  if not idx then
    io.stderr:write(
      ("CSV does not have column '%s'. Found columns: %s\n"):format(
        csv_col,
        table.concat(header, ', ')
      )
    )
    fp:close()
    os.exit(1)
  end

  local added = 0
  local skipped_existing = 0
  local line_no = 1 -- header is line 1

  -- Process each row
  for line in fp:lines() do
    line_no = line_no + 1
    if line ~= '' then
      local fields = parse_csv_line(line)
      -- Handle rows with fewer fields gracefully
      local raw_key = fields[idx] or ''
      local key = trim(raw_key)

      if key ~= '' then
        local existing = mapper:get(key)
        if existing then
          -- Already mapped; no prompt
          skipped_existing = skipped_existing + 1
        else
          -- Prompt for Excel location
          io.write(
            ("No mapping for '%s' (CSV line %d). Enter Excel location (leave blank to skip): "):format(
              key,
              line_no
            )
          )
          local val = io.read '*l' or ''
          val = trim(val)

          if val ~= '' then
            local ok, err = mapper:add(key, val)
            if not ok then
              io.stderr:write(
                ("Add failed for key '%s': %s\n"):format(key, err)
              )
              -- Continue; you might want to re-prompt or log
            else
              added = added + 1
              io.write(("Added: '%s' -> '%s'\n"):format(key, val))
            end
          else
            -- skipped
          end
        end
      end
    end
  end

  fp:close()

  -- Summary
  print(
    ('\nDone. Existing mappings used: %d, new mappings added: %d\n'):format(
      skipped_existing,
      added
    )
  )
  print(('Mapping file: %s'):format(mappings_json))
end

main()
