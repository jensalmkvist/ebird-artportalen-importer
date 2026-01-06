-- example.lua
local KeyValueMapper = require 'KeyValueMapper'

local mapper = KeyValueMapper.new 'location_mappings.json'

-- Query
local csv_key = 'Warehouse_22'
local val = mapper:get(csv_key)
if val then
  print(("Mapped '%s' -> '%s'"):format(csv_key, val))
else
  print(("No mapping for '%s'"):format(csv_key))
end

-- Add (enforces uniqueness on both the key and the value)
local ok, err = mapper:add('New_CSV_Key', 'Excel Location A')
if not ok then
  print('Add failed: ' .. err)
else
  print 'Added mapping.'
end
