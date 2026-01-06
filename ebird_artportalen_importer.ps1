# this is to set to set the concole to UTF-8 before running the actual lua script

chcp 65001 >$null
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
lua .\ebird_to_artportalen.lua 'C:\Projects\Personal\ebird-artportalen-importer\Import files/' .\out_for_excel.csv
