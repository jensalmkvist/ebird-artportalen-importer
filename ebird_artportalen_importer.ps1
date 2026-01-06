# this is to set to set the concole to UTF-8 before running the actual lua script

chcp 65001 >$null
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
