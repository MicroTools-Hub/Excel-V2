$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$appRoot = Join-Path $root "sources\excel-mcp-server-main\excel-mcp-server-main"
$python = Join-Path $appRoot ".venv\Scripts\python.exe"

$env:PYTHONPATH = Join-Path $appRoot "src"
$env:EXCEL_FILES_PATH = Join-Path $root "excel_files"
$env:PORT = if ($env:PORT) { $env:PORT } else { "10000" }

New-Item -ItemType Directory -Force -Path $env:EXCEL_FILES_PATH | Out-Null
Set-Location $root
& $python -m excel_mcp.webapp
