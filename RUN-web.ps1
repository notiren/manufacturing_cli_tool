# RUN-web.ps1
$pythonExe = "python"
$scriptPath = Join-Path $PSScriptRoot "cli_tool.py"

Set-Location $PSScriptRoot

Write-Host "Starting Manufacturing CLI Tool in Web Mode..."
Write-Host "Open your browser to http://localhost:5000"
Write-Host "Press Ctrl+C to stop the web server"

Start-Process "http://localhost:5000"

& $pythonExe "$scriptPath" --web
