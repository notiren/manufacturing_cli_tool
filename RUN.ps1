# RUN.ps1
# Full path to Python executable (if in PATH, just "python" works)
$pythonExe = "python"

# Full path to the Python script
$scriptPath = Join-Path $PSScriptRoot "cli_tool.py"

# Change working directory to script folder
Set-Location $PSScriptRoot

# Check for web mode argument
if ($args -contains "--web") {
    Write-Host "Starting Manufacturing CLI Tool in Web Mode..."
    Write-Host "Open your browser to http://localhost:5000"
    Write-Host "Press Ctrl+C to stop the web server"
    & $pythonExe "$scriptPath" --web
} else {
    & $pythonExe "$scriptPath"
}
