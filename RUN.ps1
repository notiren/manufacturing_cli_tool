# RUN.ps1
# Full path to Python executable (if in PATH, just "python" works)
$pythonExe = "python"

# Full path to the Python script
$scriptPath = Join-Path $PSScriptRoot "cli_tool.py"

# Run Python with the script
& $pythonExe "$scriptPath"
