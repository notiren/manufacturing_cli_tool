# update.ps1
# Pulls the latest changes from the remote GitHub repository.

Set-Location $PSScriptRoot

Write-Host "Checking for updates from GitHub..."

# Fetch remote changes without merging
git fetch origin main 2>&1 | Out-Null

# Compare local HEAD with remote HEAD
$localHash  = git rev-parse HEAD
$remoteHash = git rev-parse origin/main

if ($localHash -eq $remoteHash) {
    Write-Host "Already up to date. No changes pulled."
} else {
    Write-Host "Updates found. Pulling latest changes..."
    git pull origin main

    if ($LASTEXITCODE -eq 0) {
        Write-Host "Update successful."
    } else {
        Write-Host "Update failed. Check for conflicts or network issues." -ForegroundColor Red
    }
}
