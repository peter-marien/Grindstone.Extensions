$sourcePath = Join-Path $PSScriptRoot "Extensions"
$destPath = Join-Path $env:APPDATA "Grindstone 4\Extensions"

Write-Host "Deploying extensions from '$sourcePath' to '$destPath'..."

if (!(Test-Path $sourcePath)) {
    Write-Error "Source directory not found: $sourcePath"
    exit 1
}

if (!(Test-Path $destPath)) {
    Write-Host "Destination directory does not exist. Creating it..."
    New-Item -ItemType Directory -Force -Path $destPath | Out-Null
}

Copy-Item -Path "$sourcePath\*" -Destination $destPath -Recurse -Force

Write-Host "Deployment complete!" -ForegroundColor Green
