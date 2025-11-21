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

$extensions = Get-ChildItem -Path $sourcePath -Directory

foreach ($ext in $extensions) {
    $extName = $ext.Name
    $extSourcePath = $ext.FullName
    $extDestPath = Join-Path $destPath $extName

    Write-Host "Deploying '$extName'..."

    if (Test-Path $extDestPath) {
        Write-Host "  Cleaning target directory..."
        Remove-Item -Path "$extDestPath\*" -Recurse -Force
    }
    else {
        New-Item -ItemType Directory -Force -Path $extDestPath | Out-Null
    }

    Write-Host "  Copying files..."
    Copy-Item -Path "$extSourcePath\*" -Destination $extDestPath -Recurse -Force
}

Write-Host "Deployment complete!" -ForegroundColor Green
