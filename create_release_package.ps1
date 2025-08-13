# Create Release Package Script
# This script creates a zip file for GitHub releases

param(
    [string]$Version = "1.2.0",
    [string]$OutputPath = ".\"
)

$ReleaseName = "OutlookAutoArchive-v$Version"
$ReleasePath = Join-Path $OutputPath $ReleaseName

Write-Host "Creating release package for version $Version..." -ForegroundColor Green

# Create release directory
if (Test-Path $ReleasePath) {
    Remove-Item $ReleasePath -Recurse -Force
}
New-Item -ItemType Directory -Path $ReleasePath | Out-Null

# Files to include in release
$FilesToInclude = @(
    "OutlookAutoArchive.exe",
    "Run_OutlookAutoArchive.bat",
    "OutlookAutoArchive.ps1",
    "config.example.json",
    "README.md",
    "LICENSE",
    "CHANGELOG.md",
    "RELEASE_NOTES.md"
)

# Copy files to release directory
foreach ($File in $FilesToInclude) {
    if (Test-Path $File) {
        Copy-Item $File $ReleasePath
        Write-Host "✓ Added $File" -ForegroundColor Green
    } else {
        Write-Host "⚠ Warning: $File not found" -ForegroundColor Yellow
    }
}

# Create zip file
$ZipPath = "$ReleasePath.zip"
if (Test-Path $ZipPath) {
    Remove-Item $ZipPath -Force
}

try {
    Compress-Archive -Path $ReleasePath -DestinationPath $ZipPath -Force
    Write-Host "✓ Created release package: $ZipPath" -ForegroundColor Green
    
    # Get file size
    $FileSize = (Get-Item $ZipPath).Length
    $FileSizeMB = [math]::Round($FileSize / 1MB, 2)
    Write-Host "✓ Package size: $FileSizeMB MB" -ForegroundColor Green
    
} catch {
    Write-Host "✗ Error creating zip file: $_" -ForegroundColor Red
}

# Clean up release directory
Remove-Item $ReleasePath -Recurse -Force

Write-Host "`nRelease package ready for GitHub upload!" -ForegroundColor Cyan
Write-Host "File: $ZipPath" -ForegroundColor White
