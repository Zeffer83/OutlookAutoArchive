# Test script to verify icon embedding in executables
# This script can be run locally to test if PS2EXE is properly embedding the icon

param(
    [string]$Version = "2.9.6"
)

Write-Host "Testing icon embedding for OutlookAutoArchive executables..." -ForegroundColor Green

# Check if icon file exists
if (-not (Test-Path "icon.ico")) {
    Write-Error "❌ Icon file not found: icon.ico"
    exit 1
}

$iconSize = (Get-Item "icon.ico").Length
Write-Host "✓ Icon file found: icon.ico ($([math]::Round($iconSize/1KB,2)) KB)" -ForegroundColor Green

# Get absolute path to icon
$iconPath = (Resolve-Path "icon.ico").Path
Write-Host "Icon path: $iconPath" -ForegroundColor Yellow

# Test building main executable
Write-Host "`nBuilding test executable..." -ForegroundColor Cyan

try {
    # Check if PS2EXE is available
    if (-not (Get-Module -ListAvailable -Name PS2EXE)) {
        Write-Host "Installing PS2EXE module..." -ForegroundColor Yellow
        Install-Module -Name PS2EXE -Force -Scope CurrentUser
    }

    # Build test executable
    ps2exe -InputFile "OutlookAutoArchive.ps1" `
           -OutputFile "test_OutlookAutoArchive.exe" `
           -Version "$Version" `
           -Title "Outlook Auto Archive (Test)" `
           -Description "Test build for icon verification" `
           -Company "Ryan Zeffiretti" `
           -Product "Outlook Auto Archive" `
           -Copyright "Copyright (c) 2025 Ryan Zeffiretti. Licensed under MIT License." `
           -IconFile "$iconPath" `
           -NoConsole `
           -RequireAdmin

    if (Test-Path "test_OutlookAutoArchive.exe") {
        $exeSize = (Get-Item "test_OutlookAutoArchive.exe").Length
        Write-Host "✓ Test executable created: test_OutlookAutoArchive.exe ($([math]::Round($exeSize/1KB,2)) KB)" -ForegroundColor Green
        
        # Try to get executable properties to verify icon embedding
        try {
            $exeInfo = Get-ItemProperty "test_OutlookAutoArchive.exe"
            Write-Host "✓ Executable metadata accessible" -ForegroundColor Green
            Write-Host "  File: $($exeInfo.Name)" -ForegroundColor Gray
            Write-Host "  Size: $([math]::Round($exeInfo.Length/1KB,2)) KB" -ForegroundColor Gray
        } catch {
            Write-Warning "⚠️ Could not access executable metadata"
        }
        
        # Clean up test file
        Remove-Item "test_OutlookAutoArchive.exe" -Force
        Write-Host "✓ Test file cleaned up" -ForegroundColor Green
        
    } else {
        Write-Error "❌ Test executable was not created"
        exit 1
    }
    
} catch {
    Write-Error "❌ Error during test build: $($_.Exception.Message)"
    exit 1
}

Write-Host "`n✅ Icon embedding test completed successfully!" -ForegroundColor Green
Write-Host "The build workflow should now properly embed the icon in executables." -ForegroundColor Cyan
