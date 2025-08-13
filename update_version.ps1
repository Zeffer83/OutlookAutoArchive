# Version Management Script
# Updates version numbers across all project files

param(
    [Parameter(Mandatory=$true)]
    [string]$NewVersion,
    
    [string]$ProjectPath = "."
)

Write-Host "Updating version to $NewVersion across all files..." -ForegroundColor Green

# Files to update with version numbers
$FilesToUpdate = @{
    "README.md" = @{
        "**Version**: \d+\.\d+\.\d+" = "**Version**: $NewVersion"
        "version \d+\.\d+\.\d+" = "version $NewVersion"
    }
    "OutlookAutoArchive.ps1" = @{
        "# Version: \d+\.\d+\.\d+" = "# Version: $NewVersion"
    }
    "CHANGELOG.md" = @{
        "## \[$NewVersion\]" = "## [$NewVersion]"
    }
}

foreach ($File in $FilesToUpdate.Keys) {
    $FilePath = Join-Path $ProjectPath $File
    if (Test-Path $FilePath) {
        $Content = Get-Content $FilePath -Raw
        $OriginalContent = $Content
        
        foreach ($Pattern in $FilesToUpdate[$File].Keys) {
            $Replacement = $FilesToUpdate[$File][$Pattern]
            $Content = $Content -replace $Pattern, $Replacement
        }
        
        if ($Content -ne $OriginalContent) {
            Set-Content $FilePath $Content -NoNewline
            Write-Host "✓ Updated $File" -ForegroundColor Green
        } else {
            Write-Host "⚠ No changes needed in $File" -ForegroundColor Yellow
        }
    } else {
        Write-Host "✗ File not found: $File" -ForegroundColor Red
    }
}

# Update release package script
$ReleaseScriptPath = Join-Path $ProjectPath "create_release_package.ps1"
if (Test-Path $ReleaseScriptPath) {
    $Content = Get-Content $ReleaseScriptPath -Raw
    $Content = $Content -replace '\[string\]\$Version = "\d+\.\d+\.\d+"', "[string]`$Version = `"$NewVersion`""
    Set-Content $ReleaseScriptPath $Content -NoNewline
    Write-Host "✓ Updated create_release_package.ps1" -ForegroundColor Green
}

Write-Host "`nVersion update complete!" -ForegroundColor Cyan
Write-Host "Next steps:" -ForegroundColor White
Write-Host "1. Test the changes" -ForegroundColor White
Write-Host "2. Commit the version update" -ForegroundColor White
Write-Host "3. Create a git tag: git tag v$NewVersion" -ForegroundColor White
Write-Host "4. Push the tag: git push origin v$NewVersion" -ForegroundColor White
