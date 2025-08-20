# Repository Housekeeping Script
# This script performs various optimization tasks to keep the repository clean and efficient

param(
    [switch]$Force,
    [switch]$DryRun
)

Write-Host "🧹 Repository Housekeeping Script" -ForegroundColor Green
Write-Host "=================================" -ForegroundColor Green

if ($DryRun) {
    Write-Host "🔍 DRY RUN MODE - No changes will be made" -ForegroundColor Yellow
}

# Function to check if command exists
function Test-Command($cmdname) {
    return [bool](Get-Command -Name $cmdname -ErrorAction SilentlyContinue)
}

# 1. Git Garbage Collection
Write-Host "`n1. Running Git Garbage Collection..." -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "   Would run: git gc --prune=now" -ForegroundColor Gray
} else {
    git gc --prune=now
    if ($LASTEXITCODE -eq 0) {
        Write-Host "   ✅ Git garbage collection completed" -ForegroundColor Green
    } else {
        Write-Host "   ❌ Git garbage collection failed" -ForegroundColor Red
    }
}

# 2. Clean up reflog entries
Write-Host "`n2. Cleaning up reflog entries..." -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "   Would run: git reflog expire --expire=30.days.ago --expire-unreachable=30.days.ago --all" -ForegroundColor Gray
} else {
    git reflog expire --expire=30.days.ago --expire-unreachable=30.days.ago --all
    Write-Host "   ✅ Reflog cleanup completed" -ForegroundColor Green
}

# 3. Prune remote references
Write-Host "`n3. Pruning remote references..." -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "   Would run: git remote prune origin" -ForegroundColor Gray
} else {
    git remote prune origin
    Write-Host "   ✅ Remote pruning completed" -ForegroundColor Green
}

# 4. Repack objects for better compression
Write-Host "`n4. Repacking objects for better compression..." -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "   Would run: git repack -a -d --depth=250 --window=250" -ForegroundColor Gray
} else {
    git repack -a -d --depth=250 --window=250
    if ($LASTEXITCODE -eq 0) {
        Write-Host "   ✅ Object repacking completed" -ForegroundColor Green
    } else {
        Write-Host "   ❌ Object repacking failed" -ForegroundColor Red
    }
}

# 5. Check repository size
Write-Host "`n5. Repository size analysis..." -ForegroundColor Cyan
$gitSize = (Get-ChildItem -Path .git -Recurse | Measure-Object -Property Length -Sum).Sum
$gitSizeMB = [math]::Round($gitSize / 1MB, 2)
Write-Host "   📊 Git repository size: $gitSizeMB MB" -ForegroundColor Yellow

# 6. Count objects
Write-Host "`n6. Object count analysis..." -ForegroundColor Cyan
$objects = git count-objects -vH
Write-Host "   📊 Object statistics:" -ForegroundColor Yellow
$objects | ForEach-Object { Write-Host "      $_" -ForegroundColor Gray }

# 7. Find large files
Write-Host "`n7. Large file analysis..." -ForegroundColor Cyan
$largeFiles = Get-ChildItem -Recurse | Where-Object { $_.Length -gt 100KB } | Sort-Object Length -Descending
if ($largeFiles) {
    Write-Host "   📁 Large files found:" -ForegroundColor Yellow
    $largeFiles | ForEach-Object {
        $sizeKB = [math]::Round($_.Length / 1KB, 2)
        Write-Host "      $($_.Name) - $sizeKB KB" -ForegroundColor Gray
    }
} else {
    Write-Host "   ✅ No large files found" -ForegroundColor Green
}

# 8. Check for potential optimizations
Write-Host "`n8. Optimization recommendations..." -ForegroundColor Cyan

# Check if Git LFS is available
if (Test-Command "git-lfs") {
    Write-Host "   ✅ Git LFS is available" -ForegroundColor Green
    Write-Host "   💡 Consider using Git LFS for large binary files (.exe, .ico)" -ForegroundColor Blue
} else {
    Write-Host "   ⚠️ Git LFS not installed" -ForegroundColor Yellow
    Write-Host "   💡 Install Git LFS for better handling of large files" -ForegroundColor Blue
}

# Check for unnecessary files
$unnecessaryFiles = @(
    "*.tmp", "*.temp", "*.bak", "*.log", "Thumbs.db", "Desktop.ini"
)

$foundUnnecessary = @()
foreach ($pattern in $unnecessaryFiles) {
    $files = Get-ChildItem -Recurse -Filter $pattern -ErrorAction SilentlyContinue
    if ($files) {
        $foundUnnecessary += $files
    }
}

if ($foundUnnecessary) {
    Write-Host "   ⚠️ Found potentially unnecessary files:" -ForegroundColor Yellow
    $foundUnnecessary | ForEach-Object {
        Write-Host "      $($_.Name)" -ForegroundColor Gray
    }
    Write-Host "   💡 Consider adding these patterns to .gitignore" -ForegroundColor Blue
} else {
    Write-Host "   ✅ No unnecessary files found" -ForegroundColor Green
}

# 9. Final summary
Write-Host "`n📋 Housekeeping Summary" -ForegroundColor Green
Write-Host "=====================" -ForegroundColor Green

if ($DryRun) {
    Write-Host "🔍 This was a dry run. No changes were made." -ForegroundColor Yellow
    Write-Host "💡 Run without -DryRun to perform actual housekeeping." -ForegroundColor Blue
} else {
    Write-Host "✅ Repository housekeeping completed successfully!" -ForegroundColor Green
    Write-Host "📊 Repository size: $gitSizeMB MB" -ForegroundColor Yellow
}

Write-Host "`n💡 Additional recommendations:" -ForegroundColor Blue
Write-Host "   • Run this script monthly to keep the repository optimized" -ForegroundColor Gray
Write-Host "   • Consider using Git LFS for large binary files" -ForegroundColor Gray
Write-Host "   • Review .gitignore regularly to exclude unnecessary files" -ForegroundColor Gray
Write-Host "   • Use shallow clones for CI/CD to reduce download times" -ForegroundColor Gray

Write-Host "`n🏁 Housekeeping complete!" -ForegroundColor Green
