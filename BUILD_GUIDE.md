# Build Guide

This guide explains how to build the OutlookAutoArchive application using the automated GitHub Actions workflow or manually.

## 🚀 Automated Build (Recommended)

### Using GitHub Actions

The repository includes an automated build workflow that compiles PowerShell scripts to executables and creates releases automatically.

#### Trigger Methods:

1. **Tag-based Release** (Recommended):
   ```bash
   # Create and push a new version tag
   git tag -a v2.9.6 -m "Release v2.9.6"
   git push origin v2.9.6
   ```
   This automatically triggers the build and creates a GitHub release.

2. **Manual Trigger**:
   - Go to GitHub repository → Actions → "Auto Build and Release"
   - Click "Run workflow"
   - Enter version number (e.g., 2.9.6)
   - Optionally check "Create and push a new tag"
   - Click "Run workflow"

#### What the Automated Build Does:

1. **Script Validation**: Runs PSScriptAnalyzer to check for issues
2. **Executable Compilation**: Compiles PowerShell scripts to .exe files using PS2EXE
3. **Metadata Injection**: Adds version info, company details, and copyright
4. **Package Creation**: Creates a complete release package with all files
5. **Release Creation**: Automatically creates a GitHub release with assets
6. **Artifact Upload**: Saves build artifacts for 30 days

### Generated Files:

- **OutlookAutoArchive.exe** - Main application (compiled from PowerShell)
- **setup_task_scheduler.exe** - Task scheduler setup utility
- **OutlookAutoArchive.ps1** - Source PowerShell script
- **setup_task_scheduler.ps1** - Setup script source
- **icon.ico** - Application icon
- **config_example.json** - Configuration template
- **Complete package ZIP** - All files bundled together

## 🔧 Manual Build Process

If you prefer to build locally or need to customize the build process:

### Prerequisites:

1. **Windows 10/11** (required for PS2EXE)
2. **PowerShell 5.1 or later**
3. **Required PowerShell Modules**:
   ```powershell
   Install-Module -Name PS2EXE -Force -Scope CurrentUser
   Install-Module -Name PSScriptAnalyzer -Force -Scope CurrentUser
   ```

### Build Steps:

1. **Validate Script**:
   ```powershell
   $results = Invoke-ScriptAnalyzer -Path "OutlookAutoArchive.ps1" -Settings PSGallery
   if ($results.Count -gt 0) {
       Write-Host "Script analysis found issues:"
       $results | Format-Table -AutoSize
   }
   ```

2. **Build Main Executable**:
   ```powershell
   $version = "2.9.6"  # Set your version
   ps2exe -InputFile "OutlookAutoArchive.ps1" `
          -OutputFile "OutlookAutoArchive.exe" `
          -Version "$version" `
          -Title "Outlook Auto Archive" `
          -Description "Automatic email archiving functionality" `
          -Company "Ryan Zeffiretti" `
          -Product "Outlook Auto Archive" `
          -Copyright "Copyright (c) 2025 Ryan Zeffiretti. Licensed under MIT License." `
          -Trademark "Outlook Auto Archive" `
          -IconFile "icon.ico" `
          -NoConsole `
          -RequireAdmin
   ```

3. **Build Setup Executable**:
   ```powershell
   ps2exe -InputFile "setup_task_scheduler.ps1" `
          -OutputFile "setup_task_scheduler.exe" `
          -Version "$version" `
          -Title "Outlook Auto Archive Setup" `
          -Description "Task scheduler setup utility" `
          -Company "Ryan Zeffiretti" `
          -Product "Outlook Auto Archive" `
          -Copyright "Copyright (c) 2025 Ryan Zeffiretti. Licensed under MIT License." `
          -IconFile "icon.ico" `
          -NoConsole `
          -RequireAdmin
   ```

4. **Create Release Package**:
   ```powershell
   $releaseDir = "OutlookAutoArchive-v$version"
   New-Item -ItemType Directory -Path $releaseDir -Force
   
   # Copy files
   Copy-Item "OutlookAutoArchive.exe" $releaseDir/
   Copy-Item "setup_task_scheduler.exe" $releaseDir/
   Copy-Item "OutlookAutoArchive.ps1" $releaseDir/
   Copy-Item "setup_task_scheduler.ps1" $releaseDir/
   Copy-Item "icon.ico" $releaseDir/
   Copy-Item "README.md" $releaseDir/
   Copy-Item "LICENSE" $releaseDir/
   Copy-Item "CHANGELOG.md" $releaseDir/
   
   # Create config example
   $configExample = '{"DryRun": true,"ArchiveFolders": [],"LogPath": "","InstallPath": ""}'
   $configExample | Out-File -FilePath "$releaseDir/config_example.json" -Encoding UTF8
   
   # Create ZIP
   Compress-Archive -Path $releaseDir -DestinationPath "$releaseDir.zip" -Force
   ```

## 📋 Build Configuration

### PS2EXE Options Used:

- **-NoConsole**: Creates a Windows GUI application
- **-RequireAdmin**: Requires administrator privileges
- **-IconFile**: Uses custom icon for the executable
- **-Version**: Embeds version information
- **-Title/Description/Company**: Adds metadata to executable properties

### Version Management:

- Use semantic versioning (e.g., 2.9.6)
- Update version in all relevant files before building
- Tag releases with 'v' prefix (e.g., v2.9.6)

## 🔍 Troubleshooting

### Common Issues:

1. **PS2EXE Module Not Found**:
   ```powershell
   Install-Module -Name PS2EXE -Force -Scope CurrentUser
   ```

2. **Script Analysis Errors**:
   - Review PSScriptAnalyzer output
   - Fix any PowerShell best practice violations
   - Update script to follow coding standards

3. **Build Failures**:
   - Check PowerShell execution policy: `Get-ExecutionPolicy`
   - Ensure all required files exist
   - Verify icon.ico file is present

4. **Large Executable Size**:
   - This is normal for PS2EXE compiled files
   - Consider using UPX compression for release builds
   - Monitor size trends over versions

### Build Verification:

```powershell
# Check executable properties
Get-ItemProperty "OutlookAutoArchive.exe" | Select-Object Name, Length, VersionInfo

# Test executable
Start-Process "OutlookAutoArchive.exe" -Wait
```

## 📊 Build Metrics

### Typical Build Times:
- **Automated Build**: 3-5 minutes
- **Manual Build**: 1-2 minutes

### File Sizes (approximate):
- **OutlookAutoArchive.exe**: ~116 KB
- **setup_task_scheduler.exe**: ~32 KB
- **Complete package**: ~400 KB

### Quality Gates:
- ✅ Script validation passes
- ✅ Executables created successfully
- ✅ All files included in package
- ✅ Release assets uploaded

## 🔄 Continuous Integration

The automated build workflow integrates with:

- **GitHub Releases**: Automatic release creation
- **Artifact Storage**: 30-day retention of build artifacts
- **Version Tagging**: Automatic version management
- **Quality Checks**: Script validation and testing

---

*For questions or issues with the build process, please check the troubleshooting section or create an issue in the repository.*
