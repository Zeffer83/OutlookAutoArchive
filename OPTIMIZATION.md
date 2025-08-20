# Repository Optimization Guide

This guide provides comprehensive strategies for keeping the OutlookAutoArchive repository optimized, clean, and efficient.

## 📊 Current Repository Status

- **Repository Size**: ~0.35 MB (Git objects)
- **Large Files**: 
  - `icon.ico` (160.70 KB)
  - `OutlookAutoArchive.exe` (116.00 KB)
  - `CHANGELOG.md` (34 KB)
  - `OutlookAutoArchive.ps1` (91 KB)

## 🧹 Housekeeping Scripts

### Automated Housekeeping
- **GitHub Actions Workflow**: `.github/workflows/housekeeping.yml`
  - Runs monthly on the 1st day at 2 AM UTC
  - Can be triggered manually via workflow_dispatch
  - Performs garbage collection, branch cleanup, and optimization

### Manual Housekeeping
- **PowerShell Script**: `housekeeping.ps1`
  - Comprehensive repository analysis and cleanup
  - Run with `-DryRun` flag to preview changes
  - Provides optimization recommendations

## 🔧 Optimization Strategies

### 1. Git Configuration Optimizations

```bash
# Enable compression for better storage efficiency
git config --global core.compression 9

# Optimize pack settings
git config --global pack.windowMemory "100m"
git config --global pack.packSizeLimit "100m"

# Enable delta compression
git config --global pack.deltaCacheSize 2047m
git config --global pack.deltaCacheLimit 1000
```

### 2. Regular Maintenance Commands

```bash
# Garbage collection
git gc --prune=now

# Repack objects for better compression
git repack -a -d --depth=250 --window=250

# Clean up reflog entries older than 30 days
git reflog expire --expire=30.days.ago --expire-unreachable=30.days.ago --all

# Prune remote references
git remote prune origin
```

### 3. Large File Management

#### Current Large Files
- `icon.ico` (160.70 KB) - Application icon
- `OutlookAutoArchive.exe` (116.00 KB) - Compiled executable

#### Optimization Options
1. **Keep as-is**: These files are essential and reasonably sized
2. **Git LFS**: For future large files (>50MB)
3. **Compression**: Consider UPX for executables in releases

### 4. File Organization

#### Binary Files
- Tracked: `*.exe`, `*.ico` (essential for releases)
- Ignored: `*.tmp`, `*.log`, `*.bak`

#### Documentation
- Optimized: Markdown files with proper formatting
- Versioned: All documentation in repository

## 📋 Monthly Maintenance Checklist

### Automated (GitHub Actions)
- [x] Garbage collection
- [x] Branch cleanup
- [x] Object repacking
- [x] Reflog cleanup

### Manual Review
- [ ] Review `.gitignore` for new patterns
- [ ] Check for large files that shouldn't be tracked
- [ ] Update documentation
- [ ] Review release assets

## 🚀 Performance Optimizations

### 1. Clone Optimization
```bash
# Shallow clone for CI/CD
git clone --depth 1 https://github.com/Zeffer83/OutlookAutoArchive.git

# Clone specific branch only
git clone --single-branch --branch master https://github.com/Zeffer83/OutlookAutoArchive.git
```

### 2. Fetch Optimization
```bash
# Fetch only necessary data
git fetch --depth 1

# Prune during fetch
git fetch --prune
```

### 3. Repository Size Monitoring
- Monitor `.git` directory size
- Track object count trends
- Review large file additions

## 🔍 Monitoring and Alerts

### Repository Health Metrics
- **Size**: Target < 1MB for `.git` directory
- **Objects**: Monitor object count growth
- **Large Files**: Alert on files > 100KB

### Automated Checks
- GitHub Actions workflow runs monthly
- Manual script provides detailed analysis
- Size monitoring in CI/CD pipeline

## 📈 Best Practices

### 1. Commit Strategy
- Use meaningful commit messages
- Squash commits when appropriate
- Avoid committing large binary files

### 2. Branch Management
- Delete merged branches promptly
- Use feature branches for development
- Regular cleanup of stale branches

### 3. File Management
- Keep `.gitignore` updated
- Use `.gitattributes` for file handling
- Consider Git LFS for large files

## 🛠️ Troubleshooting

### Common Issues
1. **Repository size growing rapidly**
   - Check for large files: `git rev-list --objects --all | git cat-file --batch-check='%(objecttype) %(objectname) %(objectsize) %(rest)' | sed -n 's/^blob //p' | sort -nr -k 2 | head -10`
   - Run garbage collection: `git gc --aggressive --prune=now`

2. **Slow clone times**
   - Use shallow clones for CI/CD
   - Consider Git LFS for large files
   - Optimize pack settings

3. **High object count**
   - Run object repacking: `git repack -a -d --depth=250 --window=250`
   - Clean up reflog entries
   - Review commit history for unnecessary commits

## 📚 Additional Resources

- [Git Documentation](https://git-scm.com/doc)
- [GitHub Actions Documentation](https://docs.github.com/en/actions)
- [Git LFS Documentation](https://git-lfs.github.com/)

## 🔄 Maintenance Schedule

### Daily
- Monitor CI/CD performance
- Check for failed workflows

### Weekly
- Review recent commits
- Check for new large files

### Monthly
- Run automated housekeeping
- Review optimization metrics
- Update documentation

### Quarterly
- Comprehensive repository audit
- Performance review
- Strategy updates

---

*Last updated: $(Get-Date -Format "yyyy-MM-dd")*
