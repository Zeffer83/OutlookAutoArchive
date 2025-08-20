# Repository Maintenance Guide

This document outlines the maintenance procedures and optimization strategies for the OutlookAutoArchive repository.

## Repository Optimization

### Current Repository Size
- **Total size**: ~278.75 KiB (after optimization)
- **Objects**: 286
- **Packs**: 1

### Large Files Analysis
The following files contribute most to the repository size:

1. **icon.ico** (160.70 KB) - Application icon
2. **OutlookAutoArchive.exe** (116.00 KB) - Main executable
3. **OutlookAutoArchive.ps1** (90.51 KB) - Source script
4. **CHANGELOG.md** (34.24 KB) - Version history
5. **setup_task_scheduler.exe** (32.00 KB) - Setup utility

## Optimization Strategies

### 1. Git Configuration
- **Garbage Collection**: Run `git gc --prune=now` regularly
- **Object Compression**: Use `git repack -a -d` for better compression
- **Reflog Cleanup**: Remove old reflog entries with `git reflog expire`

### 2. File Optimization
- **Binary Files**: Marked as binary in `.gitattributes` to prevent unnecessary diffs
- **Text Files**: Normalized line endings for consistent diffs
- **Large Files**: Consider Git LFS for files > 50MB (not currently needed)

### 3. Automated Cleanup
- **Monthly Cleanup**: Automated workflow runs on the 1st of each month
- **Release Optimization**: Assets are optimized when releases are published
- **Branch Cleanup**: Stale branches are automatically removed

## Maintenance Procedures

### Monthly Tasks
1. **Repository Cleanup**: Automated via GitHub Actions
2. **Size Monitoring**: Check repository size trends
3. **Asset Review**: Review if large files can be optimized

### Before Releases
1. **Garbage Collection**: Run `git gc --aggressive --prune=now`
2. **Asset Optimization**: Compress executables if possible
3. **Documentation Update**: Ensure all docs are current

### Repository Health Checks
```bash
# Check repository size
git count-objects -vH

# Check for large files
git ls-files | xargs ls -la | sort -k5 -nr | head -10

# Check for loose objects
git count-objects

# Analyze repository size
git rev-list --objects --all | git cat-file --batch-check='%(objecttype) %(objectname) %(objectsize) %(rest)' | sed -n 's/^blob //p' | sort -k2nr | head -10
```

## Best Practices

### File Management
- **Keep executables small**: Use compression tools like UPX
- **Optimize images**: Compress icons and graphics
- **Minimize documentation**: Keep docs concise but comprehensive

### Git Practices
- **Regular commits**: Avoid large single commits
- **Clean history**: Use interactive rebase to clean up history
- **Branch management**: Delete merged branches promptly

### Release Management
- **Asset optimization**: Compress release assets
- **Version tagging**: Use semantic versioning
- **Release notes**: Keep changelog updated

## Monitoring

### Size Thresholds
- **Warning**: Repository > 500 KB
- **Action Required**: Repository > 1 MB
- **Critical**: Repository > 5 MB

### Performance Metrics
- **Clone time**: Should be < 30 seconds
- **Fetch time**: Should be < 10 seconds
- **Push time**: Should be < 15 seconds

## Troubleshooting

### Common Issues
1. **Large repository size**: Run garbage collection
2. **Slow operations**: Check for loose objects
3. **Large downloads**: Optimize binary files

### Recovery Procedures
1. **Repository corruption**: Use `git fsck`
2. **Lost objects**: Check reflog for recovery
3. **Size issues**: Run aggressive cleanup

## Future Optimizations

### Planned Improvements
- [ ] Implement Git LFS for large assets
- [ ] Add automated size monitoring
- [ ] Create asset compression pipeline
- [ ] Implement delta compression for releases

### Monitoring Tools
- [ ] GitHub repository size alerts
- [ ] Automated cleanup notifications
- [ ] Performance benchmarking

---

*Last updated: January 2025*
*Maintained by: Ryan Zeffiretti*
