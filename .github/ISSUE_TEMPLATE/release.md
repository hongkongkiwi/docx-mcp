---
name: Release Checklist
about: Checklist for preparing a new release
title: 'Release v[VERSION]'
labels: 'release'
assignees: ''

---

## Pre-release Checklist

- [ ] All planned features and fixes are merged
- [ ] All tests are passing on main branch
- [ ] Documentation is updated
- [ ] CHANGELOG.md is updated (if maintained separately)
- [ ] Version is updated in Cargo.toml
- [ ] No critical security vulnerabilities in dependencies

## Release Process

- [ ] Run `./scripts/release.sh [patch|minor|major|version X.Y.Z]`
- [ ] Verify all CI checks pass
- [ ] Tag is created and pushed
- [ ] GitHub release is created automatically
- [ ] Binaries are built for all platforms
- [ ] Crate is published to crates.io (for stable releases)
- [ ] Docker images are pushed

## Post-release Tasks

- [ ] Verify release artifacts are available
- [ ] Test installation from released binaries
- [ ] Update any dependent projects
- [ ] Announce release (if applicable)

## Release Notes

<!--
Add release notes here:
- New features
- Bug fixes  
- Breaking changes
- Performance improvements
- Security fixes
-->

## Verification Commands

```bash
# Test the release script (dry run)
./scripts/release.sh patch --dry-run

# Run pre-release checks
./scripts/release.sh check

# Create the actual release
./scripts/release.sh patch  # or minor/major/version X.Y.Z
```