# Release Guide

This document describes the release process for docx-mcp.

## Overview

The release process is automated using GitHub Actions and includes:

- Automated testing on multiple platforms
- Building release binaries for all supported targets
- Publishing to crates.io
- Creating GitHub releases with binaries
- Building and pushing Docker images
- Updating documentation

## Release Types

### Semantic Versioning

We follow [Semantic Versioning](https://semver.org/):

- **MAJOR**: Incompatible API changes
- **MINOR**: New features (backwards compatible)
- **PATCH**: Bug fixes (backwards compatible)

### Pre-release Versions

Pre-release versions can include suffixes like:
- `1.0.0-alpha.1` - Alpha releases
- `1.0.0-beta.1` - Beta releases  
- `1.0.0-rc.1` - Release candidates

## Quick Release Process

For most releases, use the automated release script:

```bash
# Patch release (1.0.0 -> 1.0.1)
./scripts/release.sh patch

# Minor release (1.0.0 -> 1.1.0)
./scripts/release.sh minor

# Major release (1.0.0 -> 2.0.0)  
./scripts/release.sh major

# Specific version
./scripts/release.sh version 1.5.0

# Pre-release
./scripts/release.sh version 1.0.0-beta.1
```

## Manual Release Process

If you need to create a release manually:

### 1. Pre-release Checks

```bash
# Run all checks
./scripts/release.sh check

# Or manually:
cargo fmt --all -- --check
cargo clippy --all-targets --all-features -- -D warnings
cargo test --all-features
cargo build --release --all-features
cargo package --dry-run
```

### 2. Update Version

Update the version in `Cargo.toml`:

```toml
[package]
version = "1.2.3"
```

Update `Cargo.lock`:

```bash
cargo update -p docx-mcp
```

### 3. Commit and Tag

```bash
git add Cargo.toml Cargo.lock
git commit -m "Release v1.2.3"
git tag -a "v1.2.3" -m "Release v1.2.3"
git push origin main
git push origin v1.2.3
```

### 4. GitHub Actions

The release workflow will automatically:

1. Validate the release
2. Run tests on all platforms
3. Build binaries for all targets
4. Create GitHub release
5. Publish to crates.io (stable releases only)
6. Build and push Docker images
7. Update documentation

## Supported Platforms

Release binaries are built for:

- **Linux**: x86_64-unknown-linux-gnu, x86_64-unknown-linux-musl
- **Linux ARM**: aarch64-unknown-linux-gnu, aarch64-unknown-linux-musl  
- **macOS**: x86_64-apple-darwin, aarch64-apple-darwin
- **Windows**: x86_64-pc-windows-msvc

## Docker Images

Docker images are published to:

- GitHub Container Registry: `ghcr.io/hongkongkiwi/docx-mcp`
- Docker Hub: `dockerhub-username/docx-mcp` (if configured)

Tags include:
- `latest` - Latest stable release
- `v1.2.3` - Specific version
- `1.2.3` - Semantic version
- `1.2` - Major.minor version
- `1` - Major version

## Publishing to crates.io

Stable releases (without pre-release suffixes) are automatically published to crates.io.

### Prerequisites

1. Set `CARGO_REGISTRY_TOKEN` secret in GitHub repository settings
2. Ensure you have publishing permissions for the crate

### Manual Publishing

```bash
# Dry run
cargo publish --dry-run

# Publish
cargo publish
```

## Troubleshooting

### Release Workflow Fails

1. Check the Actions tab in GitHub for detailed logs
2. Common issues:
   - Version mismatch between tag and Cargo.toml
   - Tests failing on specific platforms
   - Missing secrets (CARGO_REGISTRY_TOKEN, DOCKERHUB credentials)

### Version Already Exists

If you need to recreate a release:

1. Delete the tag: `git tag -d v1.2.3 && git push origin :v1.2.3`
2. Delete the GitHub release (if created)
3. Create the tag again

### Docker Build Fails

1. Check if all dependencies are available in the Docker environment
2. Verify Dockerfile syntax and build context
3. Test locally: `docker build -t docx-mcp:test .`

### crates.io Publishing Fails

1. Verify `CARGO_REGISTRY_TOKEN` is set and valid
2. Check if version already exists
3. Ensure all required metadata is in Cargo.toml
4. Run `cargo package --dry-run` to check for issues

## Security Considerations

### Signing Releases

Currently, releases are not cryptographically signed. Consider adding:

1. GPG signing of Git tags
2. Binary signing with platform-specific tools
3. SBOM (Software Bill of Materials) generation

### Supply Chain Security

- Dependencies are audited in CI with `cargo audit`
- Docker images use specific base image versions
- Build reproducibility is enhanced with Rust's deterministic builds

## Release Checklist

Use this checklist for important releases:

- [ ] All planned features are implemented
- [ ] All tests pass locally and in CI
- [ ] Documentation is updated
- [ ] Breaking changes are documented
- [ ] Migration guide is provided (for major releases)
- [ ] Security implications are reviewed
- [ ] Performance regression tests pass
- [ ] Cross-platform compatibility verified
- [ ] Release notes are prepared

## Post-Release Tasks

After a release:

1. **Verify Installation**: Test installation from released binaries
2. **Update Examples**: Update example configurations if needed
3. **Notify Users**: Announce significant releases
4. **Monitor Issues**: Watch for issues related to the new release
5. **Update Dependencies**: Consider updating dependent projects

## Emergency Releases

For critical security fixes:

1. Create a hotfix branch from the affected release tag
2. Apply minimal fix
3. Follow expedited release process
4. Consider yanking affected versions from crates.io if necessary

```bash
# Yank a version from crates.io (if needed)
cargo yank --version 1.2.3

# Un-yank if needed later
cargo yank --version 1.2.3 --undo
```

## Release Schedule

- **Patch releases**: As needed for bug fixes
- **Minor releases**: Monthly or when significant features accumulate  
- **Major releases**: Annually or when breaking changes are necessary

## Getting Help

- Open an issue for release-related problems
- Check GitHub Actions logs for CI failures
- Review this guide and workflow files for automation details