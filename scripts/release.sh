#!/bin/bash

# Release script for docx-mcp
# This script helps with version management and release preparation

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Helper functions
info() {
    echo -e "${BLUE}ℹ️  $1${NC}"
}

success() {
    echo -e "${GREEN}✅ $1${NC}"
}

warning() {
    echo -e "${YELLOW}⚠️  $1${NC}"
}

error() {
    echo -e "${RED}❌ $1${NC}"
}

# Check if we're in a git repository
check_git_repo() {
    if ! git rev-parse --git-dir > /dev/null 2>&1; then
        error "Not in a git repository"
        exit 1
    fi
}

# Check if working directory is clean
check_clean_working_dir() {
    if ! git diff-index --quiet HEAD --; then
        error "Working directory is not clean. Please commit or stash your changes."
        exit 1
    fi
}

# Get current version from Cargo.toml
get_current_version() {
    grep '^version = ' Cargo.toml | head -1 | sed 's/version = "\(.*\)"/\1/'
}

# Update version in Cargo.toml
update_version() {
    local new_version=$1
    info "Updating version to $new_version"
    
    # Update Cargo.toml
    sed -i.bak "s/^version = \".*\"/version = \"$new_version\"/" Cargo.toml
    rm Cargo.toml.bak
    
    # Update Cargo.lock
    cargo update -p docx-mcp
    
    success "Version updated to $new_version"
}

# Generate changelog since last tag
generate_changelog() {
    local last_tag=$(git tag --sort=-version:refname | head -1)
    local new_version=$1
    
    info "Generating changelog since $last_tag"
    
    if [ -n "$last_tag" ]; then
        git log --pretty=format:"- %s (%h)" --no-merges ${last_tag}..HEAD > CHANGELOG.tmp
    else
        git log --pretty=format:"- %s (%h)" --no-merges > CHANGELOG.tmp
    fi
    
    echo "## Release $new_version ($(date +%Y-%m-%d))"
    echo ""
    cat CHANGELOG.tmp
    echo ""
    rm CHANGELOG.tmp
}

# Run pre-release checks
run_checks() {
    info "Running pre-release checks..."
    
    # Format check
    info "Checking code formatting..."
    cargo fmt --all -- --check
    success "Code formatting is correct"
    
    # Clippy check
    info "Running Clippy..."
    cargo clippy --all-targets --all-features -- -D warnings
    success "Clippy checks passed"
    
    # Tests
    info "Running tests..."
    cargo test --all-features
    success "All tests passed"
    
    # Build check
    info "Testing release build..."
    cargo build --release --all-features
    success "Release build successful"
    
    # Package check
    info "Testing package..."
    cargo package --dry-run
    success "Package validation passed"
}

# Create and push git tag
create_tag() {
    local version=$1
    local tag="v$version"
    
    info "Creating git tag $tag"
    
    # Create annotated tag
    git tag -a "$tag" -m "Release $tag"
    
    success "Created tag $tag"
    
    # Ask if user wants to push
    echo -n "Push tag to origin? [y/N]: "
    read -r response
    if [[ "$response" =~ ^[Yy]$ ]]; then
        git push origin "$tag"
        success "Tag pushed to origin"
    else
        warning "Tag not pushed. Remember to push it manually: git push origin $tag"
    fi
}

# Show usage information
usage() {
    cat << EOF
Usage: $0 [COMMAND] [OPTIONS]

Commands:
    patch           Bump patch version (0.1.0 -> 0.1.1)
    minor           Bump minor version (0.1.0 -> 0.2.0)
    major           Bump major version (0.1.0 -> 1.0.0)
    version X.Y.Z   Set specific version
    check           Run pre-release checks only
    changelog       Generate changelog since last tag
    tag             Create git tag for current version

Options:
    --dry-run       Show what would be done without making changes
    --no-checks     Skip pre-release checks (not recommended)
    --no-tag        Don't create git tag
    --help          Show this help message

Examples:
    $0 patch                    # Bump to next patch version
    $0 version 1.0.0           # Set version to 1.0.0
    $0 check                   # Run all pre-release checks
    $0 patch --dry-run         # Show what patch release would do
EOF
}

# Parse version bump type
bump_version() {
    local current_version=$1
    local bump_type=$2
    
    # Split version into components
    IFS='.' read -ra VERSION_PARTS <<< "$current_version"
    local major=${VERSION_PARTS[0]}
    local minor=${VERSION_PARTS[1]}
    local patch=${VERSION_PARTS[2]}
    
    case $bump_type in
        "patch")
            patch=$((patch + 1))
            ;;
        "minor")
            minor=$((minor + 1))
            patch=0
            ;;
        "major")
            major=$((major + 1))
            minor=0
            patch=0
            ;;
        *)
            error "Invalid bump type: $bump_type"
            exit 1
            ;;
    esac
    
    echo "${major}.${minor}.${patch}"
}

# Validate version format
validate_version() {
    local version=$1
    if ! [[ $version =~ ^[0-9]+\.[0-9]+\.[0-9]+(-[a-zA-Z0-9.-]+)?$ ]]; then
        error "Invalid version format: $version"
        error "Expected format: X.Y.Z or X.Y.Z-suffix"
        exit 1
    fi
}

# Main script logic
main() {
    local command=$1
    local dry_run=false
    local no_checks=false
    local no_tag=false
    
    # Parse arguments
    while [[ $# -gt 0 ]]; do
        case $1 in
            --dry-run)
                dry_run=true
                shift
                ;;
            --no-checks)
                no_checks=true
                shift
                ;;
            --no-tag)
                no_tag=true
                shift
                ;;
            --help)
                usage
                exit 0
                ;;
            *)
                if [ -z "$command" ]; then
                    command=$1
                elif [ -z "$version_arg" ] && [ "$command" = "version" ]; then
                    version_arg=$1
                fi
                shift
                ;;
        esac
    done
    
    # Check if command provided
    if [ -z "$command" ]; then
        usage
        exit 1
    fi
    
    # Basic checks
    check_git_repo
    
    if [ "$dry_run" = false ]; then
        check_clean_working_dir
    fi
    
    current_version=$(get_current_version)
    info "Current version: $current_version"
    
    case $command in
        "patch"|"minor"|"major")
            new_version=$(bump_version "$current_version" "$command")
            ;;
        "version")
            if [ -z "$version_arg" ]; then
                error "Version argument required for 'version' command"
                exit 1
            fi
            new_version=$version_arg
            validate_version "$new_version"
            ;;
        "check")
            run_checks
            success "All pre-release checks passed!"
            exit 0
            ;;
        "changelog")
            generate_changelog "$current_version"
            exit 0
            ;;
        "tag")
            if [ "$dry_run" = true ]; then
                info "Would create tag v$current_version"
            else
                create_tag "$current_version"
            fi
            exit 0
            ;;
        *)
            error "Unknown command: $command"
            usage
            exit 1
            ;;
    esac
    
    info "New version will be: $new_version"
    
    if [ "$dry_run" = true ]; then
        warning "DRY RUN MODE - No changes will be made"
        info "Would update version from $current_version to $new_version"
        if [ "$no_checks" = false ]; then
            info "Would run pre-release checks"
        fi
        if [ "$no_tag" = false ]; then
            info "Would create git tag v$new_version"
        fi
        exit 0
    fi
    
    # Confirm with user
    echo -n "Proceed with release $new_version? [y/N]: "
    read -r response
    if [[ ! "$response" =~ ^[Yy]$ ]]; then
        warning "Release cancelled"
        exit 0
    fi
    
    # Run pre-release checks
    if [ "$no_checks" = false ]; then
        run_checks
    fi
    
    # Update version
    update_version "$new_version"
    
    # Commit version bump
    git add Cargo.toml Cargo.lock
    git commit -m "Release $new_version"
    success "Version bump committed"
    
    # Create tag
    if [ "$no_tag" = false ]; then
        create_tag "$new_version"
    fi
    
    # Generate changelog for reference
    info "Changelog for release:"
    generate_changelog "$new_version"
    
    success "Release $new_version completed!"
    info "Next steps:"
    info "1. Push commits: git push origin main"
    if [ "$no_tag" = false ]; then
        info "2. Push tag: git push origin v$new_version (if not done already)"
    fi
    info "3. GitHub Actions will automatically create the release"
}

# Run main function with all arguments
main "$@"