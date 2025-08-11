# Justfile for docx-mcp project
# Usage: just <command>
# Install just: https://github.com/casey/just

# Default recipe
default:
    @just --list

# Build the project
build:
    cargo build --all-features

# Build for release
build-release:
    cargo build --release --all-features

# Run all tests
test:
    ./scripts/run_tests.sh

# Run only unit tests
test-unit:
    ./scripts/run_tests.sh --unit-only

# Run only integration tests
test-integration:
    ./scripts/run_tests.sh --integration-only

# Run all tests including slow ones
test-all:
    ./scripts/run_tests.sh --all

# Run performance tests
test-performance:
    ./scripts/run_tests.sh --performance

# Generate coverage report
coverage:
    ./scripts/run_tests.sh --coverage

# Run benchmarks
bench:
    cargo bench --all-features

# Check code formatting
fmt-check:
    cargo fmt --all -- --check

# Format code
fmt:
    cargo fmt --all

# Run Clippy lints
clippy:
    cargo clippy --all-targets --all-features -- -D warnings

# Fix Clippy issues automatically where possible
clippy-fix:
    cargo clippy --all-targets --all-features --fix

# Run security audit
audit:
    cargo audit

# Check dependencies for issues
deny:
    cargo deny check

# Clean build artifacts
clean:
    cargo clean

# Update dependencies
update:
    cargo update

# Install development tools
install-dev-tools:
    cargo install cargo-audit cargo-deny cargo-llvm-cov

# Run the application in development mode
dev:
    RUST_LOG=debug cargo run --all-features

# Run the application in release mode
run:
    cargo run --release --all-features

# Generate documentation
docs:
    cargo doc --all-features --no-deps --open

# Check documentation
docs-check:
    cargo doc --all-features --no-deps

# Package the project for distribution
package:
    cargo package

# Publish to crates.io (dry run)
publish-dry:
    cargo publish --dry-run

# Publish to crates.io
publish:
    cargo publish

# Docker build
docker-build:
    docker build -t docx-mcp:latest .

# Docker run
docker-run:
    docker run -p 8080:8080 docx-mcp:latest

# Run pre-commit checks (formatting, linting, tests)
pre-commit: fmt-check clippy test-unit

# Full CI pipeline simulation
ci: pre-commit test audit

# Quick development cycle (format, build, test)
dev-cycle: fmt build test-unit

# Setup development environment
setup:
    rustup component add rustfmt clippy llvm-tools-preview
    just install-dev-tools

# Generate sample documents for testing
generate-samples:
    cargo run --bin generate-test-docs --features=test-utils

# Run stress tests
stress-test:
    STRESS_TEST=1 cargo test --release --test performance_tests -- --ignored --test-threads=1

# Profile the application
profile:
    cargo build --release --all-features
    perf record -g target/release/docx-mcp
    perf report

# Memory usage analysis
memory-check:
    cargo build --all-features
    valgrind --tool=memcheck --leak-check=full target/debug/docx-mcp

# Run with different Rust versions (requires rustup)
test-msrv:
    rustup install 1.70.0
    rustup run 1.70.0 cargo test

# Check for outdated dependencies
outdated:
    cargo install cargo-outdated
    cargo outdated

# Security scan
security-scan: audit deny

# Performance profiling with flamegraph
flamegraph:
    cargo install flamegraph
    cargo flamegraph --bin docx-mcp

# Generate changelog (requires git-cliff)
changelog:
    git cliff --output CHANGELOG.md

# Release commands using the release script

# Create a patch release (0.1.0 -> 0.1.1)
release-patch:
    ./scripts/release.sh patch

# Create a minor release (0.1.0 -> 0.2.0)
release-minor:
    ./scripts/release.sh minor

# Create a major release (0.1.0 -> 1.0.0)
release-major:
    ./scripts/release.sh major

# Create a specific version release
release-version version:
    ./scripts/release.sh version {{version}}

# Dry run of patch release (see what would happen)
release-patch-dry:
    ./scripts/release.sh patch --dry-run

# Dry run of minor release
release-minor-dry:
    ./scripts/release.sh minor --dry-run

# Dry run of major release
release-major-dry:
    ./scripts/release.sh major --dry-run

# Dry run of specific version release
release-version-dry version:
    ./scripts/release.sh version {{version}} --dry-run

# Run all pre-release checks
release-check:
    ./scripts/release.sh check

# Generate changelog since last tag
release-changelog:
    ./scripts/release.sh changelog

# Create git tag for current version
release-tag:
    ./scripts/release.sh tag

# Prepare a release (legacy command - use release-* commands above)
prepare-release version:
    @echo "⚠️  This command is deprecated. Use 'just release-version {{version}}' instead."
    @echo "The new release commands provide better automation and safety checks."
    @echo ""
    @echo "Available release commands:"
    @echo "  just release-patch        - Bump patch version"
    @echo "  just release-minor        - Bump minor version" 
    @echo "  just release-major        - Bump major version"
    @echo "  just release-version X.Y.Z - Set specific version"

# Show project statistics
stats:
    @echo "=== Project Statistics ==="
    @echo "Lines of code:"
    @find src -name "*.rs" -type f -exec wc -l {} + | tail -n 1
    @echo ""
    @echo "Test coverage:"
    @just coverage --quiet | grep "Overall coverage" || echo "Run 'just coverage' first"
    @echo ""
    @echo "Dependencies:"
    @cargo tree --depth 1 | wc -l
    @echo ""
    @echo "Binary size (release):"
    @if [ -f "target/release/docx-mcp" ]; then ls -lh target/release/docx-mcp | awk '{print $5}'; else echo "Run 'just build-release' first"; fi

# Watch for changes and run tests
watch:
    cargo install cargo-watch
    cargo watch -x "test --lib"

# Watch for changes and run specific test
watch-test test_name:
    cargo watch -x "test {{test_name}}"

# Initialize git hooks
init-hooks:
    #!/usr/bin/env bash
    cat > .git/hooks/pre-commit << 'EOF'
    #!/bin/bash
    just pre-commit
    EOF
    chmod +x .git/hooks/pre-commit
    echo "Pre-commit hook installed"

# Remove git hooks
remove-hooks:
    rm -f .git/hooks/pre-commit
    echo "Pre-commit hook removed"

# Docker commands

# Build multi-platform Docker image
docker-build-multiarch:
    docker buildx create --use --name multiarch || true
    docker buildx build --platform linux/amd64,linux/arm64 -t docx-mcp:latest .

# Build and tag Docker image for release
docker-build-release version:
    docker buildx build --platform linux/amd64,linux/arm64 \
        -t docx-mcp:{{version}} \
        -t docx-mcp:latest \
        -t ghcr.io/hongkongkiwi/docx-mcp:{{version}} \
        -t ghcr.io/hongkongkiwi/docx-mcp:latest \
        .

# Push Docker images to registry
docker-push version:
    docker push docx-mcp:{{version}}
    docker push docx-mcp:latest
    docker push ghcr.io/hongkongkiwi/docx-mcp:{{version}}
    docker push ghcr.io/hongkongkiwi/docx-mcp:latest

# Run Docker container with volume mount for testing
docker-test:
    docker run --rm -it -v $(pwd)/test-docs:/test-docs docx-mcp:latest

# Development environment commands

# Full development setup from scratch
dev-setup:
    # Install Rust if not present
    @if ! command -v rustup >/dev/null 2>&1; then \
        echo "Installing Rust..."; \
        curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y; \
        source ~/.cargo/env; \
    fi
    # Setup toolchain and tools
    just setup
    # Initialize git hooks
    just init-hooks
    # Build project
    just build
    echo "✅ Development environment ready!"

# Check system dependencies
check-deps:
    @echo "=== System Dependencies Check ==="
    @echo "Checking required tools..."
    @command -v rustc >/dev/null && echo "✅ Rust compiler found" || echo "❌ Rust compiler not found"
    @command -v cargo >/dev/null && echo "✅ Cargo found" || echo "❌ Cargo not found"
    @command -v git >/dev/null && echo "✅ Git found" || echo "❌ Git not found"
    @command -v docker >/dev/null && echo "✅ Docker found" || echo "❌ Docker not found"
    @command -v just >/dev/null && echo "✅ Just found" || echo "❌ Just not found"
    @echo ""
    @echo "Optional tools:"
    @command -v libreoffice >/dev/null && echo "✅ LibreOffice found" || echo "⚠️  LibreOffice not found (optional)"
    @command -v pdftoppm >/dev/null && echo "✅ pdftoppm found" || echo "⚠️  pdftoppm not found (optional)"
    @command -v convert >/dev/null && echo "✅ ImageMagick convert found" || echo "⚠️  ImageMagick not found (optional)"

# Cross-compilation commands

# Build for all supported targets
build-all-targets:
    # Install targets if not present
    rustup target add x86_64-unknown-linux-gnu
    rustup target add x86_64-unknown-linux-musl
    rustup target add aarch64-unknown-linux-gnu
    rustup target add x86_64-apple-darwin
    rustup target add aarch64-apple-darwin
    rustup target add x86_64-pc-windows-msvc
    # Build for each target
    cargo build --release --target x86_64-unknown-linux-gnu --all-features
    cargo build --release --target x86_64-unknown-linux-musl --all-features
    cargo build --release --target x86_64-apple-darwin --all-features
    @echo "✅ Built for all available targets"

# Build using cross for Linux targets
build-cross-linux:
    cargo install cross --git https://github.com/cross-rs/cross
    cross build --release --target x86_64-unknown-linux-gnu --all-features
    cross build --release --target x86_64-unknown-linux-musl --all-features  
    cross build --release --target aarch64-unknown-linux-gnu --all-features
    cross build --release --target aarch64-unknown-linux-musl --all-features

# Maintenance commands

# Update all dependencies to latest versions
update-deps:
    cargo update
    cargo outdated --depth 1

# Check for security vulnerabilities and update
security-update:
    cargo audit fix
    cargo update

# Clean everything (including registry cache)
clean-all:
    cargo clean
    rm -rf ~/.cargo/registry/cache
    rm -rf ~/.cargo/git/db
    docker system prune -f

# Backup project (excluding target and build artifacts)
backup:
    #!/usr/bin/env bash
    BACKUP_NAME="docx-mcp-backup-$(date +%Y%m%d-%H%M%S)"
    tar czf "${BACKUP_NAME}.tar.gz" \
        --exclude='target' \
        --exclude='.git' \
        --exclude='*.log' \
        --exclude='*.tmp' \
        .
    echo "✅ Backup created: ${BACKUP_NAME}.tar.gz"

# Development workflows

# Quick development loop (format, build, test unit, lint)
dev-loop:
    just fmt
    just build
    just test-unit
    just clippy

# Full quality check (everything CI runs)
quality-check:
    just fmt-check
    just clippy
    just test
    just docs-check
    just audit
    just deny

# Continuous development with file watching
dev-watch:
    cargo install cargo-watch
    cargo watch -w src -w tests -x "build" -x "test --lib"

# Performance analysis
perf-analysis:
    # Build optimized release
    cargo build --release --all-features
    # Run criterion benchmarks
    cargo bench --all-features
    # Generate flamegraph if available
    @if command -v flamegraph >/dev/null 2>&1; then \
        echo "Generating flamegraph..."; \
        cargo flamegraph --bin docx-mcp -- --help; \
    fi

# MCP-specific commands

# Test MCP server functionality
test-mcp:
    @echo "Testing MCP server..."
    # Build the server
    cargo build --release --all-features
    # Run basic functionality test
    python3 example/test_client.py || echo "❌ MCP test failed"

# Generate MCP documentation
mcp-docs:
    @echo "Generating MCP server documentation..."
    cargo run --bin docx-mcp -- --help > docs/CLI_REFERENCE.md
    @echo "✅ CLI reference updated"

# Example commands

# Run all examples
run-examples:
    @echo "Running all examples..."
    @if [ -f example/test_client.py ]; then python3 example/test_client.py; fi
    @if [ -f example/automation_example.py ]; then python3 example/automation_example.py; fi

# Generate test documents
gen-test-docs:
    @echo "Generating test documents..."
    mkdir -p test-docs
    # You could add commands here to generate various test DOCX files

# Utility commands

# Show detailed project info
info:
    @echo "=== Project Information ==="
    @echo "Name: docx-mcp"
    @echo "Version: $(grep '^version = ' Cargo.toml | head -1 | sed 's/version = "\(.*\)"/\1/')"
    @echo "Rust version: $(rustc --version)"
    @echo "Cargo version: $(cargo --version)"
    @echo ""
    just stats

# List all available commands with descriptions
help:
    @echo "=== Available Commands ==="
    @just --list
    @echo ""
    @echo "=== Release Commands ==="
    @echo "  release-patch      - Create patch release (0.1.0 -> 0.1.1)"
    @echo "  release-minor      - Create minor release (0.1.0 -> 0.2.0)"  
    @echo "  release-major      - Create major release (0.1.0 -> 1.0.0)"
    @echo "  release-version X  - Create specific version release"
    @echo "  release-*-dry      - Dry run versions of above commands"
    @echo ""
    @echo "=== Development Workflows ==="
    @echo "  dev-loop           - Quick development cycle"
    @echo "  quality-check      - Full quality assessment"
    @echo "  dev-setup          - Complete development environment setup"