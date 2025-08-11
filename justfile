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

# Prepare a release
prepare-release version:
    # Update version in Cargo.toml
    sed -i 's/^version = ".*"/version = "{{version}}"/' Cargo.toml
    # Update dependencies to use new version
    cargo update
    # Run full test suite
    just ci
    # Generate changelog
    just changelog
    # Commit changes
    git add .
    git commit -m "chore: prepare release {{version}}"
    git tag -a "v{{version}}" -m "Release {{version}}"

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