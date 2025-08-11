#!/bin/bash

# Comprehensive test runner script for docx-mcp
# Usage: ./scripts/run_tests.sh [OPTIONS]

set -euo pipefail

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Default options
RUN_UNIT_TESTS=true
RUN_INTEGRATION_TESTS=true
RUN_PERFORMANCE_TESTS=false
RUN_BENCHMARKS=false
RUN_SECURITY_AUDIT=true
RUN_COVERAGE=false
VERBOSE=false
QUIET=false
CLEAN_FIRST=false

# Function to print colored output
print_status() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

print_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# Function to show help
show_help() {
    cat << EOF
Usage: $0 [OPTIONS]

Test runner script for docx-mcp project

OPTIONS:
    -h, --help              Show this help message
    -u, --unit-only         Run only unit tests
    -i, --integration-only  Run only integration tests
    -p, --performance       Include performance tests (slow)
    -b, --benchmarks        Run benchmarks (slow)
    -s, --skip-security     Skip security audit
    -c, --coverage          Generate coverage report
    -v, --verbose           Verbose output
    -q, --quiet             Quiet output (errors only)
    --clean                 Clean build artifacts first
    --all                   Run all tests including slow ones

Examples:
    $0                      # Run standard test suite
    $0 -u                   # Run only unit tests
    $0 --all                # Run all tests including performance
    $0 -c --verbose         # Generate coverage with verbose output
EOF
}

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        -h|--help)
            show_help
            exit 0
            ;;
        -u|--unit-only)
            RUN_UNIT_TESTS=true
            RUN_INTEGRATION_TESTS=false
            shift
            ;;
        -i|--integration-only)
            RUN_UNIT_TESTS=false
            RUN_INTEGRATION_TESTS=true
            shift
            ;;
        -p|--performance)
            RUN_PERFORMANCE_TESTS=true
            shift
            ;;
        -b|--benchmarks)
            RUN_BENCHMARKS=true
            shift
            ;;
        -s|--skip-security)
            RUN_SECURITY_AUDIT=false
            shift
            ;;
        -c|--coverage)
            RUN_COVERAGE=true
            shift
            ;;
        -v|--verbose)
            VERBOSE=true
            shift
            ;;
        -q|--quiet)
            QUIET=true
            shift
            ;;
        --clean)
            CLEAN_FIRST=true
            shift
            ;;
        --all)
            RUN_UNIT_TESTS=true
            RUN_INTEGRATION_TESTS=true
            RUN_PERFORMANCE_TESTS=true
            RUN_BENCHMARKS=true
            shift
            ;;
        *)
            print_error "Unknown option: $1"
            show_help
            exit 1
            ;;
    esac
done

# Set up output redirection based on quiet flag
if [ "$QUIET" = true ]; then
    CARGO_OUTPUT="--quiet"
else
    CARGO_OUTPUT=""
fi

if [ "$VERBOSE" = true ]; then
    CARGO_OUTPUT="$CARGO_OUTPUT --verbose"
fi

print_status "Starting docx-mcp test suite"

# Check if we're in the right directory
if [ ! -f "Cargo.toml" ]; then
    print_error "Cargo.toml not found. Please run this script from the project root."
    exit 1
fi

# Check if Rust is installed
if ! command -v cargo &> /dev/null; then
    print_error "Cargo not found. Please install Rust first."
    exit 1
fi

# Clean build artifacts if requested
if [ "$CLEAN_FIRST" = true ]; then
    print_status "Cleaning build artifacts..."
    cargo clean $CARGO_OUTPUT
fi

# Check formatting
print_status "Checking code formatting..."
if ! cargo fmt --all -- --check; then
    print_error "Code formatting issues found. Run 'cargo fmt' to fix."
    exit 1
fi
print_success "Code formatting OK"

# Run Clippy lints
print_status "Running Clippy lints..."
if ! cargo clippy --all-targets --all-features $CARGO_OUTPUT -- -D warnings; then
    print_error "Clippy lints failed"
    exit 1
fi
print_success "Clippy lints passed"

# Build the project
print_status "Building project..."
if ! cargo build --all-features $CARGO_OUTPUT; then
    print_error "Build failed"
    exit 1
fi
print_success "Build completed"

# Initialize test results tracking
TESTS_PASSED=0
TESTS_FAILED=0
FAILED_TESTS=()

# Function to run a test and track results
run_test() {
    local test_name="$1"
    local test_command="$2"
    
    print_status "Running $test_name..."
    
    if eval $test_command; then
        print_success "$test_name passed"
        ((TESTS_PASSED++))
    else
        print_error "$test_name failed"
        ((TESTS_FAILED++))
        FAILED_TESTS+=("$test_name")
    fi
}

# Run unit tests
if [ "$RUN_UNIT_TESTS" = true ]; then
    run_test "unit tests" "cargo test --lib $CARGO_OUTPUT"
    run_test "doc tests" "cargo test --doc $CARGO_OUTPUT"
fi

# Run integration tests
if [ "$RUN_INTEGRATION_TESTS" = true ]; then
    run_test "DOCX handler tests" "cargo test --test docx_handler_tests $CARGO_OUTPUT"
    run_test "MCP integration tests" "cargo test --test mcp_integration_tests $CARGO_OUTPUT"
    run_test "security tests" "cargo test --test security_tests $CARGO_OUTPUT"
    run_test "converter tests" "cargo test --test converter_tests $CARGO_OUTPUT"
    run_test "end-to-end workflow tests" "cargo test --test e2e_workflow_tests $CARGO_OUTPUT"
fi

# Run performance tests (if requested)
if [ "$RUN_PERFORMANCE_TESTS" = true ]; then
    print_warning "Running performance tests (this may take a while)..."
    run_test "performance tests" "cargo test --test performance_tests $CARGO_OUTPUT --release"
fi

# Run benchmarks (if requested)
if [ "$RUN_BENCHMARKS" = true ]; then
    print_warning "Running benchmarks (this may take a while)..."
    run_test "benchmarks" "cargo bench --all-features $CARGO_OUTPUT"
fi

# Run security audit
if [ "$RUN_SECURITY_AUDIT" = true ]; then
    print_status "Running security audit..."
    
    # Install cargo-audit if not present
    if ! command -v cargo-audit &> /dev/null; then
        print_status "Installing cargo-audit..."
        cargo install cargo-audit
    fi
    
    run_test "security audit" "cargo audit"
    
    # Check for denied dependencies if cargo-deny is available
    if command -v cargo-deny &> /dev/null; then
        run_test "dependency check" "cargo deny check"
    else
        print_warning "cargo-deny not found, skipping dependency checks"
    fi
fi

# Generate coverage report (if requested)
if [ "$RUN_COVERAGE" = true ]; then
    print_status "Generating coverage report..."
    
    # Check if cargo-llvm-cov is installed
    if ! command -v cargo-llvm-cov &> /dev/null; then
        print_status "Installing cargo-llvm-cov..."
        cargo install cargo-llvm-cov
    fi
    
    if cargo llvm-cov --all-features --workspace --html; then
        print_success "Coverage report generated in target/llvm-cov/html/"
        
        # Also generate lcov format for CI
        if cargo llvm-cov --all-features --workspace --lcov --output-path lcov.info; then
            print_success "LCOV format generated as lcov.info"
        fi
    else
        print_error "Coverage generation failed"
        ((TESTS_FAILED++))
        FAILED_TESTS+=("coverage generation")
    fi
fi

# Test different feature configurations
print_status "Testing different feature configurations..."

run_test "minimal features" "cargo test --no-default-features $CARGO_OUTPUT"
run_test "all features" "cargo test --all-features $CARGO_OUTPUT"

# Check that package builds for release
print_status "Verifying release build..."
run_test "release build" "cargo build --release --all-features $CARGO_OUTPUT"

# Verify package can be published (dry run)
print_status "Verifying package..."
run_test "package verification" "cargo package --dry-run $CARGO_OUTPUT"

# Print final results
echo ""
print_status "============ Test Results Summary ============"

if [ $TESTS_FAILED -eq 0 ]; then
    print_success "All tests passed! ($TESTS_PASSED/$((TESTS_PASSED + TESTS_FAILED)))"
    echo ""
    print_status "Ready for deployment! 🚀"
    exit 0
else
    print_error "Some tests failed! ($TESTS_PASSED passed, $TESTS_FAILED failed)"
    echo ""
    print_error "Failed tests:"
    for test in "${FAILED_TESTS[@]}"; do
        echo -e "  ${RED}✗${NC} $test"
    done
    echo ""
    print_error "Please fix the failing tests before proceeding."
    exit 1
fi