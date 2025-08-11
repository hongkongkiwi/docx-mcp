#!/bin/bash

# Build script for DOCX MCP Server

set -e

echo "🔨 Building DOCX MCP Server (Standalone Edition)..."

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Check for Rust
if ! command -v cargo &> /dev/null; then
    echo -e "${RED}❌ Cargo not found. Please install Rust.${NC}"
    echo "Visit: https://www.rust-lang.org/tools/install"
    exit 1
fi

# Check if fonts are downloaded
if [ ! -f "assets/fonts/LiberationSans-Regular.ttf" ]; then
    echo -e "${YELLOW}📥 Fonts not found. Downloading open-source fonts...${NC}"
    if [ -f "./download_fonts.sh" ]; then
        ./download_fonts.sh
    else
        echo -e "${YELLOW}⚠️  Font files not found. The server will still work but with basic fonts.${NC}"
        echo -e "${YELLOW}   Run ./download_fonts.sh to download professional fonts.${NC}"
        mkdir -p assets/fonts
        # Create placeholder files so build doesn't fail
        touch assets/fonts/LiberationSans-Regular.ttf
        touch assets/fonts/LiberationSans-Bold.ttf
        touch assets/fonts/LiberationSans-Italic.ttf
        touch assets/fonts/LiberationMono-Regular.ttf
        touch assets/fonts/NotoSans-Regular.ttf
        touch assets/fonts/NotoSans-Bold.ttf
    fi
fi

# Build mode selection
BUILD_MODE=${1:-release}
FEATURES=${2:-}

if [ "$BUILD_MODE" = "debug" ]; then
    echo -e "${YELLOW}📦 Building in debug mode...${NC}"
    if [ -n "$FEATURES" ]; then
        cargo build --features "$FEATURES"
    else
        cargo build
    fi
    BINARY_PATH="target/debug/docx-mcp"
else
    echo -e "${YELLOW}📦 Building in release mode...${NC}"
    if [ -n "$FEATURES" ]; then
        cargo build --release --features "$FEATURES"
    else
        cargo build --release
    fi
    BINARY_PATH="target/release/docx-mcp"
fi

# Check if build succeeded
if [ -f "$BINARY_PATH" ]; then
    echo -e "${GREEN}✅ Build successful!${NC}"
    echo -e "Binary location: ${GREEN}$BINARY_PATH${NC}"
    
    # Display standalone features
    echo -e "\n${BLUE}🎯 Standalone Features Enabled:${NC}"
    echo -e "${GREEN}✓${NC} Pure Rust DOCX parsing"
    echo -e "${GREEN}✓${NC} Built-in PDF generation"
    echo -e "${GREEN}✓${NC} Embedded fonts"
    echo -e "${GREEN}✓${NC} Native image processing"
    echo -e "${GREEN}✓${NC} Zero external dependencies required"
    
    # Check for optional enhancements
    echo -e "\n${YELLOW}Optional enhancements (not required):${NC}"
    
    if command -v libreoffice &> /dev/null; then
        echo -e "${GREEN}✓${NC} LibreOffice found (enhanced PDF conversion available)"
    else
        echo -e "${YELLOW}○${NC} LibreOffice not found (using built-in PDF converter)"
        echo "   Optional: brew install libreoffice (macOS) or apt-get install libreoffice (Linux)"
    fi
    
    if command -v pdftoppm &> /dev/null; then
        echo -e "${GREEN}✓${NC} pdftoppm found (PDF to image conversion available)"
    elif command -v convert &> /dev/null; then
        echo -e "${GREEN}✓${NC} ImageMagick found (PDF to image conversion available)"
    elif command -v gs &> /dev/null; then
        echo -e "${GREEN}✓${NC} Ghostscript found (PDF to image conversion available)"
    else
        echo -e "${YELLOW}○${NC} No PDF to image converter found"
        echo "   Install one of: poppler-utils, imagemagick, or ghostscript"
    fi
    
    # Create example output directories
    mkdir -p example/output example/images example/thumbnails
    
    echo -e "\n${GREEN}🚀 Ready to run!${NC}"
    echo -e "Start the server with: ${GREEN}$BINARY_PATH${NC}"
    echo -e "Or with logging: ${GREEN}RUST_LOG=info $BINARY_PATH${NC}"
else
    echo -e "${RED}❌ Build failed!${NC}"
    exit 1
fi