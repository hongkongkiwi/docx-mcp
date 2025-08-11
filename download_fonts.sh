#!/bin/bash

# Script to download open-source fonts for embedded PDF generation
# These fonts are used when creating PDFs without external dependencies

set -e

FONTS_DIR="assets/fonts"
mkdir -p "$FONTS_DIR"

echo "📥 Downloading open-source fonts for standalone operation..."

# Liberation Fonts (Red Hat) - Open source replacements for Arial, Times New Roman, Courier
LIBERATION_VERSION="2.1.5"
LIBERATION_URL="https://github.com/liberationfonts/liberation-fonts/files/7261482/liberation-fonts-ttf-${LIBERATION_VERSION}.tar.gz"

# Download Liberation fonts
echo "Downloading Liberation fonts..."
curl -L "$LIBERATION_URL" -o /tmp/liberation-fonts.tar.gz
tar -xzf /tmp/liberation-fonts.tar.gz -C /tmp/

# Copy the fonts we need
cp "/tmp/liberation-fonts-ttf-${LIBERATION_VERSION}/LiberationSans-Regular.ttf" "$FONTS_DIR/"
cp "/tmp/liberation-fonts-ttf-${LIBERATION_VERSION}/LiberationSans-Bold.ttf" "$FONTS_DIR/"
cp "/tmp/liberation-fonts-ttf-${LIBERATION_VERSION}/LiberationSans-Italic.ttf" "$FONTS_DIR/"
cp "/tmp/liberation-fonts-ttf-${LIBERATION_VERSION}/LiberationMono-Regular.ttf" "$FONTS_DIR/"

# Noto Sans (Google) - Fallback font with wide Unicode coverage
echo "Downloading Noto Sans fonts..."
NOTO_BASE_URL="https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSans"

curl -L "${NOTO_BASE_URL}/NotoSans-Regular.ttf" -o "$FONTS_DIR/NotoSans-Regular.ttf"
curl -L "${NOTO_BASE_URL}/NotoSans-Bold.ttf" -o "$FONTS_DIR/NotoSans-Bold.ttf"

# Clean up
rm -f /tmp/liberation-fonts.tar.gz
rm -rf /tmp/liberation-fonts-ttf-${LIBERATION_VERSION}

echo "✅ Fonts downloaded successfully!"
echo ""
echo "Fonts installed in $FONTS_DIR:"
ls -la "$FONTS_DIR"/*.ttf

echo ""
echo "The application can now run completely standalone without external dependencies!"