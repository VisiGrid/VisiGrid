#!/bin/bash
set -e

# Generate macOS .icns from a source PNG
# Usage: ./generate-icons.sh [source.png]
#
# The source PNG should be at least 1024x1024 pixels.
# If no source is provided, uses assets/icon-1024.png

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
ASSETS_DIR="$PROJECT_DIR/assets"

SOURCE_PNG="${1:-$ASSETS_DIR/icon-1024.png}"
OUTPUT_ICNS="$ASSETS_DIR/AppIcon.icns"

if [[ ! -f "$SOURCE_PNG" ]]; then
    echo "Error: Source PNG not found: $SOURCE_PNG"
    echo ""
    echo "Please provide a 1024x1024 PNG icon."
    echo "Usage: $0 [path/to/icon.png]"
    echo ""
    echo "Or place your icon at: $ASSETS_DIR/icon-1024.png"
    exit 1
fi

# Check we're on macOS (required for iconutil)
if [[ "$(uname)" != "Darwin" ]]; then
    echo "Error: This script requires macOS (uses iconutil and sips)"
    exit 1
fi

echo "Generating macOS icon set from: $SOURCE_PNG"

# Create temporary iconset directory
ICONSET_DIR="$ASSETS_DIR/AppIcon.iconset"
rm -rf "$ICONSET_DIR"
mkdir -p "$ICONSET_DIR"

# Generate all required sizes
# macOS requires these specific sizes and naming
SIZES=(16 32 64 128 256 512 1024)

for size in "${SIZES[@]}"; do
    # Standard resolution
    sips -z $size $size "$SOURCE_PNG" --out "$ICONSET_DIR/icon_${size}x${size}.png" >/dev/null

    # Retina resolution (2x) - only for sizes up to 512
    if [[ $size -le 512 ]]; then
        retina_size=$((size * 2))
        half_size=$((size))
        sips -z $retina_size $retina_size "$SOURCE_PNG" --out "$ICONSET_DIR/icon_${half_size}x${half_size}@2x.png" >/dev/null
    fi
done

# Rename to match Apple's expected naming convention
cd "$ICONSET_DIR"
mv icon_16x16.png icon_16x16.png 2>/dev/null || true
mv icon_32x32.png icon_32x32.png 2>/dev/null || true
mv icon_64x64.png icon_32x32@2x.png 2>/dev/null || true
mv icon_128x128.png icon_128x128.png 2>/dev/null || true
mv icon_256x256.png icon_128x128@2x.png 2>/dev/null || true
mv icon_512x512.png icon_256x256@2x.png 2>/dev/null || true
mv icon_1024x1024.png icon_512x512@2x.png 2>/dev/null || true

# Also need 256x256 and 512x512 non-retina
sips -z 256 256 "$SOURCE_PNG" --out icon_256x256.png >/dev/null
sips -z 512 512 "$SOURCE_PNG" --out icon_512x512.png >/dev/null

cd "$ASSETS_DIR"

# Convert iconset to icns
iconutil -c icns "$ICONSET_DIR" -o "$OUTPUT_ICNS"

# Clean up
rm -rf "$ICONSET_DIR"

echo "Icon generated: $OUTPUT_ICNS"
echo ""
echo "Icon sizes included:"
echo "  16x16, 16x16@2x (32)"
echo "  32x32, 32x32@2x (64)"
echo "  128x128, 128x128@2x (256)"
echo "  256x256, 256x256@2x (512)"
echo "  512x512, 512x512@2x (1024)"
