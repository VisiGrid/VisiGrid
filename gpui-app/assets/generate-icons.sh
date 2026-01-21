#!/bin/bash
# Generate PNG icons from icon.svg
# Requires: rsvg-convert (librsvg)
# On macOS: brew install librsvg
# On Linux: apt install librsvg2-bin

set -e
cd "$(dirname "$0")"

if ! command -v rsvg-convert &> /dev/null; then
    echo "Error: rsvg-convert not found. Install librsvg first."
    exit 1
fi

echo "Generating dark icon sizes..."
for size in 16 32 64 128 256 512 1024; do
    rsvg-convert -w $size -h $size icon.svg -o icon-${size}.png
    echo "  icon-${size}.png"
done

echo "Generating light icon sizes..."
for size in 16 32 64 128 256 512 1024; do
    rsvg-convert -w $size -h $size icon-light.svg -o icon-light-${size}.png
    echo "  icon-light-${size}.png"
done

echo "Generating macOS dark iconset..."
mkdir -p AppIcon.iconset
rsvg-convert -w 16 -h 16 icon.svg -o AppIcon.iconset/icon_16x16.png
rsvg-convert -w 32 -h 32 icon.svg -o AppIcon.iconset/icon_16x16@2x.png
rsvg-convert -w 32 -h 32 icon.svg -o AppIcon.iconset/icon_32x32.png
rsvg-convert -w 64 -h 64 icon.svg -o AppIcon.iconset/icon_32x32@2x.png
rsvg-convert -w 128 -h 128 icon.svg -o AppIcon.iconset/icon_128x128.png
rsvg-convert -w 256 -h 256 icon.svg -o AppIcon.iconset/icon_128x128@2x.png
rsvg-convert -w 256 -h 256 icon.svg -o AppIcon.iconset/icon_256x256.png
rsvg-convert -w 512 -h 512 icon.svg -o AppIcon.iconset/icon_256x256@2x.png
rsvg-convert -w 512 -h 512 icon.svg -o AppIcon.iconset/icon_512x512.png
rsvg-convert -w 1024 -h 1024 icon.svg -o AppIcon.iconset/icon_512x512@2x.png

echo "Generating macOS light iconset..."
mkdir -p AppIconLight.iconset
rsvg-convert -w 16 -h 16 icon-light.svg -o AppIconLight.iconset/icon_16x16.png
rsvg-convert -w 32 -h 32 icon-light.svg -o AppIconLight.iconset/icon_16x16@2x.png
rsvg-convert -w 32 -h 32 icon-light.svg -o AppIconLight.iconset/icon_32x32.png
rsvg-convert -w 64 -h 64 icon-light.svg -o AppIconLight.iconset/icon_32x32@2x.png
rsvg-convert -w 128 -h 128 icon-light.svg -o AppIconLight.iconset/icon_128x128.png
rsvg-convert -w 256 -h 256 icon-light.svg -o AppIconLight.iconset/icon_128x128@2x.png
rsvg-convert -w 256 -h 256 icon-light.svg -o AppIconLight.iconset/icon_256x256.png
rsvg-convert -w 512 -h 512 icon-light.svg -o AppIconLight.iconset/icon_256x256@2x.png
rsvg-convert -w 512 -h 512 icon-light.svg -o AppIconLight.iconset/icon_512x512.png
rsvg-convert -w 1024 -h 1024 icon-light.svg -o AppIconLight.iconset/icon_512x512@2x.png

# Generate .icns on macOS
if command -v iconutil &> /dev/null; then
    echo "Generating AppIcon.icns..."
    iconutil -c icns AppIcon.iconset
    echo "Generating AppIconLight.icns..."
    iconutil -c icns AppIconLight.iconset
fi

echo "Done."
