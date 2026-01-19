#!/bin/bash
set -e

# VisiGrid macOS Bundle Script
# Creates a signed, notarized .app bundle and optional DMG

# Configuration
APP_NAME="VisiGrid"
BUNDLE_ID="com.visigrid.app"
BINARY_NAME="visigrid"

# Paths (relative to gpui-app directory)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
WORKSPACE_DIR="$(dirname "$PROJECT_DIR")"
BUILD_DIR="$PROJECT_DIR/build"
BUNDLE_DIR="$BUILD_DIR/$APP_NAME.app"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Parse arguments
RELEASE=true
SIGN=false
NOTARIZE=false
CREATE_DMG=false
UNIVERSAL=true

while [[ $# -gt 0 ]]; do
    case $1 in
        --debug)
            RELEASE=false
            shift
            ;;
        --sign)
            SIGN=true
            shift
            ;;
        --notarize)
            NOTARIZE=true
            SIGN=true  # Notarization requires signing
            shift
            ;;
        --dmg)
            CREATE_DMG=true
            shift
            ;;
        --x86-only)
            UNIVERSAL=false
            ARCH="x86_64-apple-darwin"
            shift
            ;;
        --arm-only)
            UNIVERSAL=false
            ARCH="aarch64-apple-darwin"
            shift
            ;;
        --help)
            echo "Usage: $0 [options]"
            echo ""
            echo "Options:"
            echo "  --debug      Build debug instead of release"
            echo "  --sign       Code sign the app (requires APPLE_SIGNING_IDENTITY env var)"
            echo "  --notarize   Notarize the app (requires APPLE_ID, APPLE_TEAM_ID, APPLE_APP_PASSWORD)"
            echo "  --dmg        Create a DMG installer"
            echo "  --x86-only   Build only for Intel Macs"
            echo "  --arm-only   Build only for Apple Silicon"
            echo "  --help       Show this help message"
            exit 0
            ;;
        *)
            echo -e "${RED}Unknown option: $1${NC}"
            exit 1
            ;;
    esac
done

# Check we're on macOS
if [[ "$(uname)" != "Darwin" ]]; then
    echo -e "${RED}Error: This script must be run on macOS${NC}"
    exit 1
fi

# Check for required tools
command -v cargo >/dev/null 2>&1 || { echo -e "${RED}Error: cargo is required${NC}"; exit 1; }

if $SIGN; then
    if [[ -z "$APPLE_SIGNING_IDENTITY" ]]; then
        echo -e "${RED}Error: APPLE_SIGNING_IDENTITY environment variable is required for signing${NC}"
        echo "Set it to your Developer ID Application certificate name, e.g.:"
        echo '  export APPLE_SIGNING_IDENTITY="Developer ID Application: Your Name (TEAMID)"'
        exit 1
    fi
fi

if $NOTARIZE; then
    if [[ -z "$APPLE_ID" || -z "$APPLE_TEAM_ID" || -z "$APPLE_APP_PASSWORD" ]]; then
        echo -e "${RED}Error: APPLE_ID, APPLE_TEAM_ID, and APPLE_APP_PASSWORD are required for notarization${NC}"
        echo "APPLE_APP_PASSWORD should be an app-specific password from appleid.apple.com"
        exit 1
    fi
fi

echo -e "${GREEN}=== Building VisiGrid for macOS ===${NC}"
echo ""

# Clean previous build
rm -rf "$BUILD_DIR"
mkdir -p "$BUILD_DIR"

# Build configuration
if $RELEASE; then
    BUILD_TYPE="release"
    CARGO_FLAGS="--release"
else
    BUILD_TYPE="debug"
    CARGO_FLAGS=""
fi

cd "$WORKSPACE_DIR"

# Build binary/binaries
if $UNIVERSAL; then
    echo -e "${YELLOW}Building universal binary (x86_64 + aarch64)...${NC}"

    # Build for Intel
    echo "  Building for x86_64-apple-darwin..."
    cargo build $CARGO_FLAGS --target x86_64-apple-darwin -p visigrid-gpui

    # Build for Apple Silicon
    echo "  Building for aarch64-apple-darwin..."
    cargo build $CARGO_FLAGS --target aarch64-apple-darwin -p visigrid-gpui

    # Create universal binary
    echo "  Creating universal binary with lipo..."
    mkdir -p "$BUILD_DIR/universal"
    lipo -create \
        "target/x86_64-apple-darwin/$BUILD_TYPE/$BINARY_NAME" \
        "target/aarch64-apple-darwin/$BUILD_TYPE/$BINARY_NAME" \
        -output "$BUILD_DIR/universal/$BINARY_NAME"

    BINARY_PATH="$BUILD_DIR/universal/$BINARY_NAME"
else
    echo -e "${YELLOW}Building for $ARCH...${NC}"
    cargo build $CARGO_FLAGS --target "$ARCH" -p visigrid-gpui
    BINARY_PATH="$WORKSPACE_DIR/target/$ARCH/$BUILD_TYPE/$BINARY_NAME"
fi

echo ""
echo -e "${YELLOW}Creating app bundle...${NC}"

# Create .app structure
mkdir -p "$BUNDLE_DIR/Contents/MacOS"
mkdir -p "$BUNDLE_DIR/Contents/Resources"

# Copy binary
cp "$BINARY_PATH" "$BUNDLE_DIR/Contents/MacOS/$BINARY_NAME"

# Copy Info.plist and update version from Cargo.toml
VERSION=$(grep '^version' "$PROJECT_DIR/Cargo.toml" | head -1 | sed 's/.*= *"\(.*\)".*/\1/' || echo "0.1.0")
if [[ "$VERSION" == *"workspace"* ]]; then
    VERSION=$(grep '^version' "$WORKSPACE_DIR/Cargo.toml" | head -1 | sed 's/.*= *"\(.*\)".*/\1/')
fi

sed -e "s/0\.1\.0/$VERSION/g" "$PROJECT_DIR/macos/Info.plist" > "$BUNDLE_DIR/Contents/Info.plist"

# Copy icon if it exists
if [[ -f "$PROJECT_DIR/assets/AppIcon.icns" ]]; then
    cp "$PROJECT_DIR/assets/AppIcon.icns" "$BUNDLE_DIR/Contents/Resources/"
else
    echo -e "${YELLOW}Warning: No AppIcon.icns found in assets/ - app will use default icon${NC}"
fi

# Create PkgInfo
echo -n "APPLVSGD" > "$BUNDLE_DIR/Contents/PkgInfo"

echo -e "${GREEN}App bundle created at: $BUNDLE_DIR${NC}"

# Code signing
if $SIGN; then
    echo ""
    echo -e "${YELLOW}Code signing...${NC}"

    codesign --force --deep --sign "$APPLE_SIGNING_IDENTITY" \
        --entitlements "$PROJECT_DIR/macos/entitlements.plist" \
        --options runtime \
        --timestamp \
        "$BUNDLE_DIR"

    # Verify signature
    echo "Verifying signature..."
    codesign --verify --deep --strict --verbose=2 "$BUNDLE_DIR"

    echo -e "${GREEN}Code signing complete${NC}"
fi

# Notarization
if $NOTARIZE; then
    echo ""
    echo -e "${YELLOW}Notarizing (this may take a few minutes)...${NC}"

    # Create a zip for notarization
    NOTARIZE_ZIP="$BUILD_DIR/$APP_NAME-notarize.zip"
    ditto -c -k --keepParent "$BUNDLE_DIR" "$NOTARIZE_ZIP"

    # Submit for notarization
    xcrun notarytool submit "$NOTARIZE_ZIP" \
        --apple-id "$APPLE_ID" \
        --team-id "$APPLE_TEAM_ID" \
        --password "$APPLE_APP_PASSWORD" \
        --wait

    # Staple the ticket
    echo "Stapling notarization ticket..."
    xcrun stapler staple "$BUNDLE_DIR"

    # Clean up
    rm "$NOTARIZE_ZIP"

    echo -e "${GREEN}Notarization complete${NC}"
fi

# Create DMG
if $CREATE_DMG; then
    echo ""
    echo -e "${YELLOW}Creating DMG...${NC}"

    DMG_NAME="$APP_NAME-$VERSION.dmg"
    DMG_PATH="$BUILD_DIR/$DMG_NAME"
    DMG_TEMP="$BUILD_DIR/dmg-temp"

    # Create temp directory for DMG contents
    mkdir -p "$DMG_TEMP"
    cp -R "$BUNDLE_DIR" "$DMG_TEMP/"

    # Create symlink to Applications
    ln -s /Applications "$DMG_TEMP/Applications"

    # Create DMG
    hdiutil create -volname "$APP_NAME" \
        -srcfolder "$DMG_TEMP" \
        -ov -format UDZO \
        "$DMG_PATH"

    # Clean up
    rm -rf "$DMG_TEMP"

    # Sign DMG if signing is enabled
    if $SIGN; then
        codesign --force --sign "$APPLE_SIGNING_IDENTITY" "$DMG_PATH"
    fi

    echo -e "${GREEN}DMG created at: $DMG_PATH${NC}"
fi

echo ""
echo -e "${GREEN}=== Build complete ===${NC}"
echo ""
echo "Output:"
echo "  App: $BUNDLE_DIR"
if $CREATE_DMG; then
    echo "  DMG: $DMG_PATH"
fi
echo ""
echo "To test locally:"
echo "  open $BUNDLE_DIR"
