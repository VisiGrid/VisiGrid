#!/bin/bash
set -e

# VisiGrid Linux Bundle Script
# Creates AppImage, .deb, and .tar.gz packages

# Configuration
APP_NAME="VisiGrid"
BINARY_NAME="visigrid"

# Paths
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
WORKSPACE_DIR="$(dirname "$PROJECT_DIR")"
BUILD_DIR="$PROJECT_DIR/build"

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

# Parse arguments
RELEASE=true
CREATE_APPIMAGE=false
CREATE_DEB=false
CREATE_TARBALL=false

while [[ $# -gt 0 ]]; do
    case $1 in
        --debug)
            RELEASE=false
            shift
            ;;
        --appimage)
            CREATE_APPIMAGE=true
            shift
            ;;
        --deb)
            CREATE_DEB=true
            shift
            ;;
        --tarball)
            CREATE_TARBALL=true
            shift
            ;;
        --all)
            CREATE_APPIMAGE=true
            CREATE_DEB=true
            CREATE_TARBALL=true
            shift
            ;;
        --help)
            echo "Usage: $0 [options]"
            echo ""
            echo "Options:"
            echo "  --debug      Build debug instead of release"
            echo "  --appimage   Create AppImage"
            echo "  --deb        Create .deb package (requires cargo-deb)"
            echo "  --tarball    Create .tar.gz archive"
            echo "  --all        Create all package formats"
            echo "  --help       Show this help message"
            exit 0
            ;;
        *)
            echo -e "${RED}Unknown option: $1${NC}"
            exit 1
            ;;
    esac
done

# Default to tarball if nothing specified
if ! $CREATE_APPIMAGE && ! $CREATE_DEB && ! $CREATE_TARBALL; then
    CREATE_TARBALL=true
fi

echo -e "${GREEN}=== Building VisiGrid for Linux ===${NC}"
echo ""

# Build configuration
if $RELEASE; then
    BUILD_TYPE="release"
    CARGO_FLAGS="--release"
else
    BUILD_TYPE="debug"
    CARGO_FLAGS=""
fi

# Clean and create build directory
rm -rf "$BUILD_DIR"
mkdir -p "$BUILD_DIR"

cd "$WORKSPACE_DIR"

# Build
echo -e "${YELLOW}Building $BUILD_TYPE binary...${NC}"
cargo build $CARGO_FLAGS -p visigrid-gpui

BINARY_PATH="$WORKSPACE_DIR/target/$BUILD_TYPE/$BINARY_NAME"

# Get version
VERSION=$(grep '^version' "$WORKSPACE_DIR/Cargo.toml" | head -1 | sed 's/.*= *"\(.*\)".*/\1/')

echo -e "${GREEN}Built version: $VERSION${NC}"

# Create tarball
if $CREATE_TARBALL; then
    echo ""
    echo -e "${YELLOW}Creating tarball...${NC}"

    TARBALL_DIR="$BUILD_DIR/$APP_NAME-$VERSION-linux-x86_64"
    mkdir -p "$TARBALL_DIR"

    # Copy files
    cp "$BINARY_PATH" "$TARBALL_DIR/"
    cp "$WORKSPACE_DIR/assets/visigrid.desktop" "$TARBALL_DIR/"
    cp "$PROJECT_DIR/assets/icon-256.png" "$TARBALL_DIR/visigrid.png"

    # Create install script
    cat > "$TARBALL_DIR/install.sh" << 'INSTALL_EOF'
#!/bin/bash
set -e

PREFIX="${PREFIX:-/usr/local}"

echo "Installing VisiGrid to $PREFIX..."

# Install binary
sudo install -Dm755 visigrid "$PREFIX/bin/visigrid"

# Install desktop file
sudo install -Dm644 visigrid.desktop "$PREFIX/share/applications/visigrid.desktop"

# Install icon
sudo install -Dm644 visigrid.png "$PREFIX/share/icons/hicolor/256x256/apps/visigrid.png"

# Update desktop database
if command -v update-desktop-database &> /dev/null; then
    sudo update-desktop-database "$PREFIX/share/applications" 2>/dev/null || true
fi

# Update icon cache
if command -v gtk-update-icon-cache &> /dev/null; then
    sudo gtk-update-icon-cache -f -t "$PREFIX/share/icons/hicolor" 2>/dev/null || true
fi

echo "VisiGrid installed successfully!"
echo "You can now run 'visigrid' from your terminal or find it in your applications menu."
INSTALL_EOF
    chmod +x "$TARBALL_DIR/install.sh"

    # Create tarball
    cd "$BUILD_DIR"
    tar -czvf "$APP_NAME-$VERSION-linux-x86_64.tar.gz" "$(basename "$TARBALL_DIR")"
    rm -rf "$TARBALL_DIR"

    echo -e "${GREEN}Tarball created: $BUILD_DIR/$APP_NAME-$VERSION-linux-x86_64.tar.gz${NC}"
fi

# Create .deb
if $CREATE_DEB; then
    echo ""
    echo -e "${YELLOW}Creating .deb package...${NC}"

    if ! command -v cargo-deb &> /dev/null; then
        echo -e "${YELLOW}Installing cargo-deb...${NC}"
        cargo install cargo-deb
    fi

    cd "$PROJECT_DIR"
    cargo deb --no-build --target-dir "$WORKSPACE_DIR/target"

    # Move .deb to build directory
    mv "$WORKSPACE_DIR/target/debian"/*.deb "$BUILD_DIR/" 2>/dev/null || true

    echo -e "${GREEN}.deb package created in $BUILD_DIR/${NC}"
fi

# Create AppImage
if $CREATE_APPIMAGE; then
    echo ""
    echo -e "${YELLOW}Creating AppImage...${NC}"

    APPDIR="$BUILD_DIR/AppDir"
    mkdir -p "$APPDIR/usr/bin"
    mkdir -p "$APPDIR/usr/share/applications"
    mkdir -p "$APPDIR/usr/share/icons/hicolor/256x256/apps"

    # Copy files
    cp "$BINARY_PATH" "$APPDIR/usr/bin/"
    cp "$WORKSPACE_DIR/assets/visigrid.desktop" "$APPDIR/usr/share/applications/"
    cp "$WORKSPACE_DIR/assets/visigrid.desktop" "$APPDIR/"
    cp "$PROJECT_DIR/assets/icon-256.png" "$APPDIR/usr/share/icons/hicolor/256x256/apps/visigrid.png"
    cp "$PROJECT_DIR/assets/icon-256.png" "$APPDIR/visigrid.png"
    ln -sf visigrid.png "$APPDIR/.DirIcon"

    # Create AppRun
    cat > "$APPDIR/AppRun" << 'APPRUN_EOF'
#!/bin/bash
SELF=$(readlink -f "$0")
HERE=${SELF%/*}
export PATH="${HERE}/usr/bin:${PATH}"
exec "${HERE}/usr/bin/visigrid" "$@"
APPRUN_EOF
    chmod +x "$APPDIR/AppRun"

    # Download appimagetool if needed
    if [[ ! -f "$BUILD_DIR/appimagetool" ]]; then
        echo "Downloading appimagetool..."
        wget -q -O "$BUILD_DIR/appimagetool" "https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage"
        chmod +x "$BUILD_DIR/appimagetool"
    fi

    # Create AppImage
    cd "$BUILD_DIR"
    ARCH=x86_64 ./appimagetool AppDir "$APP_NAME-$VERSION-x86_64.AppImage"

    # Cleanup
    rm -rf "$APPDIR"

    echo -e "${GREEN}AppImage created: $BUILD_DIR/$APP_NAME-$VERSION-x86_64.AppImage${NC}"
fi

echo ""
echo -e "${GREEN}=== Build complete ===${NC}"
echo ""
echo "Output directory: $BUILD_DIR"
ls -la "$BUILD_DIR"
