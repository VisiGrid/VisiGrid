# VisiGrid Windows Bundle Script
# Creates portable zip and optional MSI installer

param(
    [switch]$Debug,
    [switch]$Portable,
    [switch]$Msi,
    [switch]$All,
    [switch]$Help
)

$ErrorActionPreference = "Stop"

# Configuration
$AppName = "VisiGrid"
$BinaryName = "visigrid"

# Paths
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir
$WorkspaceDir = Split-Path -Parent $ProjectDir
$BuildDir = Join-Path $ProjectDir "build"

function Write-ColorOutput($ForegroundColor) {
    $fc = $host.UI.RawUI.ForegroundColor
    $host.UI.RawUI.ForegroundColor = $ForegroundColor
    if ($args) {
        Write-Output $args
    }
    $host.UI.RawUI.ForegroundColor = $fc
}

if ($Help) {
    Write-Output "Usage: .\bundle-windows.ps1 [options]"
    Write-Output ""
    Write-Output "Options:"
    Write-Output "  -Debug      Build debug instead of release"
    Write-Output "  -Portable   Create portable zip archive"
    Write-Output "  -Msi        Create MSI installer (requires WiX)"
    Write-Output "  -All        Create all package formats"
    Write-Output "  -Help       Show this help message"
    exit 0
}

# Default to portable if nothing specified
if (-not $Portable -and -not $Msi) {
    $Portable = $true
}

if ($All) {
    $Portable = $true
    $Msi = $true
}

Write-ColorOutput Green "=== Building VisiGrid for Windows ==="
Write-Output ""

# Build configuration
if ($Debug) {
    $BuildType = "debug"
    $CargoFlags = ""
} else {
    $BuildType = "release"
    $CargoFlags = "--release"
}

# Clean and create build directory
if (Test-Path $BuildDir) {
    Remove-Item -Recurse -Force $BuildDir
}
New-Item -ItemType Directory -Path $BuildDir | Out-Null

Set-Location $WorkspaceDir

# Build
Write-ColorOutput Yellow "Building $BuildType binary..."
cargo build $CargoFlags -p visigrid-gpui

$BinaryPath = Join-Path $WorkspaceDir "target\$BuildType\$BinaryName.exe"

# Get version from Cargo.toml
$CargoContent = Get-Content (Join-Path $WorkspaceDir "Cargo.toml") -Raw
if ($CargoContent -match 'version\s*=\s*"([^"]+)"') {
    $Version = $Matches[1]
} else {
    $Version = "0.1.0"
}

Write-ColorOutput Green "Built version: $Version"

# Create portable zip
if ($Portable) {
    Write-Output ""
    Write-ColorOutput Yellow "Creating portable zip..."

    $PortableDir = Join-Path $BuildDir "$AppName-$Version-windows-x64"
    New-Item -ItemType Directory -Path $PortableDir | Out-Null

    # Copy files
    Copy-Item $BinaryPath $PortableDir
    Copy-Item (Join-Path $ProjectDir "windows\visigrid.ico") $PortableDir

    # Create README
    @"
VisiGrid $Version - Portable Edition
=====================================

To run VisiGrid, double-click visigrid.exe

For command-line usage:
    visigrid.exe [file.csv]
    visigrid.exe [file.vsg]

For more information, visit: https://visigrid.com
"@ | Out-File -FilePath (Join-Path $PortableDir "README.txt") -Encoding UTF8

    # Create zip
    $ZipPath = Join-Path $BuildDir "$AppName-$Version-windows-x64.zip"
    Compress-Archive -Path $PortableDir -DestinationPath $ZipPath -Force

    # Cleanup
    Remove-Item -Recurse -Force $PortableDir

    Write-ColorOutput Green "Portable zip created: $ZipPath"
}

# Create MSI installer
if ($Msi) {
    Write-Output ""
    Write-ColorOutput Yellow "Creating MSI installer..."

    # Check for WiX
    $WixPath = Get-Command candle.exe -ErrorAction SilentlyContinue
    if (-not $WixPath) {
        Write-ColorOutput Yellow "WiX Toolset not found. Installing via cargo..."
        cargo install cargo-wix

        # Try cargo-wix instead
        Set-Location $ProjectDir
        cargo wix --no-build --nocapture

        Move-Item (Join-Path $WorkspaceDir "target\wix\*.msi") $BuildDir -ErrorAction SilentlyContinue
    } else {
        Write-ColorOutput Red "Manual WiX build not implemented. Use cargo-wix instead."
        Write-Output "Run: cargo install cargo-wix && cargo wix"
    }
}

Write-Output ""
Write-ColorOutput Green "=== Build complete ==="
Write-Output ""
Write-Output "Output directory: $BuildDir"
Get-ChildItem $BuildDir
