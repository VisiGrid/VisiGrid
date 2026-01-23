# VisiGrid Windows Bundle Script
# Creates portable zip, Inno Setup installer, and optional MSI installer

param(
    [switch]$Debug,
    [switch]$Portable,
    [switch]$Inno,
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
    Write-Output "  -Inno       Create Inno Setup installer (requires Inno Setup)"
    Write-Output "  -Msi        Create MSI installer (requires WiX)"
    Write-Output "  -All        Create all package formats"
    Write-Output "  -Help       Show this help message"
    exit 0
}

# Default to portable if nothing specified
if (-not $Portable -and -not $Inno -and -not $Msi) {
    $Portable = $true
}

if ($All) {
    $Portable = $true
    $Inno = $true
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

# Create Inno Setup installer
if ($Inno) {
    Write-Output ""
    Write-ColorOutput Yellow "Creating Inno Setup installer..."

    # Check for Inno Setup
    $InnoPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if (-not (Test-Path $InnoPath)) {
        $InnoPath = Get-Command ISCC.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
    }

    if (-not $InnoPath -or -not (Test-Path $InnoPath)) {
        Write-ColorOutput Red "Inno Setup not found!"
        Write-Output "Please install Inno Setup from: https://jrsoftware.org/isdl.php"
        Write-Output "Or add ISCC.exe to your PATH"
    } else {
        # Build CLI as well
        Write-ColorOutput Yellow "Building CLI..."
        cargo build $CargoFlags -p visigrid-cli

        # Create dist directory for output
        $DistDir = Join-Path $WorkspaceDir "dist"
        New-Item -ItemType Directory -Path $DistDir -Force | Out-Null

        # Run Inno Setup compiler
        $InstallerScript = Join-Path $WorkspaceDir "installer\VisiGrid.iss"
        & $InnoPath /DVersion=$Version $InstallerScript

        if ($LASTEXITCODE -eq 0) {
            $InstallerPath = Join-Path $DistDir "VisiGrid-Setup-x64.exe"
            if (Test-Path $InstallerPath) {
                # Move to build directory
                Move-Item $InstallerPath $BuildDir -Force
                Write-ColorOutput Green "Installer created: $(Join-Path $BuildDir 'VisiGrid-Setup-x64.exe')"
            }
        } else {
            Write-ColorOutput Red "Inno Setup compilation failed!"
        }
    }
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
