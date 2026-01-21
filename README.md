# VisiGrid

A fast, native spreadsheet built with [GPUI](https://gpui.rs) (the framework behind [Zed](https://zed.dev)).

## Features

- **GPU-accelerated rendering** - smooth scrolling and editing at any scale
- **96 spreadsheet functions** - financial, statistical, text, date/time, and more
- **Excel/CSV/TSV import** - open existing files
- **Named ranges** - organize your data with semantic names
- **Formula bar with autocomplete** - IntelliSense-style function help
- **Undo/redo** - full edit history
- **Cross-platform** - macOS, Windows, Linux

## Download

Get the latest release from the [Releases page](https://github.com/VisiGrid/VisiGrid/releases).

| Platform | Download |
|----------|----------|
| macOS (Universal) | `.dmg` |
| Windows (x64) | `.zip` |
| Linux (x86_64) | `.tar.gz` or `.AppImage` |

## Build from source

Requires [Rust](https://rustup.rs/) 1.75+.

```bash
# Clone
git clone https://github.com/VisiGrid/VisiGrid.git
cd VisiGrid

# Build
cargo build --release -p visigrid-gpui

# Run
./target/release/visigrid
```

### Linux dependencies

```bash
# Ubuntu/Debian
sudo apt-get install libgtk-3-dev libxcb-shape0-dev libxcb-xfixes0-dev \
  libxkbcommon-dev libxkbcommon-x11-dev libwayland-dev
```

## VisiGrid Pro

VisiGrid is free for personal use. [VisiGrid Pro](https://visigrid.app) adds:

- Lua scripting
- Large file performance mode
- Priority support

## License

VisiGrid is source-available under the [FSL-1.1-MIT](LICENSE.md) license. After two years, each release converts to MIT.

## Contributing

Issues and PRs welcome.
