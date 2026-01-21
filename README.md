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

## Editions

### Free

Open source. Local-first. No lock-in — your data is always yours.

- Core grid + editing
- Multi-selection + multi-edit
- Command palette + quick open
- Formula autocomplete + error checking
- Themes + keybindings (JSON)
- Export data freely (CSV, TSV, JSON)

### Pro

For power users who want leverage: speed, scale, and automation.

- Everything in Free
- Fast large-file mode (million+ rows)
- Advanced transforms (clean, split, dedupe, fill)
- Inspector (dependency graphs + formula diagnostics)
- Scripting & automation (Lua)
- Plugin runtime

**$12/month · $99/year · $249 perpetual**

One license. No account. No telemetry. Yours forever.

[Get Pro](https://visigrid.app)

### Pro+

For serious work: continuity, confidence, and collaboration.

- Everything in Pro
- Includes VisiHub Pro
- Cloud sync & backup
- Version history
- Share without exporting (publish datasets)
- Priority support

**$19/month · $149/year**

[Get Pro+](https://visigrid.app)

## License

Source-available under the [FSL-1.1-MIT](LICENSE.md) license. After two years, each version converts to MIT.

## Contributing

Issues and PRs welcome.
