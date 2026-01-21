# VisiGrid

A fast, native spreadsheet built with [GPUI](https://gpui.rs) — the GPU-accelerated UI framework behind [Zed](https://zed.dev).

VisiGrid is a local-first, native spreadsheet for people who work with real data. It prioritizes speed, correctness, and keyboard-driven workflows over cloud lock-in.

## Principles

- **Local-first**: Your data lives on your machine. No accounts required.
- **Native performance**: GPU-accelerated rendering. Smooth at any scale.
- **Serious workflows**: Built for analysts, engineers, and operators.
- **No lock-in**: Export freely. Files are standard formats, not hosted documents.

## Features

- Multi-selection and multi-edit
- Formula bar with autocomplete and signature help
- Named ranges with rename refactoring
- 96 spreadsheet functions (financial, statistical, text, date/time, logical)
- Undo/redo with full history
- Command palette and quick open
- Excel, CSV, and TSV import
- Themes and custom keybindings (JSON)
- Cross-platform: macOS, Windows, Linux

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
git clone https://github.com/VisiGrid/VisiGrid.git
cd VisiGrid
cargo build --release -p visigrid-gpui
./target/release/visigrid
```

### Linux dependencies

```bash
# Ubuntu/Debian
sudo apt-get install libgtk-3-dev libxcb-shape0-dev libxcb-xfixes0-dev \
  libxkbcommon-dev libxkbcommon-x11-dev libwayland-dev
```

## VisiGrid Pro

VisiGrid is free and fully usable.

Pro unlocks capabilities for large datasets and automation:

- **Fast large-file mode** — million+ rows without lag
- **Lua scripting** — automate transforms and workflows
- **Inspector panel** — dependency graphs and formula diagnostics
- **Advanced transforms** — clean, split, dedupe, fill
- **Plugin runtime** — extend VisiGrid with custom tools

Pro is licensed locally:
- No account required
- No telemetry
- Works offline forever
- One license, yours to keep

**$12/month · $99/year · $249 perpetual**

[Get Pro →](https://visigrid.app)

## VisiGrid Pro+

Everything in Pro, plus cloud continuity:

- **VisiHub Pro** included
- **Cloud sync & backup** — never lose work
- **Version history** — know what changed
- **Publish datasets** — share without exporting
- **Priority support**

**$19/month · $149/year**

[Get Pro+ →](https://visigrid.app)

## License

VisiGrid is open source under the [AGPLv3](LICENSE.md) with a plugin exception.

This ensures improvements remain open while allowing commercial plugins and extensions. Plugins using the public API may be licensed independently. Commercial licenses available for embedding or hosting.

See [LICENSE.md](LICENSE.md) for details.

## Contributing

Issues and PRs welcome.
