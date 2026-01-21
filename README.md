# VisiGrid

A fast, local-first spreadsheet for serious data work.

VisiGrid prioritizes speed, correctness, and keyboard-driven workflows over cloud lock-in and opaque automation.

Built as a native desktop app in Rust, powered by [GPUI](https://gpui.rs)—the GPU-accelerated UI framework behind [Zed](https://zed.dev)—for instant startup, smooth scrolling, and low-latency interaction.

## Principles

- **Local-first**: Your data lives on your machine. No accounts required.
- **Native performance**: GPU-accelerated rendering. Smooth at any scale.
- **Serious workflows**: Built for analysts, engineers, and operators.
- **No lock-in**: Export freely. Files are standard formats, not hosted documents.

## Features (Free)

- Core grid, selection, and editing
- Multi-selection and multi-edit
- Command palette
- Formula autocomplete and error reporting
- Fast search and navigation
- Export data freely (CSV, TSV, JSON)
- Cross-platform: macOS, Windows, Linux

Some advanced features are under active development. See the [roadmap](ROADMAP.md) for details.

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

For power users who want leverage: speed, scale, and automation.

- Everything in Free, plus:
- Large-file performance optimizations
- Advanced transforms (clean, split, dedupe, fill)
- Inspector tools (dependencies, diagnostics)
- Scripting and automation (Lua)
- Plugin runtime

**$12/month · $99/year · $249 perpetual**

One license. No account. No telemetry. Yours.

[Get Pro →](https://visigrid.app)

## VisiGrid Pro+

For workflows that need continuity.

- Everything in Pro, plus:
- Cloud sync and backups (via VisiHub)
- Version history
- Dataset publishing
- Priority support

**$19/month · $149/year**

[Get Pro+ →](https://visigrid.app)

## License

VisiGrid is open source under the [AGPLv3](LICENSE.md) with a plugin exception.

This ensures improvements remain open while allowing commercial plugins and extensions. Plugins using the public API may be licensed independently. Commercial licenses available for embedding or hosting.

See [LICENSE.md](LICENSE.md) for details.

## Contributing

Issues and PRs welcome.
