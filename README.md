# VisiGrid

**The Explainable Spreadsheet.**

VisiGrid is the spreadsheet where you can prove why a value exists and how it got there. Every formula has a dependency graph. Every change has provenance. Every recompute is deterministic and verifiable.

Built as a native desktop app in Rust, powered by [GPUI](https://gpui.rs)—the GPU-accelerated UI framework behind [Zed](https://zed.dev)—for instant startup, smooth scrolling, and low-latency interaction.

## What Makes It Explainable

- **Verified Mode**: Toggle F9 to guarantee all values are current. No hidden stale cells.
- **Cell Inspector**: See precedents, dependents, evaluation order, and recompute timestamps.
- **Path Tracing**: Click any input to see exactly how data flows to outputs.
- **Provenance History**: Every paste, fill, and sort generates replayable Lua code.
- **Cycle Detection**: Circular dependencies caught at edit-time, not buried in #VALUE errors.

## Principles

- **Local-first**: Your data lives on your machine. No accounts required.
- **Native performance**: GPU-accelerated rendering. Smooth at any scale.
- **Explainable by default**: Trust is free. Causality is visible.
- **No lock-in**: Export freely. Files are standard formats, not hosted documents.

## Features (Free)

**Trust:**
- Verified Mode (F9) with status bar assertion
- Cycle detection at edit-time
- Basic Inspector: type, value, 1-hop dependency counts

**Core:**
- Full grid, selection, multi-selection, and editing
- 107 formula functions with autocomplete
- Command palette and keyboard-driven navigation
- Export to CSV, TSV, JSON, XLSX
- Cross-platform: macOS, Windows, Linux

See the [roadmap](ROADMAP.md) for what's next.

## Download

Get the latest release from the [Releases page](https://github.com/VisiGrid/VisiGrid/releases).

| Platform | Download |
|----------|----------|
| macOS (Universal) | `.dmg` |
| Windows (x64) | `.zip` |
| Linux (x86_64) | `.tar.gz` or `.AppImage` |

## Build from source

Requires [Rust](https://rustup.rs/) 1.80+.

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

For power users who need to explain and defend their work.

- Everything in Free, plus:
- **Deep Inspector**: DAG visualization, path tracing, evaluation timestamps
- **Provenance History**: View Lua code for every operation (Ctrl+Shift+Y)
- **Named Range Intelligence**: Usage tracking, dependency traces
- Advanced transforms (clean, split, dedupe, fill)
- Scripting and automation (Lua console)
- Large-file performance optimizations

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
