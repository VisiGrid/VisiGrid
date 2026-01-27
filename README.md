# VisiGrid

**A spreadsheet that behaves like code.**

VisiGrid is a native desktop spreadsheet designed for correctness under change.

It makes causality explicit: where values come from, why they changed, and whether the workbook is fully up to date. Structural edits are traceable, recomputation is deterministic, and circular dependencies are caught at edit-time.

Built in Rust, powered by [GPUI](https://gpui.rs) (the GPU-accelerated UI framework behind [Zed](https://zed.dev)). Designed as a deterministic computation engine with a GPU-accelerated UI, not a scriptable notebook or macro layer.

## Why VisiGrid

Most spreadsheets fail quietly.

A wrong reference. A missed row. A filter changed weeks ago.
The number still looks right — until it isn't.

VisiGrid is built to make these failures visible before they matter.

- **Causality is explicit**: precedents, dependents, and evaluation order are inspectable.
- **Changes are safe**: structural edits generate provenance you can review and replay.
- **State is verifiable**: you always know whether values are current or stale.

## What Makes VisiGrid Explainable

- **Verified Mode (F9)** — guarantees all values are current.
- **Cell Inspector** — view formulas, values, precedents, dependents, and recompute timestamps.
- **Path Tracing** — follow data flow across sheets and ranges.
- **Provenance History** — structural edits emit replayable Lua.
- **Cycle Detection** — circular dependencies caught at edit-time.

## Editing and Navigation

- Command palette for every action
- Keyboard-first navigation and editing
- Multi-select editing across non-adjacent cells
- 100+ formula functions with autocomplete
- Instant startup and smooth scrolling

## Design Principles

- **Local-first**: Your data lives on your machine. No accounts required.
- **Native performance**: GPU-accelerated rendering. Smooth at any scale.
- **Explainable by default**: Causality is visible. Trust is earned.
- **No lock-in**: Standard formats. Export freely.

## Download

Get the latest release from [Releases](https://github.com/VisiGrid/VisiGrid/releases).

| Platform | Download |
|----------|----------|
| macOS (Universal) | `.dmg` |
| Windows (x64) | `.zip` |
| Linux (x86_64) | `.tar.gz` / `.AppImage` |

Or via package manager:

```bash
# macOS
brew install visigrid/visigrid/visigrid

# Arch Linux
yay -S visigrid-bin
```

## Build from Source

Requires [Rust](https://rustup.rs/) 1.80+.

```bash
git clone https://github.com/VisiGrid/VisiGrid.git
cd VisiGrid
cargo build --release -p visigrid-gpui
./target/release/visigrid
```

### Linux Dependencies

```bash
# Ubuntu / Debian
sudo apt-get install libgtk-3-dev libxcb-shape0-dev libxcb-xfixes0-dev \
  libxkbcommon-dev libxkbcommon-x11-dev libwayland-dev
```

## Formats

- Import/export: CSV, TSV, JSON, XLSX, XLS, ODS
- Cross-platform: macOS, Windows, Linux

## Commercial Editions

VisiGrid is fully usable for real work in its open-source edition.

A commercial license is available for users and organizations that need advanced inspection, provenance, automation, and large-file support.

Commercial features include:

- Deep dependency inspection and full path tracing
- Complete provenance history with replayable operations
- Advanced data transforms and scripting
- Performance optimizations for large workbooks

Licenses are per-user, perpetual options are available, and no account or telemetry is required.

See [visigrid.app](https://visigrid.app) for details.

## VisiHub Integration

VisiGrid integrates with [VisiHub](https://visihub.io), a public-first publishing service for versioned datasets.

VisiHub is optional and not required to use VisiGrid.

## License

VisiGrid is open source under [AGPLv3](LICENSE.md) with a plugin exception.

This ensures improvements remain open while allowing commercial plugins and extensions. Plugins using the public API may be licensed independently. Commercial licenses are available for embedding or hosting.

See [LICENSE.md](LICENSE.md) for details.

## Contributing

Issues and pull requests are welcome.

See the [Roadmap](ROADMAP.md) for what's next.
