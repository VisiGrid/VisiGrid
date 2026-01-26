# VisiGrid

**The spreadsheet that behaves like code.**

VisiGrid is a desktop spreadsheet built for people who need to move fast without losing correctness.

You can see where numbers come from.
You can tell why they changed.
You can clean and transform data without breaking everything else.

Every formula is traceable.
Every change has provenance.
Every recompute is deterministic and verifiable.

Built as a native desktop app in Rust, powered by [GPUI](https://gpui.rs) (the GPU-accelerated UI framework behind [Zed](https://zed.dev)), VisiGrid opens instantly, stays responsive at scale, and never hides what's happening.

## Why VisiGrid

Most spreadsheets fail quietly.

A wrong reference.
A missed row.
A filter changed weeks ago.

The number still looks right — until it isn't.

VisiGrid is designed to make causality visible and changes safe, so you can trust your work even as it evolves.

## Core Pillars

### Trust your numbers

- See precedents, dependents, and evaluation order
- Know when values are fully up to date
- Catch circular dependencies at edit-time
- Inspect why a value exists — not just what it is

### Clean data safely

- Validate columns against rules (types, ranges, allowed values)
- Find and fix problems before they ripple through formulas
- Preview transforms before committing
- Undo restores the entire change, not just one cell

### Work at the speed of thought

- Command palette for every action
- Keyboard-first navigation and editing
- Multi-select editing across non-adjacent cells
- Instant startup, smooth scrolling, no UI freezes

These aren't separate modes.
They're properties of the same system.

## What Makes It Explainable

**Verified Mode (F9)**
Guarantee all values are current. No hidden stale cells.

**Cell Inspector**
See formulas, values, precedents, dependents, and recompute timestamps.

**Path Tracing**
Click any input to see exactly how data flows to outputs — across sheets.

**Provenance History**
Structural edits (paste, fill, sort, transform) generate replayable Lua code.

**Cycle Detection**
Circular dependencies are caught at edit-time, not buried in #VALUE!.

## Design Principles

- **Local-first**: Your data lives on your machine. No accounts required.
- **Native performance**: GPU-accelerated rendering. Smooth at any scale.
- **Explainable by default**: Trust is free. Causality is visible.
- **No lock-in**: Files are standard formats. Export freely.

## Features (Free)

**Trust:**
- Verified Mode with status bar assertion
- Edit-time cycle detection
- Inspector: type, value, 1-hop dependency counts

**Core:**
- Full grid, selection, multi-selection, and editing
- 100+ formula functions with autocomplete
- Command palette and keyboard-driven navigation
- Import/export: CSV, TSV, JSON, XLSX
- Cross-platform: macOS, Windows, Linux

See the [Roadmap](ROADMAP.md) for what's next.

## Download

Get the latest release from [Releases](https://github.com/VisiGrid/VisiGrid/releases).

| Platform | Download |
|----------|----------|
| macOS (Universal) | `.dmg` |
| Windows (x64) | `.zip` |
| Linux (x86_64) | `.tar.gz` / `.AppImage` |

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

## VisiGrid Pro

For power users who need to explain and defend their work.

Everything in Free, plus:

- **Deep Inspector** (DAG visualization, full path tracing)
- **Full provenance history** (view Lua code for every operation)
- **Named range intelligence** and dependency tracking
- Advanced transforms (clean, split, dedupe, fill)
- Scripting and automation (Lua console)
- Large-file performance optimizations

**$12/month · $99/year · $249 perpetual**

One license.
No account.
No telemetry.

[Get Pro →](https://visigrid.app)

## VisiGrid Pro+

For workflows that need continuity.

Everything in Pro, plus:

- Cloud sync and backups (via VisiHub)
- Version history
- Dataset publishing
- Priority support

**$19/month · $149/year**

[Get Pro+ →](https://visigrid.app)

## License

VisiGrid is open source under [AGPLv3](LICENSE.md) with a plugin exception.

This ensures improvements remain open while allowing commercial plugins and extensions. Plugins using the public API may be licensed independently. Commercial licenses are available for embedding or hosting.

See [LICENSE.md](LICENSE.md) for details.

## Contributing

Issues and pull requests are welcome.
