# VisiGrid

**A deterministic spreadsheet engine for reconciliation, audit, and reproducible computation.**

VisiGrid is a native spreadsheet engine designed for correctness under change. Recomputation is deterministic, structural edits are traceable, and circular dependencies are caught at edit-time.

Built in Rust, powered by [GPUI](https://gpui.rs) (the GPU-accelerated UI framework behind [Zed](https://zed.dev)).

## Why VisiGrid

Most spreadsheets fail quietly.

A wrong reference. A missed row. A filter changed weeks ago.
The number still looks right — until it isn't.

VisiGrid is built to make these failures visible before they matter.

- **Causality is explicit**: precedents, dependents, and evaluation order are inspectable.
- **Changes are safe**: structural edits generate provenance you can review and replay.
- **State is verifiable**: you always know whether values are current or stale.

## Determinism

Everything in VisiGrid is deterministic: same inputs, same outputs, always. No volatile surprises, no hidden state, no ambient context that changes results between runs.

The CLI is the primary interface for automation, verification, and audit trails. The GUI is an inspector and editor built on the same engine.

## Headless Spreadsheet Workflows

VisiGrid ships a CLI that runs without a GUI. Same engine, no window.

```bash
# Evaluate a formula against piped data
cat sales.csv | visigrid-cli calc "=SUM(B:B)" --from csv

# Reconcile two datasets by key
visigrid-cli diff vendor.xlsx ours.csv --key Invoice --compare Total --tolerance 0.01

# Convert between formats
visigrid-cli convert data.xlsx --to csv
```

The CLI reads spreadsheet files (CSV, XLSX, JSON, TSV), runs the same formula engine and comparison logic as the GUI, and writes structured output to stdout. Exit codes are stable for scripting. Output is JSON or CSV.

**Filtering rows** (`convert --where`) — no awk required:

```bash
# Pending transactions
visigrid-cli convert rh_transactions.csv -t csv --headers --where 'Status=Pending'

# Pending charges (negative amounts)
visigrid-cli convert rh_transactions.csv -t csv --headers \
  --where 'Status=Pending' --where 'Amount<0'

# Vendor name contains
visigrid-cli convert rh_transactions.csv -t csv --headers \
  --where 'Description~"google workspace"'
```

Five operators: `=` `!=` `<` `>` `~` (contains). Typed comparisons — numeric RHS triggers numeric compare, string RHS triggers case-insensitive string compare. Lenient parsing handles `$1,200.00`. Multiple `--where` = AND.

**Reconciliation** (`diff`) compares two datasets row-by-row:
- Rows only in the left file, only in the right file, or in both with value differences
- Numeric tolerance for financial data (`$1,234.56`, `(500.00)` handled natively)
- Duplicate keys and ambiguous matches fail loudly instead of guessing

Example output:

```json
{
  "matched": 14238,
  "only_left": 12,
  "only_right": 9,
  "changed": [
    { "cell": "D412", "before": 120.00, "after": 118.75 }
  ]
}
```

## Explainability

The desktop app is the debugger for your data.

- **Cell Inspector** — view formulas, values, precedents, dependents, and recompute timestamps.
- **Path Tracing** — follow data flow across sheets and ranges.
- **Provenance History** — structural edits emit replayable scripts.
- **Cycle Detection** — circular dependencies caught at edit-time.
- **Deterministic recomputation** — explicit verification of stale vs current values.

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

## AI Without Sacrificing Trust

Most AI tools trade explainability for convenience. VisiGrid doesn't.

**The problem**: AI in spreadsheets typically means black-box automation — formulas appear, values change, and you're expected to trust the result. When something goes wrong, there's no audit trail.

**VisiGrid's approach**: AI is a witness, not an author.

### Three Layers of Explainability

1. **Cell-level truth** — Inspector shows formula, value, inputs, dependents. Deterministic, local, zero AI involvement.

2. **Change-level accountability** — Every mutation is tagged: Human vs AI (with provider and timestamp). The diff engine surfaces net effects. AI-touched filter exposes exactly where AI participated.

3. **Narrative understanding** — "Explain this change" and "Explain differences" describe what happened in plain language. Both are optional, bounded, and never modify data.

### What AI Can Do

- Answer questions about your data with formula proposals
- Summarize what changed between two points in history
- Explain individual cell changes in 2-4 sentences

### What AI Cannot Do

- Edit cells without your explicit approval
- Run automatically or in the background
- Hide its participation (provenance is always visible)
- Suggest changes from the audit view

### Why This Matters

When you open a workbook six months from now, you can answer:
- Which values came from AI?
- What exactly did it change?
- Can I verify the formula it suggested?

The answer to all three is yes. That's what "explainable" means.

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
brew install --cask visigrid/tap/visigrid

# Windows
winget install VisiGrid.VisiGrid

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

- Import: CSV, TSV, JSON, XLSX, XLS, XLSB, ODS
- Export: CSV, TSV, JSON, .sheet (XLSX export planned)
- Cross-platform: macOS, Windows, Linux

## Known Limitations (v0.4)

- **XLSX export** is not yet implemented — CLI writes CSV, TSV, JSON, .sheet
- **Replay**: layout operations (sort, column widths, merge) are hashed for fingerprint but not applied to workbook data
- **Nondeterminism detection** is conservative — `NOW()`, `TODAY()`, `RAND()`, `RANDBETWEEN()` fail `--verify` even in dead-code branches
- **Multi-sheet export** writes sheet 0 only
- **CLI `calc`** reads from stdin only; no file-path argument

## CI / Scripting

```bash
# Reconcile two files in CI — non-zero exit on differences
visigrid-cli diff expected.csv actual.csv --key SKU --tolerance 0.01 --quiet || {
  echo "Reconciliation failed"
  exit 1
}

# Verify a provenance trail hasn't been tampered with
visigrid-cli replay audit-trail.lua --verify --quiet
```

## Commercial Use

VisiGrid is fully usable under its open-source license.

Some organizations require additional guarantees around scale, support, or long-term use. Commercial licenses are available for large-file performance, team controls, and operational assurances.

See [visigrid.app/commercial](https://visigrid.app/commercial) for details.

## VisiHub Integration

VisiGrid integrates with [VisiHub](https://visihub.io), a public-first publishing service for versioned datasets.

VisiHub is optional and not required to use VisiGrid.

## License

VisiGrid is open source under [AGPLv3](LICENSE.md) with a plugin exception.

This ensures improvements remain open while allowing commercial plugins and extensions. Plugins using the public API may be licensed independently. Commercial licenses are available for organizations that require alternative terms.

See [LICENSE.md](LICENSE.md) for details.

## Contributing

Issues and pull requests are welcome.

See the [Roadmap](ROADMAP.md) for what's next.
