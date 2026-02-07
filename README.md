# VisiGrid

**A fast, keyboard-first, local-only spreadsheet for people who care about flow.**

VisiGrid is a native spreadsheet app that starts instantly, stays out of your way, and makes its work visible. It's built for Linux-first workflows, power users, and anyone tired of heavyweight, ribbon-driven spreadsheets.

Built in Rust, powered by [GPUI](https://gpui.rs) (the GPU-accelerated UI framework behind [Zed](https://zed.dev)).

## Why VisiGrid

Most spreadsheets feel heavy.

They take seconds to open. They hide actions behind menus. They encourage copy-paste instead of understanding.

VisiGrid is built to feel light, fast, and intentional:

- **Instant startup** (~300ms cold launch)
- **Local-only files** — no accounts, no cloud
- **Keyboard-first** navigation and editing
- **Command palette** for every action
- **Inspectable formulas** and dependencies
- **Works with existing files** — Excel, CSV, ODS, JSON, TSV

When correctness matters, VisiGrid also makes failure visible:

- **Deterministic recomputation** — same inputs, same outputs, always
- **Explicit dependencies** — precedents, dependents, and evaluation order are inspectable
- **Traceable changes** — structural edits generate provenance you can review and replay

For advanced workflows, VisiGrid also includes a CLI and headless mode built on the same engine.

VisiGrid was influenced by keyboard-first Linux environments such as [Omarchy](https://omarchy.com) — prioritizing speed, minimal friction, and staying in flow.

## Editing and Navigation

- Command palette for every action
- Keyboard-first navigation and editing
- Multi-select editing across non-adjacent cells
- Format Painter (single-shot and locked mode)
- 100+ formula functions with autocomplete
- Instant startup and smooth scrolling
- 5 built-in themes including System (follows OS dark/light)

## Design Principles

- **Local-first**: Your data lives on your machine. No accounts required.
- **Native performance**: GPU-accelerated rendering. Smooth at any scale.
- **Explainable by default**: Causality is visible. Trust is earned.
- **No lock-in**: Standard formats. Export freely.

## Explainability

The desktop app is the debugger for your data.

- **Cell Inspector** — view formulas, values, precedents, dependents, and recompute timestamps.
- **Path Tracing** — follow data flow across sheets and ranges.
- **Provenance History** — structural edits emit replayable scripts.
- **Cycle Detection** — circular dependencies caught at edit-time.
- **Deterministic recomputation** — explicit verification of stale vs current values.

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

## Advanced: Automation, CLI, and Reproducible Workflows

For users who treat spreadsheets as part of a larger system, VisiGrid includes a full CLI and headless execution mode built on the same engine.

### Headless Spreadsheet Workflows

VisiGrid ships a CLI that runs without a GUI. Same engine, no window.

```bash
# Evaluate a formula against piped data
cat sales.csv | visigrid-cli calc "=SUM(B:B)" --from csv

# Reconcile two datasets by key
visigrid-cli diff vendor.xlsx ours.csv --key Invoice --compare Total --tolerance 0.01

# Convert between formats
visigrid-cli convert data.xlsx --to csv

# Project a vendor export down to reconciliation columns, then diff
visigrid-cli convert vendor.xlsx -t csv --headers --select 'Invoice,Amount' | \
  visigrid-cli diff - our_export.csv --key Invoice --compare Amount --tolerance 0.01
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

**Column selection** (`convert --select`) — pick and reorder output columns:

```bash
visigrid-cli convert data.csv -t csv --headers --select 'Status,Amount'

# Filter by one column, output different ones
visigrid-cli convert data.csv -t csv --headers \
  --where 'Status=Pending' --select 'Amount,Vendor'
```

**Reconciliation** (`diff`) compares two datasets row-by-row:
- Rows only in the left file, only in the right file, or in both with value differences
- Numeric tolerance for financial data (`$1,234.56`, `(500.00)` handled natively)
- Either side can be `-` to read from stdin — pipe live exports directly into reconciliation
- Duplicate keys and ambiguous matches fail loudly instead of guessing

Example summary (from `--out json`):

```json
{
  "contract_version": 1,
  "summary": {
    "matched": 14238,
    "only_left": 12,
    "only_right": 9,
    "diff": 3,
    "diff_outside_tolerance": 1
  },
  "results": [ ... ]
}
```

### Session Control

Control a running VisiGrid GUI from the terminal. Inspect cells, apply changes, and watch state evolve — all from scripts or the command line.

```bash
# List running sessions
visigrid sessions

# View live grid snapshot (auto-refresh on changes)
visigrid view --follow

# Inspect a cell
visigrid inspect A1
# → A1 = 1234.56  (number)

# Apply operations with retry on contention
cat ops.jsonl | visigrid apply --atomic --wait

# Query server health
visigrid stats
```

**Session protocol**: TCP localhost with token auth. Protocol v1 is frozen — wire format locked by golden vectors.

**Scriptable control loop**:

```bash
# Get session
SESSION=$(visigrid sessions --json | jq -r '.[0].session_id')

# Inspect current state
REV=$(visigrid inspect workbook --json | jq '.revision')

# Apply changes with revision check (prevents stale overwrites)
visigrid apply ops.jsonl --atomic --expected-revision $REV --wait

# Verify new state
visigrid view --range A1:D10
```

**Exit codes** are stable for scripting: 0 = success, 20-29 = session errors (conflict, auth, protocol).

### Agents: Verifiable Spreadsheet Builds

VisiGrid provides a headless build loop for LLM agents and CI pipelines. Write Lua, build a `.sheet`, inspect results, verify fingerprint.

```bash
# Build from Lua script (replacement semantics — Lua is source of truth)
visigrid-cli sheet apply model.sheet --lua build.lua --json

# Inspect cells to verify results
visigrid-cli sheet inspect model.sheet B3 --json
# → {"cell":"B3","value":"220000","formula":"=SUM(B1:B2)","value_type":"formula"}

# Get fingerprint for audit trail
visigrid-cli sheet fingerprint model.sheet --json
# → {"fingerprint":"v1:42:abc123...","ops":42}

# Verify in CI (exit 0 = match, exit 1 = mismatch)
visigrid-cli sheet verify model.sheet --fingerprint v1:42:abc123...
```

**Fingerprint boundary**: `set()`, `clear()`, and `meta()` affect fingerprint. `style()` does not. Agents can format sheets without breaking verification.

**Workflow rule**: Always `apply → inspect → verify`. Never assume results.

See [Agent Tools](docs/agent-tools.json) for MCP definitions and [Claude MD Snippet](docs/claude-md-snippet.md) for copy-paste instructions.

### CI / Scripting

Exit 0 means reconciled (within tolerance). Exit 1 means material differences — missing rows or diffs outside tolerance. Exit ≥ 2 means error. No wrapper scripts needed.

**Bank statement vs ledger:**

```bash
# Reconcile bank export against your ledger
# --key-transform digits: match "INV-001" to "PO-001" by extracting "001"
visigrid-cli diff bank_export.csv ledger.csv \
  --key Reference --key-transform digits \
  --compare Amount --tolerance 0.01 \
  --out json
```

**Vendor export via stdin:**

```bash
# Pipe a live vendor export, project to reconciliation columns, then diff
curl -s https://vendor.example.com/api/export.csv | \
  visigrid-cli convert - -t csv --headers --select 'Invoice,Amount' | \
  visigrid-cli diff - our_export.csv --key Invoice --compare Amount --tolerance 0.01
```

**CI gate with strict exit:**

```bash
# In CI: fail the build if ANY value differs, even within tolerance
visigrid-cli diff expected.csv actual.csv \
  --key SKU --tolerance 0.01 --strict-exit --quiet || {
  echo "Reconciliation failed — diffs detected"
  exit 1
}
```

**More examples:**

```bash
# Quiet mode: just the exit code, no output
visigrid-cli diff expected.csv actual.csv --key id --quiet

# Pipe a live export into diff
rails runner 'Ledger.export_csv' | visigrid-cli diff - expected.csv --key id --quiet

# Verify a provenance trail hasn't been tampered with
visigrid-cli replay audit-trail.lua --verify --quiet
```

## Known Limitations (v0.4)

- **XLSX export** is not yet implemented — CLI writes CSV, TSV, JSON, .sheet
- **Replay**: layout operations (sort, column widths, merge) are hashed for fingerprint but not applied to workbook data
- **Nondeterminism detection** is conservative — `NOW()`, `TODAY()`, `RAND()`, `RANDBETWEEN()` fail `--verify` even in dead-code branches
- **Multi-sheet export** writes sheet 0 only
- **CLI `calc`** reads from stdin only; no file-path argument

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

**Diff bug reports**: if you file a bug against `visigrid diff`, include a minimal
CSV fixture that reproduces the issue. Every confirmed diff bug becomes a corpus
golden test in `tests/cli/diff/` — this is how the contract stays honest.

See the [Roadmap](docs/roadmap.md) for what's next.
