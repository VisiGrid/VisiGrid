# Changelog

## 0.3.7

### XLSX Import Overhaul

Importing Excel files with shared formulas, formula-only cells, and missing value cells now works correctly. Financial models that previously showed mass #CIRC! and #ERR errors now import cleanly.

- **calamine 0.26 → 0.32** — fixes shared formula expansion for `t="shared"` cells with ranges, absolute references, and column/row ranges
- **Topological recompute on import** — all formulas are re-evaluated in dependency order after the full graph is built, fixing stale cached values from per-cell evaluation during import
- **XML formula backfill** — cells with `<f>` elements but no cached `<v>` value (skipped by calamine) are now extracted directly from worksheet XML and backfilled
- **Shared formula follower expansion** — 2-pass XML parsing reconstructs follower formulas from master definitions with reference shifting (respects `$` absolute anchors)
- **XML value backfill** — cells stored as shared strings, inline strings, or numeric values that calamine drops are recovered from worksheet XML via shared string table resolution
- **Post-recalc error counting** — circular refs (`is_cycle_error()`) and formula errors (`Value::Error`) are counted after topo recalc, with up to 5 concrete examples in the import report
- **Import report diagnostics** — shows formula backfill count, value backfill count, shared formula groups, recalc errors with cell addresses and formulas
- **Status bar error count** — "Opened in Xms — 0 errors" or "Opened in Xms — N errors (Import Report)" with clickable access

### XLSX Formatting Import

Cell formatting from Excel files is now preserved on import.

- **styles.xml parsing** — number formats, font styles, fills, borders, alignment, column widths, and row heights extracted from XLSX style tables
- **Style deduplication** — `StyleTable` with `style_id` per cell, interning identical formats to minimize memory
- **Theme color approximation** — Excel theme colors with tint modifiers approximated to RGB with import report warnings
- **Column widths and row heights** — imported from XLSX and applied to sheet layout

### Number Format Rendering (ssfmt)

Custom Excel number format codes now render correctly via the `ssfmt` crate (ECMA-376 compliant, 99.99% SheetJS SSF compatibility).

- **Accounting formats** — zero values show `-` instead of `$-??` garbage; `_` padding and `*` fill handled correctly
- **Multi-section formats** — positive/negative/zero/text sections selected and rendered properly
- **Full format code support** — thousands separators, decimals, percent, currency, date/time tokens, conditional sections
- **Graceful fallback** — if ssfmt can't parse a code, falls back to VisiGrid's built-in formatter

### Formula Engine Additions

- **Unary plus** — `=+A1` parses correctly (common in financial models from Excel)
- **Power operator** — `A1^2` with right-associative precedence and fractional exponents
- **Percent operator** — `50%` → 0.5, works in expressions (`=A1*10%`)
- **IRR / XIRR** — Internal Rate of Return with Newton-Raphson iteration and bisection fallback
- **NPV** — Net Present Value
- **PMT / IPMT / PPMT** — Loan payment functions (total, interest, principal)
- **FV / PV** — Future Value and Present Value
- **CUMIPMT / CUMPRINC** — Cumulative interest and principal over a period range
- **IFNA / ISNA** — Error handling for #N/A values
- **XLOOKUP** — Modern lookup with exact match and custom default
- **TEXTJOIN** — Join text with delimiter and ignore-empty option
- **AVERAGEIF / AVERAGEIFS** — Conditional averaging
- **SPARKLINE** — Bar/line/winloss sparkline rendering in cells

### Recalc Engine

- **Value-typed computed cache** — `HashMap<(usize, usize), Value>` replaces String cache, eliminating lossy numeric conversions
- **No evaluate-on-cache-miss** — all getter paths return defaults on cache miss; only topo recalc and `set_value()` populate the cache
- **Correct recalc ordering** — workbook-level topological evaluation ensures upstream cells are computed before dependents

### Internal

- **Dependency upgrades** — calamine 0.32, quick-xml 0.38, zip 4, ssfmt 0.1, regex 1
- **quick-xml 0.38 migration** — `unescape()` → `decode()` in validation parser

## 0.3.6

### Tab-Chain Return (Excel-Style)

When you Tab across a row entering values, Enter returns you to the next row under the starting column — so you always know where you are in a model.

- **Origin tracking** — first Tab records the starting column; subsequent Tabs preserve it
- **Enter returns to origin** — moves down one row and snaps back to the column where the chain began
- **Shift+Enter returns upward** — same origin snap, one row up
- **Explicit chain breakers** — arrow keys, mouse clicks, Escape, dialog open, and sheet switch all reset the chain cleanly
- **Works in edit and navigation mode** — Tab without editing also builds a chain

### Navigation Snappiness

Improved keyboard navigation responsiveness. Arrow-key navigation now coalesces scroll updates per frame, reducing latency and improving feel — especially on Windows.

- **Repeat batching** — multiple arrow key repeats within a single frame batch into up to 4 cell moves, matching Excel cursor travel speed during held keys
- **Scroll coalescing** — scroll adjustment deferred to render start; multiple moves per frame compute viewport once
- **Latency instrumentation** — debug ring buffer (p50/p95) enabled via `VISIGRID_PERF=nav`, report via command palette with one-click copy
- **Measured** — key→render p50 ~5ms, p95 ~15ms; state update ~1µs

### Format Inspector Polish

Excel-grade formatting controls in the inspector panel.

- **Mixed-state visuals** — multi-cell selections show checkerboard fill chip, italic "(Mixed)" font, and "—" toggles when properties differ across cells
- **"Formats: mixed"** now checks all 10 user-facing properties (bold, italic, underline, strikethrough, font, alignment, vertical alignment, wrap, number format, background color)
- **Compact layout** — alignment section condensed to 2 rows (H-align + Wrap inline, V-align below), value preview collapses for empty cells, tighter gaps and padding
- **Keyboard hints** — platform-aware shortcut labels below Text Style (⌘B · ⌘I · ⌘U · ⌘⇧X) and Borders (⌘⇧7 Outline · ⌘⇧- Clear)
- **Clear Formatting** — resets all format properties to default, single undo step, command palette entry

### Borders Inspector

8 border presets accessible from the Format tab, matching Excel semantics.

- **Preset buttons** — None, Outline, All, Inside, Top, Bottom, Left, Right in a compact 2-row grid
- **Inside mode** — internal edges only (vertical separators as right edges, horizontal as bottom — aligned with precedence rules)
- **Single-edge presets** — Top/Bottom/Left/Right apply to the corresponding selection perimeter
- **Clear canonicalization** — clears inward-facing neighbor edges to prevent ghost borders
- **"Style: Thin" label** — signals future Medium/Thick support without UI overhead

### Format Painter

Copy formatting from one cell and apply it to others with a single click.

- **Click to capture** — reads the active cell's full CellFormat (bold, italic, fill, borders, number format, font, alignment — everything)
- **Click to apply** — next cell click applies the captured format and exits the mode
- **Esc to cancel** — exits painter mode without applying
- **Single undo step** — one Ctrl+Z reverses the entire paint operation
- **Command palette** — "Format Painter" with keywords: paint, format, brush
- **Inspector button** — side by side with Clear Formatting at the bottom of the Format tab

### Color Picker

Full color picker modal replaces the fixed swatch row in the inspector.

- **Fill Color picker** — 6x10 theme grid (tints/shades), 10 standard colors, No Fill, recent colors, hex input
- **Hex input** — type `#RRGGBB`, `#RGB`, or `rgb(R,G,B)` and press Enter to apply
- **Smart paste** — paste CSS like `color: #ff6600;` and the color token is extracted automatically
- **Shift+click** — apply a swatch without closing the picker
- **Recent colors** — last 10 picks, deduplicated, session-scoped
- **Pre-populated input** — hex field shows the current cell's color on open
- **Reusable architecture** — `ColorTarget` enum ready for text color and border color

### Internal

- **`UiState` boundary** — transient picker/dialog state separated from document view-model on `Spreadsheet`
- **`ui::text_input` helper** — shared input handling (typing, backspace, select-all, paste) for manual text fields
- **Removed ellipsis overlay** — cells clip silently via `overflow_hidden()`, matching Excel behavior

## 0.3.5

AI as witness, not author. Explainability without automation.

### Ask AI

Natural language questions about your data, answered with formulas you can verify.

- **Ask AI** (`Ctrl+Shift+A`) — describe what you want, get a formula proposal
- **Formula proposals, not edits** — AI suggests, you review, then insert with one click
- **Provenance tags** — every AI-inserted formula is marked with `MutationSource::Ai`
- **Inspector shows source** — purple "AI" badge and provider info visible in cell inspector
- **No hidden network calls** — AI only runs when you explicitly ask

### Explain Differences

Audit what changed since any point in history.

- **Right-click history entry → "Explain changes since..."** — opens diff dialog
- **Net change computation** — shows final state, not intermediate churn (A→B→C becomes A→C)
- **Grouped by type** — Values, Formulas, Structural, Named Ranges, Validation
- **AI-touched filter** — toggle to show only changes made by AI (purple badges)
- **Click to jump** — select any change to navigate to that cell
- **Copy Report** — plain text summary for Slack, email, or audit logs

### AI Summary (Optional)

- **Generate Summary** button — produces 4-8 sentence description of changes
- **Manual trigger only** — no surprise API calls
- **Copy button** — share summaries easily

### Explain This Change

- **Explain button** on selected diff entry — 2-4 sentence description of what changed
- **Cached per-entry** — no redundant API calls
- **Copy button** on each explanation
- **Tight prompt discipline** — describes only, never suggests edits

### Trust Model

Three concentric layers of explainability:

1. **Cell-level truth** — Inspector shows formula, value, inputs, dependents (deterministic, zero AI)
2. **Change-level accountability** — Every mutation tagged Human vs AI with metadata
3. **Narrative understanding** — AI describes changes but never modifies data from this context

AI is a witness, not an actor. It can explain what happened but cannot edit cells from the diff view.

### Configuration

- OpenAI provider support (more providers planned)
- API key stored in `~/.config/visigrid/settings.json`
- No telemetry, no account required

### Series Fill via Fill Handle

Excel-compatible series fill for non-formula cells.

**Pattern Detection:**
- **Single numbers copy by default** — drag `1` down → `1,1,1`. Hold **Ctrl** (Windows/Linux) or **Cmd** (macOS) to fill as series → `2,3,4`
- **Built-in lists auto-extend** — drag `Jan` → `Feb,Mar,Apr`. Hold modifier to copy instead
- **Two+ cell selections detect step** — select `1,3` and drag → `5,7,9`. Hold modifier to repeat pattern
- **Alphanumeric sequences** — `Item1` → `Item2,Item3`; negatives work (`Item-1` → `Item0`); leading zeros preserved (`001` → `002`)
- **Letter overflow** — `Row Z` → `Row AA,Row AB`; case preserved
- **Quarter/year patterns** — `Q4 2026` → `Q1 2027,Q2 2027`
- **Formulas unchanged** — `=A1` drag still uses reference adjustment (`=A2,=A3...`)
- **Single undo step** — any fill operation reverts in one undo

**Fill Handle UX Overhaul:**
- **Larger, easier to grab** — 10px visual size, 18px hit target (Fitts's Law)
- **Corner cap positioning** — handle overlaps selection border by 1px inward (feels like part of selection)
- **Solid dark fill** — uses selection border color, no transparency
- **Hover feedback** — subtle glow appears on hover, crosshair cursor confirms target
- **Border-only drag preview** — destination cells show outline only (clear "action in progress" signal)
- **Handle stays visible during drag** — anchors spatial understanding at source
- **First-use tip** — status bar hint on first drag ("Hold Ctrl/Cmd to toggle series/copy")

Behavior backed by 45 fill-related tests covering all core patterns and edge cases.

### KeyTips (macOS)

Excel-style keyboard accelerators for menu navigation, adapted for macOS:

- **Option+Space** opens KeyTips overlay showing accelerator hints
- Press a letter key to open scoped command palette for that category:
  - `F` → File (New, Open, Save, Export, VisiHub)
  - `E` → Edit (Undo, Redo, Cut, Copy, Paste, Find, Go To)
  - `V` → View (Inspector, Zen Mode, Zoom, Freeze Panes, Split)
  - `O` → Format (Bold, Italic, Fonts, Backgrounds, Borders)
  - `D` → Data (Fill, Sort, Filter, Validation)
  - `T` → Tools (Trace, Verified Mode, Ask AI)
  - `H` → Help (Shortcuts, About, Tour)
- **Stable mapping** — these letters are locked; commands may be added but categories won't move
- **Repeat last scope** — Enter or Space reopens last scoped palette (power-user speed)
- **Discovery hint** — first command palette open shows "Tip: ⌥Space shows KeyTips" (once per session)
- Auto-dismisses after 3 seconds or on Escape
- Discoverable via **Help > Keyboard Shortcuts** menu

Note: Uses Option+Space instead of double-tap Option due to gpui framework limitation (modifier-only key releases don't generate events). See ROADMAP.md for upstream tracking.

---

## 0.3.4

Excel-accurate behavior. Keyboard-first navigation. Layout is now auditable.

### Command Palette: Alt Scoping

Excel-style Alt scoping for the Command Palette:
- `Alt+A` → Data, `Alt+E` → Edit, `Alt+F` → File, `Alt+V` → View, `Alt+T` → Tools
- Scope hint bar appears on open, making Alt workflows discoverable
- Uppercase scope breadcrumb (`DATA ›`, `TOOLS ›`) for clear command context

Windows Excel users can bring their Alt muscle memory to macOS and Linux.

### Deterministic Recalculation (F9)

- `F9` now always recalculates (Excel-accurate)
- Verified Mode is no longer overloaded onto F9
- Status bar shows: `Recalculated · 1,284 cells · 12 ms · Verified`

### Navigation Ergonomics

- `Ctrl+Backspace` jumps viewport to active cell
- Arrow keys, Backspace, and typing correctly scoped inside modals
- Grid navigation fully blocked while dialogs are open

### Modal & Inspector Fixes

Inspector (`Ctrl+1`) now:
- Closes when clicking outside
- Fully captures clicks inside (no grid click-through)
- Disables grid interaction while open

Keyboard handling consistent across theme picker, font picker, named range editors, license dialogs, Find/GoTo/Rename flows.

### Excel-Accurate Text Overflow

Text spillover now matches Excel exactly:
- Only left-aligned (or General text) spills; center/right clips
- Two-pass rendering (background → overlay), deduplicated across freeze panes
- Visual ellipsis (`…`) when clipped
- Respects adjacent cell occupancy, alignment, and grid bounds

### Per-Sheet Column & Row Sizing

Column widths and row heights stored per sheet (Excel-correct):
- New sheets start with default sizing (no inheritance from active sheet)
- Duplicate Sheet = identical layout
- Correct behavior across sheet switching, named range jumps, trace navigation, session restore, async Excel import
- XLSX export writes correct per-sheet layout XML

### Layout Provenance

Layout changes are now first-class, undoable actions with full provenance:
- **Set column width** / **Set row height** actions in History
- Clear summaries: `Col A: default → 200`, `Row 5: 50 → default`
- Full undo/redo, Space-bar preview, rewind support
- Stable `SheetId` addressing (survives sheet delete/reorder)
- Provenance export as Lua:
  ```lua
  grid.set_col_width{ sheet_id=3, col="A", width=200 }
  grid.clear_row_height{ sheet_id=3, row=5 }
  ```

Layout is now part of the audit trail, not hidden UI state.

---

## 0.3.3

Explainable Spreadsheets: Trace formulas, validate inputs, navigate everything from the keyboard.

### Split View

- **Side-by-side comparison** (`Ctrl+\`) — view two regions of the same workbook simultaneously
- **Independent scroll and selection** per pane — each pane maintains its own view state
- **Shared data** — edits in one pane are immediately visible in the other
- **Close split** (`Ctrl+Shift+\`) — return to single pane, preserving active pane's state
- **Focus other pane** (`Cmd+]` on macOS, `Ctrl+`` on Windows/Linux)

### Trace Mode

- **Toggle trace** (`Alt+T` / `⌥T`) — highlight formula dependencies for the selected cell
- **Precedents** (inputs) shown in green tint, **dependents** (outputs) in purple
- **Jump to precedent** (`Ctrl+[` on Windows/Linux, `⌥[` on macOS) — cycle through inputs
- **Jump to dependent** (`Ctrl+]` on Windows/Linux, `⌥]` on macOS) — cycle through outputs
- **Return to source** (`F5` on Windows/Linux, `⌥↩` on macOS) — snap back to original cell
- **Invalid input warning** — status bar shows `⚠ N marked invalid (F8)` when precedents have validation failures
- **Theme-aware colors** — trace highlights respect your theme (Ledger Dark, Slate, Light, VisiCalc, Catppuccin)
- **Performance safe** — capped at 10,000 traced cells

### Trace + Validation Integration

- **Honest composition** — trace shows causality, validation shows integrity, no graph pollution
- **Actionable badge** — `F8` jumps to invalid inputs directly from trace mode
- **"Marked invalid"** wording — system is explicit about what it knows (no implied omniscience)
- **Modal-safe** — trace shortcuts blocked during validation dropdown or dialog

### Status Bar

Live affordance surface that teaches shortcuts:
```
Trace: A1 | 3 prec | 2 dep | Ctrl+[ ] | Back: F5 | ⚠ 2 marked invalid (F8)
```

---

## 0.3.1

Editing & Keyboard Polish

### Formula Editing

- **Multi-color reference highlighting** in the grid and formula bar with stable colors while typing (no color jumping during edit sessions)
- **Caret vs Point mode** for formula navigation:
  - Caret mode: arrow keys move the text cursor inside the formula
  - Point mode: arrow keys pick cell references
  - Auto-switching detects when you're at a ref insertion point (after `(`, `,`, operators)
  - F2 toggles modes with override latch (toggle sticks until you type)
- Arrow keys in formula mode now behave consistently across all directions

### Data Entry

- **Commit-on-arrow**: typing a value and pressing an arrow key commits the edit and moves selection — faster grid entry without extra keystrokes

### Clipboard

- **Multi-selection paste**: copying a single cell and pasting into a selected region fills all selected cells
- Formulas pasted to multiple cells adjust references relative to each destination cell

### Bug Fixes

- Sheet tab renaming is now reliable: double-click rename, cursor movement, selection, backspace/delete, enter/escape, click-away confirm all work correctly
- Fixed formula mode not activating immediately when typing `=`
- Fixed arrow keys not working after confirming a formula (no longer requires Escape first)

---

## 0.3.0

Preview the past. Rewind the workbook. Verify it in CI.

### Provenance History (Pro)

VisiGrid records *why* the grid changed — not just *that* it changed.

- **History tab** shows a git-log-style list of high-level actions (Paste, Fill, Sort, Clear, Multi-edit).
- **Select an entry** to view a canonical, read-only Lua snippet describing the action in deterministic A1 notation.
- **Copy the Lua snippet** to share, document, or audit how a sheet was produced.
- **Filter history** by label or scope.

**Pro feature:** View + Copy Lua provenance. Free users see the history list and can upgrade to export provenance.

**Shortcut:** `Ctrl+Shift+Y` opens History.

### Soft Rewind Preview (Pro)

Hold Space on a history entry to preview the workbook state *before* that action — without mutating anything.

- **Space + ↑/↓ scrubbing**: Navigate through history previews while holding Space.
- **Preview banner**: Clear "PREVIEW" status with action summary and position [N/M].
- **Preview safety gates**: Aborts with explicit errors on unsupported actions or integrity issues — no silent skips, no wrong previews.
- **Sort-aware preview**: Preview replays include sort state with lightweight `PreviewViewState` (no full app snapshot).

### Hard Rewind (Pro)

Commit to a preview and revert the workbook to that historical state.

- **"Rewind to here..."** button in history detail panel.
- **Two-phase commit**: Builds plan, validates fingerprint, then atomically applies.
- **Audit trail**: Appends a non-undoable `Rewind` record with 128-bit blake3 fingerprint.
- **Confirmation dialog**: Shows discard count and target context badge.

### CLI Replay (CI-Ready)

New command for headless provenance verification. Designed for CI pipelines: verify provenance scripts deterministically.

```bash
visigrid-cli replay script.lua --verify      # Verify fingerprint
visigrid-cli replay script.lua -o output.csv # Export to CSV/TSV/JSON
visigrid-cli replay script.lua --fingerprint # Print fingerprint only
visigrid-cli replay script.lua --verify -q   # CI mode (quiet)
```

- Full Lua bindings for all `grid.*` operations.
- Nondeterministic function detection (NOW, RAND, TODAY, RANDBETWEEN) — fails `--verify` with clear error.
- Golden test scripts in `tests/golden/`.

### Named Range Intelligence (Pro)

Named ranges are now first-class model entry points.

- **Named range details** show value preview, depth, verified status, and usage count.
- **Quick Open** (`Ctrl+P`) includes named ranges by default.
- **One-click tracing** highlights dependencies for the selected named range.

### Dialog Standardization

Reusable dialog primitives in `gpui-app/src/ui/`:

- `modal.rs` — backdrop + centering
- `dialog_frame.rs` — DialogFrame, DialogSize
- `button.rs` — consistent button styles

Migrated 10 dialogs to shared components.

### Also

- macOS: Added Help → License menu item
- Replay semantics match live app (validation replace-in-range, named range full payloads)
- History fingerprinting now cryptographic and order-sensitive (blake3)
- Theme refresh (contrast + readability improvements)
- Formula bar enhancements

---

## 0.2.7

- Pro Inspector trust-focused redesign

## 0.2.6

- Reliability & data-safety release

## 0.2.5

- AutoFilter
- Sort

## 0.2.4

- macOS polish
- Freeze panes fix

## 0.2.3

- Multi-color formula references
- F1 inspector peek

## 0.2.2

- Cell borders
- Paste Values (`Ctrl+Shift+V`)

## 0.2.1

- Fill handle
- Background colors

## 0.2.0

- Cross-sheet references
- Zoom controls
- 107 formula functions
