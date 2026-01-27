# Changelog

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
