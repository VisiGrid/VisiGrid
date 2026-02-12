# Changelog

## 0.7.0

### Lua Debugger

- **Debug panel UI** — source pane with gutter breakpoints (click to toggle, F9 at paused line), current-line highlight, call stack with frame selection, and variables pane with lazy tree expansion (locals, upvalues, nested tables). Controls bar with Continue (F5), Step Over (F10), Step In (F11), Step Out (Shift+F11), Stop (Shift+F5).
- **Debug session backend** — background thread with its own Lua VM, hook-based pause/resume, breakpoint/stepping support, variable inspection per frame, and table expansion. Timeout and instruction limits exclude paused time.
- **Console debug integration** — Run and Debug tabs in the Lua console. Shift+Enter or F5 starts a debug session. Two-pass event pump streams output, pause snapshots, and completion results. Completed sessions apply ops as a single undo group.
- **Production hardening** — frame_index in VariableExpanded prevents race conditions on frame switch, debug output ring buffer (10K cap) prevents unbounded memory, viewport_lines recomputed every frame for stable scroll across resize/maximize.

### Lua Console

- **Redesigned console panel** — tab bar with Run/Debug tabs, toolbar with Clear, Maximize/Restore, and Close buttons. Viewport-relative max height (60% of window). Empty-state hint that disappears when output exists.
- **Pro gating** — Debug tab gated by `lua_tooling` feature flag (REPL stays free). Locked panel with skeleton preview and CTA for non-Pro users.

### CLI

- **Multi-sheet inspect** — `--sheet` (index or name), `--sheets` (list mode), `--non-empty` (sparse output), `--ndjson` (streamable). `convert --sheet` now works for .sheet and .xlsx.
- **diff --match contains** — `--contains-column` for substring matching, `--no-fail` for agent-friendly exit semantics, `KeyTransform::Alnum` for fuzzy key normalization.

### UI Polish

- **Panel toggle icons** — status bar icons for Inspector, Profiler, and Console with tooltips.
- **Save As defaults to .sheet** — imported files (xlsx, csv) now suggest `.sheet` extension to prevent format confusion.

### Advanced Transforms (0.6.9)

- **Transform pipeline** — Pro-gated column transforms with diff preview and history filter.
- **Ungated Lua runtime** — mlua always compiled, Lua REPL available to all users.

## 0.6.8

- **Performance Profiler panel** — debug recalculation bottlenecks with phase timing (invalidation, topo sort, eval, Lua), hotspot analysis, cycle detection, and heuristic classification. Ctrl+Alt+P to toggle.
- **Visible-but-locked Pro features** — inspector panel now shows skeleton previews and "Request Early Access" CTAs for locked Pro functionality instead of hiding features entirely. Reusable `LockedFeaturePanel` component.
- **Trial scaffolding** — local-only 14-day trial infrastructure with explicit feature allowlist. `is_feature_enabled()` now checks dev bypass, real license, and active trial in unified flow.

## 0.6.7

### Cross-Sheet Formula Navigation

- **Cross-sheet formula navigation** — when editing a formula referencing another sheet, arrow keys jump to that sheet to edit references, then return to origin. F2 correctly toggles between caret and point modes.
- **Fixed cross-sheet formula evaluation** — SUMIF, COUNTIF, AVERAGEIF, VLOOKUP, HLOOKUP, MATCH, INDEX, and XLOOKUP now read from the correct sheet instead of the formula's home sheet.

### VisiHub CI Integration

- **`vgrid login` and `vgrid publish`** — connect spreadsheets to VisiHub for versioned verification in CI pipelines. TTY-aware output: JSON when piped, human text when interactive.
- **`vgrid fill`** — load CSV into .sheet template with strict financial parsing (rejects currency symbols, commas, malformed decimals, formula injection).
- **`vgrid peek .sheet`** — interactive multi-sheet TUI viewer with Tab/Shift+Tab for sheet navigation and cell formula display in status bar.
- **`--assert-cell` on publish** — client-attested cell assertions with engine metadata and tolerance for model verification in CI.
- **Check policy system** — `--reset-baseline` flag, `check_policy` with warn severity level, CLI schema contract validated via golden tests.
- **Save Changes dialog** — keyboard focus trapping with Tab/Shift+Tab for button cycling.

## 0.6.6

- **Persist hidden rows and columns** — hide/unhide via Ctrl+9/0 (rows), Ctrl+Shift+9/0 (unhide). Hidden state is per-sheet, undoable/redoable, and skipped during rendering. Schema v6→v7 migration.
- **Fixed comma formatting** — `format_with_commas()` now places commas correctly (`$1,234.56` was `$,1234.56`).
- **Keyboard Parity v1 (Windows/Linux)** — complete Excel-compatible shortcuts: Alt+Enter for newline in cell, Shift+F10 for context menu, F6 for pane cycling, Ctrl+Alt+=/- for zoom, Ctrl+Shift+L for AutoFilter, Alt+F11 for Lua Console, Ctrl+Shift+* to select current region, Ctrl+9/0 to hide/unhide rows.

## 0.6.5

- **CLI binary renamed to `vgrid`** — shorter, easier to type.
- **Background CSV/TSV import** — large files import on background thread with delayed-overlay pattern. Fixed unnecessary 1.9GB buffer clone.
- **Cell Styles dropdown** — "Styles" button in format bar with presets for Normal, Error, Warning, Success, Input, Total, Note.

## 0.6.3

### Iterative Calculation & Cycle Handling

- **Iterative calculation (Jacobi iteration)** — resolve circular references with max iterations and convergence tolerance. Three-phase evaluation: upstream → SCC iteration → downstream topo sort. Non-converged cells get #NUM!.
- **Empty formula arguments** — formulas can now omit arguments (e.g., `=IF(a,b,)`) via Expr::Empty with proper coercion.
- **Blank cell semantics** — formula refs to empty cells return Empty instead of Number(0), fixing comparisons.
- **NORMSDIST / NORM.S.DIST** — normal distribution CDF/PDF.
- **SUMPRODUCT** — multiply corresponding cells across equal-shaped ranges and sum products.
- **Freeze Cycle Values** — XLSX imports with circular references freeze values instead of destroying formula ASTs.
- **Cycle state UX** — dismissible top banner, clickable ITER/FROZEN/CYCLES status pills, reframed import report with iteration toggle.
- **In-app Save Changes modal** — replaced OS-native dialog with themed in-app modal.
- **Number format quick-menu** — "123" dropdown in format bar with common number format presets.
- **Ctrl+Tab window cycling** — NextWindow/PrevWindow actions.
- **System theme default** — respect OS dark/light mode preference.

## 0.6.1

- **Minimap** — vertical density strip (View > Minimap) showing data distribution with click-to-jump and drag-to-scrub navigation.
- **Auto-detect CSV delimiter** — automatically detect semicolon, tab, pipe, or comma on import.
- **ODS OpenFormula support** — strip `=of:` prefix and convert semicolon arg separators.
- **ROUNDUP, ROUNDDOWN, TRUNC** — rounding functions.

## 0.6.0

- **Peek command** — interactive Ratatui-based TUI file viewer for CSV/TSV with column packing, cursor navigation, and unicode-aware alignment.
- **Native aarch64 CI** — switched aarch64 Linux builds from QEMU emulation to native ARM runner.
- **Flatpak distribution** — published to Flathub with manifest, metainfo, and screenshots.

## 0.5.9

- **Custom Functions v1** — user-defined Lua functions callable in formulas. Write `functions.lua` in `~/.config/visigrid/` and call `=ACCRUED_INTEREST(B2,C2,D2)` in any cell. Sandbox: no os/io/network, 100k instruction limit.
- **Format Painter v1** — copy formatting from one cell and apply to others. Single-shot or locked mode, Ctrl+Shift+C/V shortcuts, range painting via drag-select, full undo support.
- **VisiCalc inverse-video theme** — fully opaque green selection with black text matching the classic aesthetic.

## 0.5.8

- **aarch64 Linux builds** — ARM64 Linux now included in release workflow.
- **System theme** — follows OS dark/light preference.
- **Semantic cell styles** — Lua helpers and inspector hook for agent-built spreadsheets.
- **Colored border UI** — border color selection.

## 0.5.7

- **Formula bar polish** — format dropdown improvements.
- **15-digit precision limit** — number display limited to float64 significant digits.
- **Fixed Cmd+V paste** in Find and GoTo dialogs.

## 0.5.6

- **Semantic verification system** — model integrity verification.
- **Cross-sheet reference fix** — use actual sheet names.
- **`grid.name_sheet` API** — name sheets from Lua.
- **Multi-sheet .sheet format** — native format now stores multiple sheets.
- **Global Cmd+O** keybinding.

## 0.5.5

- **Semantic approval system** — model verification with approval, drift detection, and "Why drifted?" panel.
- **Date string parsing** — formulas can now parse date strings.
- **Role-based auto-styling** — agent-built spreadsheets get automatic formatting.
- **Semantic metadata persistence** — fingerprint boundary enforcement.

## 0.5.0 - 0.5.3

- **Windows build fixes** — HANDLE null check, reqwest/blake3 dependency fixes, windows-sys dependency.

## 0.4.7

### Agent-Ready Verifiable Builds

- **Sheet commands** (headless build loop):
  - `sheet apply` — Lua → .sheet (replacement semantics)
  - `sheet inspect` — read cells/ranges/workbook
  - `sheet fingerprint` — compute fingerprint
  - `sheet verify` — verify fingerprint (exit 0/1)
- **Lua API for agents** — `set()`, `clear()`, `meta()`, `style()` with fingerprint-affecting semantics.
- **Agent Kit** — MCP tool definitions, copy-paste instructions, demo script.

## 0.4.6

### Spreadsheet

- **Fixed formula reference shifting on paste (Linux/Wayland)** — copying a formula like `=A1+1` and pasting it one column over now correctly adjusts to `=B1+1`. This was a regression on Linux Wayland where clipboard metadata isn't preserved; the fallback detection incorrectly treated internal clipboard data as external, skipping reference adjustment. The fix adds a third fallback tier: when the system clipboard is unavailable but internal clipboard exists, treat as internal paste.

## 0.4.5

### Spreadsheet

- **Automatic recalculation** — dependent formulas now update immediately after edits, without requiring F9 or re-entering the cell. VisiGrid uses incremental dependency-based recalc: only cells in the dirty subgraph are re-evaluated, in topological order. Multi-cell operations (paste, fill, undo/redo, Lua scripts) batch into a single recalc pass. Manual calculation mode is respected when set via document settings; status bar shows "MANUAL CALC" when active.

- **Percent entry** — typing `1%` now stores the numeric value `0.01` and auto-applies Percent formatting (when the cell format is General). Handles whitespace, commas, and negatives: `" -1,000 % "` parses as `-10.0`. Applying Percent format to an existing text cell like `50%` converts it to the numeric `0.5`.

- **Formula entry parity** — `+` now behaves identically to `=` for starting formula mode. Arrow keys immediately enter ref-picking regardless of previous F2 toggle state, and Escape cleanly exits. Fixes an intermittent bug where formula navigation state leaked between edit sessions.

### Engine

- **Session server foundation** — engine-layer infrastructure for the upcoming session server (Phase 1 of AI agent support):
  - Revision tracking: `Workbook.revision()` increments exactly once per successful batch
  - Event types: `BatchApplied`, `CellsChanged`, `RevisionChanged` with revision tagging
  - Engine harness: `EngineHarness` for testing batch operations with event collection
  - 16 invariant tests defining the protocol contract (atomic rollback, event boundaries, fingerprint encoding)

## 0.4.4

### CLI

- **`convert --where` row filtering** — filter rows by column value before writing output. Supports five operators: `=` (typed equality), `!=`, `<`, `>`, `~` (case-insensitive contains). Multiple `--where` flags combine as AND. Typed comparisons: numeric RHS triggers numeric compare, string RHS triggers case-insensitive string compare — matching user intuition with zero extra syntax. Lenient numeric parsing strips `$` and `,` so bank/ledger exports work out of the box. Quoted values (`'Entity Name="Affinity House Inc"'`) handle columns and values with special characters. Header names are matched after trimming whitespace.

  ```bash
  # Filter to pending transactions
  vgrid convert data.csv -t csv --headers --where 'Status=Pending'

  # Negative amounts only
  vgrid convert data.csv -t csv --headers --where 'Amount<0'

  # AND: pending negative amounts
  vgrid convert data.csv -t csv --headers \
    --where 'Status=Pending' --where 'Amount<0'

  # Substring search
  vgrid convert data.csv -t csv --headers --where 'Vendor~cloud'

  # Pipe into calc
  vgrid convert data.csv -t csv --headers --where 'Status=Pending' | \
    vgrid calc '=SUM(E:E)' -f csv --headers
  ```

  When a numeric operator encounters a cell that doesn't parse as a number, the row is silently skipped. After output completes, a one-line stderr note reports the count (`note: 3 rows skipped (Amount not numeric)`). Suppressed by `--quiet`.

- **`convert --select` column projection** — select and reorder output columns by name. Requires `--headers`. Comma-separated or repeatable. Order of operations: parse → `--where` filter → `--select` projection → write. `--where` can reference columns not in `--select` (filtering happens before projection). JSON objects contain only the selected keys, emitted in `--select` order.

  ```bash
  # Pick two columns, reorder them
  vgrid convert data.csv -t csv --headers --select 'Status,Amount'

  # Filter by one column, output different columns
  vgrid convert data.csv -t csv --headers \
    --where 'Status=Pending' --select 'Amount,Vendor'

  # JSON with only selected fields
  vgrid convert data.csv -t json --headers --select 'Status,Amount'
  ```

  Error handling: unknown column → exit 2 with available headers; duplicate column → exit 2; ambiguous headers (case-insensitive collision) → exit 2 for both `--where` and `--select`.

- **`convert --quiet`** — new `-q` / `--quiet` flag suppresses stderr notes (e.g. skipped-row counts from `--where`). Designed for pipelines where only stdout matters.

- **`diff -` stdin support** — either side of `diff` can now be `-` to read from stdin. Format is inferred from the other file's extension, or set explicitly with `--stdin-format`. Enables piping live exports directly into reconciliation without temp files.

  ```bash
  # Pipe an export into diff
  cat export.csv | vgrid diff - baseline.csv --key id

  # Right side from stdin
  docker exec db dump.sh | vgrid diff expected.csv - --key sku

  # Explicit format when the other side has no extension
  cat data.tsv | vgrid diff - reference.csv --key id --stdin-format tsv
  ```

- **`diff` reconciliation exit-code semantics** — exit code 1 now indicates material differences only: missing rows or value diffs outside `--tolerance`. Within-tolerance diffs are reported in JSON (with `within_tolerance: true` and `diff_outside_tolerance` counter) but do not cause a non-zero exit code. This means `--tolerance 0.01` in CI passes when the only differences are rounding — no wrapper scripts needed. Use `--strict-exit` for Unix-diff semantics (any diff → exit 1).

- **`diff --save-ambiguous <path>`** — exports ambiguous matches to CSV before exiting. Columns: `left_key`, `candidate_count`, `candidate_keys` (pipe-separated). Written even when `--on-ambiguous error` causes exit 4, so the file is always available for manual review.

- **Golden tests: 52/52** — regenerated all stale expected outputs and locked down exit-code semantics with contract test, tolerance, strict-exit, and `--select` projection tests.

### Desktop App

- **Menu accelerator keys** — every menu item now has a keyboard accelerator. When a dropdown is open, press the underlined letter to execute that item immediately. Accelerators are explicitly assigned per item (not auto-derived) and validated at startup in debug builds to catch collisions. Works across all menus: File, Edit, View, Format, Data, Help.

- **Copy/cut dashed border** — copying or cutting cells now draws a dashed border overlay around the source range, matching Excel's visual feedback. The border uses the selection accent color and is cleared on paste, escape, edit start, edit confirm, or delete. Non-interactive (clicks pass through).

- **Ctrl+W close window** — Windows/Linux keybinding to close the current window, matching platform convention.

## 0.4.2

CLI hardening and documentation truthfulness for public launch.

### CLI Error Messages

Every CLI error now prints an actionable hint below the error line. The goal: a user who hits an error can fix it without leaving the terminal.

- **`error:` / `hint:` pattern** — `CliError` gains an optional `hint` field. When present, printed on a second line after the error message. All error constructors (`args`, `io`, `parse`, `format`, `eval`) default to `hint: None`; callers chain `.with_hint()` to add guidance.
- **Empty stdin** — `no input received on stdin` now shows a working pipe example (`cat file.csv | vgrid calc '=SUM(A:A)' --from csv`).
- **Formula errors** — context-specific hints per error token: `#REF!` (out-of-range reference), `#NAME?` (suggests `list-functions`), `#DIV/0!`, `#VALUE!`, `#N/A`.
- **Unknown column in diff** — lists available columns from the header row.
- **Duplicate keys** — suggests dedup or choosing a different key column.
- **Ambiguous matches** — suggests `--on_ambiguous report`.
- **Convert stdin without `--from`** — shows a working example with `--from`.
- **XLSX export attempt** — suggests `csv` or `json` as alternatives.
- **Unknown file extension** — lists valid extensions.
- **Lua errors in replay** — syntax hints for common mistakes (`attempt to call a nil value`, malformed `grid.set{}` calls).
- **Nondeterministic functions** — lists which functions (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`) caused `--verify` to fail.
- **Fingerprint mismatch** — explains that the source data or script was modified.

### CLI Flags & Output

- **`diff --quiet`** — new `-q` / `--quiet` flag suppresses the stderr summary and warnings. Designed for CI where only the exit code matters.
- **`diff --format` alias** — `--format` is now accepted as an alias for `--out`, so both `--out json` and `--format json` work. Prevents a common papercut.
- **`.xlsb` and `.ods` extension recognition** — `infer_format` now maps `.xlsb` and `.ods` to the XLSX reader (via calamine), matching the desktop app's import support.
- **JSON trailing newline** — `write_json` and `format_diff_json` now append `\n` after the closing brace. Prevents `bash: warning: here-document delimited by end-of-file` and plays nicely with `jq`, `diff`, and other line-oriented tools.
- **`--version` enhancement** — `vgrid --version` now prints engine version and build type (debug/release) in addition to the CLI version.

### Documentation & Site Truthfulness

Every CLI example on the marketing site and docs site was audited against the actual binary. Eight discrepancies were found and fixed by correcting the documentation (not by adding features).

- **`calc` examples** — changed from file-path syntax to stdin pipe syntax (`cat file.csv | vgrid calc ...`). Removed multi-formula example that doesn't exist.
- **`diff --format`** — site examples updated to `--out` (the real flag; `--format` alias also added to CLI as a safety net).
- **`replay --stop-at`** — removed from site; flag does not exist.
- **`convert -t xlsx`** — removed XLSX export examples; not yet implemented.
- **`list-functions`** — removed "with descriptions" claim; output is names only.
- **XLS/XLSB/ODS** — scoped to "desktop app reads these via calamine" in FAQ; CLI reads XLSX.
- **Known Limitations section** — added to both README and docs CLI reference. Five items: no XLSX export, calc stdin-only, replay layout ops hashed but not applied, conservative nondeterminism detection, multi-sheet exports sheet 0 only.
- **CI snippet** — added `diff --quiet` and `replay --verify --quiet` examples to README.

## 0.4.1

### Session Restore Fix

Closed windows no longer reappear on restart. Previously, closing a window removed it from the live window registry but not from the persisted session file, so every closed window came back as a zombie on next launch.

- **Stable window identity** — each window gets a unique `window_id: u64` assigned at creation, used for session matching. Replaces fragile file-path matching which broke for untitled windows and multiple windows on the same file.
- **Single close funnel** — all close paths (clean close, save-then-close, discard-and-close, Save As close) now call `prepare_close()`, which removes the window from session state, unregisters from the window registry, and persists immediately. No close path can forget session removal.
- **Self-healing IDs** — if a window reaches a session method without an assigned ID (bug), it allocates one on the spot and proceeds normally rather than silently skipping. `debug_assert!` fires in development builds to catch the root cause.
- **Robust ID assignment on load** — respects existing `window_id` values from the session file, sets the counter to `max + 1`, and deduplicates corrupt sessions (first-seen wins, duplicates get reassigned).
- **Quit snapshots state** — the focused window commits pending edits and snapshots its session state on Quit, so scroll position and selection survive restart.
- **"Session restored" status** — restored windows show a status bar message on launch (only when windows were actually restored).
- **8 new tests** — window ID counter, remove-from-session, deserialization default, same-file-different-IDs, ID assignment with existing/duplicate IDs.

### macOS File Association ("Open With")

VisiGrid now appears in Finder's "Open With" menu for spreadsheet files and handles file-open events from Finder, Spotlight, and drag-to-dock.

- **Document type declarations** — Info.plist declares `.vgrid` and `.sheet` (Owner), `.xlsx`, `.xls`, `.csv`, `.tsv` (Alternate) with proper `LSItemContentTypes` for modern macOS. Exported UTType declarations for `com.visigrid.vgrid` and `com.visigrid.sheet`.
- **`on_open_urls` handler** — registered before `Application::run()` to catch file URLs that arrive during launch (Finder double-click starts app). Uses a shared `Arc<Mutex<Vec<String>>>` buffer bridging the platform callback (no App context) to the gpui event loop.
- **Startup drain** — URLs accumulated during initialization are processed synchronously after window creation, opening each file in a new window.
- **Adaptive polling task** — for files opened while the app is already running. Polls the buffer with adaptive sleep: 100ms after recent activity (responsive), backing off to 1s when idle (no wasted cycles). Buffer drained in one shot via `std::mem::take` for minimal lock hold time.
- **URL deduplication** — `normalize_and_dedup_urls()` parses file:// URLs, percent-decodes, canonicalizes paths (resolves symlinks), and deduplicates with a `HashSet` while preserving order. Prevents "why did it open twice?" when Finder sends duplicate or differently-encoded URLs.
- **9 new tests** — URL-to-path conversion, percent decoding (mixed, incomplete sequences), dedup (exact dupes, encoding variants, order preservation, non-file URL filtering).

### Windows & Linux File Associations

- **Windows installer** — `.xlsx`, `.xls`, `.sheet` added to the Inno Setup installer. `.sheet` is checked by default (native format). `.xlsx` and `.xls` are optional, unchecked by default, labeled "best-effort import" (respects user choice — registers in "Open with" without hijacking defaults). All six types (`.vgrid`, `.sheet`, `.csv`, `.tsv`, `.xlsx`, `.xls`) now appear in SupportedTypes for the "Open with" menu.
- **Linux desktop entry** — `visigrid.desktop` now declares MIME types for xlsx (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`), xls (`application/vnd.ms-excel`), and `.sheet` (`application/x-visigrid-sheet`). Desktop environments (GNOME, KDE) will show VisiGrid in "Open With" for these file types.
- **Multi-file open from CLI** — `visigrid a.xlsx b.csv c.tsv` now opens each file in its own window. Previously only the last file argument was opened. Required for Windows "Open with" (sends `VisiGrid.exe "%1"` per file) and Linux desktop integration (`%F` in .desktop passes multiple paths).

### Release Pipeline Hardening

- **Entitlements in CI signing** — `release.yml` now includes `--entitlements gpui-app/macos/entitlements.plist` when code signing. Previously signed with hardened runtime but without entitlements, which could block file access and network in the shipped build.
- **Info.plist verification** — both `release.yml` and `bundle-macos.sh` now verify the built app's Info.plist contains the expected UTTypes (`org.openxmlformats.spreadsheetml.sheet`, `com.visigrid.vgrid`, `UTExportedTypeDeclarations`) before signing. Fails fast with a clear error if the plist is wrong.
- **Homebrew cask URL fix** — `update-cask.yml` DMG download URL and regenerated cask URL now match the actual release artifact filename (`VisiGrid-macOS-universal.dmg`). Previously used a versioned filename that would 404.

### Native .sheet Save/Load Hardening

- **Alignment round-trip fix** — `Alignment::General` (auto: numbers right-align, text left-aligns) now survives save/load. Previously, General was incorrectly mapped to Left on load, losing smart alignment behavior.
- **`CellFormat::is_default()` helper** — replaces fragile inline field checks in save/load skip conditions. New fields added to `CellFormat` are automatically covered via derived `PartialEq`.
- **`alignment_to_db` / `alignment_from_db` helpers** — explicit encoding for all 5 alignment variants (including `CenterAcrossSelection`), replacing scattered `match` blocks and a lossy `_ =>` catch-all.
- **Bloated file auto-migration** — legacy `.sheet` files containing millions of empty default rows (from a prior save bug) are automatically cleaned on load. Rebuild triggers when >50% of rows are junk (>10k total). Explicit Left alignment and other real formatting survives the migration.
- **5 new regression tests** — `is_default`, save-only-populated-cells, General round-trip, Left persistence, bloated file cleanup.

## 0.3.9

Number format editor, thousands separators, negative styles, and currency symbols.

### Number Formatting

- **Thousands separators** — Number and Currency formats now support a thousands separator toggle. UI-applied formats default to thousands on.
- **Negative number styles** — four styles: minus (`-1,234.56`), parentheses (`(1,234.56)`), red minus, and red parentheses. Red styles render negative values in red in the grid.
- **Currency symbol selection** — preset buttons ($, €, £, ¥) in the editor. Custom symbols persist and export correctly to XLSX (quoted when needed).
- **Number Format Editor** — new modal dialog (Ctrl+1, or "Edit..." in inspector). Type pills (General / Number / Currency / Percent / Date), decimals ±, thousands toggle, negative style radio, and live preview. Keyboard-drivable: Tab cycles types, arrows cycle negative styles, +/- adjust decimals, Space toggles thousands.
- **Inspector summary** — Number/Currency formats show a summary line (e.g. "2 decimals, thousands, parens"), live preview using the active cell's value, and an "Edit..." shortcut button.
- **XLSX export** — emits full 4-section format codes (`positive;negative;zero;text`) for Number and Currency, including `[Red]` markers and parentheses for Excel compatibility.
- **Native persistence** — new `.vgrid` schema columns (`fmt_thousands`, `fmt_negative`, `fmt_currency_symbol`) with `PRAGMA user_version` migration. Old files open and render identically; saving migrates them forward.

### Inspector Panel

- **Vertical alignment icons** — Top / Middle / Bottom buttons now use three-line icons (lines positioned at top, center, or bottom of a box) instead of plain "T" / "M" / "B" text labels. Matches the horizontal alignment icon style.

### Menu Dropdowns (Windows/Linux)

- **Click-outside dismiss** — clicking anywhere outside an open menu dropdown now closes it. Uses a transparent backdrop below the menu bar so hover-to-switch between menu headers still works.
- **Keyboard navigation** — Up/Down moves highlight, Left/Right switches menus, Enter executes highlighted item, Escape closes. Insert menu (empty) is skipped during keyboard cycling.
- **Mouse-hover highlight sync** — hovering a menu item sets the keyboard highlight, so you can hover then press Enter to execute.
- **Click isolation** — clicks inside the menu dropdown no longer leak through to the grid.
- **`menu_model.rs`** — new module: typed `MenuAction` enum and `MenuEntry` descriptors as single source of truth for menu structure, item counts, and action dispatch. Decouples `app.rs` (navigation) from `views/menu_bar.rs` (rendering).

### Internal

- **`ui/popup.rs`** — new design system primitive for floating popup containers (context menus, positioned panels). Provides standard styling (bg, border, rounded, shadow), click-outside dismiss via `on_mouse_down_out`, and a `clamp_to_viewport` helper. Migrated 3 call sites: `context_menu.rs`, `status_bar.rs` (also fixed missing click-outside dismiss), and `inspector_panel.rs` history context menu.

## 0.3.8

Format bar, context menus, grid polish, and merge export correctness.

### Format Bar

Toggleable formatting toolbar between the formula bar and column headers. View → Format Bar (default: on), persisted in user settings, hidden in zen mode.

- **Font family** — clickable label shows current font name, click opens font picker. Mixed selections show "—" in italic.
- **Font size** — editable input with dropdown of common sizes (8–72), integer-only. Click-to-edit with select-all replace, Enter/Esc/Tab semantics, and Up/Down to adjust size by 1. Invalid input reverts silently.
- **Bold / Italic / Underline** — toggle buttons with tri-state: active (accent), inactive (transparent), and mixed (em dash, muted background) for heterogeneous selections.
- **Fill color** — colored swatch chip, click opens color picker (`ColorTarget::Fill`). Mixed selections show checkerboard. None shows theme default.
- **Text color** — "A" with colored underbar, click opens color picker with new `ColorTarget::Text` variant. None = Automatic (theme default text color, not transparent).
- **Alignment** — Left / Center / Right toggle buttons with alignment icons, active state reflects current selection. Shared icons also used in inspector panel.
- **Tooltips** — platform-aware keyboard shortcuts on all controls (e.g. "Bold (⌘B)" on macOS, "Bold (Ctrl+B)" elsewhere).
- **Engine** — `set_font_size()` and `set_font_color()` setters on Sheet, with selection wrappers (`set_font_size_selection`, `set_font_color_selection`) and undo/redo via `FormatActionKind::FontSize` / `FontColor`. Provenance export as Lua (`font_size=24`, `font_color="#FF0000"`).
- **Font size rendering** — custom font sizes render correctly in the grid. gpui's `TextRun` doesn't carry font_size, so sized text is wrapped in a div with `.text_size()` to cascade via the element tree.
- **Action dispatch routing** — font size input correctly captures Enter, Escape, Backspace, Tab, Delete, and arrow keys. gpui dispatches keybinding actions before `on_key_down` handlers, so action handlers in `actions_edit.rs` and `actions_nav.rs` gate on `format_bar.size_editing` to route keys to the format bar instead of the grid.
- **UI state** — `FormatBarState` on `UiState` (input buffer, focus handle, dropdown visibility) — transient UI chrome, never serialized, no undo semantics.

### Context Menus

Right-click on cells, row headers, or column headers for common actions.

- **Cell menu** — Cut, Copy, Paste, Paste Values, Format Painter, Clear Contents, Clear Formats, Inspect.
- **Row header menu** — Insert Row, Delete Row, Clear Contents, Clear Formats.
- **Column header menu** — Insert Column, Delete Column, Clear Contents, Clear Formats, Sort A→Z, Sort Z→A.
- **Smart selection** — right-click inside a multi-cell, multi-row, or multi-column selection preserves it; right-click outside moves to the target.
- **Mode-aware** — right-click commits edits in place (no cursor jump) and cancels Format Painter cleanly.
- **Paste gating** — Paste and Paste Values are disabled when the clipboard is empty. Sort is disabled when the selection is not a full column.
- **Dismissal** — click outside, Escape, or any non-modifier key closes the menu. Modifier-only keys (Shift, Ctrl) keep it open.
- **Edge clamping** — menu repositions to stay within window bounds near edges.

### Grid Click Reliability

- **No more dead clicks on gridlines** — mouse handlers moved from inner cell div to the outer wrapper div, which has exact pixel dimensions from the flex layout. Sub-pixel gaps at cell borders no longer create hit-testing dead zones. Cursor stays crosshair everywhere in the grid, including on gridline boundaries.
- **Edit commit on click** — clicking another cell while editing commits the current edit in place and exits edit mode (Excel behavior). Works for cells, row headers, column headers, and merged cell overlays. Same-cell clicks preserve edit mode.
- **Selection gridlines** — selected cells now draw interior gridlines as a child overlay (GridLines color) while keeping the selection border (accent) on outer edges. Previously, selected cells suppressed all gridlines.

### Grid Cursor

- **Crosshair over cells** — grid interior uses crosshair cursor (Excel convention), signaling spatial selection rather than UI interaction. Merged cell overlays match. Headers, toolbars, and dialogs remain arrow/pointer.

### Merge Export Correctness

Merge-hidden cells no longer leak residual data into exported files.

- **CSV export** — cells hidden by a merge are forced to empty strings. The origin cell retains its value. `.flexible(true)` handles variable-width rows from trailing empty suppression.
- **XLSX export** — `merge_range()` is now called before cell export (ordering fix), and merge-hidden cells are skipped entirely. Origin cell values overwrite the blank placeholder written by `merge_range()`.
- **Roundtrip tests** — `test_csv_export_merged_cells_no_leak` and `test_xlsx_export_merged_cells_no_leak` verify origin values survive and hidden data never appears in output.
- **Spec completion** — merged cells spec moved from `docs/features/future/` to `docs/features/done/`. All 6 phases complete.

### Internal

- **AI context** — header row detection threshold relaxed from 70% to 60% text, allowing mixed headers like "ID", "123", "Product".
- **Inspector alignment** — horizontal alignment buttons use shared `render_align_icon()` from the format bar module, replacing plain text labels with consistent icons.
- **Grid selection borders** — simplified gridline suppression: user borders and merge interiors suppress gridlines, but neighboring-cell selection state no longer factors in (handled by the new interior overlay approach).

## 0.3.7

Trust, Correctness, and Real Excel Compatibility.

### AI with Explicit Rules

Execution Contracts make AI constraints visible and verifiable at every touchpoint.

- **Execution Contracts** — every AI feature displays a named contract badge (`read_only_v1`, `single_cell_write_v1`) showing exactly what the AI can and cannot do. Tooltips explain the contract system; Copy Details includes contract identifier, human label, and write scope in the diagnostic dump.
- **Multi-provider support** — OpenAI, Google Gemini, and xAI Grok all work via OpenAI-compatible chat completions endpoints. Provider selection configures credentials and defaults; feature availability is gated by implemented capabilities.
- **Analyze with AI** — read-only data analysis for the current selection. AI describes patterns and structure but cannot modify cells. Contract: `read_only_v1`.
- **Keyring persistence fix** — API keys now persist across sessions on all platforms. The `keyring` crate v3 requires explicit platform backend features (`apple-native`, `windows-native`, `linux-native-sync-persistent`); without them, keys were stored in memory only.
- **Capability gate fix** — Diff Summary and Explain This Change correctly gate on the `analyze` capability (read-only), not `insert_formula` (write).
- **Dialog stability** — Generate Summary no longer closes the diff dialog on click (fixed event propagation). Long AI responses scroll instead of being clipped.

### Excel Import & Financial Model Fidelity

Financial models that previously showed mass #CIRC! and #ERR errors now import cleanly.

#### XLSX Import Overhaul

- **calamine 0.26 → 0.32** — fixes shared formula expansion for `t="shared"` cells with ranges, absolute references, and column/row ranges
- **Topological recompute on import** — all formulas are re-evaluated in dependency order after the full graph is built, fixing stale cached values from per-cell evaluation during import
- **XML formula backfill** — cells with `<f>` elements but no cached `<v>` value (skipped by calamine) are now extracted directly from worksheet XML and backfilled
- **Shared formula follower expansion** — 2-pass XML parsing reconstructs follower formulas from master definitions with reference shifting (respects `$` absolute anchors)
- **XML value backfill** — cells stored as shared strings, inline strings, or numeric values that calamine drops are recovered from worksheet XML via shared string table resolution
- **Post-recalc error counting** — circular refs and formula errors counted after topo recalc, with up to 5 concrete examples in the import report
- **Import report diagnostics** — shows formula backfill count, value backfill count, shared formula groups, recalc errors with cell addresses and formulas
- **Status bar error count** — "Opened in Xms — 0 errors" or "Opened in Xms — N errors (Import Report)" with clickable access

#### XLSX Formatting Import

- **styles.xml parsing** — number formats, font styles, fills, borders, alignment, column widths, and row heights extracted from XLSX style tables
- **Style deduplication** — `StyleTable` with `style_id` per cell, interning identical formats to minimize memory
- **Theme color approximation** — Excel theme colors with tint modifiers approximated to RGB with import report warnings
- **Column widths and row heights** — imported from XLSX and applied to sheet layout

#### Number Format Rendering (ssfmt)

Custom Excel number format codes now render correctly via the `ssfmt` crate (ECMA-376 compliant, 99.99% SheetJS SSF compatibility).

- **Accounting formats** — zero values show `-` instead of `$-??` garbage; `_` padding and `*` fill handled correctly
- **Multi-section formats** — positive/negative/zero/text sections selected and rendered properly
- **Full format code support** — thousands separators, decimals, percent, currency, date/time tokens, conditional sections
- **Graceful fallback** — if ssfmt can't parse a code, falls back to VisiGrid's built-in formatter

#### Formula Engine Additions

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

### Engine, UI, and Infrastructure

#### Merged Cells

Full merged cell support across the engine, rendering, navigation, clipboard, and UI.

- **Rendering** — merged regions render as unified overlays with correct z-order (above cell grid, below text spill). Spacer cells for hidden/origin cells. Spill exclusion prevents double-rendered text. Interactive overlays route click, drag, and fill to merge origin.
- **Navigation** — arrow keys, Ctrl+Arrow, selection extension, Go To, and edit Tab/Enter flows treat merges as atomic data units. `find_data_boundary()`, `jump_selection()`, `extend_jump_selection()`, `confirm_goto()`, and tab-chain Enter are all merge-aware.
- **Copy/Paste** — `InternalClipboard` stores merged regions as relative coordinates. Paste recreates merges at destination. Cut removes source merges. Overlap guard blocks partial overlap. Paste Values/Formulas/Formats ignore merge metadata. Group undo bundles value and topology changes.
- **UI** — Merge Cells (Ctrl+Shift+M) and Unmerge Cells (Ctrl+Shift+U) with data-loss confirmation dialog, overlap guard, contained merge replacement, and full undo/redo via `UndoAction::SetMerges`.
- **17 ship-gate tests** — viewport overlap, pixel rect, spill predicates, span dimensions. 476+ engine tests total.

**Known limitation:** merges fully inside a frozen region may not render correctly when scrolling (planned for Rendering Overlays v2).

#### Recalc Engine

- **Value-typed computed cache** — `HashMap<(usize, usize), Value>` replaces String cache, eliminating lossy numeric conversions
- **No evaluate-on-cache-miss** — all getter paths return defaults on cache miss; only topo recalc and `set_value()` populate the cache
- **Correct recalc ordering** — workbook-level topological evaluation ensures upstream cells are computed before dependents

#### Internal

- **`views/mod.rs` split** — 4,422-line file split into 7 focused modules (`actions_nav`, `actions_edit`, `actions_ui`, `key_handler`, `f1_help`, `named_range_dialogs`, `rewind_dialogs`)
- **Formula engine modularization** — `eval.rs` (5,431 → 595 lines) split into 11 category modules
- **Dependency upgrades** — calamine 0.32, quick-xml 0.38, zip 4, ssfmt 0.1, regex 1

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
vgrid replay script.lua --verify      # Verify fingerprint
vgrid replay script.lua -o output.csv # Export to CSV/TSV/JSON
vgrid replay script.lua --fingerprint # Print fingerprint only
vgrid replay script.lua --verify -q   # CI mode (quiet)
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
