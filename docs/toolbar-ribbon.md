# Native toolbar and ribbon

This branch starts Phase 1 of the Obsidian **VisiGrid Toolbar and Ribbon Spec**.
It is an opt-in native UI implementation, not a release of the complete spec.

## Try it

- **View → Toolbar: Ribbon**, or **Preferences → Appearance → Toolbar layout**.
- The command palette also has **Use Ribbon Toolbar**, **Use Compact Toolbar**,
  **Collapse/Expand Ribbon**, and **Show/Hide Toolbar**.
- Home, Insert, Formulas, Data, and View group existing commands. Home reuses
  the existing font picker, font-size editor, color controls, and mixed formatting states.
- Collapse keeps the tabs visible. Selecting a tab opens a temporary panel;
  Escape or clicking outside dismisses it. Double-clicking a tab toggles collapse.
- Reduced widths collapse groups into menus without adding rows of chrome.
- Compact remains the default. An existing hidden format bar remains hidden;
  choosing Ribbon does not implicitly turn it on.

No account is required and no workbook format changes are involved. Changing
the layout is a personal preference, not a document edit or undo operation.

## Alt KeyTips (Linux and Windows)

Tap and release **Alt** by itself to show hints. Existing menu letters stay
**F** File, **E** Edit, **V** View, **I** Insert, **O** Format, **D** Data,
**H** Help. In Ribbon, the tab hints are **B** Home, **N** Insert,
**M** Formulas, **A** Data, **W** View. Compact shows the menu hints.

After selecting a ribbon tab, type the group letter and command number shown
on the control. For example, **Alt, B, F, 2** toggles Bold; **Alt, B, F, 6**
opens the font-size editor. These are sequential presses, not function keys.
A collapsed group opens when its letter is typed, revealing the command hints.
Disabled commands keep their existing guards and explain the reason in the
status bar. They leave hints open so another command can be selected.

Escape/Backspace clears a partial command, then returns to the root hints,
then dismisses them. Alt again dismisses them. Hints do not expire on a timer.
A mouse click or window deactivation dismisses hints. A chord using Alt does
not activate hints when Alt is released; existing Alt shortcuts still use the
normal keymap. Hints do not take over dialogs, terminal/script input, IME
composition, or toolbar text fields. Opening/dismissing hints preserves the
cell editor's draft, caret and selection; cell-changing ribbon commands remain
disabled during editing.

**macOS is deferred:** its existing Option+Space/category behavior is unchanged.
The new Alt-tap binding and key interceptor are excluded from macOS builds.
Native menu placement, Option composition, font rendering and window scaling
must be reviewed on a Mac before deciding the Mac interaction.

Known pre-existing binding overlaps are separate follow-ups: held
`Alt+H, C, P` shares the Help prefix; `Alt+Enter` has both newline and trace
return registrations; opted-in Mac `Alt+T` overlaps Tools and Trace. This change
preserves those bindings. In the new tap-Alt hints, H opens Help immediately.
F10 remains Problems / debugger Step Over; Ctrl+F1 is not newly assigned.

## Implementation

`toolbar.rs` owns command-surface geometry and preference actions. The formula
bar is below the command groups in Ribbon and above the format bar in Compact.
The grid's origin, formula-bar hit testing, script editor's available height,
format dropdowns, and formula help placement share these coordinates. Normal
chrome remains hidden in Zen; Table controls and recovery banners keep their
existing behavior.

`views/ribbon.rs` describes tabs/groups and dispatches existing `CommandId`
operations. Row/column insertion preserves the existing whole-row/column
selection requirement. Commands that require cell navigation stay disabled
during an active edit or modal; changing layout itself does not commit or
discard a cell edit. Read-only restrictions are rechecked on invocation.

`appearance.toolbar` stores `schema_version`, `layout`, and
`ribbon_collapsed` in the user settings. Unknown fields are retained. Malformed
toolbar fields resolve independently, so they do not reset unrelated settings.
Future schemas are preserved and are not overwritten by the layout controls.
`appearance.show_format_bar` remains the visibility flag for compatibility.
Toolbar changes use the shared settings store and report save failures; user
settings saves now write a temporary file and rename it over the destination.

No keyboard shortcuts are reassigned. The new actions can be bound through
`view.compacttoolbar`, `view.ribbontoolbar`, `view.collapseribbon`, and
`view.toolbar` in the user keybindings file.

## Remaining work before calling Phase 1 complete

- Complete native interaction QA, including focus transfer, cell/formula edits,
  IME input, mixed formatting, dropdown handoff, multiple windows, and scaling.
- Validate keyboard entry into the ribbon and the complete focus order. Match
  the specified tab-strip arrow behavior and every picker's keyboard support.
- Consolidate toolbar command metadata and effective shortcut labels; finish
  context-specific availability for commands such as pivot refresh.
- Convert the flat View entries to the specified checked Toolbar submenu.
- Replace Compact's existing horizontal scrolling with group overflow.
- Anchor every popup to measured control bounds, including controls opened
  from an overflow menu; finish viewport-edge clamping.
- Complete macOS and Windows interactive validation and broader Linux native QA.
- Validate the concept with the original commenter before expanding scope.

Quick access pinning/reordering and advanced custom tabs/groups are later
phases. This branch does not expose placeholder customization controls.

## Validation

On Linux, `cargo check` and the desktop build pass. The desktop unit suite passes
with **615 passed, 0 failed, 3 ignored**. It caught and fixed a duplicate View
accelerator and a default preference serialization regression.

Linux/Wayland smoke checks at 1920×1200, scale 1: Home/Data tab rendering,
collapse and temporary expansion, palette switching in both directions,
preservation of an active cell draft when switching to Ribbon, cancelling that
draft, and the reused Bold control. These are targeted checks, not the full
acceptance matrix above. Additional Linux KeyTips checks cover Alt tap, root/tab command hints, H → Help,
Bold dispatch, disabled formatting during a cell edit, Escape preserving that
draft, the collapsed font-size editor, held Alt+F without stray hints, and
Compact showing menu-only hints, and a 941-pixel tiled window revealing the
Editing overflow group and opening Find through its KeyTip. Windows runtime and macOS review remain open.
A screenshot is saved alongside the workspace concept
at `work/visigrid-toolbar/native-ribbon.png` (outside this repository).

Settings tests cover legacy defaults, hidden-toolbar migration, malformed
fields, round trips, future-schema retention, atomic save, and write failure.
Geometry tests cover both platform chrome heights, expanded formula bars,
Compact, expanded/collapsed Ribbon, hidden toolbar, and Zen. Group-layout tests
cover fixed-height overflow. Existing menu model tests validate mouse and
keyboard command order together.

Run from the workspace root:

```sh
cargo check -p visigrid-gpui --bin visigrid
cargo test -p visigrid-gpui --bin visigrid
```

Build and run a separate binary/configuration for visual review; avoid changing
the user's normal settings or opening production workbooks during QA.
