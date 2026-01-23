# Title Bar, Menus, and Document Identity

## Summary

Implement proper document identity display, themed title bar, and unified menu system across platforms. The goal is to show document state at a glance without duplicate UI, while respecting each platform's conventions.

---

## Shipped (v0.2.4)

### macOS Transparent Titlebar

Zed-style chrome blending that makes the app feel native and professional.

**Window Options:**
```rust
WindowOptions {
    titlebar: Some(TitlebarOptions {
        title: None,
        appears_transparent: true,
        traffic_light_position: Some(point(px(9.0), px(9.0))),
    }),
    ..Default::default()
}
```

**Custom Titlebar Element:**
- Height: 34px (matches Zed)
- Left padding: 72px (clears traffic lights)
- Background: PanelBg token
- Border: 1px bottom, 50% opacity (subtle hairline)
- Draggable via `WindowControlArea::Drag`
- Double-click triggers `window.titlebar_double_click()` for zoom

**Typography (Zed-class polish):**
| Element | Size | Color | Notes |
|---------|------|-------|-------|
| Primary (filename) | 12px | TextPrimary | Full contrast |
| Secondary (provenance) | 10px | TextMuted @ 85% | Quieter, recedes |

**Chrome Scrim:**
- 8px gradient below titlebar
- Top: PanelBg @ 12% opacity
- Bottom: transparent
- Creates subtle visual separation without hard edge

**Inactive State:**
- Let macOS handle it natively
- No manual opacity changes (Zed approach)
- Works because `appears_transparent` content still gets OS-level dimming

**What We Learned from Zed:**
1. Don't over-engineer inactive state on macOS
2. Use semantic color tokens (TextPrimary, TextMuted) not opacity tricks
3. 34px minimum height lets text breathe
4. Provenance should be visually quieter than filename (size + color)
5. Typography polish > architectural complexity

### Default App Prompt (Title Bar Chip)

Zed-style "onboarding banner" for prompting users to set VisiGrid as their default spreadsheet app. Designed to feel like a helpful suggestion, not an ad.

**Copy (tight, non-marketing):**
- Label: "Open .csv files with VisiGrid" (scoped to specific file type)
- Button: "Make default"
- Success: "Default set"
- Needs Settings: "Finish in System Settings" + "Open"

**Layout:**
- Right-aligned in title bar (uses `justify_between` flexbox)
- Same 34px height as title bar content
- Follows Zed's `OnboardingBanner` pattern

**Visual Style:**
```
┌──────────────────────────────────────────────────────────────────────────┐
│ ● ● ●  data.csv                    [Open .csv files with VisiGrid │ ✕]  │
│        imported from legacy.xls                          Make default   │
└──────────────────────────────────────────────────────────────────────────┘
```

- Muted background (50% opacity, like inactive tab)
- 1px border at 20% opacity (subtle)
- 10px text (smaller than 12px filename)
- Label: muted @ 80% opacity (recedes behind filename)
- Button: muted (not bold, not loud)
- ✕: 40% opacity until hover (very quiet)
- Rounded corners (`rounded_sm`)

**Visibility Rules (guardrails):**
1. macOS only
2. File successfully loaded (has path, not showing ImportReport dialog)
3. Not an unsaved new document
4. Not a temporary file (`/tmp`, `~$`, `.Trash`, etc.)
5. File type is CSV/TSV/Excel (not native .vgrid)
6. User hasn't permanently dismissed (`TipId::DefaultAppPrompt`)
7. Not in cool-down period (7 days after ignoring)
8. Not already shown this session (prevents spam)
9. VisiGrid isn't already the default handler for this file type

**State Machine:**
- `Hidden` → Initial state, not showing
- `Showing` → Prompt visible, waiting for user action
- `Success` → User clicked, duti succeeded, show "Default set" for 2s
- `NeedsSettings` → User clicked but must complete in System Settings

**Session & Cool-down:**
- `shown_this_session` flag prevents re-showing after close+reopen another file
- `default_app_prompt_shown_at` timestamp enables 7-day cool-down
- Cool-down applies when user ignores (closes without clicking)
- Permanent dismiss via ✕ button overrides cool-down

**Post-Click Feedback:**
- Success: "Default set" chip for 2 seconds, then auto-hides
- Needs Settings: "Finish in System Settings" + "Open" button
- Both states permanently dismiss so prompt won't return

**macOS Integration:**
- Uses `duti -s com.visigrid.app {UTI} all` when available
- Falls back to opening System Settings Extensions pane
- Re-checks `is_default_handler` after duti to detect success
- Checks default via `duti -x {UTI}` output parsing

---

## Design Principles

1. **Document identity at a glance** - Always show file name, dirty state, and read-only status
2. **No duplicate menu systems** - macOS uses system menu bar; Windows/Linux use in-app menu from same definition
3. **Title bar is informational** - No toolbar, no ribbon, no button strip
4. **Cross-platform parity** - Same semantics everywhere, different placement per platform
5. **Theming integration** - Title bar blends with app (like Zed), not stock OS chrome

## Platform Mapping

| Surface | macOS | Windows | Linux |
|---------|-------|---------|-------|
| Menu commands | Global menu bar | In-app menu bar | In-app menu bar |
| Document identity | OS title bar (full) | Window title (short) | Window title (short) |
| Provenance | In title bar | Header row (future) | Header row (future) |
| Mode indicator | Bottom status bar | Bottom status bar | Bottom status bar |

---

## Phase A: DocumentMeta & Title String

### Constants

```rust
/// Native file extension for VisiGrid documents
pub const NATIVE_EXT: &str = "vgrid";

/// Returns true if the extension is considered "native" (no provenance needed)
/// Native formats: vgrid (our format), xlsx/xls (Excel, first-class support)
pub fn is_native_ext(ext: &str) -> bool {
    matches!(ext.to_lowercase().as_str(), "vgrid" | "xlsx" | "xls")
}
```

### Data Model

Add to `src/app.rs`:

```rust
/// Source of the document (for provenance display)
///
/// Only used for non-native formats that were imported/converted.
/// Native formats (vgrid, xlsx) have no provenance - they're first-class.
#[derive(Clone, Debug, PartialEq)]
pub enum DocumentSource {
    /// Imported from a non-native format (CSV, TSV, JSON)
    /// These are converted on load and need "Save As" to persist as native.
    Imported { filename: String },
    /// Recovered from session restore (unsaved work from crash/quit)
    Recovered,
}

/// Document metadata for title bar display
#[derive(Clone, Debug)]
pub struct DocumentMeta {
    /// Display name - FULL filename with extension (e.g., "budget.xlsx", not "budget")
    /// For unsaved documents, this is "Sheet1" (no extension)
    pub display_name: String,
    /// Document has been saved at least once (to native format)
    pub is_saved: bool,
    /// Document is read-only
    pub is_read_only: bool,
    /// How the document was opened/created (only for non-native sources)
    pub source: Option<DocumentSource>,
    /// Full path if saved
    pub path: Option<PathBuf>,
    /// Active sheet name (for multi-sheet display, future)
    pub active_sheet_name: Option<String>,
}
```

**Note:** `is_dirty` is NOT stored in DocumentMeta. It's derived from history state (see Dirty State Derivation below).

### Helper Functions

```rust
/// Extract display filename from path (full name with extension)
fn display_filename(path: &Path) -> String {
    path.file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("Sheet1")
        .to_string()
}

/// Extract lowercase extension from path
fn ext_lower(path: &Path) -> Option<String> {
    path.extension()
        .and_then(|e| e.to_str())
        .map(|s| s.to_lowercase())
}

### Title String Format

```rust
impl DocumentMeta {
    /// Generate the window title string for macOS (includes provenance)
    pub fn title_string_full(&self, is_dirty: bool) -> String {
        let mut title = self.display_name.clone();

        // Dirty indicator
        if is_dirty {
            title.push_str(" \u{25CF}"); // ●
        }

        // Unsaved suffix (new document, never saved)
        if !self.is_saved && self.source.is_none() {
            title.push_str(" — unsaved");
        }

        // Provenance subtitle (only for imported/recovered)
        if let Some(source) = &self.source {
            match source {
                DocumentSource::Imported { filename } => {
                    title.push_str(&format!(" — imported from {}", filename));
                }
                DocumentSource::Recovered => {
                    title.push_str(" — recovered session");
                }
            }
        }

        // Read-only indicator
        if self.is_read_only {
            title.push_str(" — read-only");
        }

        title
    }

    /// Generate the window title string for Windows/Linux (compact, no provenance)
    ///
    /// Provenance is omitted because:
    /// - Window titles get truncated aggressively on these platforms
    /// - Long titles pollute task switchers (Alt+Tab, taskbar)
    /// - Provenance can be shown in an optional header row inside the app
    pub fn title_string_short(&self, is_dirty: bool) -> String {
        let mut title = self.display_name.clone();

        // Dirty indicator
        if is_dirty {
            title.push_str(" \u{25CF}"); // ●
        }

        // Unsaved suffix
        if !self.is_saved && self.source.is_none() {
            title.push_str(" — unsaved");
        }

        // Read-only indicator (important enough to keep)
        if self.is_read_only {
            title.push_str(" — read-only");
        }

        title
    }

    /// Platform-appropriate title string
    pub fn title_string(&self, is_dirty: bool) -> String {
        #[cfg(target_os = "macos")]
        { self.title_string_full(is_dirty) }

        #[cfg(not(target_os = "macos"))]
        { self.title_string_short(is_dirty) }
    }
}
```

### Example Titles

**macOS (full):**

| State | Title |
|-------|-------|
| New unsaved | `Sheet1 — unsaved` |
| Saved native, clean | `budget.vgrid` |
| Saved native, dirty | `budget.vgrid ●` |
| Opened XLSX, clean | `report.xlsx` |
| Opened XLSX, dirty | `report.xlsx ●` |
| Imported CSV, clean | `customers.csv — imported from customers.csv` |
| Imported CSV, dirty | `customers.csv ● — imported from customers.csv` |
| Read-only | `budget.vgrid — read-only` |
| Recovered | `Sheet1 ● — recovered session` |

**Windows/Linux (short):**

| State | Title |
|-------|-------|
| New unsaved | `Sheet1 — unsaved` |
| Saved native, clean | `budget.vgrid` |
| Saved native, dirty | `budget.vgrid ●` |
| Opened XLSX, clean | `report.xlsx` |
| Imported CSV, dirty | `customers.csv ●` |
| Read-only | `budget.vgrid — read-only` |
| Recovered | `Sheet1 ●` |

**Note:** Display name includes the extension (e.g., `budget.xlsx` not `budget`). This matches user expectations in title bars and task switchers.

### Dirty State Derivation

**Invariant: Dirty state is derived, not tracked.**

```rust
impl Spreadsheet {
    /// Returns true if document has unsaved changes.
    /// Computed from history index vs save point.
    pub fn is_dirty(&self) -> bool {
        self.history.current_index() != self.history.save_point()
    }
}

impl History {
    /// Mark current position as the save point
    pub fn mark_saved(&mut self) {
        self.save_point = self.current_index;
    }

    /// Check if we're at the save point
    pub fn at_save_point(&self) -> bool {
        self.current_index == self.save_point
    }
}
```

This ensures:
- Undo past save point → dirty
- Redo back to save point → clean
- No manual dirty flag tracking that can desync

### Save Point Initialization Rules

**Critical:** If save_point is not initialized correctly, everything appears "dirty" immediately.

| Scenario | save_point | is_saved | is_dirty | Title shows |
|----------|------------|----------|----------|-------------|
| New document | `= current_index` | `false` | `false` | `Sheet1 — unsaved` |
| Load native (vgrid/xlsx) | `= current_index` | `true` | `false` | `budget.xlsx` |
| Import non-native (csv) | `= current_index` | `false` | `false` | `data.csv — imported...` |
| After first edit | unchanged | unchanged | `true` | `... ●` |
| After save | `= current_index` | `true` | `false` | no dot, no unsaved |

**Key insight:** Imported docs start CLEAN (no dirty dot) but NOT SAVED (shows provenance). The dirty dot only appears after actual edits.

### Title Update Debounce

**Invariant: Only call `set_window_title()` if the string actually changed.**

```rust
impl Spreadsheet {
    /// Update window title if it changed
    pub fn update_title_if_needed(&mut self, window: &mut Window) {
        let new_title = self.document_meta.title_string(self.is_dirty());
        if self.cached_title.as_ref() != Some(&new_title) {
            window.set_window_title(&new_title);
            self.cached_title = Some(new_title);
        }
    }
}
```

This prevents unnecessary OS calls during rapid edits.

### Integration Points

**Spreadsheet struct changes:**
- Add `document_meta: DocumentMeta` field
- Add `cached_title: Option<String>` for debounce

**Update title on:**
- `new()` - Initialize with "Sheet1", unsaved
- `load_file()` - Set display_name from path, determine source
- `save()` / `save_as()` - Update path, clear source (now native), mark saved
- Any edit - Title updates via dirty state change
- Undo/redo - Title updates via dirty state change

**Source determination on load:**

```rust
pub fn load_file(&mut self, path: &Path, cx: &mut Context<Self>) {
    let ext = ext_lower(path);
    let filename = display_filename(path);

    // Determine if this is native or an import
    let is_native = ext.as_ref().map(|e| is_native_ext(e)).unwrap_or(false);

    let (source, is_saved) = if is_native {
        // Native formats - no provenance, considered "saved"
        (None, true)
    } else {
        // Import formats - show provenance, not "saved" until Save As
        (Some(DocumentSource::Imported { filename: filename.clone() }), false)
    };

    self.document_meta.source = source;
    self.document_meta.is_saved = is_saved;
    self.document_meta.display_name = filename;  // Full filename with extension!
    self.document_meta.path = Some(path.to_path_buf());

    // ... load the actual data ...

    // CRITICAL: Set save point AFTER load completes
    // This ensures the document starts "clean" (not dirty)
    self.history.mark_saved();

    // Update title
    self.update_title_if_needed(cx);
}
```

---

## Phase B: Provenance Tracking

### Session Recovery

When restoring from session with unsaved changes:

```rust
if restored_from_session && self.document_meta.path.is_none() {
    self.document_meta.source = Some(DocumentSource::Recovered);
}
```

### Clearing Provenance

Provenance is cleared when saving to a native format. Use `is_native_ext()` for consistency.

```rust
/// Called after successful save (both Save and Save As)
fn finalize_save(&mut self, path: &Path, cx: &mut Context<Self>) {
    let ext = ext_lower(path);
    let becomes_native = ext.as_ref().map(|e| is_native_ext(e)).unwrap_or(false);

    self.document_meta.display_name = display_filename(path);
    self.document_meta.path = Some(path.to_path_buf());
    self.history.mark_saved();

    if becomes_native {
        // Saving to native format clears import provenance
        self.document_meta.source = None;
        self.document_meta.is_saved = true;
    }
    // Note: Exporting to CSV/JSON does NOT clear provenance or mark as saved

    self.update_title_if_needed(cx);
}
```

### Provenance Rules Summary

| Action | Clears provenance? | Sets is_saved? |
|--------|-------------------|----------------|
| Save As .vgrid | Yes | Yes |
| Save As .xlsx | Yes | Yes |
| Export to .csv | No | No |
| Export to .json | No | No |
| Save (to existing native path) | Yes (already clear) | Yes |

---

## Phase C: Title Bar Theming (macOS)

### GPUI APIs Available

```rust
pub struct TitlebarOptions {
    pub title: Option<SharedString>,
    pub appears_transparent: bool,  // Custom titlebar drawing
    pub traffic_light_position: Option<Point<Pixels>>,
}

pub enum WindowBackgroundAppearance {
    Opaque,        // Standard
    Transparent,   // Alpha transparency
    Blurred,       // macOS vibrancy
    MicaBackdrop,  // Windows 11
    MicaAltBackdrop,
}
```

### Implementation Options

**Option A: Native Titlebar (Recommended for v1)**

Keep default titlebar, just set title dynamically:

```rust
WindowOptions {
    titlebar: Some(TitlebarOptions {
        title: Some("VisiGrid".into()),
        appears_transparent: false,
        traffic_light_position: None,
    }),
    window_background: WindowBackgroundAppearance::Opaque,
    ..Default::default()
}
```

**Option B: Themed Titlebar (Future)**

Custom titlebar with theme integration. Only pursue after:
- Multiple documents / split views exist
- Real user feedback requests it
- Complexity is justified

---

## Phase D: Menu Definition System

### Invariants

1. **Menu definitions are independent of Alt accelerators**
   - Menus invoke commands directly
   - Alt accelerators scope the Command Palette
   - They never call each other

2. **Single source of truth**
   - `app_menus()` is the only menu definition
   - macOS native menu and future Windows/Linux in-app menu use the same definition

3. **Menu keyboard shortcuts are platform-appropriate**
   - macOS: Show ⌘ equivalents
   - Windows/Linux: Show Ctrl equivalents

### Menu Definition

Create `src/menus.rs`:

```rust
use gpui::{Menu, MenuItem};
use crate::actions::*;

/// Build the application menu bar.
///
/// This is the single source of truth for menu structure.
/// - macOS: Rendered as native global menu bar
/// - Windows/Linux: Rendered as in-app menu bar (future)
///
/// NOTE: macOS App menu (VisiGrid menu) is kept minimal.
/// Hide/HideOthers/ShowAll/Services are OS behaviors - do NOT implement
/// them as fake app actions. Let the OS handle them automatically,
/// or use GPUI's OsAction/SystemMenuType if available.
pub fn app_menus() -> Vec<Menu> {
    vec![
        // macOS App menu - minimal viable native
        // OS handles Hide/Services automatically; we only add our items
        #[cfg(target_os = "macos")]
        Menu {
            name: "VisiGrid".into(),
            items: vec![
                MenuItem::action("About VisiGrid", ShowAbout),
                MenuItem::separator(),
                MenuItem::action("Preferences…", ShowPreferences),
                MenuItem::separator(),
                // Note: Hide/HideOthers/ShowAll handled by OS
                // Only add if GPUI provides OsAction support
                MenuItem::action("Quit VisiGrid", Quit),
            ],
        },

        Menu {
            name: "File".into(),
            items: vec![
                MenuItem::action("New", NewFile),
                MenuItem::action("Open…", OpenFile),
                MenuItem::separator(),
                MenuItem::action("Save", Save),
                MenuItem::action("Save As…", SaveAs),
                MenuItem::separator(),
                MenuItem::submenu(Menu {
                    name: "Export".into(),
                    items: vec![
                        MenuItem::action("CSV…", ExportCsv),
                        MenuItem::action("TSV…", ExportTsv),
                        MenuItem::action("JSON…", ExportJson),
                        MenuItem::action("Excel (.xlsx)…", ExportXlsx),
                    ],
                }),
                MenuItem::separator(),
                MenuItem::action("Close Window", CloseWindow),
            ],
        },

        Menu {
            name: "Edit".into(),
            items: vec![
                MenuItem::action("Undo", Undo),
                MenuItem::action("Redo", Redo),
                MenuItem::separator(),
                MenuItem::action("Cut", Cut),
                MenuItem::action("Copy", Copy),
                MenuItem::action("Paste", Paste),
                MenuItem::action("Paste Values", PasteValues),
                MenuItem::separator(),
                MenuItem::action("Find…", FindInCells),
                MenuItem::action("Find and Replace…", FindReplace),
                MenuItem::action("Go To…", GoToCell),
                MenuItem::separator(),
                MenuItem::action("Select All", SelectAll),
            ],
        },

        Menu {
            name: "View".into(),
            items: vec![
                MenuItem::action("Command Palette…", ToggleCommandPalette),
                MenuItem::separator(),
                MenuItem::action("Inspector", ToggleInspector),
                MenuItem::action("Problems", ToggleProblems),
                MenuItem::separator(),
                MenuItem::action("Show Formulas", ToggleFormulaView),
                MenuItem::action("Show Zeros", ToggleShowZeros),
                MenuItem::separator(),
                MenuItem::action("Zoom In", ZoomIn),
                MenuItem::action("Zoom Out", ZoomOut),
                MenuItem::action("Reset Zoom", ZoomReset),
            ],
        },

        Menu {
            name: "Format".into(),
            items: vec![
                MenuItem::action("Bold", ToggleBold),
                MenuItem::action("Italic", ToggleItalic),
                MenuItem::action("Underline", ToggleUnderline),
                MenuItem::separator(),
                MenuItem::action("Align Left", AlignLeft),
                MenuItem::action("Align Center", AlignCenter),
                MenuItem::action("Align Right", AlignRight),
                MenuItem::separator(),
                MenuItem::submenu(Menu {
                    name: "Number Format".into(),
                    items: vec![
                        MenuItem::action("Currency", FormatCurrency),
                        MenuItem::action("Percent", FormatPercent),
                    ],
                }),
            ],
        },

        Menu {
            name: "Data".into(),
            items: vec![
                MenuItem::action("Fill Down", FillDown),
                MenuItem::action("Fill Right", FillRight),
                MenuItem::separator(),
                MenuItem::action("AutoSum", AutoSum),
            ],
        },

        Menu {
            name: "Help".into(),
            items: vec![
                MenuItem::action("About VisiGrid", ShowAbout),
            ],
        },
    ]
}
```

### Registration

In `main.rs`:

```rust
cx.set_menus(menus::app_menus());
```

### Windows/Linux In-App Menu Bar (Future)

**Constraint:** If an in-app menu bar is added on Windows/Linux, it must:

1. Be **one row** only
2. Be **text-only** (no icons in the menu bar itself)
3. Use the **same `app_menus()` definition**
4. **Never** introduce ribbons, toolbars, or grouped controls

This prevents accidental Excel-ification.

---

## Implementation Plan

### Step 1: Phase A (Now)

**A1) Add fields to Spreadsheet**
- `document_meta: DocumentMeta`
- `cached_title: Option<String>`

**A2) Add helper functions**
- `display_filename(path: &Path) -> String` - uses `file_name()`, safe fallback
- `ext_lower(path: &Path) -> Option<String>` - lowercase extension
- `is_native_ext(ext: &str) -> bool` - returns true for vgrid/xlsx/xls

**A3) Add single title update method**
```rust
fn update_title_if_needed(&mut self, window: &mut Window) {
    let title = self.document_meta.title_string(self.is_dirty());
    if self.cached_title.as_deref() != Some(&title) {
        window.set_window_title(&title);
        self.cached_title = Some(title);
    }
}
```
**This is the ONLY way titles should update.**

**A4) Wire into 6 lifecycle events**
1. `new()` - Sheet1, is_saved=false, source=None, save_point=current
2. `load_file()` - set meta + mark saved/imported + set save_point
3. `import_*()` - source=Imported, is_saved=false, save_point=current
4. `save()` - mark_saved + is_saved=true + clear source if becomes_native
5. `save_as()` - same as save
6. undo/redo - call update_title_if_needed after history move

**A5) Tests (non-negotiable)**
- Load native doc does NOT show "unsaved"
- Imported doc is clean but not saved (no dirty dot until edit, but shows provenance)
- save_point initialization correct for all entry paths

**Ship this first.** Immediately visible, low risk.

### Step 2: Phase D (Menus)

1. Create `src/menus.rs` with `app_menus()` function
2. macOS App menu: minimal (About, Preferences, Quit) - don't fake OS actions
3. Call `cx.set_menus()` in main.rs
4. Verify all actions work from menu
5. Verify keyboard shortcuts display correctly (⌘ on Mac, Ctrl elsewhere)

**Do not touch theming yet.**

### Step 3: Phase C (Future)

Evaluate Option B (themed titlebar) only after:
- Multiple documents exist
- Split views exist
- Real user feedback requests it

---

## Testing Checklist

### Phase A - Critical
- [ ] New document shows "Sheet1 — unsaved" (not dirty dot)
- [ ] **Load native doc does NOT show "unsaved"** - e.g., `budget.xlsx` not `budget.xlsx — unsaved`
- [ ] **Imported doc is clean but not saved:**
  - No dirty dot until you actually edit something
  - Still shows "unsaved" or provenance depending on platform
- [ ] After save, shows filename without "unsaved"
- [ ] Dirty indicator (●) appears immediately on edit
- [ ] Dirty indicator clears on save
- [ ] Dirty indicator clears on undo to save point
- [ ] Dirty indicator reappears on redo past save point
- [ ] Title updates are debounced (no flicker during rapid edits)
- [ ] Display name includes extension: `budget.xlsx` not `budget`

### Phase A - Platform-specific
- [ ] macOS shows provenance in title bar
- [ ] Windows/Linux does NOT show provenance in title bar (too long)

### Phase B
- [ ] CSV import shows provenance (macOS only in title bar)
- [ ] Session recovery shows "— recovered session" (macOS only)
- [ ] Provenance clears after Save As to .vgrid
- [ ] Provenance clears after Save As to .xlsx
- [ ] Provenance does NOT clear after Export to .csv
- [ ] XLSX open has no provenance (it's native enough)

### Phase D
- [ ] All menu items have working actions
- [ ] Keyboard shortcuts shown correctly (⌘ on Mac, Ctrl elsewhere)
- [ ] Disabled items shown as disabled
- [ ] Submenus expand correctly

---

## Invariants (Must Not Violate)

1. **Dirty state is derived** - Computed from `history.current_index != history.save_point`, never manually tracked
2. **save_point initialized after load** - Every load/import path must call `history.mark_saved()` after loading data
3. **Display name is full filename** - `budget.xlsx` not `budget` (use `file_name()` not `file_stem()`)
4. **Native format is canonical** - `is_native_ext()` is the single source of truth for vgrid/xlsx/xls
5. **Title reflects identity** - Provenance may be rendered elsewhere on non-macOS
6. **Menu definitions are independent of Alt accelerators** - Menus invoke commands; Alt scopes palette
7. **Single menu source** - `app_menus()` drives all menu surfaces
8. **No ribbon creep** - Windows/Linux menu bar (if added) is one text-only row
9. **Don't fake OS actions** - Hide/HideOthers/Services are OS behaviors, not app actions

---

## References

- [GPUI WindowOptions](https://docs.rs/gpui/latest/gpui/struct.WindowOptions.html)
- [Zed Title Bar Discussion](https://github.com/zed-industries/zed/discussions/6734)
- [Zed Transparent Titlebar PR](https://github.com/zed-industries/zed/pull/26403)
- [Zed Window Transparency PR](https://github.com/zed-industries/zed/pull/9610)
- Zed source: `build_window_options()` in `crates/zed/src/zed.rs`
- Zed source: `app_menus()` in `crates/zed/src/zed/app_menus.rs`
