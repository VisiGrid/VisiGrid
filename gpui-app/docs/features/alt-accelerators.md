# Alt Accelerators (Menu Namespace Shortcuts)

## Summary

Support Excel-style Alt keyboard accelerators that open the Command Palette pre-scoped to a menu namespace. This preserves Windows muscle memory without creating duplicate UI or hijacking Option during text entry.

## Design Principles

1. **One command graph, many surfaces** - Alt accelerators are a view onto the existing command system, not a second system
2. **Palette always wins** - Alt opens a scoped palette, not visual menu highlighting
3. **Opt-in only** - Disabled by default, enabled via Preferences
4. **Edit-mode safe** - Never intercept keys during cell/formula editing
5. **Mac-native compatible** - Option key stays available for character composition when disabled
6. **Alt is never stateful** - Only complete chords (alt+letter) trigger actions; Alt keydown alone does nothing and creates no transient state

## User Experience

### Default (Setting Off)
- Option key behaves normally (character composition)
- No Alt menu behavior

### With Setting Enabled

| Input | Result |
|-------|--------|
| `Alt` alone | Nothing (no UI change, no state) |
| `Alt+F` | Opens palette filtered to **File** commands |
| `Alt+E` | Opens palette filtered to **Edit** commands |
| `Alt+V` | Opens palette filtered to **View** commands |
| `Alt+O` | Opens palette filtered to **Format** commands |
| `Alt+D` | Opens palette filtered to **Data** commands |
| `Alt+H` | Opens palette filtered to **Home/Format** commands (modern Excel 2010+) |

### Palette Behavior When Scoped

Once palette is open with a menu scope:
- Type to fuzzy search within that namespace
- Arrow keys to navigate
- Enter to execute
- Escape always closes the palette and clears any active scope

### Backspace Behavior (Explicit Rules)

| State | Backspace Action |
|-------|------------------|
| Query non-empty | Delete last character from query |
| Query empty + scope active | Clear scope, return to full palette |
| Query empty + no scope | Close palette |

### Edge Case: Palette Already Open

If Command Palette is already open and user presses `Alt+F`:
1. Replace current scope with File scope
2. Clear query
3. Keep palette open

This feels natural and avoids "nothing happened" confusion.

### Multi-key Sequences (Muscle Memory)

Fast typists can chain: `Alt+F` then `S` → matches "Save" → Enter executes

This works because:
1. `Alt+F` opens palette scoped to File
2. `S` is typed into the palette search
3. "Save" is the top match
4. Enter executes

No special "sequence mode" needed - it falls out naturally from scoped search.

## Implementation Plan

### Phase 1: Command Categories

Add menu category metadata to commands.

**File: `src/search.rs`**

```rust
/// Menu category for Alt accelerator filtering
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MenuCategory {
    File,
    Edit,
    View,
    Format,
    Data,
    Help,
}
```

**Update CommandId mapping:**

```rust
impl CommandId {
    /// Returns the menu category for this command.
    /// Returns None for commands not addressable via Alt accelerators.
    ///
    /// DESIGN NOTE: Scoped palettes intentionally hide commands without
    /// a menu category. This is deliberate - Alt accelerators provide
    /// structured access to menu commands only. Non-menu commands remain
    /// accessible via the unscoped Command Palette.
    pub fn menu_category(&self) -> Option<MenuCategory> {
        match self {
            // File
            CommandId::NewFile | CommandId::OpenFile | CommandId::Save
            | CommandId::SaveAs | CommandId::ExportCsv | CommandId::ExportTsv
            | CommandId::ExportJson | CommandId::ExportXlsx | CommandId::CloseWindow
                => Some(MenuCategory::File),

            // Edit
            CommandId::Undo | CommandId::Redo | CommandId::Cut | CommandId::Copy
            | CommandId::Paste | CommandId::DeleteCell | CommandId::SelectAll
            | CommandId::StartEdit | CommandId::FindInCells | CommandId::GoToCell
                => Some(MenuCategory::Edit),

            // View
            CommandId::ToggleCommandPalette | CommandId::ToggleInspector
            | CommandId::ToggleFormulaView | CommandId::ToggleShowZeros
                => Some(MenuCategory::View),

            // Format
            CommandId::ToggleBold | CommandId::ToggleItalic | CommandId::ToggleUnderline
            | CommandId::ShowFontPicker
                => Some(MenuCategory::Format),

            // Data
            CommandId::FillDown | CommandId::FillRight
                => Some(MenuCategory::Data),

            // Help
            CommandId::ShowAbout
                => Some(MenuCategory::Help),

            // Commands not in any menu - intentionally excluded from Alt accelerators
            _ => None,
        }
    }
}
```

### Phase 2: Palette Scope (Extensible)

Use an extensible scope enum for future flexibility.

**File: `src/app.rs`**

```rust
/// Palette scope for filtering results
///
/// This abstraction supports menu scoping now and can be extended
/// for selection-scoped commands, contextual palettes, etc.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PaletteScope {
    /// Filter to commands in a specific menu category
    Menu(MenuCategory),
    // Future: Selection, Context, History, etc.
}

pub struct Spreadsheet {
    // ... existing fields ...

    /// Current scope filter for Command Palette (None = no filter)
    pub palette_scope: Option<PaletteScope>,
}
```

**Update palette methods:**

```rust
impl Spreadsheet {
    pub fn show_palette_with_scope(&mut self, scope: PaletteScope, cx: &mut Context<Self>) {
        self.palette_scope = Some(scope);
        self.palette_query.clear();
        self.mode = Mode::CommandPalette;
        cx.notify();
    }

    /// Handle Alt accelerator when palette may already be open
    pub fn apply_menu_scope(&mut self, category: MenuCategory, cx: &mut Context<Self>) {
        // Works whether palette is open or not
        self.palette_scope = Some(PaletteScope::Menu(category));
        self.palette_query.clear();
        self.mode = Mode::CommandPalette;
        cx.notify();
    }

    pub fn palette_results(&self) -> Vec<SearchItem> {
        let mut results = self.search_engine.search(&self.palette_query, 20);

        // Filter by scope if set
        if let Some(scope) = &self.palette_scope {
            match scope {
                PaletteScope::Menu(category) => {
                    results.retain(|item| {
                        if let SearchAction::RunCommand(cmd) = &item.action {
                            cmd.menu_category() == Some(*category)
                        } else {
                            false // Non-command items filtered out in menu scope
                        }
                    });
                }
            }
        }

        results
    }

    pub fn clear_palette_scope(&mut self, cx: &mut Context<Self>) {
        self.palette_scope = None;
        cx.notify();
    }

    pub fn hide_palette(&mut self, cx: &mut Context<Self>) {
        self.palette_scope = None; // Clear scope on close
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Handle backspace in palette with scope-aware logic
    pub fn palette_backspace(&mut self, cx: &mut Context<Self>) {
        if !self.palette_query.is_empty() {
            // Query has content - delete last char
            self.palette_query.pop();
            cx.notify();
        } else if self.palette_scope.is_some() {
            // Query empty but scoped - clear scope
            self.palette_scope = None;
            cx.notify();
        }
        // Query empty and no scope - do nothing (Escape closes)
    }
}
```

### Phase 3: Settings Integration

**File: `src/settings/types.rs`**

```rust
/// Whether Alt key triggers menu accelerators (Excel-style)
/// Only relevant on macOS; on Windows/Linux Alt menus are native
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum AltAccelerators {
    #[default]
    Disabled,
    Enabled,
}
```

**File: `src/settings/user.rs`**

Add to `NavigationSettings`:

```rust
pub struct NavigationSettings {
    // ... existing fields ...

    /// Excel-style Alt menu accelerators (macOS only)
    /// When enabled, Alt+F opens File commands, Alt+E opens Edit, etc.
    #[serde(default, skip_serializing_if = "Setting::is_inherit")]
    pub alt_accelerators: Setting<AltAccelerators>,
}
```

### Phase 4: Actions & Keybindings

**File: `src/actions.rs`**

```rust
// Alt accelerator actions (scoped palette open)
actions!(accelerators, [
    AltFile,    // Alt+F
    AltEdit,    // Alt+E
    AltView,    // Alt+V
    AltFormat,  // Alt+O (Format, using O like Excel)
    AltData,    // Alt+D
    AltHelp,    // Alt+H
]);
```

**File: `src/keybindings.rs`**

```rust
/// Register Alt accelerator keybindings (opt-in)
///
/// IMPORTANT: Alt is never stateful. We only bind complete chords (alt-f),
/// never Alt keydown/keyup. This prevents ghost states and ensures
/// Option key works normally for character composition.
pub fn register_alt_accelerators(cx: &mut App) {
    // These use "alt" which maps to Option on macOS
    cx.bind_keys([
        KeyBinding::new("alt-f", AltFile, Some("Spreadsheet")),
        KeyBinding::new("alt-e", AltEdit, Some("Spreadsheet")),
        KeyBinding::new("alt-v", AltView, Some("Spreadsheet")),
        KeyBinding::new("alt-o", AltFormat, Some("Spreadsheet")),
        KeyBinding::new("alt-d", AltData, Some("Spreadsheet")),
        KeyBinding::new("alt-h", AltHelp, Some("Spreadsheet")),
    ]);
}
```

**Conditional registration in `main.rs`:**

```rust
// Register Alt accelerators only if setting enabled
let alt_accel = settings::user_settings(cx)
    .navigation
    .alt_accelerators
    .as_value()
    .copied()
    .unwrap_or_default();

if alt_accel == AltAccelerators::Enabled {
    keybindings::register_alt_accelerators(cx);
}
```

### Phase 5: Action Handlers

**File: `src/views/mod.rs`**

```rust
// Alt accelerator handlers - work in Navigation OR CommandPalette mode
.on_action(cx.listener(|this, _: &AltFile, _, cx| {
    if this.mode == Mode::Navigation || this.mode == Mode::CommandPalette {
        this.apply_menu_scope(MenuCategory::File, cx);
    }
}))
.on_action(cx.listener(|this, _: &AltEdit, _, cx| {
    if this.mode == Mode::Navigation || this.mode == Mode::CommandPalette {
        this.apply_menu_scope(MenuCategory::Edit, cx);
    }
}))
.on_action(cx.listener(|this, _: &AltView, _, cx| {
    if this.mode == Mode::Navigation || this.mode == Mode::CommandPalette {
        this.apply_menu_scope(MenuCategory::View, cx);
    }
}))
.on_action(cx.listener(|this, _: &AltFormat, _, cx| {
    if this.mode == Mode::Navigation || this.mode == Mode::CommandPalette {
        this.apply_menu_scope(MenuCategory::Format, cx);
    }
}))
.on_action(cx.listener(|this, _: &AltData, _, cx| {
    if this.mode == Mode::Navigation || this.mode == Mode::CommandPalette {
        this.apply_menu_scope(MenuCategory::Data, cx);
    }
}))
.on_action(cx.listener(|this, _: &AltHelp, _, cx| {
    if this.mode == Mode::Navigation || this.mode == Mode::CommandPalette {
        this.apply_menu_scope(MenuCategory::Help, cx);
    }
}))
```

### Phase 6: Palette UI Updates

**File: `src/views/command_palette.rs`**

Show scope indicator when filtered:

```rust
fn scope_name(scope: &PaletteScope) -> &'static str {
    match scope {
        PaletteScope::Menu(cat) => match cat {
            MenuCategory::File => "File",
            MenuCategory::Edit => "Edit",
            MenuCategory::View => "View",
            MenuCategory::Format => "Format",
            MenuCategory::Data => "Data",
            MenuCategory::Help => "Help",
        }
    }
}

// In the header area, show current scope
.child(
    div()
        .flex()
        .items_center()
        .gap_2()
        // Scope badge (when scoped)
        .when(app.palette_scope.is_some(), |d| {
            d.child(
                div()
                    .px_2()
                    .py(px(2.0))
                    .bg(accent.opacity(0.2))
                    .rounded_sm()
                    .text_size(px(11.0))
                    .text_color(accent)
                    .font_weight(FontWeight::MEDIUM)
                    .child(scope_name(app.palette_scope.as_ref().unwrap()))
            )
        })
        // Input area
        .child(/* ... */)
)
```

Update placeholder text:

```rust
let placeholder = if let Some(scope) = &app.palette_scope {
    format!("{} ▸", scope_name(scope))  // "File ▸" - arrow reinforces hierarchy
} else {
    "Execute a command...".to_string()
};
```

### Phase 7: Preferences UI

**File: `src/views/preferences_panel.rs`**

Add toggle in Navigation section (macOS only):

```rust
// Excel-style Alt shortcuts (macOS only)
.when(cfg!(target_os = "macos"), |d: Div| {
    d.child(
        div()
            .flex()
            .flex_col()
            .gap(px(2.0))
            .child(
                div()
                    .flex()
                    .items_center()
                    .justify_between()
                    .child(row_label("Excel-style Alt shortcuts", text_muted))
                    .child(/* checkbox for alt_accelerators */)
            )
            .child(
                div()
                    .text_size(px(10.0))
                    .text_color(text_muted.opacity(0.7))
                    .child("Alt+F for File, Alt+E for Edit, etc.")
            )
    )
})
```

Note: Like `modifier_style`, changing this requires restart.

## Testing Checklist

- [ ] Setting disabled by default
- [ ] Alt+F does nothing when setting disabled
- [ ] Alt+F opens palette scoped to File when enabled
- [ ] Typing narrows results within scope
- [ ] Escape closes and clears scope
- [ ] Backspace with non-empty query edits query
- [ ] Backspace with empty query + scope clears scope
- [ ] Backspace with empty query + no scope does nothing
- [ ] Alt+E while palette open switches to Edit scope
- [ ] Alt accelerators ignored in Edit mode
- [ ] Alt accelerators ignored in Formula mode
- [ ] Option key still types special characters when setting disabled
- [ ] Works with Cmd/Ctrl swap setting
- [ ] Restart message shown when setting changed

## Invariants (Must Not Violate)

1. **Alt is never stateful** - No keydown tracking, no modifier state, only complete chords
2. **Scoped palettes hide non-menu commands** - By design, not by accident
3. **Single command graph** - Menu bar, palette, and Alt accelerators all read from same source
4. **Text input immunity** - Alt accelerators only fire when key context is exactly `Spreadsheet` and not inside any focused text field (protects future text inputs, plugins, modal dialogs)

## Future Enhancements

1. **Visual hint overlay** - Optional transient overlay showing available Alt shortcuts (like Excel's KeyTips)
2. **Custom accelerator mapping** - Let users remap Alt+letter to different menus
3. **Deep sequences** - Support Alt+F+O for specific items (if demand exists)
4. **Selection scope** - `PaletteScope::Selection` for commands applicable to current selection
5. **Context scope** - `PaletteScope::Context` for cell-type-specific commands

## References

- Excel Alt menu behavior: https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel
- macOS keyboard guidelines: https://developer.apple.com/design/human-interface-guidelines/keyboards
- Original design discussion: ChatGPT conversation on Reddit Mac Excel pain points
