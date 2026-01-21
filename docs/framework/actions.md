# gpui Actions & Keybindings

Actions are the primary way to handle keyboard shortcuts and command dispatch in gpui.

---

## Defining Actions

Use the `actions!` macro to define action types:

```rust
use gpui::actions;

// Syntax: actions!(namespace, [Action1, Action2, ...])
actions!(file, [NewFile, OpenFile, Save, SaveAs, Close]);
actions!(edit, [Undo, Redo, Cut, Copy, Paste]);
actions!(navigation, [MoveUp, MoveDown, MoveLeft, MoveRight]);
```

This generates zero-sized structs for each action:

```rust
// Equivalent to:
#[derive(Clone, Default, PartialEq)]
pub struct Save;
impl gpui::Action for Save { ... }
```

---

## Registering Keybindings

Register keybindings in your app initialization:

```rust
pub fn register_keybindings(cx: &mut App) {
    cx.bind_keys([
        // KeyBinding::new(key_combo, Action, context)
        KeyBinding::new("ctrl-s", Save, Some("Editor")),
        KeyBinding::new("ctrl-shift-s", SaveAs, Some("Editor")),
        KeyBinding::new("ctrl-n", NewFile, None),  // Global (any context)
    ]);
}
```

### Key Combo Format

| Component | Examples |
|-----------|----------|
| Modifiers | `ctrl`, `shift`, `alt`, `cmd`, `platform` |
| Letters | `a`, `b`, `c`, ... `z` |
| Numbers | `0`, `1`, ... `9` |
| Function keys | `f1`, `f2`, ... `f12` |
| Navigation | `up`, `down`, `left`, `right`, `home`, `end`, `pageup`, `pagedown` |
| Special | `enter`, `escape`, `tab`, `backspace`, `delete`, `space` |

**Modifier notes:**
- `platform` = `ctrl` on Linux/Windows, `cmd` on macOS
- Chain with `-`: `ctrl-shift-s`, `alt-enter`
- Order doesn't matter: `ctrl-shift-s` = `shift-ctrl-s`

### Examples

```rust
KeyBinding::new("ctrl-s", Save, Some("MyApp")),
KeyBinding::new("ctrl-shift-s", SaveAs, Some("MyApp")),
KeyBinding::new("alt-enter", Confirm, Some("Dialog")),
KeyBinding::new("f2", StartEdit, Some("Grid")),
KeyBinding::new("escape", Cancel, Some("Modal")),
KeyBinding::new("ctrl-shift-p", CommandPalette, Some("MyApp")),
```

---

## Key Context

Context determines which keybindings are active. Set it on your root element:

```rust
impl Render for MyApp {
    fn render(&mut self, ...) -> impl IntoElement {
        div()
            .key_context("MyApp")  // Must match keybinding context
            .track_focus(&self.focus_handle)
            // ...
    }
}
```

**Context matching:**
- `Some("MyApp")` - Only active when "MyApp" context is focused
- `None` - Global, always active

Nested contexts work. If "Editor" is inside "MyApp", both contexts are active.

---

## Handling Actions

### In Render (Recommended)

```rust
impl Render for MyApp {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        div()
            .key_context("MyApp")
            .track_focus(&self.focus_handle)
            .on_action(cx.listener(|this, _: &Save, _, cx| {
                this.save(cx);
            }))
            .on_action(cx.listener(|this, _: &NewFile, _, cx| {
                this.new_file(cx);
            }))
            // ... more handlers
    }
}
```

### The Listener Closure

```rust
cx.listener(|this, action, window, cx| {
    // this: &mut Self (your app state)
    // action: &ActionType (the action that was triggered)
    // window: &mut Window
    // cx: &mut Context<Self>
})
```

If you don't need all parameters:

```rust
cx.listener(|this, _: &Save, _, cx| {
    this.save(cx);
})
```

---

## Actions with Data

For actions that carry data, define them manually:

```rust
#[derive(Clone, PartialEq)]
pub struct GoToLine(pub usize);

impl gpui::Action for GoToLine {
    // ... implementation
}

// Or use the macro with impl
actions!(editor, [GoToLine]);

// Usage
div().on_action(cx.listener(|this, action: &GoToLine, _, cx| {
    this.go_to_line(action.0, cx);
}))
```

---

## Dispatching Actions Programmatically

```rust
// From within a Context
cx.dispatch_action(Box::new(Save));

// With window
window.dispatch_action(Box::new(Save), cx);
```

---

## Action Bubbling

Actions bubble up through the element tree until handled:

```
┌─────────────────────────────────┐
│  App (key_context: "App")       │
│  ┌───────────────────────────┐  │
│  │  Panel (key_context: "P") │  │
│  │  ┌─────────────────────┐  │  │
│  │  │  Editor (focused)   │  │  │
│  │  │  key_context: "Ed"  │  │  │
│  │  └─────────────────────┘  │  │
│  └───────────────────────────┘  │
└─────────────────────────────────┘

Ctrl+S pressed in Editor:
1. Check Editor handlers → not found
2. Check Panel handlers → not found
3. Check App handlers → Save found! Execute.
```

---

## Multiple Bindings for Same Action

You can bind multiple key combos to one action:

```rust
cx.bind_keys([
    KeyBinding::new("ctrl-z", Undo, Some("Editor")),
    KeyBinding::new("cmd-z", Undo, Some("Editor")),  // macOS
]);

cx.bind_keys([
    KeyBinding::new("ctrl-y", Redo, Some("Editor")),
    KeyBinding::new("ctrl-shift-z", Redo, Some("Editor")),
]);
```

---

## Raw Keyboard Events

For handling keys not bound to actions (like text input):

```rust
div()
    .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
        // Check for specific keys
        if event.keystroke.key == "enter" {
            this.confirm(cx);
            return;
        }

        if event.keystroke.key == "escape" {
            this.cancel(cx);
            return;
        }

        // Handle text input
        if let Some(key_char) = &event.keystroke.key_char {
            // Skip if modifiers are held (those are for shortcuts)
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars() {
                    this.insert_char(c, cx);
                }
            }
        }
    }))
```

### KeyDownEvent Structure

```rust
pub struct KeyDownEvent {
    pub keystroke: Keystroke,
    pub is_held: bool,  // Key is being held down
}

pub struct Keystroke {
    pub key: String,              // "a", "enter", "f1", etc.
    pub key_char: Option<String>, // The character produced (if any)
    pub modifiers: Modifiers,
}

pub struct Modifiers {
    pub control: bool,
    pub alt: bool,
    pub shift: bool,
    pub platform: bool,  // cmd on macOS, ctrl on Linux/Windows
    pub function: bool,  // Fn key
}
```

---

## Best Practices

### 1. Organize Actions by Category

```rust
// actions.rs
actions!(navigation, [MoveUp, MoveDown, MoveLeft, MoveRight, JumpToStart, JumpToEnd]);
actions!(editing, [StartEdit, ConfirmEdit, CancelEdit, Delete, Backspace]);
actions!(clipboard, [Copy, Cut, Paste]);
actions!(file, [New, Open, Save, SaveAs, Export]);
actions!(view, [ZoomIn, ZoomOut, ToggleSidebar, ToggleFullscreen]);
```

### 2. Centralize Keybindings

```rust
// keybindings.rs
pub fn register(cx: &mut App) {
    cx.bind_keys([
        // Navigation
        KeyBinding::new("up", MoveUp, Some("Grid")),
        KeyBinding::new("down", MoveDown, Some("Grid")),
        // ...
    ]);
}
```

### 3. Use Context Wisely

- Global actions (like Quit): `None`
- App-wide actions: `Some("AppName")`
- Component-specific: `Some("ComponentName")`

### 4. Handle Mode-Specific Behavior in Handlers

```rust
.on_action(cx.listener(|this, _: &Delete, _, cx| {
    if this.mode.is_editing() {
        this.delete_char(cx);
    } else {
        this.delete_selection(cx);
    }
}))
```

---

## Common Patterns

### Navigation with Optional Extend

```rust
actions!(navigation, [MoveUp, ExtendUp]);

// Same key, different modifier
KeyBinding::new("up", MoveUp, Some("Grid")),
KeyBinding::new("shift-up", ExtendUp, Some("Grid")),
```

### Mode-Dependent Actions

```rust
.on_action(cx.listener(|this, _: &ConfirmEdit, _, cx| {
    match this.mode {
        Mode::Edit => this.confirm_edit(cx),
        Mode::Dialog => this.confirm_dialog(cx),
        _ => {}
    }
}))
```

### Action in Specific Element

```rust
// Button that triggers an action
div()
    .id("save-button")
    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
        this.save(cx);  // Or dispatch: cx.dispatch_action(Box::new(Save))
    }))
    .child("Save")
```
