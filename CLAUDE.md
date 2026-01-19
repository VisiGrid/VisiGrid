# VisiGrid Project Context

A native, GPU-accelerated spreadsheet built with gpui (Zed's UI framework).

## Project Structure

```
VisiGrid/
├── gpui-app/src/           # Main gpui application
│   ├── main.rs             # Entry point, window creation
│   ├── app.rs              # Spreadsheet state, core logic
│   ├── actions.rs          # Action definitions
│   ├── keybindings.rs      # Key bindings
│   ├── mode.rs             # Navigation/Edit/GoTo/Find modes
│   ├── history.rs          # Undo/redo
│   ├── file_ops.rs         # File operations (async)
│   └── views/              # UI components
│       ├── mod.rs          # Main render + action handlers
│       ├── grid.rs         # Cell grid
│       ├── headers.rs      # Row/column headers
│       ├── formula_bar.rs
│       ├── status_bar.rs
│       ├── menu_bar.rs
│       ├── goto_dialog.rs
│       └── find_dialog.rs
├── crates/
│   ├── engine/             # Formula engine (96 functions)
│   ├── core/               # Shared types
│   ├── io/                 # File format handling
│   └── config/             # Settings
└── docs/gpui/              # Documentation
    └── framework/          # gpui reference
```

## gpui Core Concepts

### The Three Registers

gpui provides three abstraction levels:

| Register | Purpose | When to Use |
|----------|---------|-------------|
| **Entities** | Observable state | Shared state between components |
| **Views** | Declarative UI | 90% of typical UI work |
| **Elements** | Low-level rendering | Custom lists, text editors, max perf |

1. **Entities** - State owned by the App, accessed via `Entity<T>` handles
2. **Views** - Implement `Render` trait, Tailwind-inspired styling, auto re-render
3. **Elements** - Implement `Element` trait, direct GPU access, Taffy layout

### Context Hierarchy

```
App (base)
 └── VisualContext (+ windows, focus)
      └── Context<T> (+ notify, emit, observe)
```

| Feature | App | VisualContext | Context<T> |
|---------|-----|---------------|------------|
| Create entities | ✅ | ✅ | ✅ |
| Global state | ✅ | ✅ | ✅ |
| Spawn tasks | ✅ | ✅ | ✅ |
| Open windows | ❌ | ✅ | ✅ |
| `cx.notify()` | ❌ | ❌ | ✅ |
| `cx.emit()` | ❌ | ❌ | ✅ |
| `cx.observe()` | ❌ | ❌ | ✅ |

### Ownership Model

**"You don't own your data - the App does."**

`Entity<T>` is a handle, not the data itself. Cheap to clone, requires context to access.

```rust
// Create
let counter = cx.new(|_cx| Counter { count: 0 });

// Read
let count = counter.read(cx).count;

// Update
counter.update(cx, |counter, cx| {
    counter.count += 1;
    cx.notify();  // REQUIRED
});
```

## gpui Quick Reference

### Application Bootstrap

```rust
use gpui::*;

fn main() {
    Application::new().run(|cx: &mut App| {
        keybindings::register(cx);

        cx.open_window(
            WindowOptions {
                window_bounds: Some(WindowBounds::Windowed(bounds)),
                ..Default::default()
            },
            |window, cx| cx.new(|cx| MyApp::new(window, cx)),
        ).unwrap();
    });
}
```

### Defining Actions

```rust
use gpui::actions;

// Syntax: actions!(namespace, [Action1, Action2, ...])
actions!(navigation, [MoveUp, MoveDown, MoveLeft, MoveRight]);
actions!(editing, [StartEdit, ConfirmEdit, CancelEdit]);
actions!(file, [NewFile, OpenFile, Save]);
```

### Registering Keybindings

```rust
pub fn register(cx: &mut App) {
    cx.bind_keys([
        // KeyBinding::new("key-combo", Action, Some("ContextName"))
        KeyBinding::new("up", MoveUp, Some("Spreadsheet")),
        KeyBinding::new("ctrl-s", Save, Some("Spreadsheet")),
        KeyBinding::new("ctrl-shift-s", SaveAs, Some("Spreadsheet")),
    ]);
}
```

**Key combo format:**
- Modifiers: `ctrl`, `shift`, `alt`, `cmd` (macOS), `platform` (ctrl on Linux, cmd on macOS)
- Join with `-`: `ctrl-shift-s`, `alt-enter`
- Special keys: `enter`, `escape`, `tab`, `backspace`, `delete`, `space`
- F-keys: `f1` through `f12`
- Navigation: `up`, `down`, `left`, `right`, `home`, `end`, `pageup`, `pagedown`

### State Management

```rust
pub struct MyApp {
    data: Vec<String>,
    selected: usize,
    focus_handle: FocusHandle,
}

impl MyApp {
    pub fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        let focus_handle = cx.focus_handle();
        window.focus(&focus_handle, cx);
        Self { data: vec![], selected: 0, focus_handle }
    }

    // Methods take &mut self and cx: &mut Context<Self>
    pub fn do_something(&mut self, cx: &mut Context<Self>) {
        self.selected += 1;
        cx.notify();  // ALWAYS call to trigger re-render
    }
}
```

### Render Trait

```rust
impl Render for MyApp {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        div()
            .key_context("MyApp")  // Context for keybindings
            .track_focus(&self.focus_handle)
            .child(self.render_content(cx))
    }
}
```

### Element Building (Fluent API)

```rust
div()
    // Layout
    .flex()
    .flex_col()              // or .flex_row() (default)
    .flex_1()                // flex-grow: 1
    .flex_shrink_0()         // prevent shrinking
    .gap_2()                 // gap between children
    .items_center()          // align-items: center
    .justify_center()        // justify-content: center

    // Sizing
    .w(px(100.0))            // width in pixels
    .h(px(50.0))             // height
    .w_full()                // width: 100%
    .h_full()
    .size_full()             // both
    .min_w(px(50.0))
    .max_h(px(200.0))

    // Spacing
    .p_2()                   // padding (all sides)
    .px_2()                  // padding horizontal
    .py_1()                  // padding vertical
    .m_2()                   // margin
    .mx_1()

    // Visual
    .bg(rgb(0x1e1e1e))       // background color
    .text_color(rgb(0xffffff))
    .border_1()              // 1px border
    .border_color(rgb(0x3d3d3d))
    .border_b_1()            // bottom border only
    .rounded_md()            // border-radius
    .overflow_hidden()

    // Text
    .text_sm()               // font size
    .text_lg()
    .font_weight(FontWeight::BOLD)
    .italic()
    .underline()

    // Interactivity
    .cursor_pointer()
    .hover(|style| style.bg(rgb(0x404040)))
    .active(|style| style.bg(rgb(0x094771)))

    // Children
    .child(text)
    .children(vec.iter().map(|item| render_item(item)))
```

### Event Handlers

```rust
div()
    .id("my-element")  // Required for mouse events

    // Action handlers (from keybindings)
    .on_action(cx.listener(|this, _: &MyAction, _, cx| {
        this.do_something(cx);
    }))

    // Mouse events
    .on_mouse_down(MouseButton::Left, cx.listener(|this, event: &MouseDownEvent, _, cx| {
        if event.click_count == 2 {
            this.start_edit(cx);
        } else if event.modifiers.shift {
            this.extend_selection(cx);
        }
    }))

    // Keyboard events
    .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
        if event.keystroke.key == "enter" {
            this.confirm(cx);
        } else if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control {
                this.insert_text(key_char, cx);
            }
        }
    }))

    // Scroll wheel
    .on_scroll_wheel(cx.listener(|this, event: &ScrollWheelEvent, _, cx| {
        let delta = event.delta.pixel_delta(px(24.0));
        let dy: f32 = delta.y.into();
        this.scroll((-dy / 24.0).round() as i32, cx);
    }))
```

### Conditional Rendering

```rust
use gpui::prelude::FluentBuilder;

div()
    .when(condition, |div| {
        div.child(some_element)
    })
    .when(show_overlay, |div| {
        div.child(render_overlay())
    })
```

### Async Operations

```rust
impl MyApp {
    pub fn open_file(&mut self, cx: &mut Context<Self>) {
        let future = cx.prompt_for_paths(PathPromptOptions {
            files: true,
            directories: false,
            multiple: false,
            prompt: Some("Open".into()),
        });

        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(paths))) = future.await {
                if let Some(path) = paths.first() {
                    let _ = this.update(cx, |this, cx| {
                        this.load_file(path, cx);
                    });
                }
            }
        })
        .detach();
    }
}
```

### Clipboard

```rust
// Write
cx.write_to_clipboard(ClipboardItem::new_string(text));

// Read
if let Some(item) = cx.read_from_clipboard() {
    if let Some(text) = item.text() {
        // use text
    }
}
```

### Colors

```rust
rgb(0x1e1e1e)              // RGB hex
rgba(0x1e1e1e80)           // RGBA with alpha (80 = ~50%)
hsla(0.0, 0.0, 0.12, 1.0)  // HSLA

// Convert rgb to Hsla (required for some properties)
rgb(0x1e1e1e).into()
```

### Units

```rust
px(100.0)    // Pixels
rems(1.0)    // Relative to font size

// Converting Pixels to f32
let width: f32 = self.window_size.width.into();
```

### Global State

```rust
#[derive(Clone, Default)]
struct Theme { background: Rgb }
impl Global for Theme {}

// Set
cx.set_global(Theme::default());

// Read
let theme = Theme::global(cx);

// Update
cx.update_global::<Theme, _>(|theme, cx| {
    theme.background = rgb(0x1e1e1e);
});
```

### Observation Patterns

```rust
// Observe any change to an entity
cx.observe(&other_entity, |this, other, cx| {
    // other changed, react here
    cx.notify();
}).detach();  // MUST detach or store Subscription

// Subscribe to typed events
cx.subscribe(&child, |this, _child, event: &ChildEvent, cx| {
    // handle event
}).detach();
```

### Event Emission

```rust
// Make entity an event emitter
impl EventEmitter<MyEvent> for MyEntity {}

// Emit events
cx.emit(MyEvent { data: 42 });
```

## Common Gotchas

1. **Always call `cx.notify()`** after state changes to trigger re-render
2. **`id()` is required** for mouse event handlers
3. **Actions need context**: Use `.key_context("Name")` and match in keybindings
4. **Async closures**: `cx.spawn(async move |this, cx| { ... })` - `this` is `WeakEntity`
5. **Pixels to f32**: Use `.into()` since `Pixels.0` is private
6. **FluentBuilder import**: Need `use gpui::prelude::FluentBuilder` for `.when()`
7. **Return types**: `.id()` changes return from `Div` to `Stateful<Div>`
8. **Font styling**: Use `.italic()` method, not `.font_style()`
9. **Dropping subscriptions**: Must call `.detach()` or store the `Subscription`
10. **Reference cycles**: Use `WeakEntity<T>` for bidirectional or self references
11. **Cannot store contexts**: Contexts are temporary params, never struct fields

## Engine API

```rust
// Get cell
sheet.get_cell(row, col)           // Returns &Cell
sheet.get_display(row, col)        // Returns formatted String
sheet.get_raw(row, col)            // Returns raw String (formula or value)

// Set cell
sheet.set_value(row, col, "text")  // Auto-detects formulas (=...)

// Formatting
sheet.get_format(row, col)         // Returns CellFormat
sheet.toggle_bold(row, col)
sheet.toggle_italic(row, col)
sheet.toggle_underline(row, col)

// Check emptiness
cell.value.raw_display().is_empty()
```

## Build & Run

```bash
cd gpui-app
cargo run              # Debug build
cargo build --release  # Release build
```

Binary installs to: `~/.local/bin/visigrid`

## External Resources

- **gpui Cheat Book**: https://gpui.ramones.dev/ (community reference)
- **Zed Source**: https://github.com/zed-industries/zed/tree/main/crates/gpui
- **gpui Examples**: `zed/crates/gpui/examples/`
