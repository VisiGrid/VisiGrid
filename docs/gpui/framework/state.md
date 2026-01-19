# gpui State Management

How to manage application state in gpui.

---

## Entity System

gpui uses an entity system for state management. Your application state lives in an `Entity<T>`.

```rust
pub struct MyApp {
    count: i32,
    items: Vec<String>,
    selected: Option<usize>,
    focus_handle: FocusHandle,
}
```

---

## Creating State

### In Window Creation

```rust
fn main() {
    Application::new().run(|cx: &mut App| {
        cx.open_window(options, |window, cx| {
            cx.new(|cx| MyApp::new(window, cx))  // Creates Entity<MyApp>
        }).unwrap();
    });
}
```

### The Constructor

```rust
impl MyApp {
    pub fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        // Create focus handle for keyboard input
        let focus_handle = cx.focus_handle();

        // Set initial focus
        window.focus(&focus_handle, cx);

        Self {
            count: 0,
            items: vec![],
            selected: None,
            focus_handle,
        }
    }
}
```

---

## Context Types

### `App` - Global Context

Available in `Application::new().run(|cx: &mut App| ...)`:

```rust
cx.bind_keys([...]);           // Register keybindings
cx.open_window(options, ...);   // Create windows
```

### `Context<T>` - Entity Context

Available in entity methods and render:

```rust
impl MyApp {
    fn do_something(&mut self, cx: &mut Context<Self>) {
        cx.notify();                // Trigger re-render
        cx.focus_handle();          // Create focus handle
        cx.spawn(async move |this, cx| ...);  // Spawn async
        cx.write_to_clipboard(...); // Clipboard
        cx.read_from_clipboard();
        cx.prompt_for_paths(...);   // File dialogs
    }
}
```

### `Window` - Window Context

Available in render and some methods:

```rust
impl Render for MyApp {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let size = window.viewport_size();
        window.focus(&self.focus_handle, cx);
        // ...
    }
}
```

### `AsyncApp` - Async Context

Available in spawned async closures:

```rust
cx.spawn(async move |this, cx: AsyncApp| {
    // this: WeakEntity<MyApp>
    // cx: AsyncApp

    let _ = this.update(cx, |this, cx| {
        // Now we have &mut MyApp and &mut Context<MyApp>
        this.data = new_data;
        cx.notify();
    });
})
```

---

## Updating State

### Direct Mutation

```rust
impl MyApp {
    pub fn increment(&mut self, cx: &mut Context<Self>) {
        self.count += 1;
        cx.notify();  // REQUIRED: triggers re-render
    }

    pub fn add_item(&mut self, item: String, cx: &mut Context<Self>) {
        self.items.push(item);
        cx.notify();
    }

    pub fn select(&mut self, index: usize, cx: &mut Context<Self>) {
        self.selected = Some(index);
        cx.notify();
    }
}
```

### From Async

```rust
impl MyApp {
    pub fn load_data(&mut self, cx: &mut Context<Self>) {
        cx.spawn(async move |this, cx| {
            let data = fetch_data().await;

            // Update state back on main thread
            let _ = this.update(cx, |this, cx| {
                this.items = data;
                cx.notify();
            });
        })
        .detach();
    }
}
```

---

## The Notify Pattern

**Critical:** Always call `cx.notify()` after state changes.

```rust
// WRONG - UI won't update
fn increment(&mut self, _cx: &mut Context<Self>) {
    self.count += 1;
    // Missing cx.notify()!
}

// CORRECT
fn increment(&mut self, cx: &mut Context<Self>) {
    self.count += 1;
    cx.notify();
}
```

### When to Call notify()

- After any state change that affects the UI
- At the end of a method that modifies state
- After async operations complete

### When NOT to Call notify()

- In methods that don't change state
- Before returning early (no changes)
- Multiple times in one method (once at end is enough)

---

## Focus System

Focus determines which element receives keyboard input.

### Creating a Focus Handle

```rust
pub struct MyApp {
    focus_handle: FocusHandle,
}

impl MyApp {
    pub fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        let focus_handle = cx.focus_handle();
        window.focus(&focus_handle, cx);
        Self { focus_handle }
    }
}
```

### Using Focus in Render

```rust
impl Render for MyApp {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        div()
            .track_focus(&self.focus_handle)
            .key_context("MyApp")
            // ...
    }
}
```

### Multiple Focus Handles

For complex UIs with multiple focusable areas:

```rust
pub struct MyApp {
    main_focus: FocusHandle,
    sidebar_focus: FocusHandle,
    dialog_focus: FocusHandle,
}

// Focus specific areas
window.focus(&self.dialog_focus, cx);
```

---

## Child Entities

For complex state, use child entities:

```rust
pub struct MyApp {
    editor: Entity<Editor>,
    sidebar: Entity<Sidebar>,
}

impl MyApp {
    pub fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        Self {
            editor: cx.new(|cx| Editor::new(cx)),
            sidebar: cx.new(|cx| Sidebar::new(cx)),
        }
    }
}
```

### Updating Child Entities

```rust
impl MyApp {
    fn update_editor(&mut self, cx: &mut Context<Self>) {
        self.editor.update(cx, |editor, cx| {
            editor.do_something(cx);
        });
    }
}
```

### Reading Child Entities

```rust
impl MyApp {
    fn get_editor_value(&self, cx: &Context<Self>) -> String {
        self.editor.read(cx).get_value()
    }
}
```

---

## WeakEntity

For references that don't prevent cleanup:

```rust
pub struct Dialog {
    parent: WeakEntity<MyApp>,
}

impl Dialog {
    fn close(&mut self, cx: &mut Context<Self>) {
        if let Some(parent) = self.parent.upgrade() {
            parent.update(cx, |parent, cx| {
                parent.close_dialog(cx);
            });
        }
    }
}
```

---

## Async State Updates

### Spawning Async Work

```rust
impl MyApp {
    pub fn fetch_data(&mut self, cx: &mut Context<Self>) {
        let url = self.current_url.clone();

        cx.spawn(async move |this, cx| {
            // Async work
            let response = reqwest::get(&url).await?;
            let data: Vec<Item> = response.json().await?;

            // Update state
            let _ = this.update(cx, |this, cx| {
                this.items = data;
                this.loading = false;
                cx.notify();
            });

            Ok::<_, anyhow::Error>(())
        })
        .detach();

        // Set loading state
        self.loading = true;
        cx.notify();
    }
}
```

### File Dialogs

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
                    let path = path.clone();
                    let _ = this.update(cx, |this, cx| {
                        this.load_file(&path, cx);
                    });
                }
            }
        })
        .detach();
    }
}
```

---

## Global State

For app-wide state accessible anywhere:

```rust
#[derive(Clone, Default)]
struct Theme {
    background: Rgb,
    foreground: Rgb,
}

impl Global for Theme {}
```

### Setting Global State

```rust
// In app initialization
cx.set_global(Theme::default());
```

### Reading Global State

```rust
let theme = Theme::global(cx);
let bg = theme.background;
```

### Updating Global State

```rust
cx.update_global::<Theme, _>(|theme, cx| {
    theme.background = rgb(0x1e1e1e);
});
```

---

## Observation & Events

### Observing Entity Changes

React to any change in another entity:

```rust
impl MyApp {
    fn new(cx: &mut Context<Self>) -> Self {
        let settings = cx.new(|_| Settings::default());

        // Observe settings changes
        cx.observe(&settings, |this, settings, cx| {
            // settings changed, update our state
            this.apply_settings(settings.read(cx), cx);
        }).detach();  // MUST detach!

        Self { settings }
    }
}
```

**Important:** Always call `.detach()` or store the returned `Subscription`.

### Subscribing to Typed Events

For structured event handling:

```rust
// Define event type
struct ItemSelected { index: usize }

// Make entity an emitter
impl EventEmitter<ItemSelected> for ListView {}

// In ListView
impl ListView {
    fn select(&mut self, index: usize, cx: &mut Context<Self>) {
        self.selected = index;
        cx.emit(ItemSelected { index });  // Emit event
        cx.notify();
    }
}

// Parent subscribes
cx.subscribe(&list_view, |this, _list, event: &ItemSelected, cx| {
    this.handle_selection(event.index, cx);
}).detach();
```

### Observe vs Subscribe

| Pattern | Use When |
|---------|----------|
| `cx.observe()` | React to any change, don't need event data |
| `cx.subscribe()` | Need structured event data, specific events |

---

## Common Patterns

### Mode Enum

```rust
#[derive(Clone, Copy, PartialEq)]
pub enum Mode {
    Normal,
    Edit,
    Select,
    Command,
}

pub struct MyApp {
    mode: Mode,
}

impl MyApp {
    pub fn enter_edit_mode(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Edit;
        cx.notify();
    }

    pub fn is_editing(&self) -> bool {
        self.mode == Mode::Edit
    }
}
```

### Selection State

```rust
pub struct MyApp {
    selected: (usize, usize),          // Anchor
    selection_end: Option<(usize, usize)>,  // For range selection
}

impl MyApp {
    pub fn selection_range(&self) -> ((usize, usize), (usize, usize)) {
        let start = self.selected;
        let end = self.selection_end.unwrap_or(start);
        let min = (start.0.min(end.0), start.1.min(end.1));
        let max = (start.0.max(end.0), start.1.max(end.1));
        (min, max)
    }
}
```

### History (Undo/Redo)

```rust
pub struct History {
    undo_stack: Vec<HistoryEntry>,
    redo_stack: Vec<HistoryEntry>,
}

impl History {
    pub fn record(&mut self, entry: HistoryEntry) {
        self.undo_stack.push(entry);
        self.redo_stack.clear();  // Clear redo on new action
    }

    pub fn undo(&mut self) -> Option<HistoryEntry> {
        let entry = self.undo_stack.pop()?;
        self.redo_stack.push(entry.clone());
        Some(entry)
    }

    pub fn redo(&mut self) -> Option<HistoryEntry> {
        let entry = self.redo_stack.pop()?;
        self.undo_stack.push(entry.clone());
        Some(entry)
    }
}
```

### Status Messages

```rust
pub struct MyApp {
    status_message: Option<String>,
}

impl MyApp {
    pub fn set_status(&mut self, msg: impl Into<String>, cx: &mut Context<Self>) {
        self.status_message = Some(msg.into());
        cx.notify();
    }

    pub fn clear_status(&mut self, cx: &mut Context<Self>) {
        self.status_message = None;
        cx.notify();
    }
}
```
