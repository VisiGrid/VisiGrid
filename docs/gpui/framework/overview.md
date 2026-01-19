# gpui Framework Overview

gpui is Zed's GPU-accelerated UI framework for building native desktop applications in Rust.

---

## What is gpui?

gpui is **not** a web renderer. It's a native UI framework that:

- Renders directly to GPU (Metal on macOS, Vulkan on Linux)
- Has Wayland-native support
- Uses immediate-mode rendering with retained state
- Provides a fluent builder API for UI construction
- Has built-in accessibility support
- Handles platform integration (windows, clipboard, file dialogs)

---

## Core Concepts

### 1. Application

The entry point. Creates the event loop and manages windows.

```rust
Application::new().run(|cx: &mut App| {
    // cx is the App context - global state
    cx.open_window(options, |window, cx| {
        cx.new(|cx| MyApp::new(window, cx))
    });
});
```

### 2. Entity / Model

Your application state. Managed by gpui's entity system.

```rust
pub struct MyApp {
    count: i32,
    focus_handle: FocusHandle,
}

// Created via cx.new()
let entity = cx.new(|cx| MyApp::new(cx));
```

### 3. Context

Different context types for different scopes:

| Type | Scope | Used For |
|------|-------|----------|
| `App` | Global | Window creation, global keybindings |
| `Context<T>` | Entity | State access, spawning, notifications |
| `Window` | Window | Focus, viewport size, platform features |

### 4. Render Trait

Makes your entity renderable as a window root.

```rust
impl Render for MyApp {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        div().child("Hello")
    }
}
```

### 5. Elements

The UI building blocks. Built with a fluent API.

```rust
div()           // Container element
text("Hello")   // Text element
svg()           // SVG element
img()           // Image element
canvas()        // Custom drawing
```

### 6. Actions

Typed commands that can be triggered by keybindings or programmatically.

```rust
actions!(app, [Save, Quit]);

// Register keybinding
cx.bind_keys([
    KeyBinding::new("ctrl-s", Save, Some("MyApp")),
]);

// Handle in render
div().on_action(cx.listener(|this, _: &Save, _, cx| {
    this.save(cx);
}))
```

---

## Rendering Model

gpui uses **immediate-mode rendering with retained state**:

1. You call `render()` which returns a tree of elements
2. gpui diffs against previous render
3. Only changed elements are re-rendered to GPU
4. Call `cx.notify()` to trigger a re-render

```rust
impl MyApp {
    fn increment(&mut self, cx: &mut Context<Self>) {
        self.count += 1;
        cx.notify();  // Triggers render()
    }
}
```

---

## Coordinate System

- Origin: Top-left of window
- Units: Pixels (`px(100.0)`) or rems (`rems(1.0)`)
- Y increases downward
- Sizes/positions use `Pixels` type

```rust
let width = px(100.0);
let size = Size { width: px(200.0), height: px(100.0) };
let point = Point::new(px(50.0), px(50.0));
let bounds = Bounds { origin: point, size };
```

---

## Focus System

Focus determines which element receives keyboard input.

```rust
struct MyApp {
    focus_handle: FocusHandle,
}

impl MyApp {
    fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        let focus_handle = cx.focus_handle();
        window.focus(&focus_handle, cx);  // Initial focus
        Self { focus_handle }
    }
}

impl Render for MyApp {
    fn render(&mut self, ...) -> impl IntoElement {
        div()
            .track_focus(&self.focus_handle)
            .key_context("MyApp")  // Required for keybindings
    }
}
```

---

## Event Flow

1. **Platform event** (keyboard, mouse, etc.)
2. **gpui processes** - matches keybindings, finds target element
3. **Action dispatch** - if keybinding matched
4. **Event handlers** - `on_key_down`, `on_mouse_down`, etc.
5. **Bubbling** - events bubble up through parent elements

```
Keyboard Event
    ↓
Match Keybinding → Dispatch Action → .on_action() handler
    ↓ (if no match)
Fire KeyDownEvent → .on_key_down() handler
    ↓
Bubble to parent
```

---

## Async Model

gpui has its own async runtime. Use `cx.spawn()` for async operations:

```rust
cx.spawn(async move |this, cx| {
    // `this` is WeakEntity<MyApp>
    // `cx` is AsyncApp

    // Do async work
    let data = fetch_data().await;

    // Update state
    let _ = this.update(cx, |this, cx| {
        this.data = data;
        cx.notify();
    });
})
.detach();  // Run independently
```

**Important:** The closure receives `WeakEntity<T>` and `AsyncApp`, not direct references.

---

## Platform Integration

### File Dialogs

```rust
// Open file
let future = cx.prompt_for_paths(PathPromptOptions {
    files: true,
    directories: false,
    multiple: false,
    prompt: Some("Open".into()),
});

// Save file
let future = cx.prompt_for_new_path(&directory, Some("filename.txt"));
```

### Clipboard

```rust
// Write
cx.write_to_clipboard(ClipboardItem::new_string(text));

// Read
if let Some(item) = cx.read_from_clipboard() {
    let text = item.text();
}
```

### Window

```rust
// Get viewport size
let size = window.viewport_size();
let width: f32 = size.width.into();

// Set title (if supported)
window.set_title("My App");
```

---

## Dependencies

```toml
[dependencies]
gpui = { git = "https://github.com/zed-industries/zed", branch = "main" }
```

Note: gpui is part of the Zed monorepo. You typically depend on it via git.

---

## Resources

- **gpui Cheat Book**: https://gpui.ramones.dev/ (community reference, WIP)
- **Zed source code**: https://github.com/zed-industries/zed/tree/main/crates/gpui
- **gpui examples**: `zed/crates/gpui/examples/`
- **docs.rs/gpui**: API documentation (sparse)
