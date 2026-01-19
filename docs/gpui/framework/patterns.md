# gpui Patterns & Gotchas

Common patterns and pitfalls when building with gpui.

---

## Common Gotchas

### 1. Missing cx.notify()

**Problem:** UI doesn't update after state change.

```rust
// WRONG
fn increment(&mut self, _cx: &mut Context<Self>) {
    self.count += 1;
    // Missing notify!
}

// CORRECT
fn increment(&mut self, cx: &mut Context<Self>) {
    self.count += 1;
    cx.notify();
}
```

### 2. Missing Element ID for Mouse Events

**Problem:** Mouse events don't fire.

```rust
// WRONG
div()
    .on_mouse_down(...)  // Won't work without id!

// CORRECT
div()
    .id("clickable")
    .on_mouse_down(...)
```

### 3. Return Type Changes with .id()

**Problem:** Type mismatch errors.

```rust
// This returns Div
fn render_simple() -> Div { div() }

// This returns Stateful<Div>
fn render_with_id() -> Stateful<Div> { div().id("foo") }

// Solution: Use impl IntoElement
fn render() -> impl IntoElement { div().id("foo") }
```

### 4. FluentBuilder Import for .when()

**Problem:** `when` method not found.

```rust
// WRONG
div().when(cond, |d| d.child("hi"))  // Error: method not found

// CORRECT
use gpui::prelude::FluentBuilder;
div().when(cond, |d| d.child("hi"))
```

### 5. Pixels Private Field

**Problem:** Can't access `Pixels.0` directly.

```rust
// WRONG
let width = self.window_size.width.0;  // Error: private field

// CORRECT
let width: f32 = self.window_size.width.into();
```

### 6. Font Style Method

**Problem:** No `font_style()` method.

```rust
// WRONG
div().font_style(FontStyle::Italic)  // Doesn't exist

// CORRECT
div().italic()
```

### 7. Async Closure Types

**Problem:** Confusion about async closure parameters.

```rust
// The closure receives WeakEntity and AsyncApp
cx.spawn(async move |this, cx| {
    // this: WeakEntity<MyApp>  (not &mut MyApp)
    // cx: AsyncApp             (not Context)

    // Must use .update() to access state
    let _ = this.update(cx, |this, cx| {
        // Now: this is &mut MyApp, cx is &mut Context<MyApp>
        this.data = new_data;
        cx.notify();
    });
})
```

### 8. rgba() Alpha Format

**Problem:** Wrong alpha value format.

```rust
// WRONG - thinking 0.5 = 50%
rgba(0x1e1e1e, 0.5)  // Wrong signature

// CORRECT - alpha in hex (00-ff)
rgba(0x1e1e1e80)  // 0x80 ≈ 50% opacity (128/255)
rgba(0x1e1e1eff)  // Full opacity
rgba(0x1e1e1e00)  // Fully transparent
```

### 9. Action Context Mismatch

**Problem:** Keybinding doesn't trigger action.

```rust
// Keybinding
KeyBinding::new("ctrl-s", Save, Some("Editor"))

// WRONG - context doesn't match
div().key_context("MyApp")  // "MyApp" != "Editor"

// CORRECT
div().key_context("Editor")
```

### 10. Children Order Matters

**Problem:** Elements render in unexpected order.

```rust
// Renders top-to-bottom in flex-col
div()
    .flex_col()
    .child(header)   // Renders first (top)
    .child(content)  // Renders second (middle)
    .child(footer)   // Renders third (bottom)
```

---

## Common Patterns

### Modal Dialog Overlay

```rust
impl Render for MyApp {
    fn render(&mut self, ...) -> impl IntoElement {
        div()
            .relative()
            .size_full()
            .child(main_content)
            .when(self.show_dialog, |div| {
                div.child(
                    // Overlay backdrop
                    div()
                        .absolute()
                        .inset_0()
                        .bg(rgba(0x00000080))
                        .z_index(50)
                        .child(
                            // Centered dialog
                            div()
                                .absolute()
                                .top(px(100.0))
                                .left_1_2()
                                .w(px(400.0))
                                .bg(rgb(0x2d2d2d))
                                .rounded_md()
                                .p_4()
                                .child(dialog_content)
                        )
                )
            })
    }
}
```

### Toolbar with Dividers

```rust
fn render_toolbar(cx: &mut Context<App>) -> impl IntoElement {
    div()
        .flex()
        .h(px(32.0))
        .items_center()
        .bg(rgb(0x2d2d2d))
        .gap_1()
        .child(toolbar_button("New", cx, |app, cx| app.new(cx)))
        .child(toolbar_button("Open", cx, |app, cx| app.open(cx)))
        .child(toolbar_button("Save", cx, |app, cx| app.save(cx)))
        .child(divider())
        .child(toolbar_button("Undo", cx, |app, cx| app.undo(cx)))
        .child(toolbar_button("Redo", cx, |app, cx| app.redo(cx)))
}

fn toolbar_button(
    label: &'static str,
    cx: &mut Context<App>,
    action: impl Fn(&mut App, &mut Context<App>) + 'static,
) -> impl IntoElement {
    div()
        .id(ElementId::Name(label.into()))
        .px_2()
        .py_1()
        .rounded_sm()
        .cursor_pointer()
        .hover(|s| s.bg(rgb(0x404040)))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            action(this, cx);
        }))
        .child(label)
}

fn divider() -> impl IntoElement {
    div()
        .w(px(1.0))
        .h(px(16.0))
        .mx_1()
        .bg(rgb(0x404040))
}
```

### Grid Layout

```rust
fn render_grid(app: &App, cx: &mut Context<App>) -> impl IntoElement {
    div()
        .flex()
        .flex_col()
        .children(
            (0..app.visible_rows()).map(|row_idx| {
                div()
                    .flex()
                    .h(px(CELL_HEIGHT))
                    .children(
                        (0..app.visible_cols()).map(|col_idx| {
                            render_cell(app, row_idx, col_idx, cx)
                        })
                    )
            })
        )
}
```

### Scrollable List with Selection

```rust
fn render_list(app: &App, cx: &mut Context<App>) -> impl IntoElement {
    div()
        .flex_1()
        .overflow_y_scroll()
        .children(
            app.items.iter().enumerate().map(|(idx, item)| {
                let is_selected = app.selected == Some(idx);

                div()
                    .id(ElementId::NamedInteger("item".into(), idx))
                    .w_full()
                    .px_2()
                    .py_1()
                    .bg(if is_selected { rgb(0x094771) } else { rgb(0x1e1e1e) })
                    .hover(|s| s.bg(rgb(0x2a2d2e)))
                    .cursor_pointer()
                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                        this.select(idx, cx);
                    }))
                    .child(item.name.clone())
            })
        )
}
```

### Status Bar with Sections

```rust
fn render_status_bar(app: &App) -> impl IntoElement {
    div()
        .flex()
        .flex_shrink_0()
        .h(px(24.0))
        .w_full()
        .bg(rgb(0x007acc))
        .text_color(rgb(0xffffff))
        .text_sm()
        // Left section
        .child(
            div()
                .flex()
                .flex_1()
                .items_center()
                .px_2()
                .child(format!("{}", app.mode))
        )
        // Center section
        .child(
            div()
                .flex()
                .items_center()
                .child(app.status_message.clone().unwrap_or_default())
        )
        // Right section
        .child(
            div()
                .flex()
                .flex_1()
                .items_center()
                .justify_end()
                .px_2()
                .child(format!("Row {}, Col {}", app.row + 1, app.col + 1))
        )
}
```

### Input Field

```rust
pub struct InputField {
    value: String,
    placeholder: String,
    focus_handle: FocusHandle,
}

impl Render for InputField {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let display_value = if self.value.is_empty() {
            self.placeholder.clone()
        } else {
            self.value.clone()
        };

        let text_color = if self.value.is_empty() {
            rgb(0x808080)  // Placeholder color
        } else {
            rgb(0xffffff)
        };

        div()
            .id("input")
            .track_focus(&self.focus_handle)
            .w_full()
            .px_2()
            .py_1()
            .bg(rgb(0x3c3c3c))
            .border_1()
            .border_color(rgb(0x007acc))
            .rounded_sm()
            .text_color(text_color)
            .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                if event.keystroke.key == "backspace" {
                    this.value.pop();
                    cx.notify();
                } else if let Some(key_char) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control {
                        this.value.push_str(key_char);
                        cx.notify();
                    }
                }
            }))
            .child(display_value)
    }
}
```

### Hover Card / Tooltip

```rust
fn render_with_tooltip(
    content: impl IntoElement,
    tooltip: &str,
) -> impl IntoElement {
    div()
        .relative()
        .group("tooltip-trigger")
        .child(content)
        .child(
            div()
                .absolute()
                .bottom_full()
                .left_0()
                .mb_1()
                .px_2()
                .py_1()
                .bg(rgb(0x3c3c3c))
                .border_1()
                .border_color(rgb(0x555555))
                .rounded_sm()
                .text_sm()
                .invisible()
                .group_hover("tooltip-trigger", |s| s.visible())
                .child(tooltip)
        )
}
```

---

## Performance Tips

### 1. Minimize Children Rebuilds

```rust
// SLOW - rebuilds all children every render
div().children(
    (0..10000).map(|i| div().child(format!("Item {}", i)))
)

// BETTER - only render visible items
div().children(
    (scroll_offset..scroll_offset + visible_count).map(|i| {
        div().child(format!("Item {}", i))
    })
)
```

### 2. Cache Computed Values

```rust
// SLOW - recomputes in render
fn render(&mut self, ...) -> impl IntoElement {
    let total = self.items.iter().sum();  // Computed every render
    div().child(format!("Total: {}", total))
}

// BETTER - cache in state
pub struct MyApp {
    items: Vec<i32>,
    cached_total: i32,
}

impl MyApp {
    fn add_item(&mut self, item: i32, cx: &mut Context<Self>) {
        self.items.push(item);
        self.cached_total += item;  // Update cache
        cx.notify();
    }
}
```

### 3. Use Static Strings Where Possible

```rust
// SLOW - allocates every render
div().child(format!("Label"))

// BETTER - static string
div().child("Label")

// OK when dynamic data needed
div().child(format!("Count: {}", self.count))
```

### 4. Avoid Deep Nesting

Deeply nested element trees are slower to diff. Flatten where possible.

---

## Debugging Tips

### 1. Print State Changes

```rust
fn update_selection(&mut self, new_sel: (usize, usize), cx: &mut Context<Self>) {
    eprintln!("Selection: {:?} -> {:?}", self.selected, new_sel);
    self.selected = new_sel;
    cx.notify();
}
```

### 2. Verify Actions Fire

```rust
.on_action(cx.listener(|this, _: &MyAction, _, cx| {
    eprintln!("MyAction triggered!");
    this.handle_action(cx);
}))
```

### 3. Check Key Context

```rust
// Temporary: print context
div()
    .key_context("MyContext")
    .on_key_down(cx.listener(|_, event: &KeyDownEvent, _, _| {
        eprintln!("Key in MyContext: {:?}", event.keystroke.key);
    }))
```

### 4. Verify Focus

```rust
impl Render for MyApp {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let has_focus = self.focus_handle.is_focused(window);
        eprintln!("Has focus: {}", has_focus);
        // ...
    }
}
```
