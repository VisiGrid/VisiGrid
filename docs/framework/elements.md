# gpui Elements Reference

Elements are the building blocks of gpui UIs. They're constructed using a fluent builder API.

---

## Base Elements

### div()

The primary container element. Similar to HTML div.

```rust
div()
    .child("Hello")
    .child(other_element)
```

### text()

For text content. Usually just pass strings to `.child()`.

```rust
// These are equivalent for simple text:
div().child("Hello")
div().child(text("Hello"))
```

### svg()

For SVG icons and graphics.

```rust
svg()
    .path("icons/save.svg")
    .size_4()
    .text_color(rgb(0xffffff))
```

### img()

For images.

```rust
img("path/to/image.png")
    .w(px(100.0))
    .h(px(100.0))
```

### canvas()

For custom drawing.

```rust
canvas(
    |bounds, window, cx| {
        // Prepaint - runs before paint
    },
    |bounds, window, cx| {
        // Paint - do actual drawing
    }
)
```

---

## Layout Properties

### Flexbox

gpui uses flexbox for layout.

```rust
div()
    .flex()              // Enable flexbox (display: flex)
    .flex_row()          // Default direction
    .flex_col()          // Column direction
    .flex_wrap()         // Allow wrapping
    .flex_1()            // flex: 1 (grow to fill)
    .flex_shrink_0()     // Don't shrink
    .flex_grow()         // Allow growing
    .flex_none()         // Don't flex at all
```

### Alignment

```rust
div()
    .items_start()       // align-items: flex-start
    .items_center()      // align-items: center
    .items_end()         // align-items: flex-end
    .items_stretch()     // align-items: stretch

    .justify_start()     // justify-content: flex-start
    .justify_center()    // justify-content: center
    .justify_end()       // justify-content: flex-end
    .justify_between()   // justify-content: space-between
```

### Gap

```rust
div()
    .gap(px(8.0))        // Gap between flex items
    .gap_1()             // Predefined gap sizes
    .gap_2()
    .gap_4()
```

### Self Alignment

```rust
div()
    .self_start()        // align-self: flex-start
    .self_center()       // align-self: center
    .self_end()          // align-self: flex-end
```

---

## Sizing

### Width & Height

```rust
div()
    .w(px(100.0))        // width: 100px
    .h(px(50.0))         // height: 50px
    .w_full()            // width: 100%
    .h_full()            // height: 100%
    .size(px(100.0))     // Both width and height
    .size_full()         // 100% both
```

### Min/Max Constraints

```rust
div()
    .min_w(px(50.0))     // min-width
    .max_w(px(500.0))    // max-width
    .min_h(px(20.0))     // min-height
    .max_h(px(200.0))    // max-height
```

### Predefined Sizes

```rust
div()
    .size_4()            // 1rem (typically 16px)
    .size_6()            // 1.5rem
    .size_8()            // 2rem
    .w_48()              // 12rem
    .h_12()              // 3rem
```

---

## Spacing

### Padding

```rust
div()
    .p(px(8.0))          // All sides
    .p_2()               // Predefined (0.5rem)
    .px(px(16.0))        // Horizontal (left + right)
    .py(px(8.0))         // Vertical (top + bottom)
    .px_2()              // Predefined horizontal
    .py_1()              // Predefined vertical
    .pt(px(4.0))         // Top only
    .pb(px(4.0))         // Bottom only
    .pl(px(4.0))         // Left only
    .pr(px(4.0))         // Right only
```

### Margin

```rust
div()
    .m(px(8.0))          // All sides
    .m_2()               // Predefined
    .mx(px(16.0))        // Horizontal
    .my(px(8.0))         // Vertical
    .mx_auto()           // Center horizontally
    .mt(px(4.0))         // Top
    .mb(px(4.0))         // Bottom
    .ml(px(4.0))         // Left
    .mr(px(4.0))         // Right
```

---

## Visual Properties

### Colors

```rust
div()
    .bg(rgb(0x1e1e1e))           // Background color
    .text_color(rgb(0xffffff))    // Text color
```

**Color formats:**
```rust
rgb(0x1e1e1e)              // RGB hex
rgba(0x1e1e1e80)           // RGBA (80 = ~50% alpha)
hsla(0.0, 0.0, 0.12, 1.0)  // HSLA
gpui::white()              // Predefined
gpui::black()
gpui::transparent_black()
```

### Borders

```rust
div()
    .border_1()                  // 1px border all sides
    .border_2()                  // 2px border
    .border_color(rgb(0x3d3d3d))
    .border_t_1()                // Top only
    .border_b_1()                // Bottom only
    .border_l_1()                // Left only
    .border_r_1()                // Right only
```

### Border Radius

```rust
div()
    .rounded(px(4.0))      // Custom radius
    .rounded_sm()          // 2px
    .rounded_md()          // 6px
    .rounded_lg()          // 8px
    .rounded_xl()          // 12px
    .rounded_full()        // Fully rounded (circle if square)
```

### Overflow

```rust
div()
    .overflow_hidden()     // Hide overflow
    .overflow_scroll()     // Show scrollbars
    .overflow_x_hidden()
    .overflow_y_scroll()
```

### Visibility

```rust
div()
    .visible()
    .invisible()           // Takes space but not visible
```

### Opacity

```rust
div()
    .opacity(0.5)          // 50% opacity
```

---

## Typography

### Font Size

```rust
div()
    .text_xs()             // Extra small
    .text_sm()             // Small
    .text_base()           // Base (1rem)
    .text_lg()             // Large
    .text_xl()             // Extra large
    .text_2xl()            // 2x large
```

### Font Weight

```rust
div()
    .font_weight(FontWeight::NORMAL)
    .font_weight(FontWeight::MEDIUM)
    .font_weight(FontWeight::SEMIBOLD)
    .font_weight(FontWeight::BOLD)
```

### Font Style

```rust
div()
    .italic()              // Italic text
    .underline()           // Underlined text
    // Note: There's no .font_style() method
```

### Text Alignment

```rust
div()
    .text_left()           // left align
    .text_center()         // center align
    .text_right()          // right align
```

### Line Height

```rust
div()
    .line_height(px(24.0))
```

---

## Positioning

### Position Type

```rust
div()
    .relative()            // position: relative (default for layout)
    .absolute()            // position: absolute
```

### Absolute Positioning

```rust
div()
    .absolute()
    .top(px(10.0))
    .left(px(10.0))
    .right(px(10.0))
    .bottom(px(10.0))
    .inset_0()             // All sides 0
```

### Z-Index

```rust
div()
    .z_index(10)
```

---

## Interactivity

### Cursor

```rust
div()
    .cursor_pointer()
    .cursor_default()
    .cursor_text()
    .cursor_not_allowed()
```

### Hover State

```rust
div()
    .hover(|style| {
        style
            .bg(rgb(0x404040))
            .text_color(rgb(0xffffff))
    })
```

### Active State (Pressed)

```rust
div()
    .active(|style| {
        style.bg(rgb(0x094771))
    })
```

### Focus State

```rust
div()
    .focus(|style| {
        style.border_color(rgb(0x007acc))
    })
```

---

## Children

### Single Child

```rust
div().child("text")
div().child(other_element)
```

### Multiple Children

```rust
div()
    .child(element1)
    .child(element2)
    .child(element3)
```

### Dynamic Children

```rust
div().children(
    items.iter().map(|item| {
        div().child(item.name.clone())
    })
)
```

### Conditional Children

```rust
use gpui::prelude::FluentBuilder;

div()
    .when(condition, |div| {
        div.child(conditional_element)
    })
    .when_some(optional_value, |div, value| {
        div.child(format!("Value: {}", value))
    })
```

---

## Element IDs

Required for mouse events and some interactive features.

```rust
div()
    .id("my-unique-id")
    .id(ElementId::Name("my-id".into()))
    .id(ElementId::NamedInteger("item".into(), index))
```

**Note:** Adding `.id()` changes the return type from `Div` to `Stateful<Div>`.

---

## Common Patterns

### Clickable Button

```rust
div()
    .id("button")
    .px_4()
    .py_2()
    .rounded_md()
    .bg(rgb(0x007acc))
    .text_color(rgb(0xffffff))
    .cursor_pointer()
    .hover(|s| s.bg(rgb(0x005a9e)))
    .active(|s| s.bg(rgb(0x004578)))
    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
        this.on_click(cx);
    }))
    .child("Click Me")
```

### Input Field

```rust
div()
    .id("input")
    .w_full()
    .px_2()
    .py_1()
    .bg(rgb(0x3c3c3c))
    .border_1()
    .border_color(rgb(0x007acc))
    .text_color(rgb(0xffffff))
    .child(&self.input_value)
```

### Scrollable List

```rust
div()
    .flex_1()
    .overflow_y_scroll()
    .children(
        items.iter().map(|item| render_item(item))
    )
```

### Centered Content

```rust
div()
    .size_full()
    .flex()
    .items_center()
    .justify_center()
    .child(content)
```

### Header/Content/Footer Layout

```rust
div()
    .flex()
    .flex_col()
    .size_full()
    .child(header)       // Fixed height
    .child(
        div()
            .flex_1()    // Takes remaining space
            .child(content)
    )
    .child(footer)       // Fixed height
```
