//! Floating popup containers for context menus and positioned panels.
//!
//! # Usage
//!
//! `popup()` returns a styled `Stateful<Div>` that the caller positions and populates:
//!
//! ```ignore
//! popup("my-menu", panel_bg, panel_border, |this, cx| this.dismiss(cx), cx)
//!     .left(px(x))
//!     .top(px(y))
//!     .w(px(200.0))
//!     .children(items)
//! ```
//!
//! # What this provides
//!
//! - `.id()` (required for `on_mouse_down_out`)
//! - `.absolute()` positioning base
//! - Standard styling: bg, border, rounded_md, shadow_lg
//! - Flex column layout with vertical padding
//! - Click-outside dismiss via `on_mouse_down_out`
//!
//! # What the caller provides
//!
//! - Position (`.left()`, `.top()`, `.right()`, `.bottom()`)
//! - Width (`.w()`)
//! - Children (`.child()`, `.children()`)
//! - Viewport clamping (use `clamp_to_viewport` helper if needed)
//!
//! # What does NOT belong here
//!
//! - **Modal dialogs** → use `ui/modal.rs`
//! - **Menu item styling** → caller's responsibility
//! - **Anchor positioning / flip logic / arrow tips** → not yet needed

use gpui::*;
use crate::app::Spreadsheet;

/// Creates a popup container with standard styling and click-outside dismiss.
///
/// Returns a positioned `Stateful<Div>` — caller chains `.left()/.top()/.w()/.children()`.
pub fn popup<F>(
    id: impl Into<SharedString>,
    bg: Hsla,
    border: Hsla,
    on_dismiss: F,
    cx: &mut Context<Spreadsheet>,
) -> Stateful<Div>
where
    F: Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
{
    div()
        .id(ElementId::Name(id.into()))
        .absolute()
        .bg(bg)
        .border_1()
        .border_color(border)
        .rounded_md()
        .shadow_lg()
        .flex()
        .flex_col()
        .py_1()
        .on_mouse_down_out(cx.listener(move |this, _, _, cx| {
            on_dismiss(this, cx);
        }))
}

/// Clamp a popup position to keep it within window bounds.
///
/// Returns `(x, y)` clamped so the popup doesn't overflow the viewport.
pub fn clamp_to_viewport(
    x: f32,
    y: f32,
    popup_w: f32,
    popup_h: f32,
    viewport_w: f32,
    viewport_h: f32,
) -> (f32, f32) {
    let x = x.min(viewport_w - popup_w).max(0.0);
    let y = y.min(viewport_h - popup_h).max(0.0);
    (x, y)
}
