//! Modal overlay components for centered dialogs with backdrop.
//!
//! # Contract
//!
//! These helpers provide:
//! - Full-screen semi-transparent backdrop (50% black)
//! - Centered content
//! - Stop-propagation on content area (clicks inside don't dismiss)
//! - Unique element ID assignment
//!
//! ## `modal_overlay` (dismissable)
//! - Click outside dismisses via `on_dismiss` callback
//! - Escape handling is the caller's responsibility (add `.on_key_down()` to content)
//!
//! ## `modal_backdrop` (non-dismissable)
//! - No click-outside dismiss (for tours, confirmations, etc.)
//! - Caller handles all dismissal (buttons, escape key)
//!
//! # What does NOT belong here
//!
//! - **Positioned/floating panels** (find dialog, preferences) → use `ui/popup.rs` later
//! - **Panels with different opacity** (preferences uses 0.4) → panel primitive later
//! - **Escape key handling** → caller's `.on_key_down()` on content div
//! - **Focus management** → caller's responsibility
//!
//! # ID Convention
//!
//! Use `{feature}-dialog` format: `"export-report-dialog"`, `"license-dialog"`, etc.

use gpui::*;
use crate::app::Spreadsheet;

/// Creates a dismissable modal overlay.
///
/// Click outside the content area calls `on_dismiss`. Escape handling is the caller's
/// responsibility - add `.on_key_down()` to the content div if needed.
pub fn modal_overlay<F, E>(
    id: impl Into<SharedString>,
    on_dismiss: F,
    content: E,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement
where
    F: Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static + Clone,
    E: IntoElement,
{
    let dismiss = on_dismiss.clone();
    let id: SharedString = id.into();

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            dismiss(this, cx);
        }))
        .child(
            div()
                .id(ElementId::Name(id))
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                .child(content)
        )
}

/// Creates a non-dismissable modal overlay.
///
/// No click-outside dismiss. Use for guided tours, confirmations, or dialogs where
/// accidental dismissal would be disruptive. Caller handles all dismissal via buttons
/// or keyboard handlers.
pub fn modal_backdrop<E: IntoElement>(
    id: impl Into<SharedString>,
    content: E,
) -> impl IntoElement {
    let id: SharedString = id.into();

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        .bg(hsla(0.0, 0.0, 0.0, 0.5))
        .child(
            div()
                .id(ElementId::Name(id))
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                .child(content)
        )
}
