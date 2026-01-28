//! Modal overlay components for centered dialogs with backdrop.
//!
//! # Architecture
//!
//! Backdrop and dialog are SIBLINGS, not parent-child. This is critical:
//! - Backdrop: full-screen, catches clicks (dismiss) and scroll (consume)
//! - Dialog: sits on top, receives all pointer events normally
//!
//! If the dialog were a CHILD of the backdrop with stop_propagation, the
//! backdrop's mouse handler would eat press events before children see them,
//! causing double-click-to-select bugs.
//!
//! ## `modal_overlay` (dismissable)
//! - Click on backdrop dismisses via `on_dismiss` callback
//! - Escape handling is the caller's responsibility (add `.on_key_down()` to content)
//!
//! ## `modal_backdrop` (non-dismissable)
//! - No click-outside dismiss (for tours, confirmations, etc.)
//! - Caller handles all dismissal (buttons, escape key)
//!
//! # What does NOT belong here
//!
//! - **Positioned/floating panels** (find dialog, preferences) → use `ui/popup.rs` later
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
/// Click on the backdrop (outside dialog) calls `on_dismiss`.
/// The dialog content receives all pointer events normally (no interception).
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

    // Outer container: layout only, no event handlers
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        // Backdrop layer: full-screen, behind dialog, consumes events
        .child(
            div()
                .id(ElementId::Name(format!("{}-backdrop", id).into()))
                .absolute()
                .inset_0()
                .bg(hsla(0.0, 0.0, 0.0, 0.5))
                // Click backdrop = dismiss
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    dismiss(this, cx);
                }))
                // Consume scroll so grid doesn't move
                .on_scroll_wheel(|_, _, cx| {
                    cx.stop_propagation();
                })
        )
        // Dialog layer: sits on top, receives events normally
        .child(
            div()
                .id(ElementId::Name(id))
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

    // Outer container: layout only
    div()
        .absolute()
        .inset_0()
        .flex()
        .items_center()
        .justify_center()
        // Backdrop layer: event sink
        .child(
            div()
                .id(ElementId::Name(format!("{}-backdrop", id).into()))
                .absolute()
                .inset_0()
                .bg(hsla(0.0, 0.0, 0.0, 0.5))
                // Consume mouse + scroll events
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                .on_scroll_wheel(|_, _, cx| {
                    cx.stop_propagation();
                })
        )
        // Dialog layer: sits on top, receives events normally
        .child(
            div()
                .id(ElementId::Name(id))
                .child(content)
        )
}
