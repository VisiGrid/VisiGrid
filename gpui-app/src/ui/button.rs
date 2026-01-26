//! Button components for dialogs and panels.
//!
//! Provides consistent button styling across the app. Does NOT handle click events -
//! caller adds `.on_mouse_down()` (and uses `.when()` for disabled state).
//!
//! ## API Freeze
//!
//! Primary + Secondary only. No new variants unless 3+ call sites would migrate.
//! Tertiary/custom styles stay local until pattern repeats 3+ times.
//!
//! # Usage
//!
//! ```rust
//! Button::new("close-btn", "Close")
//!     .primary(accent, text_inverse)
//!     .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
//!         this.close(cx);
//!     }))
//!
//! Button::new("cancel-btn", "Cancel")
//!     .secondary(panel_border, text_muted)
//!     .on_mouse_down(...)
//!
//! // Disabled button - caller handles conditional click
//! let is_empty = input.is_empty();
//! Button::new("submit-btn", "Submit")
//!     .disabled(is_empty)
//!     .primary(accent, text_inverse)
//!     .when(!is_empty, |b| b.on_mouse_down(...))
//! ```

use gpui::*;

/// Button builder with consistent styling.
pub struct Button {
    id: ElementId,
    label: SharedString,
    disabled: bool,
}

impl Button {
    /// Create a new button.
    pub fn new(id: impl Into<ElementId>, label: impl Into<SharedString>) -> Self {
        Self {
            id: id.into(),
            label: label.into(),
            disabled: false,
        }
    }

    /// Mark button as disabled (affects styling only; caller handles click gating).
    pub fn disabled(mut self, disabled: bool) -> Self {
        self.disabled = disabled;
        self
    }

    /// Render as primary button (accent background, prominent).
    pub fn primary(self, accent: Hsla, text_color: Hsla) -> Stateful<Div> {
        let bg = if self.disabled {
            accent.opacity(0.5)
        } else {
            accent
        };

        let mut btn = div()
            .id(self.id)
            .px_4()
            .py(px(6.0))
            .bg(bg)
            .rounded_md()
            .text_size(px(12.0))
            .font_weight(FontWeight::MEDIUM)
            .text_color(text_color)
            .child(self.label);

        if !self.disabled {
            btn = btn
                .cursor_pointer()
                .hover(|s| s.opacity(0.9));
        }

        btn
    }

    /// Render as secondary button (bordered, subdued).
    pub fn secondary(self, border_color: Hsla, text_color: Hsla) -> Stateful<Div> {
        let mut btn = div()
            .id(self.id)
            .px_4()
            .py(px(6.0))
            .border_1()
            .border_color(border_color)
            .rounded_md()
            .text_size(px(12.0))
            .text_color(text_color)
            .child(self.label);

        if !self.disabled {
            btn = btn
                .cursor_pointer()
                .hover(move |s| s.bg(border_color.opacity(0.3)));
        }

        btn
    }
}
