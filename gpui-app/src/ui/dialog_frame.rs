//! Dialog frame component - container with header/body/footer slots.
//!
//! # What this handles
//! - Container styling (bg, border, rounded, shadow)
//! - Width presets
//! - Header row (optional, with border)
//! - Body with standard padding
//! - Footer row (optional, with border)
//!
//! # What this does NOT handle
//! - Click-outside dismiss (use `modal_overlay`)
//! - Keyboard handling (caller adds `.on_key_down()` to modal content)
//! - Button styling (coming in `ui/button.rs`)

use gpui::*;

/// Width presets matching common dialog sizes
#[derive(Clone, Copy)]
pub enum DialogSize {
    /// 300px - Go To, small prompts
    Sm,
    /// 380-400px - About, simple dialogs
    Md,
    /// 480-500px - License, reports
    Lg,
    /// 520px+ - Complex dialogs
    Xl,
}

impl DialogSize {
    pub fn width(self) -> Pixels {
        match self {
            DialogSize::Sm => px(300.0),
            DialogSize::Md => px(380.0),
            DialogSize::Lg => px(500.0),
            DialogSize::Xl => px(520.0),
        }
    }
}

/// Dialog frame builder.
///
/// Wraps content in a standard dialog container with optional header and footer.
/// Does not include backdrop - wrap with `modal_overlay` or `modal_backdrop`.
pub struct DialogFrame<B: IntoElement> {
    body: B,
    header: Option<AnyElement>,
    footer: Option<AnyElement>,
    width: Pixels,
    max_height: Option<Pixels>,
    panel_bg: Hsla,
    panel_border: Hsla,
}

impl<B: IntoElement> DialogFrame<B> {
    /// Create a new dialog frame with body content.
    pub fn new(body: B, panel_bg: Hsla, panel_border: Hsla) -> Self {
        Self {
            body,
            header: None,
            footer: None,
            width: DialogSize::Md.width(),
            max_height: None,
            panel_bg,
            panel_border,
        }
    }

    /// Set dialog width using a preset.
    pub fn size(mut self, size: DialogSize) -> Self {
        self.width = size.width();
        self
    }

    /// Set custom width in pixels.
    pub fn width(mut self, width: Pixels) -> Self {
        self.width = width;
        self
    }

    /// Set maximum height (enables scroll in body).
    pub fn max_height(mut self, max_height: Pixels) -> Self {
        self.max_height = Some(max_height);
        self
    }

    /// Add a header element (rendered with bottom border).
    pub fn header(mut self, header: impl IntoElement) -> Self {
        self.header = Some(header.into_any_element());
        self
    }

    /// Add a footer element (rendered with top border).
    pub fn footer(mut self, footer: impl IntoElement) -> Self {
        self.footer = Some(footer.into_any_element());
        self
    }
}

impl<B: IntoElement> IntoElement for DialogFrame<B> {
    type Element = <Div as IntoElement>::Element;

    fn into_element(self) -> Self::Element {
        let mut container = div()
            .w(self.width)
            .bg(self.panel_bg)
            .border_1()
            .border_color(self.panel_border)
            .rounded_lg()
            .shadow_xl()
            .overflow_hidden()
            .flex()
            .flex_col();

        if let Some(max_h) = self.max_height {
            container = container.max_h(max_h);
        }

        // Header (optional)
        if let Some(header) = self.header {
            container = container.child(
                div()
                    .px_4()
                    .py_3()
                    .border_b_1()
                    .border_color(self.panel_border)
                    .child(header)
            );
        }

        // Body
        let mut body_container = div()
            .p_4()
            .flex()
            .flex_col()
            .gap_4();

        // If max_height is set, make body scrollable
        if self.max_height.is_some() {
            body_container = body_container.flex_1().overflow_hidden();
        }

        container = container.child(body_container.child(self.body));

        // Footer (optional)
        if let Some(footer) = self.footer {
            container = container.child(
                div()
                    .w_full()
                    .px_4()
                    .py_3()
                    .border_t_1()
                    .border_color(self.panel_border)
                    .child(footer)
            );
        }

        container.into_element()
    }
}

// Helper functions for common header patterns

/// Simple header with title only.
pub fn dialog_header_simple(title: impl Into<SharedString>, text_primary: Hsla) -> impl IntoElement {
    div()
        .flex()
        .items_center()
        .child(
            div()
                .text_size(px(14.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_primary)
                .child(title.into())
        )
}

/// Header with title and subtitle (e.g., filename).
pub fn dialog_header_with_subtitle(
    title: impl Into<SharedString>,
    subtitle: impl Into<SharedString>,
    text_primary: Hsla,
    text_muted: Hsla,
) -> impl IntoElement {
    div()
        .flex()
        .items_center()
        .justify_between()
        .child(
            div()
                .text_size(px(14.0))
                .font_weight(FontWeight::SEMIBOLD)
                .text_color(text_primary)
                .child(title.into())
        )
        .child(
            div()
                .text_size(px(12.0))
                .text_color(text_muted)
                .child(subtitle.into())
        )
}
