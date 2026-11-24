//! Shared count picker. Confirmation delegates to the existing structural edits.
use crate::{
    app::{Spreadsheet, NUM_COLS, NUM_ROWS},
    mode::Mode,
    theme::TokenKey,
};
use gpui::*;
use visigrid_engine::structural::Axis;

pub(crate) struct CountDialog {
    pub title: &'static str,
    pub axis: Axis,
    pub at: usize,
    pub max: usize,
    pub input: String,
    pub anchor: Option<Point<Pixels>>,
    pub grow_only: bool,
}

impl CountDialog {
    pub fn count(&self) -> Option<usize> {
        self.input
            .trim()
            .parse::<usize>()
            .ok()
            .filter(|n| *n >= 1 && *n <= self.max)
    }
}

impl Spreadsheet {
    pub(crate) fn show_insert_count(&mut self, axis: Axis, after: bool, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) {
            return;
        }
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Insert not supported with multiple selections".into());
            cx.notify();
            return;
        }
        let ((r1, c1), (r2, c2)) = self.selection_range();
        let (at, max, title) = match axis {
            Axis::Row => {
                let at = self.view_to_data(if after { r2 } else { r1 }, cx) + usize::from(after);
                (
                    at,
                    NUM_ROWS.saturating_sub(at),
                    if after {
                        "Insert rows below"
                    } else {
                        "Insert rows above"
                    },
                )
            }
            Axis::Col => {
                let at = if after { c2 + 1 } else { c1 };
                (
                    at,
                    NUM_COLS.saturating_sub(at),
                    if after {
                        "Insert columns right"
                    } else {
                        "Insert columns left"
                    },
                )
            }
        };
        if max == 0 {
            self.status_message = Some("There is no remaining space to insert here.".into());
            cx.notify();
            return;
        }
        self.count_dialog = Some(CountDialog {
            title,
            axis,
            at,
            max,
            input: String::new(),
            anchor: None,
            grow_only: false,
        });
        self.mode = Mode::CountDialog;
        cx.notify();
    }

    pub(crate) fn close_count_dialog(&mut self, cx: &mut Context<Self>) {
        self.count_dialog = None;
        self.mode = Mode::Navigation;
        cx.notify();
    }

    pub(crate) fn confirm_count_dialog(&mut self, cx: &mut Context<Self>) {
        let Some(count) = self.count_dialog.as_ref().and_then(CountDialog::count) else {
            return;
        };
        let dialog = self.count_dialog.take().unwrap();
        self.mode = Mode::Navigation;
        if dialog.grow_only {
            match dialog.axis {
                Axis::Row => self.grow_grid(count, 0, cx),
                Axis::Col => self.grow_grid(0, count, cx),
            }
            return;
        }
        match dialog.axis {
            Axis::Row => self.insert_rows(dialog.at, count, cx),
            Axis::Col => self.insert_cols(dialog.at, count, cx),
        }
        cx.notify();
    }

    pub(crate) fn count_dialog_key(&mut self, event: &KeyDownEvent, cx: &mut Context<Self>) {
        match event.keystroke.key.as_str() {
            "enter" => self.confirm_count_dialog(cx),
            "escape" => self.close_count_dialog(cx),
            "backspace" => {
                self.count_dialog.as_mut().unwrap().input.pop();
            }
            _ => {
                if !event.keystroke.modifiers.modified() {
                    if let Some(text) = &event.keystroke.key_char {
                        self.count_dialog
                            .as_mut()
                            .unwrap()
                            .input
                            .extend(text.chars().filter(char::is_ascii_digit));
                    }
                }
            }
        }
        cx.notify();
    }
}

pub(crate) fn render(
    app: &Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let dialog = app.count_dialog.as_ref().unwrap();
    let valid = dialog.count().is_some();
    let cell = app.active_cell_rect();
    let (origin_x, origin_y) = app.grid_layout.grid_body_origin;
    let anchor = dialog.anchor.unwrap_or_else(|| match dialog.axis {
        Axis::Row => point(px(origin_x + 8.), px(origin_y + cell.y)),
        Axis::Col => point(px(origin_x + cell.x), px(origin_y)),
    });
    let viewport = window.viewport_size();
    let (x, y) = crate::ui::clamp_to_viewport(
        anchor.x.into(),
        anchor.y.into(),
        248.,
        112.,
        viewport.width.into(),
        viewport.height.into(),
    );
    let hint = if !dialog.input.is_empty() && !valid {
        format!("Enter a count from 1 to {}", dialog.max)
    } else if dialog.grow_only {
        "Enter to add · Esc to cancel".into()
    } else {
        "Enter to insert · Esc to cancel".into()
    };
    crate::ui::popup(
        "count-dialog",
        app.token(TokenKey::PanelBg),
        app.token(TokenKey::PanelBorder),
        |app, cx| app.close_count_dialog(cx),
        cx,
    )
    .debug_selector(|| "count-popover".into())
    .left(px(x))
    .top(px(y))
    .w(px(248.))
    .p_3()
    .gap_2()
    .text_size(px(12.))
    .text_color(app.token(TokenKey::TextPrimary))
    .child(
        div()
            .flex()
            .items_center()
            .justify_between()
            .child(dialog.title)
            .child(
                div()
                    .id("count-cancel")
                    .debug_selector(|| "count-cancel".into())
                    .px_1()
                    .cursor_pointer()
                    .text_color(app.token(TokenKey::TextMuted))
                    .child("×")
                    .on_mouse_down(
                        MouseButton::Left,
                        cx.listener(|app, _, _, cx| {
                            cx.stop_propagation();
                            app.close_count_dialog(cx);
                        }),
                    ),
            ),
    )
    .child(
        div()
            .flex()
            .items_center()
            .gap_2()
            .child(
                div()
                    .id("count-input")
                    .flex_1()
                    .min_w_0()
                    .overflow_hidden()
                    .h(px(28.))
                    .px_2()
                    .flex()
                    .items_center()
                    .border_1()
                    .rounded_sm()
                    .border_color(app.token(TokenKey::Accent))
                    .bg(app.token(TokenKey::EditorBg))
                    .child(format!("{}│", dialog.input)),
            )
            .child(
                div()
                    .id("count-confirm")
                    .flex_shrink_0()
                    .debug_selector(|| "count-confirm".into())
                    .h(px(28.))
                    .px_3()
                    .flex()
                    .items_center()
                    .rounded_sm()
                    .cursor_pointer()
                    .bg(app
                        .token(TokenKey::Accent)
                        .opacity(if valid { 0.15 } else { 0.05 }))
                    .text_color(app.token(if valid {
                        TokenKey::Accent
                    } else {
                        TokenKey::TextDisabled
                    }))
                    .child(if dialog.grow_only { "Add" } else { "Insert" })
                    .on_mouse_down(
                        MouseButton::Left,
                        cx.listener(|app, _, _, cx| {
                            cx.stop_propagation();
                            app.confirm_count_dialog(cx);
                        }),
                    ),
            ),
    )
    .child(
        div()
            .text_size(px(11.))
            .text_color(app.token(TokenKey::TextMuted))
            .child(hint),
    )
}
