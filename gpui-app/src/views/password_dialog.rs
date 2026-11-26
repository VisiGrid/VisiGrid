use std::path::PathBuf;

use crate::{app::Spreadsheet, mode::Mode, theme::TokenKey, ui::modal_overlay};
use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_io::xlsx::ImportOptions;
use visigrid_io::xlsx_crypto::Zeroizing;

/// In-memory state for opening a password-protected Excel workbook.
/// Cancelling drops and wipes the input. Confirming moves it to the import task.
pub struct PasswordPromptState {
    pub path: PathBuf,
    pub filename: String,
    pub options: ImportOptions,
    pub restore_sheet: Option<usize>,
    pub input: Zeroizing<String>,
    pub error: Option<String>,
    pub show_password: bool,
}

impl PasswordPromptState {
    pub fn new(
        path: PathBuf,
        filename: String,
        mut options: ImportOptions,
        restore_sheet: Option<usize>,
        error: Option<String>,
    ) -> Self {
        options.password = None;
        Self {
            path,
            filename,
            options,
            restore_sheet,
            input: Zeroizing::new(String::new()),
            error,
            show_password: false,
        }
    }
}

impl Spreadsheet {
    pub(crate) fn show_password_prompt(
        &mut self,
        path: PathBuf,
        filename: String,
        options: ImportOptions,
        restore_sheet: Option<usize>,
        error: Option<String>,
        cx: &mut Context<Self>,
    ) {
        self.password_prompt = Some(PasswordPromptState::new(
            path,
            filename,
            options,
            restore_sheet,
            error,
        ));
        self.mode = Mode::PasswordPrompt;
        cx.notify();
    }

    pub(crate) fn cancel_password_prompt(&mut self, cx: &mut Context<Self>) {
        self.password_prompt = None;
        if self.mode == Mode::PasswordPrompt {
            self.mode = Mode::Navigation;
        }
        self.status_message = Some("Cancelled opening password-protected workbook".to_string());
        cx.notify();
    }

    pub(crate) fn confirm_password_prompt(&mut self, cx: &mut Context<Self>) {
        let Some(state) = self.password_prompt.take() else {
            return;
        };
        if self.mode == Mode::PasswordPrompt {
            self.mode = Mode::Navigation;
        }
        let mut options = state.options;
        options.password = Some(state.input);
        if let Some(restore_sheet) = state.restore_sheet {
            self.start_excel_import_with_options(&state.path, options, restore_sheet, cx);
        } else {
            self.start_excel_import_with_password(&state.path, options, cx);
        }
    }

    pub(crate) fn password_prompt_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if let Some(state) = &mut self.password_prompt {
            state.input.push(c);
            state.error = None;
            cx.notify();
        }
    }

    pub(crate) fn password_prompt_backspace(&mut self, cx: &mut Context<Self>) {
        if let Some(state) = &mut self.password_prompt {
            state.input.pop();
            state.error = None;
            cx.notify();
        }
    }

    pub(crate) fn password_prompt_toggle_visibility(&mut self, cx: &mut Context<Self>) {
        if let Some(state) = &mut self.password_prompt {
            state.show_password = !state.show_password;
            cx.notify();
        }
    }
}

pub fn render_password_dialog(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let error_color = app.token(TokenKey::Error);

    let (filename, display_value, empty, error, show_password) = match &app.password_prompt {
        Some(state) => {
            let empty = state.input.is_empty();
            let display = if state.show_password {
                state.input.as_str().to_string()
            } else {
                "\u{2022}".repeat(state.input.chars().count())
            };
            (
                state.filename.clone(),
                display,
                empty,
                state.error.clone(),
                state.show_password,
            )
        }
        None => (String::new(), String::new(), true, None, false),
    };

    let has_error = error.is_some();

    let content = div()
        .w(px(460.0))
        .p_6()
        .bg(app.token(TokenKey::PanelBg))
        .border_1()
        .border_color(border)
        .rounded_md()
        .shadow_lg()
        .flex()
        .flex_col()
        .gap_5()
        .child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .child(
                    div()
                        .text_size(px(18.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text)
                        .child("Password required"),
                )
                .child(
                    div()
                        .text_size(px(13.0))
                        .text_color(muted)
                        .child(format!(
                            "Enter the password to open {}.",
                            if filename.is_empty() { "this workbook" } else { &filename }
                        )),
                ),
        )
        .child(
            div()
                .flex()
                .flex_col()
                .gap_2()
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(muted)
                                .child("Password"),
                        )
                        .child(
                            div()
                                .id("password-toggle-visibility")
                                .px_2()
                                .py_1()
                                .rounded_sm()
                                .bg(app.token(TokenKey::EditorBg))
                                .text_size(px(11.0))
                                .text_color(muted)
                                .cursor_pointer()
                                .on_mouse_down(
                                    MouseButton::Left,
                                    cx.listener(|s, _, _, cx| {
                                        cx.stop_propagation();
                                        s.password_prompt_toggle_visibility(cx);
                                    }),
                                )
                                .child(if show_password { "Hide" } else { "Show" }),
                        ),
                )
                .child(
                    div()
                        .id("password-prompt-input")
                        .h(px(44.0))
                        .px_3()
                        .bg(app.token(TokenKey::EditorBg))
                        .border_1()
                        .rounded_sm()
                        .border_color(if has_error { error_color } else { accent })
                        .flex()
                        .items_center()
                        .overflow_hidden()
                        .cursor_text()
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(|s, _, window, cx| {
                                cx.stop_propagation();
                                window.focus(&s.focus_handle, cx);
                            }),
                        )
                        .child(
                            div()
                                .min_w_0()
                                .overflow_hidden()
                                .text_ellipsis()
                                .text_size(px(16.0))
                                .text_color(text)
                                .child(display_value),
                        )
                        .child(div().w(px(1.0)).h(px(18.0)).flex_shrink_0().bg(accent))
                        .when(empty, |s| {
                            s.child(
                                div()
                                    .ml_1()
                                    .text_size(px(14.0))
                                    .text_color(muted)
                                    .child("Enter password"),
                            )
                        }),
                )
                .child(
                    div()
                        .text_size(px(12.0))
                        .text_color(if has_error { error_color } else { muted })
                        .child(error.unwrap_or_else(|| {
                            "Decrypted in memory only. Password is never saved to disk.".to_string()
                        })),
                ),
        )
        .child(
            div()
                .flex()
                .items_center()
                .justify_end()
                .gap_2()
                .child(
                    div()
                        .id("password-prompt-cancel")
                        .px_4()
                        .py_2()
                        .rounded_sm()
                        .border_1()
                        .border_color(border)
                        .text_size(px(13.0))
                        .text_color(text)
                        .cursor_pointer()
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(|s, _, _, cx| {
                                cx.stop_propagation();
                                s.cancel_password_prompt(cx);
                            }),
                        )
                        .child("Cancel"),
                )
                .child(
                    div()
                        .id("password-prompt-confirm")
                        .px_4()
                        .py_2()
                        .rounded_sm()
                        .bg(accent)
                        .text_size(px(13.0))
                        .font_weight(FontWeight::MEDIUM)
                        .text_color(rgb(0xffffff))
                        .cursor_pointer()
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(|s, _, _, cx| {
                                cx.stop_propagation();
                                s.confirm_password_prompt(cx);
                            }),
                        )
                        .child("Open"),
                ),
        );

    modal_overlay("password-prompt-dialog", |s, cx| s.cancel_password_prompt(cx), content, cx)
}

pub(crate) fn render_export_prompt(
    app: &Spreadsheet,
    focused: usize,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    use crate::ui::{Button, DialogFrame, DialogSize, dialog_header_simple};
    let text = app.token(TokenKey::TextPrimary);
    let border = app.token(TokenKey::PanelBorder);
    let choices = [
        ("keep-password", "Keep Password"),
        ("export-plain", "Export Unprotected"),
        ("cancel-export", "Cancel"),
    ];
    let footer = div().flex().justify_end().gap_2().children(
        choices.into_iter().enumerate().map(|(index, (id, label))| {
            let button = Button::new(id, label);
            let button = if index == focused {
                button.primary(app.token(TokenKey::Accent), app.token(TokenKey::TextInverse))
            } else {
                button.secondary(border, text)
            };
            button.on_mouse_down(MouseButton::Left, cx.listener(move |app, _, _, cx| {
                cx.stop_propagation();
                app.finish_excel_export_prompt(index, cx);
            }))
        }),
    );
    let body = div().text_sm().text_color(text).child(
        "This workbook was opened from a password-protected Excel file. Keeping the password protects the exported file with the same password.",
    );
    modal_overlay("excel-export-password", |app, cx| app.finish_excel_export_prompt(2, cx),
        DialogFrame::new(body, app.token(TokenKey::PanelBg), border)
            .size(DialogSize::Lg)
            .header(dialog_header_simple("Keep the Excel password?", text))
            .footer(footer),
        cx,
    )
}
