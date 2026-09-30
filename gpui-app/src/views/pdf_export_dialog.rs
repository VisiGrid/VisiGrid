use crate::ui::{dialog_header_with_subtitle, modal_overlay, Button, DialogFrame};
use crate::{app::Spreadsheet, theme::TokenKey};
use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_engine::print_setup::{PrintPaper as Paper, PrintScale as Scale};

pub fn render(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some(state) = &app.pdf_export else {
        return div().into_any_element();
    };
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let warn = app.token(TokenKey::Warn);
    let panel = app.token(TokenKey::PanelBg);
    let width = (f32::from(app.window_size.width) - 48.0).clamp(480.0, 1280.0);
    let height = (f32::from(app.window_size.height) - 48.0).clamp(340.0, 960.0);
    let body_height = height - 136.0;
    let compact = width < 800.0;
    let preview_width = if compact { width - 32.0 } else { width - 332.0 };
    let preview_height = if compact {
        (body_height - 196.0).max(100.0)
    } else {
        body_height
    };
    let canvas_height = (preview_height - 66.0).max(60.0);
    let labels = [
        (
            "Scope",
            if state.use_print_area {
                "Print area"
            } else if state.selection_only {
                "Selected range"
            } else {
                "Active sheet"
            },
        ),
        (
            "Paper",
            match state.setup.paper {
                Paper::A4 => "A4",
                Paper::Letter => "Letter",
                Paper::Legal => "Legal",
            },
        ),
        (
            "Orientation",
            if state.setup.landscape {
                "Landscape"
            } else {
                "Portrait"
            },
        ),
        (
            "Scaling",
            match state.setup.scale {
                Scale::FitColumns => "Fit columns",
                Scale::FitSheet => "Fit sheet",
                _ => "Actual size (100%)",
            },
        ),
        (
            "Page numbers",
            if state.setup.page_numbers {
                "On"
            } else {
                "Off"
            },
        ),
        (
            "Print gridlines",
            if state.setup.gridlines { "On" } else { "Off" },
        ),
    ];
    let stale = state.captured_revision != app.workbook.read(cx).revision();
    let settings = div().id("pdf-settings-body").flex().flex_col().gap_3()
        .min_h_0().overflow_y_scroll().track_scroll(&state.settings_scroll).flex_shrink_0()
        .when(!compact, |d| d.w(px(284.0)).h_full())
        .when(compact, |d| d.w_full().h(px(180.0)))
        .child(div().text_sm().font_weight(FontWeight::SEMIBOLD).text_color(text).child("Page setup"))
        .children(labels.into_iter().enumerate().map(|(index, (label, value))| {
            div().flex().items_center().justify_between().gap_2()
                .child(div().text_sm().text_color(text).child(label))
                .child(Button::new(ElementId::Name(format!("pdf-option-{index}").into()), format!("{value} ›"))
                    .disabled(state.busy).secondary(if state.focus == index { accent } else { border }, text)
                    .w(px(154.0)).px_2().text_center()
                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                        this.pdf_cycle_option(index, cx); cx.stop_propagation();
                    })))
        }))
        .child(div().mt_2().flex().flex_col().gap_2()
            .child(div().text_sm().text_color(text).child("Repeat top rows"))
            .child(div().text_xs().text_color(muted).child(state.setup.repeat_rows.map_or("None".into(), |r| format!("Rows {}:{}", r.start + 1, r.end + 1))))
            .child(div().flex().gap_2().children([(6, "+"), (7, "−"), (8, "Clear")].into_iter().map(|(index,label)| {
                Button::new(ElementId::Name(format!("pdf-setup-{index}").into()), label)
                    .disabled(state.busy).secondary(if state.focus == index { accent } else { border }, text)
                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| { this.pdf_cycle_option(index, cx); cx.stop_propagation(); }))
            })))
            .child(div().text_xs().text_color(muted).child("+ adds the next visible row. A header cannot split a merged cell.")))
        .child(div().flex().flex_col().gap_2()
            .child(div().text_sm().text_color(text).child("Print area"))
            .child(div().text_xs().text_color(muted).child(state.setup.area.map_or("Automatic bounds".into(), |a| format!("{}{}:{}{}", Spreadsheet::col_to_letter(a.start_col), a.start_row + 1, Spreadsheet::col_to_letter(a.end_col), a.end_row + 1))))
            .children([(9, "Use selection"), (10, "Clear print area")].into_iter().map(|(index,label)| {
                Button::new(ElementId::Name(format!("pdf-setup-{index}").into()), label)
                    .disabled(state.busy).secondary(if state.focus == index { accent } else { border }, text)
                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| { this.pdf_cycle_option(index, cx); cx.stop_propagation(); }))
            })))
        .child(Button::new("pdf-save-setup", "Save setup to sheet")
            .disabled(state.busy || state.summary.is_err() || state.setup == state.saved_setup)
            .secondary(if state.focus == 11 { accent } else { border }, text)
            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| { this.pdf_cycle_option(11, cx); cx.stop_propagation(); })))
        .child(div().text_xs().text_color(muted).child(if state.setup == state.saved_setup {
            "Setup matches this sheet. Save the workbook to keep it on disk."
        } else { "Preview changes are temporary until you save setup to the sheet." }))
        .child(div().mt_2().p_3().border_1().border_color(border).rounded_md().flex().flex_col().gap_2()
            .children(state.summary.clone().unwrap_or_else(|e| e).lines().enumerate().map(|(i, line)| {
                div().text_sm().text_color(if state.summary.is_err() || i > 0 { warn } else { text }).child(line.to_string())
            })))
        .children(state.notices.iter().map(|notice| div().text_sm().text_color(warn).child(notice.clone())))
        .when(state.error.is_some(), |d| d.child(div().text_sm().text_color(warn).child(state.error.clone().unwrap())))
        .when(state.report.is_some(), |d| d.child(div().p_3().border_1().border_color(border).rounded_md().flex().flex_col().gap_2()
            .children(state.report.iter().flatten().map(|line| div().text_sm().text_color(text).child(line.clone())))
            .when(state.saved_path.is_some(), |d| d.child(div().text_xs().text_color(muted).child(state.saved_path.as_ref().unwrap().display().to_string())))))
        .child(div().text_xs().text_color(muted).child("Half-inch margins. Uses visible rows and columns. Select a range to include blank outer cells."))
        .child(div().text_xs().text_color(if stale { warn } else { muted }).child(if stale {
            "Sheet changed. Refresh to include edits; export keeps the shown snapshot."
        } else { "Preview uses a snapshot of the sheet's current calculated values." }))
        .child(Button::new("pdf-refresh", "Refresh preview").disabled(state.busy).secondary(border, text)
            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| { this.refresh_pdf_preview(cx); cx.stop_propagation(); })))
        .child(div().text_xs().text_color(muted).child(if state.print_mode { "Tab: setting · Space: change · PgUp/PgDn: page · Enter/Ctrl+P: print · Esc: close" } else if cfg!(target_os = "linux") { "Tab: setting · Space: change · PgUp/PgDn: page · Enter: export · Ctrl+P: print · Esc: close" } else { "Tab: setting · Space: change · PgUp/PgDn: page · Enter: export · Esc: close" }));

    let (paper_width, paper_height) = (
        state.page_size.0 as f32 * 4.0 / 3.0,
        state.page_size.1 as f32 * 4.0 / 3.0,
    );
    let fit = ((preview_width - 40.0) / paper_width)
        .min((canvas_height - 40.0) / paper_height)
        .max(0.05);
    let zoom = state.zoom.unwrap_or(fit);
    let (page_width, page_height) = (paper_width * zoom, paper_height * zoom);
    let has_pages = state.output.is_some();
    let toolbar = div()
        .flex()
        .items_center()
        .flex_wrap()
        .gap_2()
        .flex_shrink_0()
        .child(
            Button::new("pdf-prev-page", "Previous")
                .disabled(!has_pages || state.page == 0)
                .secondary(border, text)
                .px_2()
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(|this, _, _, cx| {
                        this.pdf_preview_page(-1, cx);
                        cx.stop_propagation();
                    }),
                ),
        )
        .child(
            div()
                .text_sm()
                .text_color(text)
                .child(if state.page_count > 0 {
                    format!("Page {} of {}", state.page + 1, state.page_count)
                } else {
                    "No pages".into()
                }),
        )
        .child(
            Button::new("pdf-next-page", "Next")
                .disabled(!has_pages || state.page + 1 >= state.page_count)
                .secondary(border, text)
                .px_2()
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(|this, _, _, cx| {
                        this.pdf_preview_page(1, cx);
                        cx.stop_propagation();
                    }),
                ),
        )
        .child(div().flex_1())
        .child(
            Button::new("pdf-zoom-out", "-")
                .disabled(zoom <= 0.5)
                .secondary(border, text)
                .px_2()
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(move |this, _, _, cx| {
                        if zoom > 0.5 {
                            this.pdf_preview_zoom(zoom * 0.8, cx);
                        }
                        cx.stop_propagation();
                    }),
                ),
        )
        .child(
            div().text_xs().text_color(muted).child(
                state
                    .zoom
                    .map_or("Fit".into(), |z| format!("{:.0}%", z * 100.0)),
            ),
        )
        .child(
            Button::new("pdf-zoom-in", "+")
                .disabled(zoom >= 2.0)
                .secondary(border, text)
                .px_2()
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(move |this, _, _, cx| {
                        if zoom < 2.0 {
                            this.pdf_preview_zoom(zoom * 1.25, cx);
                        }
                        cx.stop_propagation();
                    }),
                ),
        )
        .child(
            Button::new("pdf-fit-page", "Fit page")
                .secondary(border, text)
                .px_2()
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(|this, _, _, cx| {
                        this.pdf_preview_zoom(0.0, cx);
                        cx.stop_propagation();
                    }),
                ),
        );
    let page = if let Some(image) = &state.image {
        div()
            .w(px(page_width + 32.0))
            .min_w_full()
            .h(px(page_height + 32.0))
            .min_h_full()
            .p_4()
            .flex()
            .justify_center()
            .items_start()
            .child(
                div()
                    .flex_shrink_0()
                    .w(px(page_width))
                    .h(px(page_height))
                    .bg(rgb(0xffffff))
                    .shadow_md()
                    .child(img(image.clone()).w_full().h_full()),
            )
            .into_any_element()
    } else {
        div()
            .size_full()
            .flex()
            .items_center()
            .justify_center()
            .p_4()
            .text_sm()
            .text_color(text)
            .child(if let Some(error) = &state.preview_error {
                error.clone()
            } else if let Err(error) = &state.summary {
                error.clone()
            } else if state.preparing {
                "Preparing document…".into()
            } else {
                "Rendering page…".into()
            })
            .into_any_element()
    };
    let preview = div()
        .flex()
        .flex_col()
        .gap_2()
        .min_w_0()
        .min_h_0()
        .flex_1()
        .child(toolbar)
        .child(
            div()
                .id("pdf-preview-canvas")
                .flex_1()
                .min_h_0()
                .overflow_scroll()
                .border_1()
                .border_color(border)
                .rounded_md()
                .bg(border.opacity(0.25))
                .child(page),
        )
        .child(
            div()
                .text_xs()
                .text_color(muted)
                .child("Preview zoom changes the view only. Export uses this PDF document."),
        );
    let body = div()
        .flex()
        .gap_4()
        .h(px(body_height))
        .min_h_0()
        .when(compact, |d| d.flex_col())
        .child(settings)
        .child(preview);
    let print_button = Button::new("pdf-print", "Print…").disabled(
        !cfg!(target_os = "linux") || state.busy || state.preparing || state.output.is_none(),
    );
    let export_button = Button::new(
        "pdf-save",
        if state.report.is_some() {
            "Export again…"
        } else {
            "Export PDF…"
        },
    )
    .disabled(state.busy || state.preparing || state.output.is_none());
    let (print_button, export_button) = if state.print_mode {
        (
            print_button.primary(accent, app.token(TokenKey::TextInverse)),
            export_button.secondary(border, text),
        )
    } else {
        (
            print_button.secondary(border, text),
            export_button.primary(accent, app.token(TokenKey::TextInverse)),
        )
    };
    let footer = div()
        .flex()
        .items_center()
        .justify_between()
        .gap_2()
        .child(
            div().text_sm().text_color(muted).max_w(px(350.0)).child(
                (if state.printing {
                    "Use the system print dialog to continue or cancel."
                } else if let Some(message) = &state.print_message {
                    message.as_str()
                } else if state.choosing_path {
                    "Choose a location in the save dialog."
                } else if state.busy {
                    "Saving PDF…"
                } else {
                    ""
                })
                .to_owned(),
            ),
        )
        .child(
            div()
                .flex()
                .gap_2()
                .when(state.saved_path.is_some() && !state.busy, |d| {
                    d.child(
                        Button::new("pdf-open", "Open PDF")
                            .secondary(border, text)
                            .on_mouse_down(
                                MouseButton::Left,
                                cx.listener(|this, _, _, cx| {
                                    if let Some(path) =
                                        this.pdf_export.as_ref().and_then(|s| s.saved_path.as_ref())
                                    {
                                        if let Err(error) = open::that(path) {
                                            if let Some(state) = this.pdf_export.as_mut() {
                                                state.error =
                                                    Some(format!("Could not open PDF: {error}"));
                                            }
                                            cx.notify();
                                        }
                                    }
                                    cx.stop_propagation();
                                }),
                            ),
                    )
                })
                .child(
                    Button::new(
                        "pdf-cancel",
                        if state.busy && !state.printing {
                            "Cancel"
                        } else {
                            "Close"
                        },
                    )
                    .disabled(state.printing)
                    .secondary(border, muted)
                    .on_mouse_down(
                        MouseButton::Left,
                        cx.listener(|this, _, _, cx| {
                            this.close_pdf_export(cx);
                            cx.stop_propagation();
                        }),
                    ),
                )
                .child(print_button.on_mouse_down(
                    MouseButton::Left,
                    cx.listener(|this, _, _, cx| {
                        this.print_pdf(cx);
                        cx.stop_propagation();
                    }),
                ))
                .child(export_button.on_mouse_down(
                    MouseButton::Left,
                    cx.listener(|this, _, _, cx| {
                        this.save_pdf(cx);
                        cx.stop_propagation();
                    }),
                )),
        );
    modal_overlay(
        "pdf-export-dialog",
        |this, cx| this.close_pdf_export(cx),
        DialogFrame::new(body, panel, border)
            .width(px(width))
            .max_height(px(height))
            .header(dialog_header_with_subtitle(
                "Print Preview",
                state
                    .snapshot
                    .as_ref()
                    .map_or("Worksheet".to_string(), |s| s.name.clone()),
                text,
                muted,
            ))
            .footer(footer),
        cx,
    )
    .into_any_element()
}
