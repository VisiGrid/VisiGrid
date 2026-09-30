use crate::ui::{dialog_header_with_subtitle, modal_overlay, Button, DialogFrame, DialogSize};
use crate::{app::Spreadsheet, theme::TokenKey};
use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_print::{Paper, Scale};

pub fn render(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some(state) = &app.pdf_export else {
        return div().into_any_element();
    };
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let labels = [
        (
            "Scope",
            if state.selection_only {
                "Selected range"
            } else {
                "Active sheet"
            },
        ),
        (
            "Paper",
            match state.settings.paper {
                Paper::A4 => "A4",
                Paper::Letter => "Letter",
                Paper::Legal => "Legal",
            },
        ),
        (
            "Orientation",
            if state.settings.landscape {
                "Landscape"
            } else {
                "Portrait"
            },
        ),
        (
            "Scaling",
            match state.settings.scale {
                Scale::FitColumns => "Fit columns on one page",
                Scale::FitSheet => "Fit sheet on one page",
                _ => "Actual size (100%)",
            },
        ),
        (
            "Page numbers",
            if state.settings.footer { "On" } else { "Off" },
        ),
    ];
    let body = div().flex().flex_col().gap_3()
        .children(labels.into_iter().enumerate().map(|(index, (label, value))| {
            div().flex().items_center().justify_between().gap_4()
                .child(div().text_sm().text_color(text).child(label))
                .child(Button::new(ElementId::Name(format!("pdf-option-{index}").into()), format!("{value}  ›"))
                    .disabled(state.busy)
                    .secondary(if state.focus == index { accent } else { border }, text)
                    .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                        this.pdf_cycle_option(index, cx);
                        cx.stop_propagation();
                    })))
        }))
        .child(div().mt_2().p_3().border_1().border_color(border).rounded_md().text_sm()
            .text_color(if state.summary.is_err() { app.token(TokenKey::Warn) } else { text })
            .child(if state.busy { "Preparing PDF… Cancel stops this export.".to_string() } else {
                state.summary.clone().unwrap_or_else(|e| e)
            }))
        .when(state.error.is_some(), |element| element.child(div().text_sm().text_color(app.token(TokenKey::Warn)).child(state.error.clone().unwrap())))
        .when(state.report.is_some(), |element| element.child(div().text_sm().text_color(text).child(state.report.clone().unwrap())))
        .child(div().text_xs().text_color(muted).child("Half-inch margins. Uses current calculated values and visible rows/columns. Blank outer formatting is excluded; select a range to include it."))
        .child(div().text_xs().text_color(muted).child("Tab: choose a setting · Space: change · Enter: export · Esc: cancel"));
    let footer = div()
        .flex()
        .justify_end()
        .gap_2()
        .when(state.saved_path.is_some() && !state.busy, |div| {
            div.child(
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
                                        state.error = Some(format!("Could not open PDF: {error}"));
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
                if state.report.is_some() {
                    "Close"
                } else {
                    "Cancel"
                },
            )
            .secondary(border, muted)
            .on_mouse_down(
                MouseButton::Left,
                cx.listener(|this, _, _, cx| {
                    this.close_pdf_export(cx);
                    cx.stop_propagation();
                }),
            ),
        )
        .child(
            Button::new("pdf-save", "Export PDF…")
                .disabled(state.busy || state.summary.is_err())
                .primary(accent, app.token(TokenKey::TextInverse))
                .on_mouse_down(
                    MouseButton::Left,
                    cx.listener(|this, _, _, cx| {
                        this.save_pdf(cx);
                        cx.stop_propagation();
                    }),
                ),
        );
    modal_overlay(
        "pdf-export-dialog",
        |this, cx| this.close_pdf_export(cx),
        DialogFrame::new(body, app.token(TokenKey::PanelBg), border)
            .size(DialogSize::Lg)
            .header(dialog_header_with_subtitle(
                "Export PDF",
                "Choose the page layout for this sheet",
                text,
                muted,
            ))
            .footer(footer),
        cx,
    )
    .into_any_element()
}
