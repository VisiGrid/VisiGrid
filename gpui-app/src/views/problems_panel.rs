//! Problems panel (F10): every formula error in the workbook, grouped by
//! sheet, click-to-jump. VS Code vocabulary — the engine already knows
//! what's broken; this is the chrome that surfaces it.

use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::app::Spreadsheet;
use crate::theme::TokenKey;

/// Column index → letters (A, B, ..., Z, AA, ...).
fn col_letters(col: usize) -> String {
    let mut letters = String::new();
    let mut c = col + 1;
    while c > 0 {
        let rem = (c - 1) % 26;
        letters.insert(0, (b'A' + rem as u8) as char);
        c = (c - 1) / 26;
    }
    letters
}

pub(super) fn render_problems_panel(
    app: &Spreadsheet,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let error_color = app.token(TokenKey::Error);

    let (problems, truncated) = app.collect_problems(cx);
    let count = problems.len();

    let header = div()
        .flex()
        .items_center()
        .gap_2()
        .px_2()
        .py_1()
        .border_b_1()
        .border_color(panel_border)
        .child(
            div()
                .text_size(px(11.0))
                .font_weight(FontWeight::BOLD)
                .text_color(if count > 0 { error_color } else { text_muted })
                .child(if count == 0 {
                    "No problems".to_string()
                } else if truncated {
                    format!("{}+ problems (showing first {})", count, count)
                } else {
                    format!("{} problem{}", count, if count == 1 { "" } else { "s" })
                }),
        )
        .child(div().flex_1())
        .child(
            div()
                .text_size(px(10.0))
                .text_color(text_muted)
                .child("click a row to jump — F10 closes"),
        );

    let mut list = div()
        .id("problems-list")
        .flex()
        .flex_col()
        .h(px(160.0))
        .overflow_y_scroll()
        .bg(panel_bg);

    if count == 0 {
        list = list.child(
            div()
                .p_3()
                .text_size(px(11.0))
                .text_color(text_muted)
                .child("All formulas evaluate cleanly."),
        );
    } else {
        let mut last_sheet: Option<usize> = None;
        for p in problems {
            if last_sheet != Some(p.sheet_idx) {
                last_sheet = Some(p.sheet_idx);
                list = list.child(
                    div()
                        .px_2()
                        .pt_2()
                        .pb_1()
                        .text_size(px(10.0))
                        .font_weight(FontWeight::BOLD)
                        .text_color(text_muted)
                        .child(p.sheet_name.clone()),
                );
            }
            let cell_ref = format!("{}{}", col_letters(p.col), p.row + 1);
            let (sheet_idx, row, col) = (p.sheet_idx, p.row, p.col);
            let formula = if p.formula.len() > 80 {
                format!("{}…", &p.formula[..p.formula.len().min(80)])
            } else {
                p.formula.clone()
            };
            list = list.child(
                div()
                    .id(SharedString::from(format!("problem-{}-{}-{}", sheet_idx, row, col)))
                    .flex()
                    .items_center()
                    .gap_2()
                    .px_3()
                    .py(px(3.0))
                    .cursor_pointer()
                    .hover(|s| s.bg(hsla(0.0, 0.0, 0.5, 0.12)))
                    .on_mouse_down(
                        MouseButton::Left,
                        cx.listener(move |this, _, _, cx| {
                            this.reveal_cell(sheet_idx, row, col, cx);
                        }),
                    )
                    .child(
                        div()
                            .w(px(56.0))
                            .flex_shrink_0()
                            .text_size(px(11.0))
                            .font_weight(FontWeight::MEDIUM)
                            .text_color(text_primary)
                            .child(cell_ref),
                    )
                    .child(
                        div()
                            .w(px(80.0))
                            .flex_shrink_0()
                            .text_size(px(11.0))
                            .text_color(error_color)
                            .child(p.error.clone()),
                    )
                    .child(
                        div()
                            .flex_1()
                            .text_size(px(11.0))
                            .text_color(text_muted)
                            .overflow_hidden()
                            .child(formula),
                    ),
            );
        }
        if truncated {
            list = list.child(
                div()
                    .px_3()
                    .py_1()
                    .text_size(px(10.0))
                    .text_color(text_muted)
                    .child("…more problems not shown (fix some and the list refreshes)"),
            );
        }
    }

    div().flex().flex_col().child(header).child(list)
}
