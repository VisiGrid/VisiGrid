use crate::{app::Spreadsheet, theme::TokenKey};
use gpui::*;

pub fn render(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> AnyElement {
    let Some((_, issue)) = &app.pending_table_recovery else {
        return div().into_any_element();
    };
    let title = match issue {
        visigrid_io::table_recovery::TableLoadIssue::FutureVersion(_) => {
            "Upgrade VisiGrid to open these Tables"
        }
        visigrid_io::table_recovery::TableLoadIssue::Corrupt(_) => "Table metadata is corrupt",
    };
    let content = div().w(px(540.0)).p_6().rounded_md().border_1()
        .bg(app.token(TokenKey::PanelBg)).border_color(app.token(TokenKey::PanelBorder))
        .text_color(app.token(TokenKey::TextPrimary)).flex().flex_col().gap_4()
        .child(div().text_size(px(18.0)).child(title))
        .child(div().text_size(px(13.0)).child(issue.to_string()))
        .child(div().text_size(px(13.0)).child("You can inspect the cells read-only. Table definitions will not be active. Saved formula results may be stale or unavailable. Editing, Save, Save As and export will be disabled; the original file stays unchanged."))
        .child(div().flex().justify_end().gap_3()
            .child(div().id("table-recovery-cancel").px_3().py_2().cursor_pointer().child("Cancel")
                .on_mouse_down(MouseButton::Left, cx.listener(|s, _, _, cx| { cx.stop_propagation(); s.pending_table_recovery = None; cx.notify(); })))
            .child(div().id("table-recovery-open").px_3().py_2().rounded_md().cursor_pointer()
                .bg(app.token(TokenKey::Accent)).text_color(app.token(TokenKey::TextInverse)).child("Open read-only")
                .on_mouse_down(MouseButton::Left, cx.listener(|s, _, _, cx| { cx.stop_propagation(); s.confirm_table_recovery(cx); }))));
    crate::ui::modal_overlay(
        "table-recovery",
        |s, cx| {
            s.pending_table_recovery = None;
            cx.notify();
        },
        content,
        cx,
    )
    .into_any_element()
}
