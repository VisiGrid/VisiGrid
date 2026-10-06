//! Traditional cell comments. Drafts stay outside the workbook until Save.
use crate::{
    app::Spreadsheet,
    history::{CommentPatch, UndoAction},
    ime::ImeBuffer,
    mode::Mode,
    theme::TokenKey,
};
use gpui::prelude::FluentBuilder;
use gpui::*;
use visigrid_engine::cell::CellComment;
use visigrid_engine::sheet::SheetId;

#[path = "comment_plan.rs"]
pub(crate) mod plan;

pub struct CommentEditor {
    pub revision: u64,
    pub error: Option<String>,
    pub sheet_id: SheetId,
    pub row: usize,
    pub col: usize,
    pub text: ImeBuffer,
    pub author: ImeBuffer,
    pub author_focused: bool,
    pub existing: bool,
    pub layout: std::cell::RefCell<Option<TextLayout>>,
    pub dragging: bool,
}
impl CommentEditor {
    pub fn buffer(&self) -> &ImeBuffer {
        if self.author_focused {
            &self.author
        } else {
            &self.text
        }
    }
    pub fn buffer_mut(&mut self) -> &mut ImeBuffer {
        if self.author_focused {
            &mut self.author
        } else {
            &mut self.text
        }
    }
}
impl Spreadsheet {
    pub fn toggle_comment_previews(&mut self, cx: &mut Context<Self>) {
        let enabled = !previews_enabled(cx);
        crate::settings::update_user_settings(cx, |settings| {
            settings.appearance.show_comment_previews = crate::settings::Setting::Value(enabled);
        });
        self.status_message = Some(
            if enabled {
                "Comment previews on"
            } else {
                "Comment previews off"
            }
            .into(),
        );
        cx.notify();
    }

    pub fn open_comment(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        self.sync_table_view(cx);
        let (r, c) = self.view_state.selected;
        self.open_cell_comment(self.row_view.view_to_data(r), c, window, cx);
    }
    pub fn open_cell_comment(
        &mut self,
        row: usize,
        col: usize,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        if !self.mode.is_navigation() || (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        self.sync_table_view(cx);
        let sheet = self.sheet(cx);
        let (row, col) = sheet
            .get_merge(row, col)
            .map(|m| m.start)
            .unwrap_or((row, col));
        if let Err(error) = plan::validate_target(sheet, row, col, true, false) {
            self.status_message = Some(error);
            cx.notify();
            return;
        }
        if self.row_view.data_to_view(row).is_none() || self.is_row_hidden(row) || self.is_col_hidden(col) {
            self.status_message = Some("Reveal this cell before editing its comment.".into());
            cx.notify();
            return;
        }
        let comment = sheet.comment(row, col).cloned();
        let text = comment.as_ref().map(|c| c.text.clone()).unwrap_or_default();
        let author = comment
            .as_ref()
            .map(|c| c.author.clone())
            .unwrap_or_default();
        let sheet_id = sheet.id;
        self.comment_reader = None;
        self.end_drag_selection(cx);
        if let Some(view_row) = self.row_view.data_to_view(row) {
            self.select_cell(view_row, col, false, cx);
        }
        self.comment_editor = Some(CommentEditor {
            revision: self.wb(cx).revision(),
            error: None,
            sheet_id,
            row,
            col,
            existing: comment.is_some(),
            author_focused: false,
            layout: Default::default(),
            dragging: false,
            text: ImeBuffer {
                cursor: text.len(),
                text,
                ..Default::default()
            },
            author: ImeBuffer {
                cursor: author.len(),
                text: author,
                ..Default::default()
            },
        });
        self.mode = Mode::Comment;
        window.focus(&self.focus_handle, cx);
        cx.notify();
    }
    pub fn close_comment(&mut self, cx: &mut Context<Self>) {
        self.comment_editor = None;
        self.mode = Mode::Navigation;
        cx.notify();
    }
    pub fn save_comment(&mut self, cx: &mut Context<Self>) {
        let Some(editor) = &self.comment_editor else {
            return;
        };
        if editor.text.text.trim().is_empty() {
            return;
        }
        let Some(index) = self
            .wb(cx)
            .sheets()
            .iter()
            .position(|s| s.id == editor.sheet_id)
        else {
            self.comment_error("The comment sheet no longer exists. Copy your draft before closing it.".into(), cx);
            return;
        };
        let (row, col) = (editor.row, editor.col);
        let after = Some(CellComment {
            text: editor.text.text.clone(),
            author: editor.author.text.trim().to_owned(),
        });
        if self.change_comment(index, row, col, after, cx) {
            self.close_comment(cx);
        }
    }
    pub fn delete_comment(&mut self, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        self.sync_table_view(cx);
        let (index, row, col) = if let Some(editor) = &self.comment_editor {
            let Some(index) = self
                .wb(cx)
                .sheets()
                .iter()
                .position(|s| s.id == editor.sheet_id)
            else {
                return;
            };
            (index, editor.row, editor.col)
        } else {
            let (r, c) = self.view_state.selected;
            let r = self.row_view.view_to_data(r);
            let (r, c) = self
                .sheet(cx)
                .get_merge(r, c)
                .map(|m| m.start)
                .unwrap_or((r, c));
            (self.sheet_index(cx), r, c)
        };
        if self.change_comment(index, row, col, None, cx) && self.comment_editor.is_some() {
            self.close_comment(cx);
        }
    }
    fn comment_error(&mut self, error: String, cx: &mut Context<Self>) {
        if let Some(editor) = &mut self.comment_editor { editor.error = Some(error.clone()); }
        self.status_message = Some(error);
        cx.notify();
    }
    fn change_comment(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        after: Option<CellComment>,
        cx: &mut Context<Self>,
    ) -> bool {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) { return false; }
        let result = plan::validate_edit(self.wb(cx), sheet_index, row, col, after.is_some(),
            self.comment_editor.as_ref().map(|e| e.revision));
        if let Err(error) = result {
            self.comment_error(error, cx);
            return false;
        }
        if sheet_index == self.sheet_index(cx)
            && (self.row_view.data_to_view(row).is_none() || self.is_row_hidden(row) || self.is_col_hidden(col))
        {
            self.comment_error("Reveal this cell before editing its comment.".into(), cx);
            return false;
        }
        let remove_cell_on_undo = self.wb(cx).sheet(sheet_index)
            .is_some_and(|s| s.get_cell_opt(row, col).is_none());
        let before = self
            .wb(cx)
            .sheet(sheet_index)
            .and_then(|s| s.comment(row, col))
            .cloned();
        if before == after {
            return true;
        }
        let description = if after.is_none() {
            "Delete comment"
        } else if before.is_none() {
            "Add comment"
        } else {
            "Edit comment"
        }
        .to_owned();
        self.workbook.update(cx, |wb, _| {
            if let Some(s) = wb.sheet_mut(sheet_index) {
                s.set_comment(row, col, after.clone());
            }
            wb.bump_revision_for_structure();
        });
        self.history.record_action_with_provenance(
            UndoAction::Comments {
                sheet_index,
                patches: vec![CommentPatch {
                    remove_cell_on_undo,
                    row,
                    col,
                    before,
                    after,
                }],
                description: description.clone(),
            },
            None,
        );
        self.is_modified = true;
        self.bump_cells_rev();
        self.request_title_refresh(cx);
        self.status_message = Some(description);
        cx.notify();
        true
    }
    pub fn comment_key(&mut self, event: &KeyDownEvent, cx: &mut Context<Self>) {
        let key = event.keystroke.key.as_str();
        let mods = event.keystroke.modifiers;
        let command = mods.platform || mods.control;
        if key == "escape" {
            self.close_comment(cx);
            cx.stop_propagation();
            return;
        }
        if key == "enter" && command {
            self.save_comment(cx);
            cx.stop_propagation();
            return;
        }
        let Some(editor) = self.comment_editor.as_mut() else {
            return;
        };
        if key == "tab" {
            editor.author_focused = !editor.author_focused;
            cx.stop_propagation();
            cx.notify();
            return;
        }
        let author = editor.author_focused;
        let layout = editor.layout.borrow().clone();
        let buf = editor.buffer_mut();
        match key {
            "a" if command => {
                buf.selection_anchor = Some(0);
                buf.cursor = buf.text.len();
            }
            "c" | "x" if command => {
                let range = buf.selection();
                if !range.is_empty() {
                    cx.write_to_clipboard(ClipboardItem::new_string(
                        buf.text[range.clone()].to_owned(),
                    ));
                    if key == "x" {
                        buf.replace(None, "");
                    }
                }
            }
            "v" if command => {
                if let Some(text) = cx.read_from_clipboard().and_then(|c| c.text()) {
                    buf.replace(
                        None,
                        &if author {
                            text.replace(['\r', '\n'], " ")
                        } else {
                            text.replace("\r\n", "\n")
                        },
                    );
                }
            }
            "enter" => {
                if !author {
                    buf.replace(None, "\n");
                }
            }
            "backspace" | "delete" => {
                if buf.selection().is_empty() {
                    let cursor = buf.cursor;
                    let other = if key == "backspace" {
                        buf.text[..cursor]
                            .char_indices()
                            .next_back()
                            .map(|(i, _)| i)
                            .unwrap_or(0)
                    } else {
                        buf.text[cursor..]
                            .chars()
                            .next()
                            .map(|c| cursor + c.len_utf8())
                            .unwrap_or(cursor)
                    };
                    buf.selection_anchor = Some(other);
                }
                buf.replace(None, "");
            }
            "left" | "right" | "home" | "end" | "up" | "down" => {
                let old = buf.cursor;
                let sel = buf.selection();
                let next = match key {
                    "left" if !mods.shift && !sel.is_empty() => sel.start,
                    "right" if !mods.shift && !sel.is_empty() => sel.end,
                    "left" => buf.text[..old]
                        .char_indices()
                        .next_back()
                        .map(|(i, _)| i)
                        .unwrap_or(0),
                    "right" => buf.text[old..]
                        .chars()
                        .next()
                        .map(|c| old + c.len_utf8())
                        .unwrap_or(old),
                    "up" | "down" => layout
                        .as_ref()
                        .and_then(|layout| {
                            layout.position_for_index(old).map(|mut pos| {
                                pos.y +=
                                    layout.line_height() * if key == "up" { -1.0 } else { 1.0 };
                                layout
                                    .index_for_position(pos)
                                    .unwrap_or_else(|i| i)
                                    .min(buf.text.len())
                            })
                        })
                        .unwrap_or(old),
                    "home" => buf.text[..old].rfind('\n').map(|i| i + 1).unwrap_or(0),
                    _ => buf.text[old..]
                        .find('\n')
                        .map(|i| old + i)
                        .unwrap_or(buf.text.len()),
                };
                buf.selection_anchor = if mods.shift {
                    Some(buf.selection_anchor.unwrap_or(old))
                } else {
                    None
                };
                buf.cursor = next;
                buf.marked = None;
            }
            _ => {
                // Committed and composed text is delivered by EntityInputHandler.
                if !command {
                    cx.propagate();
                    return;
                }
            }
        }
        cx.stop_propagation();
        cx.notify();
    }
}

pub fn render(
    app: &Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    use crate::ui::Button;
    let editor = app.comment_editor.as_ref().expect("comment dialog");
    let bg = app.token(TokenKey::PanelBg);
    let border = app.token(TokenKey::PanelBorder);
    let accent = app.token(TokenKey::Accent);
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let error = app.token(TokenKey::Error);
    let save_enabled = !editor.text.text.trim().is_empty();
    let sheet_name = app
        .wb(cx)
        .sheets()
        .iter()
        .find(|s| s.id == editor.sheet_id)
        .map(|s| s.name.as_str())
        .unwrap_or("Sheet");
    let location = format!(
        "{} · {}{}",
        sheet_name,
        Spreadsheet::col_letter(editor.col),
        editor.row + 1
    );
    let header = div()
        .flex()
        .items_start()
        .justify_between()
        .gap_3()
        .child(
            div()
                .flex()
                .flex_col()
                .gap_1()
                .child(
                    div()
                        .text_size(px(17.))
                        .font_weight(FontWeight::SEMIBOLD)
                        .text_color(text)
                        .child(if editor.existing {
                            "Edit comment"
                        } else {
                            "New comment"
                        }),
                )
                .child(div().text_size(px(12.)).text_color(muted).child(location)),
        )
        .child(
            div()
                .id("comment-close")
                .size(px(28.))
                .flex()
                .items_center()
                .justify_center()
                .rounded_md()
                .text_color(muted)
                .text_size(px(20.))
                .cursor_pointer()
                .hover(move |d| d.bg(border.opacity(0.35)).text_color(text))
                .child("×")
                .on_click(cx.listener(|this, _, _, cx| this.close_comment(cx))),
        );
    let footer = div()
        .flex()
        .items_center()
        .justify_between()
        .gap_3()
        .child(div().when(editor.existing, |d| {
            d.child(
                div()
                    .id("comment-delete")
                    .px_2()
                    .py(px(6.))
                    .rounded_md()
                    .text_size(px(12.))
                    .text_color(muted)
                    .cursor_pointer()
                    .hover(move |d| d.bg(error.opacity(0.08)).text_color(error))
                    .child("Delete comment")
                    .on_click(cx.listener(|this, _, _, cx| this.delete_comment(cx))),
            )
        }))
        .child(
            div()
                .flex()
                .items_center()
                .gap_2()
                .child(
                    Button::new("comment-cancel", "Cancel")
                        .secondary(border, text)
                        .on_click(cx.listener(|this, _, _, cx| this.close_comment(cx))),
                )
                .child(
                    Button::new(
                        "comment-save",
                        if editor.existing {
                            "Save changes"
                        } else {
                            "Add comment"
                        },
                    )
                    .disabled(!save_enabled)
                    .primary(accent, gpui::white())
                    .when(save_enabled, |b| {
                        b.on_click(cx.listener(|this, _, _, cx| this.save_comment(cx)))
                    }),
                ),
        );
    crate::ui::modal_backdrop(
        "comment-dialog",
        div()
            .w(px(480.))
            .max_w(window.viewport_size().width - px(32.))
            .bg(bg)
            .border_1()
            .border_color(border)
            .rounded_lg()
            .shadow_xl()
            .flex()
            .flex_col()
            .overflow_hidden()
            .child(
                div()
                    .p(px(20.))
                    .flex()
                    .flex_col()
                    .gap(px(18.))
                    .child(header)
                    .when_some(editor.error.as_ref(), |d, message| {
                        d.child(div().text_size(px(12.)).text_color(error).child(message.clone()))
                    })
                    .child(
                        div()
                            .flex()
                            .flex_col()
                            .gap_2()
                            .child(render_field(app, false, window, cx))
                            .child(div().text_color(muted).text_size(px(11.)).child(
                                if cfg!(target_os = "macos") {
                                    "⌘Enter to save · Esc to cancel"
                                } else {
                                    "Ctrl+Enter to save · Esc to cancel"
                                },
                            )),
                    )
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .gap_3()
                            .child(
                                div()
                                    .flex_shrink_0()
                                    .flex()
                                    .flex_col()
                                    .gap(px(2.))
                                    .child(
                                        div().text_color(text).text_size(px(12.)).child("Author"),
                                    )
                                    .child(
                                        div()
                                            .text_color(muted)
                                            .text_size(px(11.))
                                            .child("Optional"),
                                    ),
                            )
                            .child(
                                div()
                                    .flex_1()
                                    .min_w_0()
                                    .child(render_field(app, true, window, cx)),
                            ),
                    ),
            )
            .child(
                div()
                    .border_t_1()
                    .border_color(border)
                    .px(px(20.))
                    .py(px(14.))
                    .child(footer),
            ),
    )
}

pub fn previews_enabled(cx: &App) -> bool {
    !matches!(
        crate::settings::user_settings(cx)
            .appearance
            .show_comment_previews,
        crate::settings::Setting::Value(false)
    )
}

/// Build on demand, so rendering a grid does not clone every visible comment.
/// GPUI supplies the 500ms hover delay, leave grace period and viewport fitting.
pub fn preview(
    app: &Spreadsheet,
    row: usize,
    col: usize,
    cx: &Context<Spreadsheet>,
) -> impl Fn(&mut Window, &mut App) -> AnyView + 'static {
    let owner = cx.entity().downgrade();
    let sheet_id = app.sheet(cx).id;
    move |_, cx| {
        let owner = owner.clone();
        cx.new(|_| CommentPreview {
            owner,
            sheet_id,
            row,
            col,
        })
        .into()
    }
}

struct CommentPreview {
    owner: WeakEntity<Spreadsheet>,
    sheet_id: SheetId,
    row: usize,
    col: usize,
}

impl Render for CommentPreview {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let Some(owner) = self.owner.upgrade() else {
            return div().into_any_element();
        };
        let app = owner.read(cx);
        // A hover is only a view of the current sheet, never a retained draft.
        if !previews_enabled(cx)
            || reader_visible(app, cx)
            || !app.mode.is_navigation()
            || app.sheet(cx).id != self.sheet_id
        {
            return div().into_any_element();
        }
        let Some(comment) = app.sheet(cx).comment(self.row, self.col) else {
            return div().into_any_element();
        };
        let bg = app.token(TokenKey::PanelBg);
        let border = app.token(TokenKey::PanelBorder);
        let text = app.token(TokenKey::TextPrimary);
        let muted = app.token(TokenKey::TextMuted);
        let accent = app.token(TokenKey::Accent);
        let location = format!(
            "{} · {}{}",
            app.sheet(cx).name,
            Spreadsheet::col_letter(self.col),
            self.row + 1
        );
        let author = comment.author.clone();
        let body = comment.text.clone();
        div()
            .id("comment-preview")
            .w(px(300.))
            .max_w(window.viewport_size().width - px(24.))
            .flex()
            .flex_col()
            .bg(bg)
            .border_1()
            .border_color(border)
            .rounded_lg()
            .shadow_lg()
            .text_color(text)
            .text_size(px(13.))
            .cursor_default()
            .overflow_hidden()
            .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation())
            .on_scroll_wheel(|_, _, cx| cx.stop_propagation())
            .child(
                div()
                    .px_3()
                    .pt_3()
                    .pb_2()
                    .flex()
                    .flex_col()
                    .gap_1()
                    .child(div().text_size(px(11.)).text_color(muted).child(location))
                    .when(!author.is_empty(), |d| {
                        d.child(div().font_weight(FontWeight::SEMIBOLD).child(author))
                    }),
            )
            .child(
                div()
                    .id("comment-preview-body")
                    .px_3()
                    .pb_3()
                    .max_h(px(200.).min(window.viewport_size().height * 0.4))
                    .overflow_y_scroll()
                    .child(body),
            )
            .child(
                div().border_t_1().border_color(border).px_3().py_2().child(
                    div()
                        .id("comment-preview-edit")
                        .text_size(px(12.))
                        .text_color(accent)
                        .cursor_pointer()
                        .hover(|d| d.underline())
                        .child("Edit comment")
                        .on_click(cx.listener(|this, _, window, cx| {
                            let _ = this.owner.update(cx, |app, cx| {
                                if app.sheet(cx).id == this.sheet_id {
                                    app.open_cell_comment(this.row, this.col, window, cx);
                                }
                            });
                        })),
                ),
            )
            .into_any_element()
    }
}

pub fn indicator(row: usize, col: usize, cx: &Context<Spreadsheet>) -> impl IntoElement {
    div()
        .id(ElementId::Name(format!("comment-{row}-{col}").into()))
        .absolute()
        .top_0()
        .right_0()
        .w(px(9.))
        .h(px(9.))
        .child(
            canvas(
                |_, _, _| (),
                |bounds, _, window, _| {
                    let mut path = PathBuilder::fill();
                    path.move_to(bounds.origin);
                    path.line_to(bounds.origin + point(bounds.size.width, px(0.)));
                    path.line_to(bounds.origin + point(bounds.size.width, bounds.size.height));
                    path.close();
                    if let Ok(path) = path.build() {
                        window.paint_path(path, rgb(0xd94667));
                    }
                },
            )
            .size_full(),
        )
        .cursor_pointer()
        .on_mouse_down(
            MouseButton::Left,
            cx.listener(move |this, _, window, cx| {
                cx.stop_propagation();
                this.open_cell_comment(row, col, window, cx);
            }),
        )
}

/// Text layout is shared with hit testing and the IME candidate-window anchor.
fn render_field(
    app: &Spreadsheet,
    author: bool,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let editor = app.comment_editor.as_ref().unwrap();
    let buf = if author { &editor.author } else { &editor.text };
    let focused = editor.author_focused == author;
    let accent = app.token(TokenKey::Accent);
    let mut style = window.text_style();
    style.color = app.token(TokenKey::TextPrimary);
    let sel = buf.selection();
    let mut highlights = Vec::new();
    if focused && !sel.is_empty() {
        highlights.push((
            sel,
            HighlightStyle {
                background_color: Some(app.token(TokenKey::SelectionBg)),
                ..Default::default()
            },
        ));
    }
    if let Some(marked) = &buf.marked {
        highlights.push((
            marked.clone(),
            HighlightStyle {
                underline: Some(UnderlineStyle {
                    thickness: px(1.),
                    ..Default::default()
                }),
                ..Default::default()
            },
        ));
    }
    // The trailing space supplies layout for the caret after a final newline.
    let styled =
        StyledText::new(format!("{} ", buf.text)).with_default_highlights(&style, highlights);
    let layout = styled.layout().clone();
    if focused {
        *editor.layout.borrow_mut() = Some(layout.clone());
    }
    let click_layout = layout.clone();
    let drag_layout = layout.clone();
    let cursor = buf.cursor;
    div()
        .id(if author {
            "comment-author"
        } else {
            "comment-text"
        })
        .relative()
        .px_3()
        .py_2()
        .text_size(px(13.))
        .bg(app.token(TokenKey::AppBg))
        .border_1()
        .border_color(if focused {
            accent
        } else {
            app.token(TokenKey::PanelBorder)
        })
        .rounded_md()
        .cursor_text()
        .when(!author, |d| d.h(px(164.)).overflow_y_scroll())
        .child(styled)
        .when(buf.text.is_empty(), |d| {
            d.child(
                div()
                    .absolute()
                    .left(px(12.))
                    .top(px(8.))
                    .text_color(app.token(TokenKey::TextMuted))
                    .child(if author { "Name" } else { "Write a comment…" }),
            )
        })
        .child(
            canvas(
                |_, _, _| (),
                move |_, _, window, _| {
                    if focused {
                        if let Some(pos) = layout.position_for_index(cursor) {
                            window.paint_quad(fill(
                                Bounds::new(pos, size(px(1.), layout.line_height())),
                                accent,
                            ));
                        }
                    }
                },
            )
            .absolute()
            .inset_0(),
        )
        .on_mouse_down(
            MouseButton::Left,
            cx.listener(move |this, event: &MouseDownEvent, _, cx| {
                if let Some(e) = &mut this.comment_editor {
                    e.author_focused = author;
                    e.dragging = true;
                    let buf = e.buffer_mut();
                    let index = click_layout
                        .index_for_position(event.position)
                        .unwrap_or_else(|i| i)
                        .min(buf.text.len());
                    let index = buf.text.floor_char_boundary(index);
                    buf.selection_anchor = Some(if event.modifiers.shift {
                        buf.selection_anchor.unwrap_or(buf.cursor)
                    } else {
                        index
                    });
                    buf.cursor = index;
                    buf.marked = None;
                }
                cx.stop_propagation();
                cx.notify();
            }),
        )
        .on_mouse_move(cx.listener(move |this, event: &MouseMoveEvent, _, cx| {
            if let Some(e) = &mut this.comment_editor {
                if e.dragging
                    && e.author_focused == author
                    && event.pressed_button == Some(MouseButton::Left)
                {
                    let buf = e.buffer_mut();
                    let index = drag_layout
                        .index_for_position(event.position)
                        .unwrap_or_else(|i| i)
                        .min(buf.text.len());
                    buf.cursor = buf.text.floor_char_boundary(index);
                    cx.notify();
                }
            }
        }))
        .on_mouse_up(
            MouseButton::Left,
            cx.listener(|this, _, _, cx| {
                if let Some(e) = &mut this.comment_editor {
                    e.dragging = false;
                }
                cx.stop_propagation();
            }),
        )
}

pub struct CommentReader {
    sheet_id: SheetId,
    row: usize,
    col: usize,
    selection: (usize, usize),
    revision: u64,
    ordinal: usize,
    total: usize,
    hidden: bool,
}

// Workbook tab order, then canonical row/column order. Strict comparisons skip
// the current cell, and a one-comment workbook deliberately wraps to itself.
fn comment_step(
    locations: &[(usize, usize, usize)],
    cursor: (usize, usize, usize),
    forward: bool,
) -> Option<usize> {
    if locations.is_empty() {
        return None;
    }
    Some(if forward {
        locations
            .iter()
            .position(|location| *location > cursor)
            .unwrap_or(0)
    } else {
        locations
            .iter()
            .rposition(|location| *location < cursor)
            .unwrap_or(locations.len() - 1)
    })
}

impl Spreadsheet {
    pub fn dismiss_stale_comment_reader(&mut self, cx: &App) {
        if self.comment_reader.as_ref().is_some_and(|r| {
            r.sheet_id != self.sheet(cx).id
                || r.selection != self.view_state.selected
                || r.revision != self.cells_rev
        }) {
            self.comment_reader = None;
        }
    }

    pub fn navigate_comment(&mut self, forward: bool, window: &mut Window, cx: &mut Context<Self>) {
        if self.is_previewing() || self.review_mode.is_some() { return; }
        self.sync_table_view(cx);
        if !self.mode.is_navigation() && self.mode != Mode::Command {
            return;
        }
        let mut locations: Vec<_> = self
            .wb(cx)
            .sheets()
            .iter()
            .enumerate()
            .flat_map(|(sheet, data)| {
                data.comments()
                    .map(move |((row, col), _)| (sheet, row, col))
            })
            .collect();
        locations.sort_unstable();
        let current_sheet = self.sheet_index(cx);
        let current = self
            .comment_reader
            .as_ref()
            .filter(|r| {
                r.sheet_id == self.sheet(cx).id
                    && r.selection == self.view_state.selected
                    && r.revision == self.cells_rev
            })
            .map(|r| (current_sheet, r.row, r.col))
            .unwrap_or_else(|| {
                let (row, col) = self.view_state.selected;
                (current_sheet, self.view_to_data(row, cx), col)
            });
        let Some(index) = comment_step(&locations, current, forward) else {
            self.comment_reader = None;
            self.status_message = Some("No comments in this workbook".into());
            cx.notify();
            return;
        };
        let (sheet, row, col) = locations[index];
        self.jump_to_comment(sheet, row, col, index + 1, locations.len(), window, cx);
    }

    pub(crate) fn jump_to_comment(
        &mut self,
        sheet: usize,
        row: usize,
        col: usize,
        ordinal: usize,
        total: usize,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        if self.is_previewing() || self.review_mode.is_some()
            || (!self.mode.is_navigation() && self.mode != Mode::Command)
        {
            return;
        }
        if self
            .wb(cx)
            .sheet(sheet)
            .and_then(|s| s.comment(row, col))
            .is_none()
        {
            return;
        }
        let current_sheet = self.sheet_index(cx);
        if sheet != current_sheet {
            self.goto_sheet(sheet, cx);
            if self.sheet_index(cx) != sheet {
                return;
            }
        }
        self.sync_table_view(cx);
        self.mode = Mode::Navigation;
        self.end_drag_selection(cx);
        let view_row = self.data_to_view(row, cx);
        let hidden = view_row.is_none() || self.is_row_hidden(row) || self.is_col_hidden(col);
        if let Some(view_row) = view_row.filter(|_| !hidden) {
            self.select_cell(view_row, col, false, cx);
            self.ensure_visible(cx);
        }
        self.comment_reader = Some(CommentReader {
            sheet_id: self.sheet(cx).id,
            row,
            col,
            selection: self.view_state.selected,
            revision: self.cells_rev,
            ordinal,
            total,
            hidden,
        });
        self.status_message = Some(format!(
            "Comment {} of {} · {}!{}{}",
            ordinal,
            total,
            self.sheet(cx).name,
            Self::col_letter(col),
            row + 1
        ));
        window.focus(&self.focus_handle, cx);
        cx.notify();
    }
}

pub fn reader_visible(app: &Spreadsheet, cx: &App) -> bool {
    app.mode.is_navigation()
        && app.comment_reader.as_ref().is_some_and(|r| {
            r.sheet_id == app.sheet(cx).id
                && r.selection == app.view_state.selected
                && r.revision == app.cells_rev
                && app.sheet(cx).comment(r.row, r.col).is_some()
        })
}

pub fn render_reader(
    app: &Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let reader = app.comment_reader.as_ref().expect("visible comment reader");
    let comment = app
        .sheet(cx)
        .comment(reader.row, reader.col)
        .expect("existing comment");
    let bg = app.token(TokenKey::PanelBg);
    let border = app.token(TokenKey::PanelBorder);
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let (row, col) = (reader.row, reader.col);
    let location = format!(
        "{} · {}{}",
        app.sheet(cx).name,
        Spreadsheet::col_letter(col),
        row + 1
    );
    div()
        .id("comment-reader")
        .absolute()
        .right(px(if app.comments_sidebar_visible {
            crate::comment_sidebar::WIDTH + 16.
        } else {
            16.
        }))
        .top(px(app.grid_layout.grid_body_origin.1 + 12.))
        .w(px(320.))
        .max_w(window.viewport_size().width - px(32.))
        .flex()
        .flex_col()
        .bg(bg)
        .border_1()
        .border_color(border)
        .rounded_lg()
        .shadow_lg()
        .text_color(text)
        .text_size(px(13.))
        .overflow_hidden()
        .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        .on_scroll_wheel(|_, _, cx| cx.stop_propagation())
        .child(
            div()
                .p_3()
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
                                .text_size(px(11.))
                                .text_color(muted)
                                .child(format!("Comment {} of {}", reader.ordinal, reader.total)),
                        )
                        .child(
                            div()
                                .id("comment-reader-close")
                                .px_2()
                                .cursor_pointer()
                                .text_color(muted)
                                .child("×")
                                .on_click(cx.listener(|this, _, _, cx| {
                                    this.comment_reader = None;
                                    cx.notify();
                                })),
                        ),
                )
                .child(div().font_weight(FontWeight::SEMIBOLD).child(location))
                .when(!comment.author.is_empty(), |d| {
                    d.child(
                        div()
                            .text_size(px(12.))
                            .text_color(muted)
                            .child(comment.author.clone()),
                    )
                }),
        )
        .when(reader.hidden, |d| {
            d.child(
                div()
                    .px_3()
                    .pb_2()
                    .text_size(px(11.))
                    .text_color(muted)
                    .child("This cell is hidden or filtered. Its visibility is unchanged."),
            )
        })
        .child(
            div()
                .id(("comment-reader-text", reader.ordinal))
                .px_3()
                .pb_3()
                .max_h(
                    px(200.).min(
                        (window.viewport_size().height
                            - px(app.grid_layout.grid_body_origin.1 + 180.))
                        .max(px(40.)),
                    ),
                )
                .overflow_y_scroll()
                .child(comment.text.clone()),
        )
        .child(
            div()
                .border_t_1()
                .border_color(border)
                .px_3()
                .py_2()
                .flex()
                .items_center()
                .justify_between()
                .child(
                    div()
                        .id("comment-reader-edit")
                        .text_size(px(12.))
                        .text_color(accent)
                        .when(reader.hidden, |d| d.text_color(muted))
                        .child(if reader.hidden { "Reveal cell to edit" } else { "Edit comment" })
                        .when(!reader.hidden, |d| d.cursor_pointer()
                            .hover(|d| d.underline())
                            .on_click(cx.listener(move |this, _, window, cx| {
                                this.open_cell_comment(row, col, window, cx)
                            }))),
                )
                .child(
                    div()
                        .flex()
                        .gap_2()
                        .child(
                            crate::ui::Button::new("comment-reader-prev", "Previous")
                                .secondary(border, text)
                                .on_click(cx.listener(|this, _, window, cx| {
                                    this.navigate_comment(false, window, cx)
                                })),
                        )
                        .child(
                            crate::ui::Button::new("comment-reader-next", "Next")
                                .secondary(border, text)
                                .on_click(cx.listener(|this, _, window, cx| {
                                    this.navigate_comment(true, window, cx)
                                })),
                        ),
                ),
        )
}

#[cfg(test)]
mod navigation_tests {
    use super::comment_step;
    #[test]
    fn navigation_orders_cells_and_wraps_across_sheets() {
        let cells = [(0, 2, 1), (0, 7, 5), (0, 7, 6), (1, 0, 0)];
        assert_eq!(comment_step(&cells, (0, 0, 0), true), Some(0));
        assert_eq!(comment_step(&cells, (0, 2, 1), true), Some(1));
        assert_eq!(comment_step(&cells, (0, 7, 5), true), Some(2));
        assert_eq!(comment_step(&cells, (0, 7, 6), true), Some(3));
        assert_eq!(comment_step(&cells, (1, 0, 0), true), Some(0));
        assert_eq!(comment_step(&cells, (0, 2, 1), false), Some(3));
        assert_eq!(comment_step(&cells, (1, 0, 0), false), Some(2));
        assert_eq!(comment_step(&cells, (0, 5, 0), false), Some(0));
    }
    #[test]
    fn navigation_handles_empty_and_single_comment_workbooks() {
        assert_eq!(comment_step(&[], (0, 0, 0), true), None);
        assert_eq!(comment_step(&[], (0, 0, 0), false), None);
        for forward in [true, false] {
            assert_eq!(comment_step(&[(0, 0, 0)], (0, 0, 0), forward), Some(0));
        }
    }
}
