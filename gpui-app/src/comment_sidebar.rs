//! Searchable workbook comments; the search field owns keyboard/IME input.
use crate::{
    app::Spreadsheet,
    ime::{byte_range_to_utf16, utf16_range_to_byte, ImeBuffer},
    mode::Mode,
    theme::TokenKey,
};
use gpui::{prelude::FluentBuilder, *};
use std::{cell::RefCell, ops::Range};

pub const WIDTH: f32 = 340.;

#[derive(Clone)]
pub(crate) struct Entry {
    sheet: usize,
    row: usize,
    col: usize,
    sheet_name: String,
    address: String,
    text: String,
    author: String,
    ordinal: usize,
}

fn matches(entry: &Entry, query: &str) -> bool {
    let haystack = format!(
        "{} {} {} {}",
        entry.sheet_name, entry.address, entry.author, entry.text
    )
    .to_lowercase();
    query
        .split_whitespace()
        .all(|term| haystack.contains(&term.to_lowercase()))
}

fn entries(app: &Spreadsheet, cx: &App) -> (Vec<Entry>, usize) {
    let query = &app.comment_search.read(cx).buffer.text;
    collect_entries(
        app.wb(cx),
        app.sheet_index(cx),
        app.comments_all_sheets,
        query,
    )
}

fn collect_entries(
    workbook: &visigrid_engine::workbook::Workbook,
    current_sheet: usize,
    all_sheets: bool,
    query: &str,
) -> (Vec<Entry>, usize) {
    let mut all = Vec::new();
    for (sheet, data) in workbook.sheets().iter().enumerate() {
        let mut comments: Vec<_> = data.comments().collect();
        comments.sort_unstable_by_key(|(pos, _)| *pos);
        for ((row, col), comment) in comments {
            all.push(Entry {
                sheet,
                row,
                col,
                sheet_name: data.name.clone(),
                address: format!("{}{}", Spreadsheet::col_letter(col), row + 1),
                text: comment.text.clone(),
                author: comment.author.clone(),
                ordinal: all.len() + 1,
            });
        }
    }
    let total = all.len();
    all.retain(|e| (all_sheets || e.sheet == current_sheet) && matches(e, query));
    (all, total)
}

impl Spreadsheet {
    pub fn toggle_comments_sidebar(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if !self.mode.is_navigation() && self.mode != Mode::Command {
            return;
        }
        self.mode = Mode::Navigation;
        self.comments_sidebar_visible = !self.comments_sidebar_visible;
        self.comment_reader = None;
        if self.comments_sidebar_visible {
            self.inspector_visible = false;
            self.profiler_visible = false;
            window.focus(&self.comment_search.read(cx).focus.clone(), cx);
        } else {
            window.focus(&self.focus_handle, cx);
        }
        cx.notify();
    }

    fn jump_first_comment_result(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let (results, total) = entries(self, cx);
        if let Some(e) = results.first() {
            self.jump_to_comment(e.sheet, e.row, e.col, e.ordinal, total, window, cx);
        }
    }
}

pub fn render(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let bg = app.token(TokenKey::PanelBg);
    let border = app.token(TokenKey::PanelBorder);
    let text = app.token(TokenKey::TextPrimary);
    let muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let (results, total) = entries(app, cx);
    let has_query = !app.comment_search.read(cx).buffer.text.trim().is_empty();
    let count = results.len();
    let (vr, selected_col) = app.view_state.selected;
    let selected_row = app.view_to_data(vr, cx);
    let selected_sheet = app.sheet_index(cx);
    div()
        .id("comments-sidebar")
        .w(px(WIDTH))
        .flex_shrink_0()
        .h_full()
        .min_h(px(0.))
        .flex()
        .flex_col()
        .bg(bg)
        .border_l_1()
        .border_color(border)
        .text_color(text)
        .text_size(px(13.))
        .on_mouse_down(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        .on_mouse_up(MouseButton::Left, |_, _, cx| cx.stop_propagation())
        .on_scroll_wheel(|_, _, cx| cx.stop_propagation())
        .child(
            div()
                .px_4()
                .pt_4()
                .pb_3()
                .flex()
                .flex_col()
                .gap_3()
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .text_size(px(16.))
                                .font_weight(FontWeight::SEMIBOLD)
                                .child("Comments"),
                        )
                        .child(
                            div()
                                .id("comments-sidebar-close")
                                .px_2()
                                .cursor_pointer()
                                .text_color(muted)
                                .hover(|d| d.text_color(text))
                                .on_click(cx.listener(|this, _, window, cx| {
                                    this.toggle_comments_sidebar(window, cx)
                                }))
                                .child("×"),
                        ),
                )
                .child(app.comment_search.clone())
                .child(
                    div()
                        .flex()
                        .gap_1()
                        .p_1()
                        .rounded_md()
                        .bg(app.token(TokenKey::AppBg))
                        .children(
                            [(false, "This sheet"), (true, "All sheets")]
                                .into_iter()
                                .map(|(all, label)| {
                                    div()
                                        .id(if all {
                                            "comments-all-sheets"
                                        } else {
                                            "comments-this-sheet"
                                        })
                                        .flex_1()
                                        .py_1()
                                        .text_center()
                                        .rounded_sm()
                                        .cursor_pointer()
                                        .text_color(if app.comments_all_sheets == all {
                                            accent
                                        } else {
                                            muted
                                        })
                                        .when(app.comments_all_sheets == all, |d| {
                                            d.bg(bg).shadow_sm()
                                        })
                                        .on_click(cx.listener(move |this, _, _, cx| {
                                            this.comments_all_sheets = all;
                                            this.comment_list_scroll
                                                .set_offset(point(px(0.), px(0.)));
                                            cx.notify();
                                        }))
                                        .child(label)
                                }),
                        ),
                )
                .child(
                    div()
                        .flex()
                        .justify_between()
                        .items_center()
                        .child(div().text_size(px(11.)).text_color(muted).child(format!(
                            "{count} {}",
                            if has_query {
                                if count == 1 {
                                    "result"
                                } else {
                                    "results"
                                }
                            } else if count == 1 {
                                "comment"
                            } else {
                                "comments"
                            }
                        )))
                        .child(
                            div()
                                .id("comments-add-edit")
                                .cursor_pointer()
                                .text_color(accent)
                                .text_size(px(12.))
                                .hover(|d| d.underline())
                                .on_click(
                                    cx.listener(|this, _, window, cx| {
                                        this.open_comment(window, cx)
                                    }),
                                )
                                .child("Add / edit selected"),
                        ),
                ),
        )
        .child(
            div()
                .id("comments-results")
                .flex_1()
                .min_h(px(0.))
                .overflow_y_scroll()
                .track_scroll(&app.comment_list_scroll)
                .px_3()
                .pb_3()
                .when(results.is_empty(), |d| {
                    d.child(
                        div()
                            .px_3()
                            .py_6()
                            .flex()
                            .flex_col()
                            .gap_2()
                            .child(div().font_weight(FontWeight::SEMIBOLD).child(if has_query {
                                "No matching comments"
                            } else {
                                "No comments here yet"
                            }))
                            .child(div().text_color(muted).child(if has_query {
                                "Try a word, author, or cell address."
                            } else if !app.comments_all_sheets && total > 0 {
                                "Switch to All sheets to explore the workbook."
                            } else {
                                "Select a cell, then choose Add / edit selected."
                            }))
                            .when(has_query, |d| {
                                d.child(
                                    div()
                                        .id("comments-clear-filter")
                                        .cursor_pointer()
                                        .text_color(accent)
                                        .child("Clear search")
                                        .on_click(cx.listener(|this, _, window, cx| {
                                            this.comment_search.update(cx, |search, cx| {
                                                search.buffer = ImeBuffer::default();
                                                cx.notify();
                                            });
                                            window.focus(
                                                &this.comment_search.read(cx).focus.clone(),
                                                cx,
                                            );
                                        })),
                                )
                            }),
                    )
                })
                .children(results.into_iter().map(move |e| {
                    let active =
                        e.sheet == selected_sheet && e.row == selected_row && e.col == selected_col;
                    let mut snippet: String = e.text.chars().take(180).collect();
                    if snippet.len() < e.text.len() {
                        snippet.push('…');
                    }
                    div()
                        .id(("comment-result", e.ordinal))
                        .mb_2()
                        .p_3()
                        .flex()
                        .flex_col()
                        .gap_2()
                        .rounded_md()
                        .border_1()
                        .border_color(if active { accent.opacity(0.6) } else { border })
                        .bg(if active {
                            accent.opacity(0.07)
                        } else {
                            app.token(TokenKey::AppBg)
                        })
                        .cursor_pointer()
                        .hover(|d| d.border_color(accent))
                        .on_click(cx.listener(move |this, _, window, cx| {
                            this.jump_to_comment(
                                e.sheet, e.row, e.col, e.ordinal, total, window, cx,
                            )
                        }))
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .justify_between()
                                .gap_2()
                                .child(
                                    div()
                                        .text_size(px(11.))
                                        .text_color(muted)
                                        .overflow_hidden()
                                        .text_ellipsis()
                                        .child(e.sheet_name),
                                )
                                .child(
                                    div()
                                        .text_size(px(11.))
                                        .text_color(accent)
                                        .font_weight(FontWeight::SEMIBOLD)
                                        .child(e.address),
                                ),
                        )
                        .when(!e.author.is_empty(), |d| {
                            d.child(
                                div()
                                    .text_size(px(12.))
                                    .font_weight(FontWeight::SEMIBOLD)
                                    .child(e.author),
                            )
                        })
                        .child(div().max_h(px(72.)).overflow_hidden().child(snippet))
                })),
        )
        .child(
            div()
                .border_t_1()
                .border_color(border)
                .px_4()
                .py_2()
                .text_size(px(11.))
                .text_color(muted)
                .child("Click a comment to jump to its cell"),
        )
}

pub struct CommentSearch {
    owner: WeakEntity<Spreadsheet>,
    pub focus: FocusHandle,
    pub buffer: ImeBuffer,
    layout: RefCell<Option<TextLayout>>,
}

impl CommentSearch {
    pub fn new(owner: WeakEntity<Spreadsheet>, cx: &mut Context<Self>) -> Self {
        Self {
            owner,
            focus: cx.focus_handle(),
            buffer: ImeBuffer::default(),
            layout: RefCell::new(None),
        }
    }

    fn key(&mut self, event: &KeyDownEvent, window: &mut Window, cx: &mut Context<Self>) {
        let key = event.keystroke.key.as_str();
        let mods = event.keystroke.modifiers;
        let command = mods.platform || mods.control;
        if matches!(key, "escape" | "enter" | "tab") && self.buffer.marked.is_none() {
            let owner = self.owner.clone();
            // Defer: jumping reads this search entity; it cannot be read while
            // its input handler holds the mutable entity borrow.
            let key = key.to_owned();
            window.defer(cx, move |window, cx| {
                let _ = owner.update(cx, |app, cx| match key.as_str() {
                    "escape" => app.toggle_comments_sidebar(window, cx),
                    "enter" => app.jump_first_comment_result(window, cx),
                    _ => window.focus(&app.focus_handle, cx),
                });
            });
            cx.stop_propagation();
            return;
        }
        let buf = &mut self.buffer;
        match key {
            "a" if command => {
                buf.selection_anchor = Some(0);
                buf.cursor = buf.text.len();
            }
            "c" | "x" if command => {
                let r = buf.selection();
                if !r.is_empty() {
                    cx.write_to_clipboard(ClipboardItem::new_string(buf.text[r].to_owned()));
                    if key == "x" {
                        buf.replace(None, "");
                    }
                }
            }
            "v" if command => {
                if let Some(text) = cx.read_from_clipboard().and_then(|c| c.text()) {
                    buf.replace(None, &text.replace(['\r', '\n'], " "));
                }
            }
            "backspace" | "delete" => {
                if buf.selection().is_empty() {
                    let other = if key == "backspace" {
                        buf.text[..buf.cursor]
                            .char_indices()
                            .next_back()
                            .map_or(0, |(i, _)| i)
                    } else {
                        buf.text[buf.cursor..]
                            .chars()
                            .next()
                            .map_or(buf.cursor, |c| buf.cursor + c.len_utf8())
                    };
                    buf.selection_anchor = Some(other);
                }
                buf.replace(None, "");
            }
            "left" | "right" | "home" | "end" => {
                let old = buf.cursor;
                let sel = buf.selection();
                buf.cursor = match key {
                    "home" => 0,
                    "end" => buf.text.len(),
                    "left" if !mods.shift && !sel.is_empty() => sel.start,
                    "right" if !mods.shift && !sel.is_empty() => sel.end,
                    "left" => buf.text[..old]
                        .char_indices()
                        .next_back()
                        .map_or(0, |(i, _)| i),
                    _ => buf.text[old..]
                        .chars()
                        .next()
                        .map_or(old, |c| old + c.len_utf8()),
                };
                buf.selection_anchor = if mods.shift {
                    Some(buf.selection_anchor.unwrap_or(old))
                } else {
                    None
                };
                buf.marked = None;
            }
            _ if !command => {
                cx.propagate();
                return;
            }
            _ => {}
        }
        cx.stop_propagation();
        cx.notify();
    }
}

impl Render for CommentSearch {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let Some(owner) = self.owner.upgrade() else {
            return div().into_any_element();
        };
        let app = owner.read(cx);
        let text = app.token(TokenKey::TextPrimary);
        let muted = app.token(TokenKey::TextMuted);
        let accent = app.token(TokenKey::Accent);
        let bg = app.token(TokenKey::AppBg);
        let border = app.token(TokenKey::PanelBorder);
        let focused = self.focus.is_focused(window);
        let mut style = window.text_style();
        style.color = text;
        let mut highlights = Vec::new();
        if !self.buffer.selection().is_empty() && focused {
            highlights.push((
                self.buffer.selection(),
                HighlightStyle {
                    background_color: Some(app.token(TokenKey::SelectionBg)),
                    ..Default::default()
                },
            ));
        }
        if let Some(r) = &self.buffer.marked {
            highlights.push((
                r.clone(),
                HighlightStyle {
                    underline: Some(UnderlineStyle {
                        thickness: px(1.),
                        ..Default::default()
                    }),
                    ..Default::default()
                },
            ));
        }
        let styled = StyledText::new(format!("{} ", self.buffer.text))
            .with_default_highlights(&style, highlights);
        let layout = styled.layout().clone();
        *self.layout.borrow_mut() = Some(layout.clone());
        let click_layout = layout.clone();
        let cursor = self.buffer.cursor;
        let focus = self.focus.clone();
        let entity = cx.entity();
        div()
            .id("comment-search")
            .relative()
            .track_focus(&self.focus)
            .key_context("CommentSearch")
            .px_3()
            .pr(px(30.))
            .py_2()
            .h(px(36.))
            .overflow_hidden()
            .rounded_md()
            .border_1()
            .border_color(if focused { accent } else { border })
            .bg(bg)
            .text_size(px(13.))
            .cursor_text()
            .on_key_down(cx.listener(Self::key))
            .on_mouse_down(
                MouseButton::Left,
                cx.listener(move |this, event: &MouseDownEvent, window, cx| {
                    window.focus(&this.focus, cx);
                    let i = click_layout
                        .index_for_position(event.position)
                        .unwrap_or_else(|i| i)
                        .min(this.buffer.text.len());
                    let i = this.buffer.text.floor_char_boundary(i);
                    this.buffer.selection_anchor = if event.modifiers.shift {
                        Some(this.buffer.selection_anchor.unwrap_or(this.buffer.cursor))
                    } else {
                        None
                    };
                    this.buffer.cursor = i;
                    this.buffer.marked = None;
                    cx.stop_propagation();
                    cx.notify();
                }),
            )
            .child(styled)
            .when(self.buffer.text.is_empty(), |d| {
                d.child(
                    div()
                        .absolute()
                        .left(px(12.))
                        .top(px(8.))
                        .text_color(muted)
                        .child("Search text, author, or cell…"),
                )
            })
            .child(
                canvas(
                    |_, _, _| (),
                    move |bounds, _, window, cx| {
                        window.handle_input(&focus, ElementInputHandler::new(bounds, entity), cx);
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
            .when(!self.buffer.text.is_empty(), |d| {
                d.child(
                    div()
                        .id("comment-search-clear")
                        .absolute()
                        .right(px(6.))
                        .top(px(7.))
                        .px_1()
                        .bg(bg)
                        .text_color(muted)
                        .cursor_pointer()
                        .hover(|d| d.text_color(text))
                        .child("×")
                        .on_mouse_down(
                            MouseButton::Left,
                            cx.listener(|this, _, window, cx| {
                                this.buffer = ImeBuffer::default();
                                window.focus(&this.focus, cx);
                                cx.stop_propagation();
                                cx.notify();
                            }),
                        ),
                )
            })
            .into_any_element()
    }
}

impl EntityInputHandler for CommentSearch {
    fn accepts_text_input(&self, _: &mut Window, _: &mut Context<Self>) -> bool {
        true
    }
    fn text_for_range(
        &mut self,
        r: Range<usize>,
        adjusted: &mut Option<Range<usize>>,
        _: &mut Window,
        _: &mut Context<Self>,
    ) -> Option<String> {
        let r = utf16_range_to_byte(&self.buffer.text, r);
        *adjusted = Some(byte_range_to_utf16(&self.buffer.text, r.clone()));
        Some(self.buffer.text[r].to_owned())
    }
    fn selected_text_range(
        &mut self,
        _: bool,
        _: &mut Window,
        _: &mut Context<Self>,
    ) -> Option<UTF16Selection> {
        let r = self.buffer.selection();
        Some(UTF16Selection {
            reversed: self.buffer.cursor == r.start,
            range: byte_range_to_utf16(&self.buffer.text, r),
        })
    }
    fn marked_text_range(&self, _: &mut Window, _: &mut Context<Self>) -> Option<Range<usize>> {
        self.buffer
            .marked
            .clone()
            .map(|r| byte_range_to_utf16(&self.buffer.text, r))
    }
    fn unmark_text(&mut self, _: &mut Window, cx: &mut Context<Self>) {
        self.buffer.marked = None;
        cx.notify();
    }
    fn replace_text_in_range(
        &mut self,
        r: Option<Range<usize>>,
        text: &str,
        _: &mut Window,
        cx: &mut Context<Self>,
    ) {
        let r = r.map(|r| utf16_range_to_byte(&self.buffer.text, r));
        self.buffer.replace(r, &text.replace(['\r', '\n'], " "));
        cx.notify();
    }
    fn replace_and_mark_text_in_range(
        &mut self,
        r: Option<Range<usize>>,
        text: &str,
        selected: Option<Range<usize>>,
        _: &mut Window,
        cx: &mut Context<Self>,
    ) {
        let r = r.map(|r| utf16_range_to_byte(&self.buffer.text, r));
        self.buffer.replace_and_mark(r, text, selected);
        cx.notify();
    }
    fn bounds_for_range(
        &mut self,
        _: Range<usize>,
        bounds: Bounds<Pixels>,
        _: &mut Window,
        _: &mut Context<Self>,
    ) -> Option<Bounds<Pixels>> {
        Some(
            self.layout
                .borrow()
                .as_ref()
                .and_then(|l| {
                    l.position_for_index(self.buffer.cursor)
                        .map(|p| Bounds::new(p, size(px(1.), l.line_height())))
                })
                .unwrap_or(bounds),
        )
    }
    fn character_index_for_point(
        &mut self,
        _: Point<Pixels>,
        _: &mut Window,
        _: &mut Context<Self>,
    ) -> Option<usize> {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::{collect_entries, matches, Entry};
    #[test]
    fn filters_current_sheet_without_changing_workbook_order_or_counts() {
        use visigrid_engine::{cell::CellComment, workbook::Workbook};
        let mut wb = Workbook::new();
        wb.add_sheet();
        for (sheet, row, col) in [(1, 0, 0), (0, 80, 12), (0, 2, 1)] {
            wb.sheet_mut(sheet).unwrap().set_comment(
                row,
                col,
                Some(CellComment {
                    text: "Review amount".into(),
                    author: "QA".into(),
                }),
            );
        }
        let (all, total) = collect_entries(&wb, 1, true, "review");
        assert_eq!(total, 3);
        assert_eq!(
            all.iter()
                .map(|e| (e.sheet, e.row, e.col, e.ordinal))
                .collect::<Vec<_>>(),
            vec![(0, 2, 1, 1), (0, 80, 12, 2), (1, 0, 0, 3)]
        );
        let (current, _) = collect_entries(&wb, 1, false, "");
        assert_eq!(current.len(), 1);
        assert_eq!(current[0].ordinal, 3);
        assert!(collect_entries(&wb, 1, false, "B3").0.is_empty());
        assert_eq!(collect_entries(&wb, 0, false, "B3 QA").0.len(), 1);
        wb.sheet_mut(0).unwrap().set_comment(2, 1, None);
        assert!(collect_entries(&wb, 0, true, "B3").0.is_empty());
        assert_eq!(collect_entries(&wb, 0, true, "").1, 2);
    }
    #[test]
    fn searches_text_author_address_and_sheet_with_all_terms() {
        let e = Entry {
            sheet: 1,
            row: 2,
            col: 1,
            sheet_name: "Other sheet".into(),
            address: "B3".into(),
            text: "Check café 日本語 amount".into(),
            author: "Zelda".into(),
            ordinal: 1,
        };
        for query in [
            "",
            "  ",
            "zelDA",
            "b3",
            "日本語",
            "OTHER amount",
            "café Zelda",
        ] {
            assert!(matches(&e, query), "{query}");
        }
        for query in ["missing", "B3 missing", "A1"] {
            assert!(!matches(&e, query), "{query}");
        }
    }
}
