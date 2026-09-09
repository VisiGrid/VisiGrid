//! IME (Input Method Editor) support for composed text entry.
//!
//! Without this, every keystroke reaches the app through `keystroke.key_char` and is
//! appended one character at a time. That works for Latin scripts but breaks every
//! input method that composes: Korean, Japanese, Chinese, and dead-key layouts.
//!
//! Typing "한글" on a 2-Set Korean keyboard used to land `ㅎㅏㄴㄱㅡㄹ` in the cell:
//! six standalone compatibility jamo (U+314E U+314F U+3134 U+3131 U+3161 U+3139)
//! instead of two composed syllables (U+D55C U+AE00). The damage is not recoverable
//! afterwards. Unicode normalization does not fix it, because compatibility jamo are
//! distinct characters from the conjoining jamo that NFC composes; NFKC makes it worse
//! by fusing them wrongly (`하ᄂ그ᄅ`), having lost the information about which jamo were
//! leading and which were trailing.
//!
//! The fix is to implement `EntityInputHandler` so the platform can drive composition:
//! it sends provisional ("marked") text while the user is mid-syllable and replaces it
//! with the committed text when the syllable closes. GPUI already exposes the whole
//! protocol; this module just connects the spreadsheet's edit buffer to it.
//!
//! Once a handler is registered, committed text arrives here for every layout, not only
//! for composing ones, so this is the typing path on macOS. Committed text therefore
//! goes through `insert_char`, which owns the formula behaviours. How the key-down path
//! and this one share the work is explained at the cell branch in `key_handler.rs`.
//!
//! Offsets are the fiddly part. The edit buffer indexes by UTF-8 byte, the platform
//! protocol indexes by UTF-16 code unit, and the two only coincide for ASCII. Every
//! boundary crossing goes through the converters below.
//!
//! The composition state machine lives in [`ImeBuffer`], deliberately free of `Window`
//! and `Context` so it can be exercised without a running app. The parts that are ours
//! rather than the platform's are tested here; the live keyboard is reserved for
//! confirming that the platform calls into this handler at all.

use gpui::{Bounds, Context, EntityInputHandler, Pixels, UTF16Selection, Window};
use std::ops::Range;

use crate::app::Spreadsheet;
use crate::mode::{InspectorTab, Mode};
use crate::rewind_state::EditorSurface;

/// UTF-8 byte offset to UTF-16 code unit offset.
///
/// Counts rather than slices so an offset inside a multi-byte character (the editor's
/// cursor can land there after autocomplete, which stores a char index) rounds down
/// instead of panicking.
fn byte_to_utf16(s: &str, byte: usize) -> usize {
    s.char_indices()
        .take_while(|(idx, ch)| idx + ch.len_utf8() <= byte)
        .map(|(_, ch)| ch.len_utf16())
        .sum()
}

/// UTF-16 code unit offset to UTF-8 byte offset.
///
/// Offsets that land inside a surrogate pair round down to the start of that character,
/// which keeps the result on a char boundary and so safe to slice with.
fn utf16_to_byte(s: &str, utf16: usize) -> usize {
    let mut seen = 0usize;
    for (byte_idx, ch) in s.char_indices() {
        if seen >= utf16 {
            return byte_idx;
        }
        let next = seen + ch.len_utf16();
        if next > utf16 {
            // The offset falls inside this character, which only happens between the two
            // halves of a surrogate pair. Round down to the character start so the result
            // stays a valid slice index.
            return byte_idx;
        }
        seen = next;
    }
    s.len()
}

fn byte_range_to_utf16(s: &str, r: Range<usize>) -> Range<usize> {
    byte_to_utf16(s, r.start)..byte_to_utf16(s, r.end)
}

fn utf16_range_to_byte(s: &str, r: Range<usize>) -> Range<usize> {
    utf16_to_byte(s, r.start)..utf16_to_byte(s, r.end)
}

/// The slice of edit state that composition touches, and the rules for mutating it.
///
/// All offsets are UTF-8 byte indices, matching the surrounding editor. Conversion to
/// and from the platform's UTF-16 offsets happens at the trait boundary below, so this
/// type never sees a UTF-16 index except for the caret hint inside marked text, which
/// the platform expresses relative to the marked run.
#[derive(Default)]
pub(crate) struct ImeBuffer {
    pub text: String,
    pub cursor: usize,
    pub selection_anchor: Option<usize>,
    /// Provisional text the input method may still revise. `None` when not composing.
    pub marked: Option<Range<usize>>,
}

impl ImeBuffer {
    fn selection(&self) -> Range<usize> {
        match self.selection_anchor {
            Some(anchor) if anchor != self.cursor => {
                anchor.min(self.cursor)..anchor.max(self.cursor)
            }
            _ => self.cursor..self.cursor,
        }
    }

    /// What the platform means by a `None` range: wherever input would go right now.
    ///
    /// During composition that is the marked run, because the next keystroke of the same
    /// syllable revises it rather than appending after it. Otherwise it is the selection,
    /// which collapses to the caret when nothing is selected.
    fn default_range(&self) -> Range<usize> {
        self.marked.clone().unwrap_or_else(|| self.selection())
    }

    /// Remove what a commit replaces and leave the caret where the committed text goes:
    /// the explicit range if the platform gave one, else the marked run. Ends any
    /// composition in progress.
    ///
    /// A plain selection is left alone: `insert_char`, which types the committed text
    /// afterwards, deletes the selection itself, and doing it there keeps the IME path
    /// and the key-down path on one implementation.
    pub(crate) fn begin_commit(&mut self, range: Option<Range<usize>>) {
        let Some(r) = range.or_else(|| self.marked.take()) else {
            return;
        };
        self.text.replace_range(r.clone(), "");
        self.cursor = r.start;
        self.selection_anchor = None;
        self.marked = None;
    }

    /// Test model of a full commit: [`Self::begin_commit`] followed by a plain insert
    /// over the selection, standing in for the app's per-character `insert_char` calls.
    #[cfg(test)]
    fn replace(&mut self, range: Option<Range<usize>>, text: &str) {
        self.begin_commit(range);
        let sel = self.selection();
        self.text.replace_range(sel.clone(), text);
        self.cursor = sel.start + text.len();
        self.selection_anchor = None;
    }

    /// Write provisional text and keep it marked so the next revision replaces it.
    ///
    /// `caret_utf16` is the platform's requested caret position measured from the start
    /// of the new text, not from the start of the buffer. Empty text means the input
    /// method abandoned the composition, which clears the mark rather than leaving a
    /// zero-length one behind.
    pub(crate) fn replace_and_mark(
        &mut self,
        range: Option<Range<usize>>,
        text: &str,
        caret_utf16: Option<Range<usize>>,
    ) {
        let r = range.unwrap_or_else(|| self.default_range());
        let start = r.start;
        self.text.replace_range(r, text);
        self.selection_anchor = None;

        if text.is_empty() {
            self.marked = None;
            self.cursor = start;
            return;
        }

        let end = start + text.len();
        self.marked = Some(start..end);
        self.cursor = match caret_utf16 {
            Some(sel) => start + utf16_to_byte(&self.text[start..end], sel.end),
            None => end,
        };
    }
}

impl Spreadsheet {
    /// The marked run, but only while it can mean anything.
    ///
    /// Some paths leave edit mode without going through `reset_edit_state` (opening
    /// Find or GoTo mid-edit, F12 to a named range), and buffer edits that bypass this
    /// module can move the bytes it points at. A range that is not inside the current
    /// buffer on char boundaries, or that belongs to an edit that has ended, is treated
    /// as no composition rather than used.
    fn ime_marked_range(&self) -> Option<Range<usize>> {
        let r = self.edit_marked_range.clone()?;
        let text = &self.edit_value;
        let valid = self.mode.is_editing()
            && r.start <= r.end
            && text.is_char_boundary(r.start)
            && text.is_char_boundary(r.end);
        valid.then_some(r)
    }

    /// Whether typed text belongs to the cell editor right now.
    ///
    /// Several text boxes live as flags on Navigation rather than as modes: sheet
    /// rename, the filter and validation dropdown searches, an open menu, KeyTips, the
    /// close-confirm dialog, the Lua console, and the surfaces with their own focus
    /// handle (terminal, format bar). Each takes printable keys in `handle_key_down`
    /// before the cell branch, and the platform still offers the IME's text to this
    /// handler in those states, so the handler has to make the same call and decline.
    /// This guards both `replace_*` entry points, which is what protects Linux, where
    /// `accepts_text_input` is not consulted.
    fn cell_owns_text_input(&self, window: &Window) -> bool {
        if self.terminal_has_focus(window)
            || self.ui.format_bar.is_active(window)
            || self.keytips_active
            || self.open_menu.is_some()
            || self.lua_console.visible
            || self.close_confirm_visible
            || self.filter_dropdown_col.is_some()
            || self.is_validation_dropdown_open()
            || self.renaming_sheet.is_some()
        {
            return false;
        }
        matches!(self.mode, Mode::Navigation | Mode::Edit | Mode::Formula)
    }

    /// Composition on a Ready cell starts an edit with an empty buffer, the way the
    /// first plain keystroke does. Returns false when the cell refused (spill receiver,
    /// preview), in which case the input is dropped rather than written into a buffer
    /// nobody will commit.
    ///
    /// `start_edit_clear` is deliberate, for the committed path too: it applies the
    /// merged-cell redirect and the preview block that the inline first-keystroke
    /// branch of `insert_char` skips, so a Latin and a composed first keystroke land
    /// in the same cell. Do not "fix" the asymmetry by routing through `insert_char`.
    fn ime_ensure_editing(&mut self, cx: &mut Context<Self>) -> bool {
        if !self.mode.is_editing() {
            self.start_edit_clear(cx);
        }
        self.mode.is_editing()
    }

    /// Lift the composition state out of the app so [`ImeBuffer`] can operate on it.
    fn ime_take_buffer(&mut self) -> ImeBuffer {
        let marked = self.ime_marked_range();
        ImeBuffer {
            text: std::mem::take(&mut self.edit_value),
            cursor: self.edit_cursor,
            selection_anchor: self.edit_selection_anchor,
            marked,
        }
    }

    fn ime_put_buffer(&mut self, buf: ImeBuffer) {
        self.edit_value = buf.text;
        self.edit_cursor = buf.cursor;
        self.edit_selection_anchor = buf.selection_anchor;
        self.edit_marked_range = buf.marked;
    }

    /// Everything the rest of the app expects after the buffer changed under it.
    /// Mirrors the tail of `insert_char`, for the paths that do not go through it.
    fn ime_after_buffer_change(&mut self, cx: &mut Context<Self>) {
        // Typing '=' first should still drop into formula mode, and a composition that
        // deletes back past the '=' should leave it.
        self.recompute_edit_mode();
        if self.mode.is_formula() {
            self.update_formula_refs();
            self.clear_formula_nav_override();
            self.update_formula_nav_mode();
        }

        self.autocomplete_suppressed = false;
        self.reset_caret_activity();
        self.edit_scroll_dirty = true;
        self.formula_bar_cache_dirty = true;
        self.update_autocomplete(cx);
        cx.notify();
    }
}

impl EntityInputHandler for Spreadsheet {
    fn accepts_text_input(&self, window: &mut Window, cx: &mut Context<Self>) -> bool {
        // On macOS this decides whether an active IME may compose a printable key
        // before the app sees it. A Ready cell has to say yes, or composition cannot
        // start on the first keystroke (Excel's behaviour, and the whole point here).
        //
        // Beyond the states where the cell does not own text at all, decline while a
        // Navigation letter is a command the key-down path must see first: vim motions,
        // the Names inspector filter, and Space with a history entry selected (hold to
        // peek). Those keys still reach the cell handler when they fall through, so
        // digits and operators in vim mode keep starting an edit.
        if !self.cell_owns_text_input(window) {
            return false;
        }
        if self.mode.is_navigation() {
            let names_filter =
                self.inspector_visible && self.inspector_tab == InspectorTab::Names;
            let history_peek =
                self.selected_history_id.is_some() && self.history_highlight_range.is_some();
            if names_filter || history_peek || self.vim_mode_enabled(cx) {
                return false;
            }
        }
        true
    }

    fn text_for_range(
        &mut self,
        range_utf16: Range<usize>,
        adjusted_range: &mut Option<Range<usize>>,
        _window: &mut Window,
        _cx: &mut Context<Self>,
    ) -> Option<String> {
        let byte_range = utf16_range_to_byte(&self.edit_value, range_utf16);
        *adjusted_range = Some(byte_range_to_utf16(&self.edit_value, byte_range.clone()));
        Some(self.edit_value[byte_range].to_string())
    }

    fn selected_text_range(
        &mut self,
        _ignore_disabled_input: bool,
        _window: &mut Window,
        _cx: &mut Context<Self>,
    ) -> Option<UTF16Selection> {
        if !self.mode.is_editing() {
            // Outside an edit the buffer is empty and a selection anchor may be stale
            // (some exits skip `reset_edit_state`), so do not consult it. An empty
            // selection at the origin, not None: None tells the platform there is no
            // text field, and some input methods then refuse to start composing.
            return Some(UTF16Selection { range: 0..0, reversed: false });
        }
        let (start, end) = self
            .edit_selection_range()
            .unwrap_or((self.edit_cursor, self.edit_cursor));
        let reversed = self.edit_cursor <= start && start != end;
        Some(UTF16Selection {
            range: byte_range_to_utf16(&self.edit_value, start..end),
            reversed,
        })
    }

    fn marked_text_range(
        &self,
        _window: &mut Window,
        _cx: &mut Context<Self>,
    ) -> Option<Range<usize>> {
        let marked = self.ime_marked_range()?;
        Some(byte_range_to_utf16(&self.edit_value, marked))
    }

    fn unmark_text(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        self.edit_marked_range = None;
        cx.notify();
    }

    fn replace_text_in_range(
        &mut self,
        range_utf16: Option<Range<usize>>,
        text: &str,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        if !self.cell_owns_text_input(window) {
            return;
        }

        // Same filter as the key-down path: control characters are commands, not text.
        let text: String = text.chars().filter(|c| !c.is_control()).collect();

        // An empty commit that replaces nothing is noise, not typing: Windows sends one
        // whenever a Japanese IME reports no composition string. It must not open an
        // edit on a Ready cell, or the next Enter would commit "" over the value.
        if text.is_empty() && range_utf16.is_none() && self.ime_marked_range().is_none() {
            return;
        }

        if !self.ime_ensure_editing(cx) {
            return;
        }

        // Clear what the commit replaces (the marked run, usually), then type the
        // replacement through insert_char so formula entry keeps its behaviours: an
        // operator finalises the pending reference, a leading `=` switches to Formula
        // mode, autocomplete suppression resets.
        let mut buf = self.ime_take_buffer();
        let range = range_utf16.map(|r| utf16_range_to_byte(&buf.text, r));
        buf.begin_commit(range);
        self.ime_put_buffer(buf);

        let mut inserted = false;
        for c in text.chars() {
            self.insert_char(c, cx);
            inserted = true;
        }
        if !inserted {
            // Nothing typed, so nothing ran the post-edit work for the removal above.
            self.ime_after_buffer_change(cx);
        }
        self.update_edit_scroll(window);
    }

    fn replace_and_mark_text_in_range(
        &mut self,
        range_utf16: Option<Range<usize>>,
        new_text: &str,
        new_selected_range_utf16: Option<Range<usize>>,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        if !self.cell_owns_text_input(window) {
            return;
        }
        // An empty preedit while not editing is a composition being abandoned before
        // it wrote anything; there is nothing to open an edit for.
        if new_text.is_empty() && !self.mode.is_editing() {
            return;
        }
        if !self.ime_ensure_editing(cx) {
            return;
        }

        // Provisional text is typing as far as reference picking is concerned: a
        // non-operator keystroke ends the pending reference in insert_char, and this
        // does the same.
        if self.mode.is_formula() {
            self.finalize_formula_reference();
        }

        let mut buf = self.ime_take_buffer();
        let range = range_utf16.map(|r| utf16_range_to_byte(&buf.text, r));
        buf.replace_and_mark(range, new_text, new_selected_range_utf16);
        self.ime_put_buffer(buf);
        self.ime_after_buffer_change(cx);
        self.update_edit_scroll(window);
    }

    fn bounds_for_range(
        &mut self,
        _range_utf16: Range<usize>,
        _element_bounds: Bounds<Pixels>,
        _window: &mut Window,
        _cx: &mut Context<Self>,
    ) -> Option<Bounds<Pixels>> {
        // Positions the input method's candidate window. `element_bounds` is useless here
        // because the handler is registered on a zero-size canvas, so the rectangle comes
        // from whichever surface is being edited. The grid case is the cell, not the
        // caret within it, which is where `open_context_menu` anchors too.
        let rect = match self.active_editor {
            EditorSurface::FormulaBar => self.formula_bar_text_rect,
            _ => {
                let cell = self.active_cell_rect();
                let (x, y) = self.grid_layout.grid_body_origin;
                Bounds {
                    origin: gpui::point(gpui::px(x + cell.x), gpui::px(y + cell.y)),
                    size: gpui::size(gpui::px(cell.width), gpui::px(cell.height)),
                }
            }
        };
        Some(rect)
    }

    fn character_index_for_point(
        &mut self,
        _point: gpui::Point<Pixels>,
        _window: &mut Window,
        _cx: &mut Context<Self>,
    ) -> Option<usize> {
        // Used for mouse-driven queries into marked text (dictionary lookup and the
        // like). Declining is safe; composition does not depend on it.
        None
    }
}

#[cfg(test)]
mod tests {
    use super::{byte_to_utf16, utf16_to_byte, ImeBuffer};

    // ---- offset conversion ------------------------------------------------------

    #[test]
    fn offsets_round_trip_for_ascii() {
        let s = "abc";
        for b in 0..=s.len() {
            assert_eq!(utf16_to_byte(s, byte_to_utf16(s, b)), b);
        }
    }

    #[test]
    fn offsets_round_trip_for_hangul() {
        // Each syllable is 3 UTF-8 bytes and 1 UTF-16 unit, so the two indices diverge.
        let s = "한글";
        assert_eq!(s.len(), 6);
        assert_eq!(byte_to_utf16(s, 6), 2);
        assert_eq!(byte_to_utf16(s, 3), 1);
        assert_eq!(utf16_to_byte(s, 1), 3);
        assert_eq!(utf16_to_byte(s, 2), 6);
    }

    #[test]
    fn offsets_round_trip_for_astral_chars() {
        // An emoji is 4 UTF-8 bytes and 2 UTF-16 units (a surrogate pair).
        let s = "a\u{1F600}b";
        assert_eq!(byte_to_utf16(s, s.len()), 4);
        // Landing inside the pair rounds down to the start of the emoji.
        assert_eq!(utf16_to_byte(s, 2), 1);
        assert_eq!(utf16_to_byte(s, 3), 5);
    }

    #[test]
    fn offsets_past_the_end_or_inside_a_char_do_not_panic() {
        let s = "한";
        assert_eq!(utf16_to_byte(s, 99), s.len());
        assert_eq!(byte_to_utf16(s, 99), 1);
        // A byte offset inside the 3-byte syllable rounds down to its start.
        assert_eq!(byte_to_utf16(s, 1), 0);
        assert_eq!(byte_to_utf16("한글", 4), 1);
    }

    // ---- composition lifecycle --------------------------------------------------
    //
    // The sequences below are what macOS sends through NSTextInputClient while a
    // 2-Set Korean keyboard composes. `None` for the range means "the marked run if
    // composing, the selection otherwise", which is why most steps pass None.

    /// Type g, k, s, r, m, f, which is the 2-Set Korean spelling of "한글".
    ///
    /// The interesting step is the fourth: pressing the leading consonant of the second
    /// syllable commits the first one and opens a new composition, so the buffer has to
    /// stop treating the old marked run as the write target.
    #[test]
    fn korean_two_syllables_compose() {
        let mut b = ImeBuffer::default();

        b.replace_and_mark(None, "ㅎ", None);
        assert_eq!(b.text, "ㅎ");
        assert_eq!(b.marked, Some(0..3));

        b.replace_and_mark(None, "하", None);
        assert_eq!(b.text, "하");
        assert_eq!(b.marked, Some(0..3));

        b.replace_and_mark(None, "한", None);
        assert_eq!(b.text, "한");

        // 'r' closes the first syllable and starts the second.
        b.replace(None, "한");
        assert_eq!(b.marked, None);
        b.replace_and_mark(None, "ㄱ", None);
        assert_eq!(b.text, "한ㄱ");
        assert_eq!(b.marked, Some(3..6));

        b.replace_and_mark(None, "그", None);
        b.replace_and_mark(None, "글", None);
        b.replace(None, "글");

        assert_eq!(b.text, "한글");
        assert_eq!(b.text.chars().count(), 2);
        assert_eq!(b.marked, None);
        assert_eq!(b.cursor, b.text.len());
    }

    /// What the app does on commit: clear the marked run, then type through
    /// `insert_char`. The caret must sit where the run was so the typed characters land
    /// in its place, and the mark must be gone so a stray commit afterwards appends
    /// instead of replacing.
    #[test]
    fn begin_commit_removes_the_marked_run_and_parks_the_caret_there() {
        let mut b = ImeBuffer {
            text: "매출 ".to_string(),
            cursor: "매출 ".len(),
            ..Default::default()
        };
        b.replace_and_mark(None, "합", None);
        assert_eq!(b.text, "매출 합");

        b.begin_commit(None);
        assert_eq!(b.text, "매출 ");
        assert_eq!(b.cursor, "매출 ".len());
        assert_eq!(b.marked, None);
        assert_eq!(b.selection_anchor, None);
    }

    /// With no composition a commit removes nothing, which is the case every ASCII
    /// keystroke takes. A selection is deliberately left in place for `insert_char`.
    #[test]
    fn begin_commit_without_composition_leaves_the_buffer_alone() {
        let mut b = ImeBuffer {
            text: "=A1+".to_string(),
            cursor: 4,
            selection_anchor: Some(1),
            ..Default::default()
        };
        b.begin_commit(None);
        assert_eq!(b.text, "=A1+");
        assert_eq!(b.cursor, 4);
        assert_eq!(b.selection_anchor, Some(1));
    }

    #[test]
    fn backspacing_mid_composition_shrinks_the_marked_run() {
        let mut b = ImeBuffer::default();
        b.replace_and_mark(None, "한", None);
        assert_eq!(b.marked, Some(0..3));

        // The input method decomposes rather than deleting a whole character.
        b.replace_and_mark(None, "하", None);
        assert_eq!(b.text, "하");
        assert_eq!(b.marked, Some(0..3));
    }

    #[test]
    fn abandoning_a_composition_clears_the_mark() {
        let mut b = ImeBuffer::default();
        b.replace_and_mark(None, "ㅎ", None);
        b.replace_and_mark(None, "", None);

        assert_eq!(b.text, "");
        assert_eq!(b.marked, None);
        assert_eq!(b.cursor, 0);
    }

    #[test]
    fn composing_after_existing_text_appends() {
        let mut b = ImeBuffer {
            text: "매출 ".to_string(),
            cursor: "매출 ".len(),
            ..Default::default()
        };
        b.replace_and_mark(None, "합", None);
        b.replace(None, "합");

        assert_eq!(b.text, "매출 합");
        assert_eq!(b.cursor, b.text.len());
    }

    #[test]
    fn composing_over_a_selection_replaces_it() {
        let mut b = ImeBuffer {
            text: "abcdef".to_string(),
            cursor: 5,
            selection_anchor: Some(1),
            ..Default::default()
        };
        b.replace_and_mark(None, "한", None);

        assert_eq!(b.text, "a한f");
        assert_eq!(b.marked, Some(1..4));
        assert_eq!(b.selection_anchor, None);
    }

    #[test]
    fn caret_hint_inside_marked_text_is_honoured() {
        let mut b = ImeBuffer::default();
        // Japanese input often parks the caret inside a longer marked run.
        b.replace_and_mark(None, "にほんご", Some(0..2));

        assert_eq!(b.text, "にほんご");
        assert_eq!(b.marked, Some(0..12));
        // Two UTF-16 units into the run is two 3-byte characters in.
        assert_eq!(b.cursor, 6);
    }

    #[test]
    fn explicit_range_overrides_the_default_target() {
        let mut b = ImeBuffer {
            text: "abcdef".to_string(),
            cursor: 6,
            ..Default::default()
        };
        b.replace(Some(1..3), "X");

        assert_eq!(b.text, "aXdef");
        assert_eq!(b.cursor, 2);
    }

}
