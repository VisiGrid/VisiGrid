//! Conditional formatting quick-add: the menuless authoring surface.
//!
//! Select a range, invoke "Add Conditional Format" (command palette or
//! Format menu), type a rule, Enter. The rule text is:
//!
//!     =PREDICATE -> STYLE
//!
//! The target range is the selection at the moment the dialog opened.
//! The predicate is a formula anchored at the selection's top-left cell
//! (relative refs shift per cell, `$` refs don't — Excel semantics).
//!
//! STYLE is one of:
//!   - a named style:  error | warning | success | input | total | note
//!   - inline properties, comma-separated:
//!       bold, italic, underline, strikethrough,
//!       bg=#RRGGBB, fg=#RRGGBB
//!   - like(A1) — copy the explicit formatting of a template cell
//!     (snapshotted now, not tracked live)
//!
//! Examples:
//!     =A1>100 -> warning
//!     =$C1="overdue" -> bold, fg=#B71C1C, bg=#FDE2E2
//!     =ISBLANK(A1) -> like(Z1)

use gpui::Context;

use visigrid_engine::cell::{CellFormat, CellFormatOverride, CellStyle};
use visigrid_engine::cond_format::{CondFormatRule, CondStyle};
use visigrid_engine::validation::CellRange;

use crate::app::Spreadsheet;
use crate::mode::Mode;

// ============================================================================
// Rule text parsing
// ============================================================================

/// Parse the style half of a rule ("warning", "bold, bg=#FFEB3B", "like(Z1)").
/// `template_lookup` resolves a like() reference to the cell's explicit format.
pub fn parse_style(
    text: &str,
    template_lookup: &dyn Fn(usize, usize) -> CellFormat,
) -> Result<CondStyle, String> {
    let text = text.trim();
    if text.is_empty() {
        return Err("Missing style after '->'".into());
    }

    // Named styles ("good"/"bad"/"neutral" are Excel's preset vocabulary)
    let named = match text.to_ascii_lowercase().as_str() {
        "error" | "bad" | "red" => Some(CellStyle::Error),
        "warning" | "neutral" | "yellow" => Some(CellStyle::Warning),
        "success" | "good" | "green" => Some(CellStyle::Success),
        "input" => Some(CellStyle::Input),
        "total" => Some(CellStyle::Total),
        "note" => Some(CellStyle::Note),
        _ => None,
    };
    if let Some(style) = named {
        return Ok(CondStyle::Named(style));
    }

    // like(REF)
    let lower = text.to_ascii_lowercase();
    if lower.starts_with("like(") && text.ends_with(')') {
        let inner = &text[5..text.len() - 1];
        let (row, col) = parse_cell_ref(inner.trim())
            .ok_or_else(|| format!("Invalid cell reference in like(): '{}'", inner.trim()))?;
        let format = template_lookup(row, col);
        let snapshot = override_from_explicit(&format);
        if snapshot == CellFormatOverride::default() {
            return Err(format!("Template cell {} has no explicit formatting", inner.trim()));
        }
        return Ok(CondStyle::Like { source: (row, col), snapshot });
    }

    // Inline property list
    let mut ov = CellFormatOverride::default();
    for part in text.split(',') {
        let part = part.trim();
        if part.is_empty() {
            continue;
        }
        let lower = part.to_ascii_lowercase();
        match lower.as_str() {
            "bold" => ov.bold = Some(true),
            "italic" => ov.italic = Some(true),
            "underline" => ov.underline = Some(true),
            "strikethrough" => ov.strikethrough = Some(true),
            _ => {
                if let Some(hex) = lower.strip_prefix("bg=") {
                    ov.background_color = Some(Some(parse_hex_color(hex)?));
                } else if let Some(hex) = lower.strip_prefix("fg=") {
                    ov.font_color = Some(Some(parse_hex_color(hex)?));
                } else {
                    return Err(format!(
                        "Unknown style '{}' — use a named style (warning, error, …), \
                         properties (bold, bg=#RRGGBB, fg=#RRGGBB), or like(A1)",
                        part
                    ));
                }
            }
        }
    }
    if ov == CellFormatOverride::default() {
        return Err("Style produced no formatting".into());
    }
    Ok(CondStyle::Inline(ov))
}

/// Parse a full rule line: "=PREDICATE -> STYLE".
/// Returns (predicate, style).
pub fn parse_rule_input(
    input: &str,
    template_lookup: &dyn Fn(usize, usize) -> CellFormat,
) -> Result<(String, CondStyle), String> {
    let input = input.trim();
    let Some((pred, style_text)) = input.rsplit_once("->") else {
        return Err("Expected: =PREDICATE -> STYLE  (e.g. =A1>100 -> warning)".into());
    };
    let pred = pred.trim();
    if pred.is_empty() {
        return Err("Missing predicate before '->'".into());
    }
    let pred = if pred.starts_with('=') {
        pred.to_string()
    } else {
        format!("={}", pred)
    };
    // Validate the predicate parses as a formula now, so errors surface in
    // the dialog instead of a silently-inert rule.
    visigrid_engine::formula::parser::parse(&pred)
        .map_err(|e| format!("Predicate error: {}", e))?;
    let style = parse_style(style_text, template_lookup)?;
    Ok((pred, style))
}

/// "#RRGGBB" or "RRGGBB" → RGBA (alpha 255).
fn parse_hex_color(hex: &str) -> Result<[u8; 4], String> {
    let hex = hex.trim().trim_start_matches('#');
    if hex.len() != 6 || !hex.chars().all(|c| c.is_ascii_hexdigit()) {
        return Err(format!("Invalid color '#{}' — expected #RRGGBB", hex));
    }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap();
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap();
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap();
    Ok([r, g, b, 255])
}

/// Parse an A1-style reference ("Z1", "$Z$1") into (row, col).
fn parse_cell_ref(s: &str) -> Option<(usize, usize)> {
    let s = s.replace('$', "");
    let split = s.find(|c: char| c.is_ascii_digit())?;
    let (col_part, row_part) = s.split_at(split);
    if col_part.is_empty() || row_part.is_empty() {
        return None;
    }
    let mut col = 0usize;
    for c in col_part.chars() {
        let c = c.to_ascii_uppercase();
        if !c.is_ascii_uppercase() {
            return None;
        }
        col = col * 26 + (c as usize - 'A' as usize + 1);
    }
    let row: usize = row_part.parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((row - 1, col - 1))
}

/// Explicit (non-default) format properties of a cell, as an override.
/// This is the like() snapshot: only what the user actually set.
fn override_from_explicit(format: &CellFormat) -> CellFormatOverride {
    let default = CellFormat::default();
    let mut ov = CellFormatOverride::default();
    macro_rules! diff {
        ($field:ident) => {
            if format.$field != default.$field {
                ov.$field = Some(format.$field.clone());
            }
        };
    }
    diff!(bold);
    diff!(italic);
    diff!(underline);
    diff!(strikethrough);
    diff!(alignment);
    diff!(vertical_alignment);
    diff!(number_format);
    diff!(font_family);
    diff!(font_size);
    diff!(font_color);
    diff!(background_color);
    diff!(cell_style);
    ov
}

// ============================================================================
// Dialog state + commit (Spreadsheet impl)
// ============================================================================

#[path = "cond_format_plan.rs"]
pub(crate) mod plan;

impl Spreadsheet {
    fn cf_selection_targets(
        &mut self,
        cx: &mut Context<Self>,
    ) -> Result<(Vec<CellRange>, (usize, usize)), String> {
        self.sync_table_view(cx);
        if self.wb(cx).has_table_criteria()
            || self.row_view.is_sorted()
            || self.row_view.is_filtered()
            || self.display_hidden_rows().is_some_and(|h| !h.is_empty())
            || self
                .hidden_cols
                .get(&self.cached_sheet_id())
                .is_some_and(|h| !h.is_empty())
        {
            self.validate_saved_view_layout(self.wb(cx))?;
            plan::targets(
                self.sheet(cx),
                &self.row_view,
                self.display_hidden_rows(),
                self.hidden_cols.get(&self.cached_sheet_id()),
                &self.all_selection_ranges(),
            )
        } else {
            let ((start_row, start_col), (end_row, end_col)) = self.selection_range();
            Ok((
                vec![CellRange {
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                }],
                (start_row, start_col),
            ))
        }
    }

    pub fn show_add_cond_format(&mut self, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        if self.mode.is_editing() {
            return;
        }
        let (ranges, anchor) = match self.cf_selection_targets(cx) {
            Ok(targets) => targets,
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                return;
            }
        };
        self.cf_target = ranges.clone();
        self.cf_draft = Some(plan::Draft::new(
            self.wb(cx),
            self.sheet_index(cx),
            ranges,
            anchor,
            None,
        ));
        self.cf_input.clear();
        self.cf_input_error = None;
        self.cf_preview_matches = None;
        self.mode = Mode::AddCondFormat;
        self.bump_cf_rules_rev();
        cx.notify();
    }

    pub fn hide_add_cond_format(&mut self, cx: &mut Context<Self>) {
        // Cancellation only discards private state, even if the workbook changed.
        self.cf_draft = None;
        self.bump_cf_rules_rev();
        self.cf_preview_matches = None;
        self.mode = Mode::Navigation;
        self.cf_input.clear();
        self.cf_input_error = None;
        cx.notify();
    }

    pub fn cf_input_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.cf_input.push(c);
        self.update_cf_preview(cx);
        cx.notify();
    }

    pub fn cf_input_backspace(&mut self, cx: &mut Context<Self>) {
        self.cf_input.pop();
        self.update_cf_preview(cx);
        cx.notify();
    }

    /// Draft rules are visible only to rendering; saves and session readers see
    /// the committed store until confirm publishes the complete replacement.
    fn update_cf_preview(&mut self, cx: &mut Context<Self>) {
        if let Some(draft) = &mut self.cf_draft {
            draft.preview = None;
        }
        self.cf_preview_matches = None;
        self.cf_input_error = None;
        self.bump_cf_rules_rev();
        let result = (|| {
            let draft = self
                .cf_draft
                .as_ref()
                .ok_or("Reopen the conditional-format editor.")?;
            let index = draft.validate(self.wb(cx))?;
            let sheet = self.wb(cx).sheet(index).unwrap();
            let (predicate, style) =
                parse_rule_input(&self.cf_input, &|r, c| sheet.get_format(r, c))?;
            let store = draft.build(&predicate, style)?;
            for rule in store.iter() {
                plan::validate_rule(sheet, rule)?;
            }
            Ok::<_, String>(store)
        })();
        match result {
            Ok(store) => {
                let draft = self.cf_draft.as_ref().unwrap();
                let edited_id = draft.editing.as_ref().map(|(_, r)| r.id);
                let mut scanned = 0;
                let mut matching = 0;
                'scan: for rule in store
                    .iter()
                    .filter(|r| edited_id == Some(r.id) || draft.before.get(r.id).is_none())
                {
                    for range in &rule.ranges {
                        for row in range.start_row..=range.end_row {
                            for col in range.start_col..=range.end_col {
                                if scanned >= 10_000 {
                                    break 'scan;
                                }
                                scanned += 1;
                                if rule.matches(row, col, self.sheet(cx)) {
                                    matching += 1;
                                }
                            }
                        }
                    }
                }
                self.cf_preview_matches = Some((matching, scanned));
                self.cf_draft.as_mut().unwrap().preview = Some(store);
            }
            Err(error) => {
                if !self.cf_input.trim().is_empty() {
                    self.cf_input_error = Some(error);
                }
            }
        }
    }

    pub fn confirm_add_cond_format(&mut self, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        self.update_cf_preview(cx);
        let Some(draft) = self.cf_draft.as_ref() else {
            return;
        };
        let Some(store) = draft.preview.clone() else {
            if self.cf_input_error.is_none() {
                self.cf_input_error = Some("Enter a formula and style.".into());
            }
            cx.notify();
            return;
        };
        let editing = draft.editing.is_some();
        let before = draft.before.iter().cloned().collect();
        let index = self.sheet_index(cx);
        if !store.iter().eq(self.sheet(cx).cond_formats.iter()) {
            self.publish_cf_store(
                index,
                store,
                before,
                if editing {
                    "Edit conditional format"
                } else {
                    "Add conditional format"
                },
                cx,
            );
        }
        self.hide_add_cond_format(cx);
    }

    pub fn clear_cond_formats_in_selection(&mut self, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        let result = self
            .cf_selection_targets(cx)
            .and_then(|(ranges, _)| plan::clear(&self.sheet(cx).cond_formats, &ranges));
        let store = match result {
            Ok(store) => store,
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                return;
            }
        };
        let before = self.cf_rules_snapshot(cx);
        if store.iter().eq(before.iter()) {
            self.status_message = Some("No conditional formats in selection".into());
        } else {
            self.publish_cf_store(
                self.sheet_index(cx),
                store,
                before,
                "Clear conditional formatting",
                cx,
            );
        }
        cx.notify();
    }

    fn publish_cf_store(
        &mut self,
        index: usize,
        store: visigrid_engine::cond_format::CondFormatStore,
        before: Vec<CondFormatRule>,
        description: &str,
        cx: &mut Context<Self>,
    ) {
        self.workbook.update(cx, |wb, _| {
            wb.sheet_mut(index).unwrap().cond_formats = store;
            wb.bump_revision_for_structure();
        });
        self.bump_cf_rules_rev();
        self.record_cf_list_change(index, before, description, cx);
        self.request_title_refresh(cx);
        self.status_message = Some(format!("{description} (Ctrl+Z to undo)"));
    }
}

/// Serialize a rule's style back to the typed syntax (for edit prefill and
/// the rules panel display).
pub fn style_to_text(style: &CondStyle) -> String {
    match style {
        CondStyle::Named(s) => match s {
            CellStyle::Error => "bad".into(),
            CellStyle::Warning => "neutral".into(),
            CellStyle::Success => "good".into(),
            CellStyle::Input => "input".into(),
            CellStyle::Total => "total".into(),
            CellStyle::Note => "note".into(),
            CellStyle::None => "note".into(),
        },
        CondStyle::Inline(ov) => {
            let mut parts: Vec<String> = Vec::new();
            if ov.bold == Some(true) { parts.push("bold".into()); }
            if ov.italic == Some(true) { parts.push("italic".into()); }
            if ov.underline == Some(true) { parts.push("underline".into()); }
            if ov.strikethrough == Some(true) { parts.push("strikethrough".into()); }
            if let Some(Some(bg)) = ov.background_color {
                parts.push(format!("bg=#{:02X}{:02X}{:02X}", bg[0], bg[1], bg[2]));
            }
            if let Some(Some(fg)) = ov.font_color {
                parts.push(format!("fg=#{:02X}{:02X}{:02X}", fg[0], fg[1], fg[2]));
            }
            if parts.is_empty() { "note".into() } else { parts.join(", ") }
        }
        CondStyle::Like { source, .. } => {
            format!("like({}{})", col_letter(source.1), source.0 + 1)
        }
    }
}

impl Spreadsheet {
    /// Snapshot-based undo for any rules-list mutation: records a Group of
    /// existing CondFormatsCleared/CondFormatAdded actions that transforms
    /// `before` into the store's current state on redo, and restores
    /// `before` (including order) on undo. No new undo variants needed.
    fn record_cf_list_change(
        &mut self,
        sheet_index: usize,
        before: Vec<CondFormatRule>,
        description: &str,
        cx: &Context<Self>,
    ) {
        let after: Vec<CondFormatRule> = self.sheet(cx).cond_formats.iter().cloned().collect();
        let mut actions = vec![crate::history::UndoAction::CondFormatsCleared {
            sheet_index,
            rules: before,
        }];
        for rule in after {
            actions.push(crate::history::UndoAction::CondFormatAdded { sheet_index, rule });
        }
        self.history.record_action_with_provenance(
            crate::history::UndoAction::Group {
                actions,
                description: description.to_string(),
            },
            None,
        );
        self.is_modified = true;
    }

    fn cf_rules_snapshot(&self, cx: &Context<Self>) -> Vec<CondFormatRule> {
        self.sheet(cx).cond_formats.iter().cloned().collect()
    }

    pub fn toggle_cf_panel(&mut self, cx: &mut Context<Self>) {
        self.cf_panel_visible = !self.cf_panel_visible;
        cx.notify();
    }

    fn change_cf_rules(
        &mut self,
        description: &str,
        cx: &mut Context<Self>,
        change: impl FnOnce(&mut visigrid_engine::cond_format::CondFormatStore),
    ) {
        let before = self.cf_rules_snapshot(cx);
        let mut store = self.sheet(cx).cond_formats.clone();
        change(&mut store);
        if store.iter().eq(before.iter()) {
            return;
        }
        for rule in store.iter() {
            if let Err(error) = plan::validate_rule(self.sheet(cx), rule) {
                self.status_message = Some(error);
                cx.notify();
                return;
            }
        }
        self.publish_cf_store(self.sheet_index(cx), store, before, description, cx);
        cx.notify();
    }

    pub fn toggle_cf_rule(&mut self, id: u64, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        self.change_cf_rules("Toggle conditional format", cx, |store| {
            if let Some(rule) = store.get_mut(id) {
                rule.enabled = !rule.enabled;
            }
        });
    }

    pub fn delete_cf_rule(&mut self, id: u64, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        self.change_cf_rules("Delete conditional format", cx, |store| {
            store.remove(id);
        });
    }

    pub fn move_cf_rule(&mut self, id: u64, delta: i32, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        self.change_cf_rules("Reorder conditional formats", cx, |store| {
            let pos = store.iter().position(|r| r.id == id);
            if let Some(pos) = pos {
                if let Some(next) = pos
                    .checked_add_signed(delta as isize)
                    .filter(|n| *n < store.len())
                {
                    store.reorder(id, next);
                }
            }
        });
    }

    pub fn edit_cf_rule(&mut self, id: u64, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        let rules = self.cf_rules_snapshot(cx);
        let Some(pos) = rules.iter().position(|r| r.id == id) else {
            return;
        };
        let rule = rules[pos].clone();
        let Some(first) = rule.ranges.first() else {
            return;
        };
        let anchor = (first.start_row, first.start_col);
        self.cf_target = rule.ranges.clone();
        self.cf_input = format!("{} -> {}", rule.predicate, style_to_text(&rule.style));
        self.cf_draft = Some(plan::Draft::new(
            self.wb(cx),
            self.sheet_index(cx),
            rule.ranges.clone(),
            anchor,
            Some((pos, rule)),
        ));
        self.mode = Mode::AddCondFormat;
        self.update_cf_preview(cx);
        cx.notify();
    }
}

pub(crate) fn format_range_label(ranges: &[CellRange]) -> String {
    let label = ranges
        .iter()
        .take(6)
        .map(|r| {
            format!(
                "{}{}:{}{}",
                col_letter(r.start_col),
                r.start_row + 1,
                col_letter(r.end_col),
                r.end_row + 1
            )
        })
        .collect::<Vec<_>>()
        .join(", ");
    if ranges.len() > 6 {
        format!("{label}, +{} ranges", ranges.len() - 6)
    } else {
        label
    }
}

fn col_letter(mut col: usize) -> String {
    let mut s = String::new();
    loop {
        s.insert(0, (b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    s
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    fn no_template(_r: usize, _c: usize) -> CellFormat {
        CellFormat::default()
    }

    #[test]
    fn parses_named_style_rule() {
        let (pred, style) = parse_rule_input("=A1>100 -> warning", &no_template).unwrap();
        assert_eq!(pred, "=A1>100");
        assert_eq!(style, CondStyle::Named(CellStyle::Warning));
    }

    #[test]
    fn parses_inline_style_rule() {
        let (pred, style) =
            parse_rule_input("A1>100 -> bold, bg=#FFEB3B, fg=#823C00", &no_template).unwrap();
        assert_eq!(pred, "=A1>100", "leading = added when omitted");
        let CondStyle::Inline(ov) = style else { panic!("expected inline") };
        assert_eq!(ov.bold, Some(true));
        assert_eq!(ov.background_color, Some(Some([255, 235, 59, 255])));
        assert_eq!(ov.font_color, Some(Some([130, 60, 0, 255])));
    }

    #[test]
    fn parses_like_rule_with_snapshot() {
        let template = |_r: usize, _c: usize| {
            let mut f = CellFormat::default();
            f.bold = true;
            f.background_color = Some([1, 2, 3, 255]);
            f
        };
        let (_, style) = parse_rule_input("=ISBLANK(A1) -> like(Z1)", &template).unwrap();
        let CondStyle::Like { source, snapshot } = style else { panic!("expected like") };
        assert_eq!(source, (0, 25));
        assert_eq!(snapshot.bold, Some(true));
        assert_eq!(snapshot.background_color, Some(Some([1, 2, 3, 255])));
        assert_eq!(snapshot.italic, None, "unset properties not snapshotted");
    }

    #[test]
    fn arrow_inside_predicate_string_is_ok() {
        // rsplit_once means a literal "->" inside the predicate text would
        // mis-split; but ">" comparisons parse fine
        let (pred, _) = parse_rule_input("=A1>-5 -> error", &no_template).unwrap();
        assert_eq!(pred, "=A1>-5");
    }

    #[test]
    fn rejects_bad_input() {
        assert!(parse_rule_input("no arrow here", &no_template).is_err());
        assert!(parse_rule_input("=A1>1 -> ", &no_template).is_err());
        assert!(parse_rule_input(" -> warning", &no_template).is_err());
        assert!(parse_rule_input("=A1>1 -> shiny", &no_template).is_err());
        assert!(parse_rule_input("=A1>1 -> bg=#XYZXYZ", &no_template).is_err());
        assert!(parse_rule_input("=SUM(( -> warning", &no_template).is_err());
    }

    #[test]
    fn cell_ref_parsing() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("Z1"), Some((0, 25)));
        assert_eq!(parse_cell_ref("AA10"), Some((9, 26)));
        assert_eq!(parse_cell_ref("$B$2"), Some((1, 1)));
        assert_eq!(parse_cell_ref("1A"), None);
        assert_eq!(parse_cell_ref(""), None);
    }
}
