//! Provenance export: convert history actions to deterministic Lua scripts.
//!
//! Phase 9A: Export Provenance (MVP)
//! - Every history entry can be exported as Lua
//! - Whole history can be exported as a deterministic script
//! - Output is stable and suitable for CI/CD replay

use crate::history::{CellChange, CellFormatPatch, HistoryEntry, HistoryFingerprint, UndoAction};

/// Lua API version for provenance scripts.
/// Increment when breaking changes are made to the API surface.
pub const LUA_API_VERSION: &str = "v1";

/// Convert a cell reference (row, col) to A1 notation.
/// row and col are 0-indexed.
fn cell_ref(row: usize, col: usize) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

/// Convert 0-indexed column to letter(s): 0=A, 25=Z, 26=AA, etc.
fn col_to_letter(col: usize) -> String {
    let mut s = String::new();
    let mut n = col;
    loop {
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    s
}

/// Convert a range to A1:B2 notation.
fn range_ref(start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> String {
    if start_row == end_row && start_col == end_col {
        cell_ref(start_row, start_col)
    } else {
        format!(
            "{}:{}",
            cell_ref(start_row, start_col),
            cell_ref(end_row, end_col)
        )
    }
}

/// Stable string representation of FormatActionKind for fingerprint hashing.
fn format_kind_hash_str(kind: &crate::history::FormatActionKind) -> &'static str {
    use crate::history::FormatActionKind;
    match kind {
        FormatActionKind::Bold => "bold",
        FormatActionKind::Italic => "italic",
        FormatActionKind::Underline => "underline",
        FormatActionKind::Strikethrough => "strikethrough",
        FormatActionKind::Font => "font",
        FormatActionKind::Alignment => "alignment",
        FormatActionKind::VerticalAlignment => "valignment",
        FormatActionKind::TextOverflow => "overflow",
        FormatActionKind::NumberFormat => "numfmt",
        FormatActionKind::DecimalPlaces => "decimals",
        FormatActionKind::BackgroundColor => "bgcolor",
        FormatActionKind::FontSize => "fontsize",
        FormatActionKind::FontColor => "fontcolor",
        FormatActionKind::Border => "border",
        FormatActionKind::PasteFormats => "paste_formats",
        FormatActionKind::ClearFormatting => "clear",
        FormatActionKind::CellStyle => "cell_style",
    }
}

/// Escape a string for Lua (handle quotes, newlines, etc.)
fn lua_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if c.is_control() => {
                out.push_str(&format!("\\x{:02x}", c as u32));
            }
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

/// Generate Lua for a single UndoAction.
/// Returns None for actions that can't be represented as Lua (audit-only).
impl UndoAction {
    /// Convert this action to its Lua representation.
    /// Returns None for audit-only actions (Rewind) that are represented as comments.
    pub fn to_lua(&self) -> Option<String> {
        match self {
            UndoAction::Values { sheet_index, changes } => {
                Some(values_to_lua(*sheet_index, changes))
            }
            UndoAction::Format { sheet_index, patches, kind, .. } => {
                Some(format_to_lua(*sheet_index, patches, kind))
            }
            UndoAction::NamedRangeCreated { named_range } => {
                Some(named_range_created_to_lua(named_range))
            }
            UndoAction::NamedRangeDeleted { named_range } => {
                Some(format!("grid.undefine_name{{ name={} }}", lua_escape(&named_range.name)))
            }
            UndoAction::NamedRangeRenamed { old_name, new_name } => {
                Some(format!(
                    "grid.rename_name{{ from={}, to={} }}",
                    lua_escape(old_name),
                    lua_escape(new_name)
                ))
            }
            UndoAction::NamedRangeDescriptionChanged { name, new_description, .. } => {
                let desc = new_description.as_ref().map(|s| lua_escape(s)).unwrap_or_else(|| "nil".to_string());
                Some(format!(
                    "grid.set_name_description{{ name={}, description={} }}",
                    lua_escape(name),
                    desc
                ))
            }
            UndoAction::Group { actions, description } => {
                Some(group_to_lua(actions, description))
            }
            UndoAction::RowsInserted { sheet_index, at_row, count } => {
                Some(format!(
                    "grid.insert_rows{{ sheet={}, at={}, count={} }}",
                    sheet_index + 1,
                    at_row + 1,
                    count
                ))
            }
            UndoAction::RowsDeleted { sheet_index, at_row, count, .. } => {
                Some(format!(
                    "grid.delete_rows{{ sheet={}, at={}, count={} }}",
                    sheet_index + 1,
                    at_row + 1,
                    count
                ))
            }
            UndoAction::ColsInserted { sheet_index, at_col, count } => {
                Some(format!(
                    "grid.insert_cols{{ sheet={}, at={}, count={} }}",
                    sheet_index + 1,
                    at_col + 1,
                    count
                ))
            }
            UndoAction::ColsDeleted { sheet_index, at_col, count, .. } => {
                Some(format!(
                    "grid.delete_cols{{ sheet={}, at={}, count={} }}",
                    sheet_index + 1,
                    at_col + 1,
                    count
                ))
            }
            UndoAction::ColumnWidthSet { sheet_id, col, new, .. } => {
                // Use sheet_id (stable) instead of sheet_index (fragile)
                let col_letter = col_to_letter(*col);
                if let Some(width) = new {
                    Some(format!(
                        "grid.set_col_width{{ sheet_id={}, col=\"{}\", width={:.0} }}",
                        sheet_id.0,
                        col_letter,
                        width
                    ))
                } else {
                    Some(format!(
                        "grid.clear_col_width{{ sheet_id={}, col=\"{}\" }}",
                        sheet_id.0,
                        col_letter
                    ))
                }
            }
            UndoAction::RowHeightSet { sheet_id, row, new, .. } => {
                // Use sheet_id (stable) instead of sheet_index (fragile)
                if let Some(height) = new {
                    Some(format!(
                        "grid.set_row_height{{ sheet_id={}, row={}, height={:.0} }}",
                        sheet_id.0,
                        row + 1,
                        height
                    ))
                } else {
                    Some(format!(
                        "grid.clear_row_height{{ sheet_id={}, row={} }}",
                        sheet_id.0,
                        row + 1
                    ))
                }
            }
            UndoAction::RowVisibilityChanged { sheet_id, rows, hidden } => {
                let row_list: Vec<String> = rows.iter().map(|r| (r + 1).to_string()).collect();
                Some(format!(
                    "grid.set_row_visibility{{ sheet_id={}, rows={{{}}}, hidden={} }}",
                    sheet_id.0,
                    row_list.join(","),
                    hidden
                ))
            }
            UndoAction::ColVisibilityChanged { sheet_id, cols, hidden } => {
                let col_list: Vec<String> = cols.iter().map(|c| col_to_letter(*c)).collect();
                Some(format!(
                    "grid.set_col_visibility{{ sheet_id={}, cols={{{}}}, hidden={} }}",
                    sheet_id.0,
                    col_list.iter().map(|c| format!("\"{}\"", c)).collect::<Vec<_>>().join(","),
                    hidden
                ))
            }
            UndoAction::SortApplied { sheet_index, new_sort_state, .. } => {
                let (col, ascending) = new_sort_state;
                Some(format!(
                    "grid.sort{{ sheet={}, col={}, ascending={} }}",
                    sheet_index + 1,
                    col + 1,
                    ascending
                ))
            }
            UndoAction::SortCleared { sheet_index, .. } => {
                Some(format!(
                    "grid.clear_sort{{ sheet={} }}",
                    sheet_index + 1
                ))
            }
            UndoAction::ValidationSet { sheet_index, range, new_rule, .. } => {
                Some(validation_set_to_lua(*sheet_index, range, new_rule))
            }
            UndoAction::ValidationCleared { sheet_index, range, .. } => {
                Some(format!(
                    "grid.clear_validation{{ sheet={}, range=\"{}\" }}",
                    sheet_index + 1,
                    range_ref(range.start_row, range.start_col, range.end_row, range.end_col)
                ))
            }
            UndoAction::ValidationExcluded { sheet_index, range } => {
                Some(format!(
                    "grid.exclude_validation{{ sheet={}, range=\"{}\" }}",
                    sheet_index + 1,
                    range_ref(range.start_row, range.start_col, range.end_row, range.end_col)
                ))
            }
            UndoAction::ValidationExclusionCleared { sheet_index, range, .. } => {
                Some(format!(
                    "grid.clear_exclusion{{ sheet={}, range=\"{}\" }}",
                    sheet_index + 1,
                    range_ref(range.start_row, range.start_col, range.end_row, range.end_col)
                ))
            }
            UndoAction::FreezePanesChanged { .. } => {
                // View-only - no executable Lua
                None
            }
            UndoAction::Rewind { .. } => {
                // Rewind is audit-only - no executable Lua
                None
            }
            UndoAction::SetMerges { sheet_index, description, .. } => {
                Some(format!(
                    "grid.{}{{ sheet={} }}",
                    if description.starts_with("Unmerge") { "unmerge" } else { "merge" },
                    sheet_index + 1
                ))
            }
        }
    }

    /// Convert this action to a Lua comment (for audit trail).
    /// Used for Rewind and other audit-only entries.
    pub fn to_lua_comment(&self) -> String {
        match self {
            UndoAction::Rewind {
                target_entry_id,
                target_action_summary,
                discarded_count,
                timestamp_utc,
                preview_replay_count,
                preview_build_ms,
                ..
            } => {
                format!(
                    "-- REWIND: #{} Before \"{}\" | Discarded {} | Replay {} actions | {}ms | UTC {}",
                    target_entry_id,
                    target_action_summary,
                    discarded_count,
                    preview_replay_count,
                    preview_build_ms,
                    timestamp_utc
                )
            }
            _ => format!("-- {}", self.label()),
        }
    }
}

// ============================================================================
// Replay-Compatible Fingerprint
// ============================================================================

/// Replay-compatible fingerprint for provenance scripts.
///
/// This uses the same content-level hashing scheme as the CLI replay engine:
/// each operation is normalized to a deterministic string and fed into blake3.
/// This ensures that `--verify` works across the GUI export → CLI replay boundary.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ReplayFingerprint {
    pub len: usize,
    pub hash_hi: u64,
    pub hash_lo: u64,
}

impl std::fmt::Display for ReplayFingerprint {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "v1:{}:{:016x}{:016x}", self.len, self.hash_hi, self.hash_lo)
    }
}

impl UndoAction {
    /// Produce the replay-compatible hash strings for this action.
    ///
    /// Returns a vec because Group actions expand into multiple hash entries.
    /// Returns empty vec for audit-only actions (Rewind) that don't affect replay.
    ///
    /// The format of each string matches exactly what the CLI replay engine
    /// hashes in `register_grid_api`, ensuring fingerprint alignment.
    pub fn to_replay_hashes(&self) -> Vec<String> {
        match self {
            UndoAction::Values { sheet_index, changes } => {
                let sheet = sheet_index + 1; // 1-indexed in Lua API
                if changes.len() == 1 {
                    let c = &changes[0];
                    vec![format!("set:{}:{}:{}:{}", sheet, c.row, c.col, c.new_value)]
                } else {
                    let parts: Vec<String> = changes.iter()
                        .map(|c| format!("{}:{}:{}", c.row, c.col, c.new_value))
                        .collect();
                    vec![format!("set_batch:{}:{}", sheet, parts.join("|"))]
                }
            }
            UndoAction::Format { sheet_index, patches, kind, .. } => {
                if patches.is_empty() {
                    return vec![];
                }
                let (min_row, min_col, max_row, max_col) = patches.iter().fold(
                    (usize::MAX, usize::MAX, 0, 0),
                    |(r1, c1, r2, c2), p| {
                        (r1.min(p.row), c1.min(p.col), r2.max(p.row), c2.max(p.col))
                    },
                );
                let range = range_ref(min_row, min_col, max_row, max_col);
                let kind_str = format_kind_hash_str(kind);
                vec![format!("format:{}:{}:{}", sheet_index + 1, range, kind_str)]
            }
            UndoAction::RowsInserted { sheet_index, at_row, count } => {
                vec![format!("insert_rows:{}:{}:{}", sheet_index + 1, at_row + 1, count)]
            }
            UndoAction::RowsDeleted { sheet_index, at_row, count, .. } => {
                vec![format!("delete_rows:{}:{}:{}", sheet_index + 1, at_row + 1, count)]
            }
            UndoAction::ColsInserted { sheet_index, at_col, count } => {
                vec![format!("insert_cols:{}:{}:{}", sheet_index + 1, at_col + 1, count)]
            }
            UndoAction::ColsDeleted { sheet_index, at_col, count, .. } => {
                vec![format!("delete_cols:{}:{}:{}", sheet_index + 1, at_col + 1, count)]
            }
            UndoAction::ColumnWidthSet { sheet_id, col, new, .. } => {
                let col_letter = col_to_letter(*col);
                if let Some(width) = new {
                    vec![format!("set_col_width:{}:{}:{}", sheet_id.0, col_letter, *width as i64)]
                } else {
                    vec![format!("clear_col_width:{}:{}", sheet_id.0, col_letter)]
                }
            }
            UndoAction::RowHeightSet { sheet_id, row, new, .. } => {
                if let Some(height) = new {
                    vec![format!("set_row_height:{}:{}:{}", sheet_id.0, row + 1, *height as i64)]
                } else {
                    vec![format!("clear_row_height:{}:{}", sheet_id.0, row + 1)]
                }
            }
            UndoAction::RowVisibilityChanged { sheet_id, rows, hidden } => {
                let row_list: Vec<String> = rows.iter().map(|r| (r + 1).to_string()).collect();
                vec![format!("set_row_visibility:{}:{}:{}", sheet_id.0, row_list.join(","), hidden)]
            }
            UndoAction::ColVisibilityChanged { sheet_id, cols, hidden } => {
                let col_list: Vec<String> = cols.iter().map(|c| col_to_letter(*c)).collect();
                vec![format!("set_col_visibility:{}:{}:{}", sheet_id.0, col_list.join(","), hidden)]
            }
            UndoAction::SortApplied { sheet_index, new_sort_state, .. } => {
                let (col, ascending) = new_sort_state;
                vec![format!("sort:{}:{}:{}", sheet_index + 1, col + 1, ascending)]
            }
            UndoAction::SortCleared { sheet_index, .. } => {
                vec![format!("clear_sort:{}", sheet_index + 1)]
            }
            UndoAction::ValidationSet { sheet_index, range, new_rule, .. } => {
                use visigrid_engine::validation::ValidationType;
                let range_str = range_ref(range.start_row, range.start_col, range.end_row, range.end_col);
                let vtype = match &new_rule.rule_type {
                    ValidationType::WholeNumber(_) => "whole_number",
                    ValidationType::Decimal(_) => "decimal",
                    ValidationType::List(_) => "list",
                    ValidationType::Date(_) => "date",
                    ValidationType::Time(_) => "time",
                    ValidationType::TextLength(_) => "text_length",
                    ValidationType::Custom(_) => "custom",
                };
                vec![format!("validate:{}:{}:{}", sheet_index + 1, range_str, vtype)]
            }
            UndoAction::ValidationCleared { sheet_index, range, .. } => {
                let range_str = range_ref(range.start_row, range.start_col, range.end_row, range.end_col);
                vec![format!("clear_validation:{}:{}", sheet_index + 1, range_str)]
            }
            UndoAction::ValidationExcluded { sheet_index, range } => {
                let range_str = range_ref(range.start_row, range.start_col, range.end_row, range.end_col);
                vec![format!("exclude_validation:{}:{}", sheet_index + 1, range_str)]
            }
            UndoAction::ValidationExclusionCleared { sheet_index, range, .. } => {
                let range_str = range_ref(range.start_row, range.start_col, range.end_row, range.end_col);
                vec![format!("clear_exclusion:{}:{}", sheet_index + 1, range_str)]
            }
            UndoAction::NamedRangeCreated { named_range } => {
                use visigrid_engine::named_range::NamedRangeTarget;
                let (sheet, range_str) = match &named_range.target {
                    NamedRangeTarget::Cell { sheet, row, col } => {
                        (*sheet + 1, cell_ref(*row, *col))
                    }
                    NamedRangeTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
                        (*sheet + 1, range_ref(*start_row, *start_col, *end_row, *end_col))
                    }
                };
                vec![format!("define_name:{}:{}:{}", named_range.name, sheet, range_str)]
            }
            UndoAction::NamedRangeDeleted { named_range } => {
                vec![format!("undefine_name:{}", named_range.name)]
            }
            UndoAction::NamedRangeRenamed { old_name, new_name } => {
                vec![format!("rename_name:{}:{}", old_name, new_name)]
            }
            UndoAction::NamedRangeDescriptionChanged { name, new_description, .. } => {
                vec![format!("set_name_description:{}:{}", name, new_description.as_deref().unwrap_or(""))]
            }
            UndoAction::Group { actions, .. } => {
                // Groups expand: each sub-action contributes its own hash entries
                actions.iter().flat_map(|a| a.to_replay_hashes()).collect()
            }
            UndoAction::SetMerges { sheet_index, description, .. } => {
                let op = if description.starts_with("Unmerge") { "unmerge" } else { "merge" };
                vec![format!("{}:{}", op, sheet_index + 1)]
            }
            UndoAction::FreezePanesChanged { .. } => {
                // View-only — not hashed
                vec![]
            }
            UndoAction::Rewind { .. } => {
                // Audit-only — not replayed, not hashed
                vec![]
            }
        }
    }
}

/// Compute a replay-compatible fingerprint from history entries.
///
/// Uses the same blake3 content-hashing scheme as the CLI replay engine,
/// ensuring that fingerprints match across the GUI export → CLI replay boundary.
pub fn compute_replay_fingerprint(entries: &[HistoryEntry]) -> ReplayFingerprint {
    use blake3::Hasher;

    let mut hasher = Hasher::new();
    let mut count = 0usize;

    for entry in entries {
        for hash_str in entry.action.to_replay_hashes() {
            hasher.update(hash_str.as_bytes());
            hasher.update(b"\n");
            count += 1;
        }
    }

    let hash = hasher.finalize();
    let bytes = hash.as_bytes();
    let hash_hi = u64::from_be_bytes(bytes[0..8].try_into().unwrap());
    let hash_lo = u64::from_be_bytes(bytes[8..16].try_into().unwrap());

    ReplayFingerprint { len: count, hash_hi, hash_lo }
}

/// Convert Values action to Lua.
fn values_to_lua(sheet_index: usize, changes: &[CellChange]) -> String {
    if changes.len() == 1 {
        let c = &changes[0];
        format!(
            "grid.set{{ sheet={}, cell=\"{}\", value={} }}",
            sheet_index + 1,
            cell_ref(c.row, c.col),
            lua_escape(&c.new_value)
        )
    } else {
        // Batch set
        let mut cells = String::new();
        cells.push_str("{\n");
        for c in changes {
            cells.push_str(&format!(
                "    {{ cell=\"{}\", value={} }},\n",
                cell_ref(c.row, c.col),
                lua_escape(&c.new_value)
            ));
        }
        cells.push_str("  }");
        format!(
            "grid.set_batch{{ sheet={}, cells={} }}",
            sheet_index + 1,
            cells
        )
    }
}

/// Convert Format action to Lua.
fn format_to_lua(
    sheet_index: usize,
    patches: &[CellFormatPatch],
    kind: &crate::history::FormatActionKind,
) -> String {
    use crate::history::FormatActionKind;

    if patches.is_empty() {
        return "-- (empty format change)".to_string();
    }

    // Compute range from patches
    let (min_row, min_col, max_row, max_col) = patches.iter().fold(
        (usize::MAX, usize::MAX, 0, 0),
        |(r1, c1, r2, c2), p| {
            (r1.min(p.row), c1.min(p.col), r2.max(p.row), c2.max(p.col))
        },
    );
    let range = range_ref(min_row, min_col, max_row, max_col);

    // Get the format change from first patch (they should all be the same kind)
    let first = &patches[0];
    let format_str = match kind {
        FormatActionKind::Bold => format!("bold={}", first.after.bold),
        FormatActionKind::Italic => format!("italic={}", first.after.italic),
        FormatActionKind::Underline => format!("underline={}", first.after.underline),
        FormatActionKind::Strikethrough => format!("strikethrough={}", first.after.strikethrough),
        FormatActionKind::Font => {
            if let Some(ref font) = first.after.font_family {
                format!("font={}", lua_escape(font))
            } else {
                "font=nil".to_string()
            }
        }
        FormatActionKind::Alignment => {
            format!("align=\"{}\"", format_alignment(&first.after))
        }
        FormatActionKind::VerticalAlignment => {
            format!("valign=\"{}\"", format_valignment(&first.after))
        }
        FormatActionKind::TextOverflow => {
            format!("overflow=\"{}\"", format_overflow(&first.after))
        }
        FormatActionKind::NumberFormat | FormatActionKind::DecimalPlaces => {
            format_number_format(&first.after.number_format)
        }
        FormatActionKind::BackgroundColor => {
            if let Some(color) = first.after.background_color {
                format!("bg=\"#{}\"", format_color(color))
            } else {
                "bg=nil".to_string()
            }
        }
        FormatActionKind::FontSize => {
            if let Some(size) = first.after.font_size {
                format!("font_size={}", size as u32)
            } else {
                "font_size=nil".to_string()
            }
        }
        FormatActionKind::FontColor => {
            if let Some(color) = first.after.font_color {
                format!("font_color=\"#{}\"", format_color(color))
            } else {
                "font_color=nil".to_string()
            }
        }
        FormatActionKind::Border => {
            "border=true".to_string() // Simplified; full border state is complex
        }
        FormatActionKind::PasteFormats => {
            // Full format paste - show multiple properties
            let mut props = Vec::new();
            if first.after.bold { props.push("bold=true"); }
            if first.after.italic { props.push("italic=true"); }
            if first.after.underline { props.push("underline=true"); }
            if first.after.strikethrough { props.push("strikethrough=true"); }
            if first.after.background_color.is_some() { props.push("bg=..."); }
            if !props.is_empty() {
                props.join(", ")
            } else {
                "format=default".to_string()
            }
        }
        FormatActionKind::ClearFormatting => {
            "format=default".to_string()
        }
        FormatActionKind::CellStyle => {
            format!("cell_style=\"{}\"", first.after.cell_style.label())
        }
    };

    format!(
        "grid.format{{ sheet={}, range=\"{}\", kind=\"{}\", {} }}",
        sheet_index + 1,
        range,
        format_kind_hash_str(kind),
        format_str
    )
}

/// Format horizontal alignment for Lua.
fn format_alignment(fmt: &visigrid_engine::cell::CellFormat) -> &'static str {
    use visigrid_engine::cell::Alignment;
    match fmt.alignment {
        Alignment::General => "general",
        Alignment::Left => "left",
        Alignment::Center => "center",
        Alignment::Right => "right",
        Alignment::CenterAcrossSelection => "center_across",
    }
}

/// Format vertical alignment for Lua.
fn format_valignment(fmt: &visigrid_engine::cell::CellFormat) -> &'static str {
    use visigrid_engine::cell::VerticalAlignment;
    match fmt.vertical_alignment {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Middle => "middle",
        VerticalAlignment::Bottom => "bottom",
    }
}

/// Format text overflow for Lua.
fn format_overflow(fmt: &visigrid_engine::cell::CellFormat) -> &'static str {
    use visigrid_engine::cell::TextOverflow;
    match fmt.text_overflow {
        TextOverflow::Clip => "clip",
        TextOverflow::Wrap => "wrap",
        TextOverflow::Overflow => "overflow",
    }
}

/// Format RGBA color as hex string (without #).
fn format_color(rgba: [u8; 4]) -> String {
    format!("{:02x}{:02x}{:02x}{:02x}", rgba[0], rgba[1], rgba[2], rgba[3])
}

/// Format number format for Lua.
fn format_number_format(fmt: &visigrid_engine::cell::NumberFormat) -> String {
    use visigrid_engine::cell::NumberFormat;
    match fmt {
        NumberFormat::General => "number_format=\"general\"".to_string(),
        NumberFormat::Number { decimals, .. } => format!("number_format=\"number\", decimals={}", decimals),
        NumberFormat::Currency { decimals, .. } => format!("number_format=\"currency\", decimals={}", decimals),
        NumberFormat::Percent { decimals } => format!("number_format=\"percent\", decimals={}", decimals),
        NumberFormat::Date { .. } => "number_format=\"date\"".to_string(),
        NumberFormat::Time => "number_format=\"time\"".to_string(),
        NumberFormat::DateTime => "number_format=\"datetime\"".to_string(),
        NumberFormat::Custom(code) => format!("number_format=\"custom\", code=\"{}\"", code),
    }
}

/// Convert NamedRangeCreated to Lua.
fn named_range_created_to_lua(nr: &visigrid_engine::named_range::NamedRange) -> String {
    use visigrid_engine::named_range::NamedRangeTarget;

    match &nr.target {
        NamedRangeTarget::Cell { sheet, row, col } => {
            format!(
                "grid.define_name{{ name={}, sheet={}, range=\"{}\" }}",
                lua_escape(&nr.name),
                sheet + 1,
                cell_ref(*row, *col)
            )
        }
        NamedRangeTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
            format!(
                "grid.define_name{{ name={}, sheet={}, range=\"{}\" }}",
                lua_escape(&nr.name),
                sheet + 1,
                range_ref(*start_row, *start_col, *end_row, *end_col)
            )
        }
    }
}

/// Convert Group action to Lua.
fn group_to_lua(actions: &[UndoAction], description: &str) -> String {
    let mut lines = Vec::new();
    lines.push(format!("-- BEGIN GROUP: {}", description));
    for action in actions {
        if let Some(lua) = action.to_lua() {
            lines.push(lua);
        } else {
            lines.push(action.to_lua_comment());
        }
    }
    lines.push(format!("-- END GROUP: {}", description));
    lines.join("\n")
}

/// Convert ValidationSet to Lua.
fn validation_set_to_lua(
    sheet_index: usize,
    range: &visigrid_engine::validation::CellRange,
    rule: &visigrid_engine::validation::ValidationRule,
) -> String {
    use visigrid_engine::validation::{ListSource, ValidationType};

    let range_str = range_ref(range.start_row, range.start_col, range.end_row, range.end_col);

    let type_str = match &rule.rule_type {
        ValidationType::WholeNumber(constraint) => {
            format!("type=\"whole_number\", {}", constraint_to_lua(constraint))
        }
        ValidationType::Decimal(constraint) => {
            format!("type=\"decimal\", {}", constraint_to_lua(constraint))
        }
        ValidationType::List(source) => {
            let source_str = match source {
                ListSource::Inline(items) => {
                    let items_str = items.iter().map(|s| lua_escape(s)).collect::<Vec<_>>().join(", ");
                    format!("source={{{}}}", items_str)
                }
                ListSource::Range(r) => format!("source={}", lua_escape(r)),
                ListSource::NamedRange(n) => format!("source_name={}", lua_escape(n)),
            };
            format!("type=\"list\", {}", source_str)
        }
        ValidationType::Date(constraint) => {
            format!("type=\"date\", {}", constraint_to_lua(constraint))
        }
        ValidationType::Time(constraint) => {
            format!("type=\"time\", {}", constraint_to_lua(constraint))
        }
        ValidationType::TextLength(constraint) => {
            format!("type=\"text_length\", {}", constraint_to_lua(constraint))
        }
        ValidationType::Custom(formula) => {
            format!("type=\"custom\", formula={}", lua_escape(formula))
        }
    };

    format!(
        "grid.validate{{ sheet={}, range=\"{}\", {} }}",
        sheet_index + 1,
        range_str,
        type_str
    )
}

/// Convert NumericConstraint to Lua parameters.
fn constraint_to_lua(constraint: &visigrid_engine::validation::NumericConstraint) -> String {
    use visigrid_engine::validation::ComparisonOperator;

    let val1 = constraint_value_to_lua(&constraint.value1);
    let val2 = constraint.value2.as_ref().map(constraint_value_to_lua);

    match constraint.operator {
        ComparisonOperator::Between => {
            format!("min={}, max={}", val1, val2.unwrap_or_default())
        }
        ComparisonOperator::NotBetween => {
            format!("not_between=true, min={}, max={}", val1, val2.unwrap_or_default())
        }
        ComparisonOperator::EqualTo => format!("equal={}", val1),
        ComparisonOperator::NotEqualTo => format!("not_equal={}", val1),
        ComparisonOperator::GreaterThan => format!("gt={}", val1),
        ComparisonOperator::LessThan => format!("lt={}", val1),
        ComparisonOperator::GreaterThanOrEqual => format!("gte={}", val1),
        ComparisonOperator::LessThanOrEqual => format!("lte={}", val1),
    }
}

/// Convert a ConstraintValue to its Lua representation.
fn constraint_value_to_lua(val: &visigrid_engine::validation::ConstraintValue) -> String {
    use visigrid_engine::validation::ConstraintValue;
    match val {
        ConstraintValue::Number(n) => format!("{}", n),
        ConstraintValue::CellRef(r) => lua_escape(r),
        ConstraintValue::Formula(f) => lua_escape(f),
    }
}

// ============================================================================
// Script Export
// ============================================================================

/// Options for exporting a provenance script.
#[derive(Clone, Debug, Default)]
pub struct ExportOptions {
    /// Include header with metadata
    pub include_header: bool,
    /// Include expected fingerprint for verification
    pub include_fingerprint: bool,
    /// Filter to specific sheet (None = all sheets)
    pub sheet_filter: Option<usize>,
}

/// Export a complete provenance script from history entries.
///
/// The fingerprint embedded in the script uses the replay-compatible content-level
/// hashing scheme (matching the CLI replay engine), not the GUI's structural
/// `HistoryFingerprint`. This ensures `vgrid replay --verify` works on
/// scripts exported from the GUI.
pub fn export_script(
    entries: &[HistoryEntry],
    _fingerprint: HistoryFingerprint,
    workbook_name: Option<&str>,
    options: &ExportOptions,
) -> String {
    // Compute replay-compatible fingerprint (matches CLI replay engine)
    let replay_fp = compute_replay_fingerprint(entries);

    let mut lines = Vec::new();

    // Header
    if options.include_header {
        lines.push(format!("-- api={}", LUA_API_VERSION));
        lines.push("-- VisiGrid Provenance Script".to_string());
        lines.push(format!("-- Generated: {}", crate::app::chrono_lite_utc()));
        if let Some(name) = workbook_name {
            lines.push(format!("-- Workbook: {}", name));
        }
        lines.push(format!("-- Actions: {}", entries.len()));
        if options.include_fingerprint {
            lines.push(format!("-- Expected fingerprint: {}", replay_fp));
        }
        lines.push(String::new());
    }

    // Actions
    for (i, entry) in entries.iter().enumerate() {
        // Filter by sheet if requested
        if let Some(sheet_filter) = options.sheet_filter {
            if !action_affects_sheet(&entry.action, sheet_filter) {
                continue;
            }
        }

        // Comment with entry ID and summary
        let summary = entry.action.label();
        lines.push(format!("-- #{} {}", entry.id, summary));

        // Lua code or audit comment
        if let Some(lua) = entry.action.to_lua() {
            lines.push(lua);
        } else {
            lines.push(entry.action.to_lua_comment());
        }

        // Blank line between entries (except last)
        if i < entries.len() - 1 {
            lines.push(String::new());
        }
    }

    // Footer with fingerprint verification
    if options.include_fingerprint {
        lines.push(String::new());
        lines.push(format!(
            "-- END: {} actions | Fingerprint {}",
            entries.len(),
            replay_fp
        ));
    }

    lines.join("\n")
}

/// Check if an action affects a specific sheet.
fn action_affects_sheet(action: &UndoAction, sheet_index: usize) -> bool {
    match action {
        UndoAction::Values { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::Format { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::RowsInserted { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::RowsDeleted { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::ColsInserted { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::ColsDeleted { sheet_index: s, .. } => *s == sheet_index,
        // Layout actions use SheetId, not sheet_index - need to resolve at call site
        // For now, always include layout actions (they're rare in per-sheet exports)
        UndoAction::ColumnWidthSet { .. } => true,
        UndoAction::RowHeightSet { .. } => true,
        UndoAction::RowVisibilityChanged { .. } => true,
        UndoAction::ColVisibilityChanged { .. } => true,
        UndoAction::SortApplied { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::SortCleared { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::ValidationSet { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::ValidationCleared { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::ValidationExcluded { sheet_index: s, .. } => *s == sheet_index,
        UndoAction::ValidationExclusionCleared { sheet_index: s, .. } => *s == sheet_index,
        // Named ranges are global, always include
        UndoAction::NamedRangeCreated { .. } => true,
        UndoAction::NamedRangeDeleted { .. } => true,
        UndoAction::NamedRangeRenamed { .. } => true,
        UndoAction::NamedRangeDescriptionChanged { .. } => true,
        // Groups: check if any sub-action affects the sheet
        UndoAction::Group { actions, .. } => {
            actions.iter().any(|a| action_affects_sheet(a, sheet_index))
        }
        // View-only, include everywhere
        UndoAction::FreezePanesChanged { .. } => true,
        // Rewind is audit-only, always include
        UndoAction::Rewind { .. } => true,
        UndoAction::SetMerges { sheet_index: s, .. } => *s == sheet_index,
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::history::CellChange;

    #[test]
    fn test_cell_ref() {
        assert_eq!(cell_ref(0, 0), "A1");
        assert_eq!(cell_ref(0, 25), "Z1");
        assert_eq!(cell_ref(0, 26), "AA1");
        assert_eq!(cell_ref(99, 27), "AB100");
    }

    #[test]
    fn test_range_ref() {
        assert_eq!(range_ref(0, 0, 0, 0), "A1");
        assert_eq!(range_ref(0, 0, 9, 3), "A1:D10");
        assert_eq!(range_ref(4, 2, 4, 2), "C5");
    }

    #[test]
    fn test_lua_escape() {
        assert_eq!(lua_escape("hello"), "\"hello\"");
        assert_eq!(lua_escape("say \"hi\""), "\"say \\\"hi\\\"\"");
        assert_eq!(lua_escape("line1\nline2"), "\"line1\\nline2\"");
        assert_eq!(lua_escape("path\\to\\file"), "\"path\\\\to\\\\file\"");
    }

    #[test]
    fn test_values_single_to_lua() {
        let action = UndoAction::Values {
            sheet_index: 0,
            changes: vec![CellChange {
                row: 0,
                col: 0,
                old_value: "".to_string(),
                new_value: "Hello".to_string(),
            }],
        };
        let lua = action.to_lua().unwrap();
        assert_eq!(lua, "grid.set{ sheet=1, cell=\"A1\", value=\"Hello\" }");
    }

    #[test]
    fn test_values_batch_to_lua() {
        let action = UndoAction::Values {
            sheet_index: 0,
            changes: vec![
                CellChange { row: 0, col: 0, old_value: "".to_string(), new_value: "A".to_string() },
                CellChange { row: 0, col: 1, old_value: "".to_string(), new_value: "B".to_string() },
            ],
        };
        let lua = action.to_lua().unwrap();
        assert!(lua.contains("grid.set_batch"));
        assert!(lua.contains("cell=\"A1\""));
        assert!(lua.contains("cell=\"B1\""));
    }

    #[test]
    fn test_rows_inserted_to_lua() {
        let action = UndoAction::RowsInserted {
            sheet_index: 0,
            at_row: 4,
            count: 3,
        };
        let lua = action.to_lua().unwrap();
        assert_eq!(lua, "grid.insert_rows{ sheet=1, at=5, count=3 }");
    }

    #[test]
    fn test_sort_to_lua() {
        let action = UndoAction::SortApplied {
            sheet_index: 0,
            previous_row_order: vec![],
            previous_sort_state: None,
            new_row_order: vec![0, 1, 2],
            new_sort_state: (2, true), // Column C, ascending
        };
        let lua = action.to_lua().unwrap();
        assert_eq!(lua, "grid.sort{ sheet=1, col=3, ascending=true }");
    }

    #[test]
    fn test_rewind_is_audit_only() {
        let action = UndoAction::Rewind {
            target_entry_id: 42,
            target_index: 10,
            target_action_summary: "Sort Column C".to_string(),
            discarded_count: 5,
            old_history_len: 15,
            new_history_len: 11,
            timestamp_utc: "1706000000".to_string(),
            preview_replay_count: 10,
            preview_build_ms: 25,
        };
        // to_lua returns None for audit-only actions
        assert!(action.to_lua().is_none());
        // But to_lua_comment provides the audit trail
        let comment = action.to_lua_comment();
        assert!(comment.starts_with("-- REWIND:"));
        assert!(comment.contains("Sort Column C"));
        assert!(comment.contains("Discarded 5"));
    }

    #[test]
    fn test_replay_fingerprint_matches_cli_golden() {
        // Verify that compute_replay_fingerprint produces the same fingerprint
        // as the CLI replay engine for the simple.lua golden test.
        // This is the critical cross-boundary alignment test.
        use std::time::Instant;
        use crate::history::MutationSource;
        let make_entry = |id, action| HistoryEntry {
            id,
            action,
            timestamp: Instant::now(),
            provenance: None,
            source: MutationSource::Human,
        };
        let entries = vec![
            make_entry(1, UndoAction::Values {
                sheet_index: 0,
                changes: vec![CellChange { row: 0, col: 0, old_value: "".into(), new_value: "Name".into() }],
            }),
            make_entry(2, UndoAction::Values {
                sheet_index: 0,
                changes: vec![CellChange { row: 0, col: 1, old_value: "".into(), new_value: "Value".into() }],
            }),
            make_entry(3, UndoAction::Values {
                sheet_index: 0,
                changes: vec![CellChange { row: 1, col: 0, old_value: "".into(), new_value: "Alpha".into() }],
            }),
            make_entry(4, UndoAction::Values {
                sheet_index: 0,
                changes: vec![CellChange { row: 1, col: 1, old_value: "".into(), new_value: "100".into() }],
            }),
            make_entry(5, UndoAction::Values {
                sheet_index: 0,
                changes: vec![CellChange { row: 2, col: 0, old_value: "".into(), new_value: "Beta".into() }],
            }),
            make_entry(6, UndoAction::Values {
                sheet_index: 0,
                changes: vec![CellChange { row: 2, col: 1, old_value: "".into(), new_value: "200".into() }],
            }),
        ];

        let fp = compute_replay_fingerprint(&entries);
        // Must match the golden fingerprint from tests/golden/simple.lua
        assert_eq!(fp.len, 6);
        assert_eq!(
            fp.to_string(),
            "v1:6:b38b632d7f38dedf50ebea7b52785554",
            "GUI fingerprint must match CLI golden test fingerprint"
        );
    }

    #[test]
    fn test_replay_hash_values_single() {
        let action = UndoAction::Values {
            sheet_index: 0,
            changes: vec![CellChange { row: 0, col: 0, old_value: "".into(), new_value: "Hello".into() }],
        };
        let hashes = action.to_replay_hashes();
        assert_eq!(hashes, vec!["set:1:0:0:Hello"]);
    }

    #[test]
    fn test_replay_hash_values_batch() {
        let action = UndoAction::Values {
            sheet_index: 0,
            changes: vec![
                CellChange { row: 0, col: 0, old_value: "".into(), new_value: "A".into() },
                CellChange { row: 0, col: 1, old_value: "".into(), new_value: "B".into() },
            ],
        };
        let hashes = action.to_replay_hashes();
        assert_eq!(hashes, vec!["set_batch:1:0:0:A|0:1:B"]);
    }

    #[test]
    fn test_replay_hash_rewind_is_empty() {
        let action = UndoAction::Rewind {
            target_entry_id: 1,
            target_index: 0,
            target_action_summary: "test".into(),
            discarded_count: 1,
            old_history_len: 2,
            new_history_len: 1,
            timestamp_utc: "now".into(),
            preview_replay_count: 1,
            preview_build_ms: 10,
        };
        assert!(action.to_replay_hashes().is_empty(), "Rewind should not contribute to fingerprint");
    }

    #[test]
    fn test_exported_fingerprint_has_v1_prefix() {
        use std::time::Instant;
        use crate::history::MutationSource;
        let entries = vec![
            HistoryEntry {
                id: 1,
                action: UndoAction::Values {
                    sheet_index: 0,
                    changes: vec![CellChange { row: 0, col: 0, old_value: "".into(), new_value: "test".into() }],
                },
                timestamp: Instant::now(),
                provenance: None,
                source: MutationSource::Human,
            },
        ];
        let options = ExportOptions {
            include_header: true,
            include_fingerprint: true,
            sheet_filter: None,
        };
        let script = export_script(
            &entries,
            HistoryFingerprint { len: 0, hash_hi: 0, hash_lo: 0 }, // ignored
            Some("Test"),
            &options,
        );
        // Fingerprint in header must have v1: prefix
        assert!(
            script.contains("-- Expected fingerprint: v1:"),
            "Exported fingerprint must have v1: prefix, got:\n{}",
            script
        );
    }
}
