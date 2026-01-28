//! Lua provenance for spreadsheet operations (Phase 4).
//!
//! Provenance captures user intent as structured operations that can be
//! expressed as human-readable Lua snippets. This is NOT replay infrastructure -
//! it's read-only audit trail for explainability.
//!
//! Design principles:
//! - One MutationOp per user intent (not per cell)
//! - Lua uses 1-based A1 notation (human-friendly)
//! - Deterministic output (sorted cells, stable key ordering)
//! - No implementation details leak into Lua

use crate::sheet::SheetId;

/// A mutation operation representing user intent.
///
/// Each variant maps to a single user action that may affect multiple cells.
/// This is the "what happened" that gets converted to Lua.
#[derive(Debug, Clone)]
pub enum MutationOp {
    /// Single cell edit (shown as "Manual edit", minimal Lua)
    SetCell {
        sheet: SheetId,
        row: usize,
        col: usize,
        value: String,
    },

    /// Paste operation (values, formulas, or both)
    Paste {
        sheet: SheetId,
        dst_row: usize,
        dst_col: usize,
        values: Vec<Vec<String>>,
        mode: PasteMode,
    },

    /// Fill operation (handle drag, Ctrl+D, Ctrl+R)
    Fill {
        sheet: SheetId,
        src_start_row: usize,
        src_start_col: usize,
        src_end_row: usize,
        src_end_col: usize,
        dst_start_row: usize,
        dst_start_col: usize,
        dst_end_row: usize,
        dst_end_col: usize,
        direction: FillDirection,
        mode: FillMode,
    },

    /// Multi-edit (Ctrl+Enter on selection)
    MultiEdit {
        sheet: SheetId,
        cells: Vec<(usize, usize)>,
        value: String,
    },

    /// Sort operation
    Sort {
        sheet: SheetId,
        range_start_row: usize,
        range_start_col: usize,
        range_end_row: usize,
        range_end_col: usize,
        keys: Vec<SortKey>,
        has_header: bool,
    },

    /// Clear range
    Clear {
        sheet: SheetId,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        mode: ClearMode,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PasteMode {
    Values,
    Formulas,
    Formats,
    Both,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FillDirection {
    Down,
    Right,
    Up,
    Left,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FillMode {
    Values,
    Formulas,
    Both,
}

#[derive(Debug, Clone)]
pub struct SortKey {
    pub col: usize,
    pub ascending: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ClearMode {
    Values,
    Formats,
    All,
}

/// Provenance record attached to a history entry.
#[derive(Debug, Clone)]
pub struct Provenance {
    /// The operation that was performed
    pub op: MutationOp,
    /// Lua representation of the operation
    pub lua: String,
    /// Human-readable label for UI
    pub label: String,
    /// Scope description (e.g., "3x3 range", "47 cells")
    pub scope: String,
}

impl MutationOp {
    /// Convert to Lua snippet.
    ///
    /// Uses 1-based A1 notation for human readability.
    /// Output is deterministic (sorted, stable ordering).
    pub fn to_lua(&self, sheet_name: &str) -> String {
        match self {
            MutationOp::SetCell { row, col, value, .. } => {
                let cell_ref = cell_to_a1(*row, *col);
                let escaped = escape_lua_string(value);
                format!(
                    "grid.set{{\n  sheet = \"{}\",\n  cell = \"{}\",\n  value = \"{}\"\n}}",
                    sheet_name, cell_ref, escaped
                )
            }

            MutationOp::Paste { dst_row, dst_col, values, mode, .. } => {
                let cell_ref = cell_to_a1(*dst_row, *dst_col);
                let mode_str = match mode {
                    PasteMode::Values => "values",
                    PasteMode::Formulas => "formulas",
                    PasteMode::Formats => "formats",
                    PasteMode::Both => "both",
                };

                // Build values array
                let values_lua = values_to_lua_array(values);

                format!(
                    "grid.paste{{\n  sheet = \"{}\",\n  at = \"{}\",\n  mode = \"{}\",\n  values = {}\n}}",
                    sheet_name, cell_ref, mode_str, values_lua
                )
            }

            MutationOp::Fill {
                src_start_row, src_start_col, src_end_row, src_end_col,
                dst_start_row, dst_start_col, dst_end_row, dst_end_col,
                direction, mode, ..
            } => {
                let src_range = range_to_a1(*src_start_row, *src_start_col, *src_end_row, *src_end_col);
                let dst_range = range_to_a1(*dst_start_row, *dst_start_col, *dst_end_row, *dst_end_col);
                let dir_str = match direction {
                    FillDirection::Down => "down",
                    FillDirection::Right => "right",
                    FillDirection::Up => "up",
                    FillDirection::Left => "left",
                };
                let mode_str = match mode {
                    FillMode::Values => "values",
                    FillMode::Formulas => "formulas",
                    FillMode::Both => "both",
                };

                format!(
                    "grid.fill{{\n  sheet = \"{}\",\n  src = \"{}\",\n  dst = \"{}\",\n  direction = \"{}\",\n  mode = \"{}\"\n}}",
                    sheet_name, src_range, dst_range, dir_str, mode_str
                )
            }

            MutationOp::MultiEdit { cells, value, .. } => {
                let escaped = escape_lua_string(value);
                // Sort cells for determinism
                let mut sorted_cells = cells.clone();
                sorted_cells.sort_by(|a, b| (a.0, a.1).cmp(&(b.0, b.1)));

                let cells_lua: Vec<String> = sorted_cells
                    .iter()
                    .map(|(r, c)| format!("\"{}\"", cell_to_a1(*r, *c)))
                    .collect();

                format!(
                    "grid.multi_edit{{\n  sheet = \"{}\",\n  cells = {{ {} }},\n  value = \"{}\"\n}}",
                    sheet_name,
                    cells_lua.join(", "),
                    escaped
                )
            }

            MutationOp::Sort {
                range_start_row, range_start_col, range_end_row, range_end_col,
                keys, has_header, ..
            } => {
                let range = range_to_a1(*range_start_row, *range_start_col, *range_end_row, *range_end_col);

                let keys_lua: Vec<String> = keys
                    .iter()
                    .map(|k| {
                        let col_letter = col_to_letter(k.col);
                        let order = if k.ascending { "asc" } else { "desc" };
                        format!("{{ col = \"{}\", order = \"{}\" }}", col_letter, order)
                    })
                    .collect();

                format!(
                    "grid.sort{{\n  sheet = \"{}\",\n  range = \"{}\",\n  keys = {{\n    {}\n  }},\n  header = {}\n}}",
                    sheet_name,
                    range,
                    keys_lua.join(",\n    "),
                    has_header
                )
            }

            MutationOp::Clear {
                start_row, start_col, end_row, end_col, mode, ..
            } => {
                let range = range_to_a1(*start_row, *start_col, *end_row, *end_col);
                let mode_str = match mode {
                    ClearMode::Values => "values",
                    ClearMode::Formats => "formats",
                    ClearMode::All => "all",
                };

                format!(
                    "grid.clear{{\n  sheet = \"{}\",\n  range = \"{}\",\n  mode = \"{}\"\n}}",
                    sheet_name, range, mode_str
                )
            }
        }
    }

    /// Generate human-readable label for history UI.
    pub fn label(&self) -> String {
        match self {
            MutationOp::SetCell { .. } => "Edit cell".to_string(),
            MutationOp::Paste { values, .. } => {
                let rows = values.len();
                let cols = values.first().map(|r| r.len()).unwrap_or(0);
                format!("Paste {}x{}", rows, cols)
            }
            MutationOp::Fill { direction, .. } => {
                let dir = match direction {
                    FillDirection::Down => "down",
                    FillDirection::Right => "right",
                    FillDirection::Up => "up",
                    FillDirection::Left => "left",
                };
                format!("Fill {}", dir)
            }
            MutationOp::MultiEdit { cells, .. } => {
                format!("Edit {} cells", cells.len())
            }
            MutationOp::Sort { keys, .. } => {
                format!("Sort by {} column{}", keys.len(), if keys.len() == 1 { "" } else { "s" })
            }
            MutationOp::Clear { mode, .. } => {
                let what = match mode {
                    ClearMode::Values => "values",
                    ClearMode::Formats => "formats",
                    ClearMode::All => "contents",
                };
                format!("Clear {}", what)
            }
        }
    }

    /// Generate scope description for history UI.
    pub fn scope(&self) -> String {
        match self {
            MutationOp::SetCell { row, col, .. } => {
                cell_to_a1(*row, *col)
            }
            MutationOp::Paste { dst_row, dst_col, values, .. } => {
                let rows = values.len();
                let cols = values.first().map(|r| r.len()).unwrap_or(0);
                let end_row = dst_row + rows.saturating_sub(1);
                let end_col = dst_col + cols.saturating_sub(1);
                range_to_a1(*dst_row, *dst_col, end_row, end_col)
            }
            MutationOp::Fill { dst_start_row, dst_start_col, dst_end_row, dst_end_col, .. } => {
                range_to_a1(*dst_start_row, *dst_start_col, *dst_end_row, *dst_end_col)
            }
            MutationOp::MultiEdit { cells, .. } => {
                format!("{} cells", cells.len())
            }
            MutationOp::Sort { range_start_row, range_start_col, range_end_row, range_end_col, .. } => {
                range_to_a1(*range_start_row, *range_start_col, *range_end_row, *range_end_col)
            }
            MutationOp::Clear { start_row, start_col, end_row, end_col, .. } => {
                range_to_a1(*start_row, *start_col, *end_row, *end_col)
            }
        }
    }

    /// Create a Provenance record from this operation.
    pub fn to_provenance(&self, sheet_name: &str) -> Provenance {
        Provenance {
            lua: self.to_lua(sheet_name),
            label: self.label(),
            scope: self.scope(),
            op: self.clone(),
        }
    }
}

// ============================================================================
// A1 Notation Helpers
// ============================================================================

/// Convert 0-based column index to letter (0 -> A, 25 -> Z, 26 -> AA)
fn col_to_letter(col: usize) -> String {
    let mut result = String::new();
    let mut c = col;
    loop {
        result.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    result
}

/// Convert 0-based row/col to A1 notation (0,0 -> A1)
fn cell_to_a1(row: usize, col: usize) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

/// Convert 0-based range to A1 notation
fn range_to_a1(start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> String {
    if start_row == end_row && start_col == end_col {
        cell_to_a1(start_row, start_col)
    } else {
        format!("{}:{}", cell_to_a1(start_row, start_col), cell_to_a1(end_row, end_col))
    }
}

/// Escape a string for Lua (handle quotes, newlines, etc.)
fn escape_lua_string(s: &str) -> String {
    s.replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}

/// Convert 2D values array to Lua table syntax
fn values_to_lua_array(values: &[Vec<String>]) -> String {
    if values.is_empty() {
        return "{}".to_string();
    }

    let rows: Vec<String> = values
        .iter()
        .map(|row| {
            let cells: Vec<String> = row
                .iter()
                .map(|v| format!("\"{}\"", escape_lua_string(v)))
                .collect();
            format!("    {{ {} }}", cells.join(", "))
        })
        .collect();

    format!("{{\n{}\n  }}", rows.join(",\n"))
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(51), "AZ");
        assert_eq!(col_to_letter(52), "BA");
        assert_eq!(col_to_letter(701), "ZZ");
        assert_eq!(col_to_letter(702), "AAA");
    }

    #[test]
    fn test_cell_to_a1() {
        assert_eq!(cell_to_a1(0, 0), "A1");
        assert_eq!(cell_to_a1(0, 1), "B1");
        assert_eq!(cell_to_a1(9, 0), "A10");
        assert_eq!(cell_to_a1(0, 26), "AA1");
    }

    #[test]
    fn test_range_to_a1() {
        assert_eq!(range_to_a1(0, 0, 0, 0), "A1");
        assert_eq!(range_to_a1(0, 0, 2, 2), "A1:C3");
        assert_eq!(range_to_a1(4, 1, 9, 3), "B5:D10");
    }

    #[test]
    fn test_paste_lua() {
        let op = MutationOp::Paste {
            sheet: SheetId(1),
            dst_row: 1,  // B2
            dst_col: 1,
            values: vec![
                vec!["10".to_string(), "11".to_string()],
                vec!["20".to_string(), "21".to_string()],
            ],
            mode: PasteMode::Values,
        };

        let lua = op.to_lua("Sheet1");
        assert!(lua.contains("grid.paste"));
        assert!(lua.contains("at = \"B2\""));
        assert!(lua.contains("mode = \"values\""));
        assert!(lua.contains("\"10\""));
        assert!(lua.contains("\"21\""));
    }

    #[test]
    fn test_fill_lua() {
        let op = MutationOp::Fill {
            sheet: SheetId(1),
            src_start_row: 0,
            src_start_col: 0,
            src_end_row: 4,
            src_end_col: 0,
            dst_start_row: 5,
            dst_start_col: 0,
            dst_end_row: 49,
            dst_end_col: 0,
            direction: FillDirection::Down,
            mode: FillMode::Formulas,
        };

        let lua = op.to_lua("Sheet1");
        assert!(lua.contains("grid.fill"));
        assert!(lua.contains("src = \"A1:A5\""));
        assert!(lua.contains("dst = \"A6:A50\""));
        assert!(lua.contains("direction = \"down\""));
        assert!(lua.contains("mode = \"formulas\""));
    }

    #[test]
    fn test_sort_lua() {
        let op = MutationOp::Sort {
            sheet: SheetId(1),
            range_start_row: 0,
            range_start_col: 0,
            range_end_row: 199,
            range_end_col: 3,
            keys: vec![
                SortKey { col: 1, ascending: true },
                SortKey { col: 3, ascending: false },
            ],
            has_header: true,
        };

        let lua = op.to_lua("Sheet1");
        assert!(lua.contains("grid.sort"));
        assert!(lua.contains("range = \"A1:D200\""));
        assert!(lua.contains("col = \"B\", order = \"asc\""));
        assert!(lua.contains("col = \"D\", order = \"desc\""));
        assert!(lua.contains("header = true"));
    }

    #[test]
    fn test_multi_edit_lua() {
        let op = MutationOp::MultiEdit {
            sheet: SheetId(1),
            cells: vec![(2, 1), (0, 0), (1, 1)],  // Out of order - should be sorted
            value: "=SUM(A1:A10)".to_string(),
        };

        let lua = op.to_lua("Sheet1");
        assert!(lua.contains("grid.multi_edit"));
        // Should be sorted: A1, B2, B3
        assert!(lua.contains("\"A1\""));
        assert!(lua.contains("\"B2\""));
        assert!(lua.contains("\"B3\""));
        assert!(lua.contains("value = \"=SUM(A1:A10)\""));
    }

    #[test]
    fn test_escape_lua_string() {
        assert_eq!(escape_lua_string("hello"), "hello");
        assert_eq!(escape_lua_string("say \"hi\""), "say \\\"hi\\\"");
        assert_eq!(escape_lua_string("line1\nline2"), "line1\\nline2");
        assert_eq!(escape_lua_string("path\\file"), "path\\\\file");
    }
}
