/// Diff engine for Explain Differences feature
///
/// Computes what changed between current state and a historical point.
/// Groups changes by category: Values, Formulas, Structural.
/// Supports filtering for AI-touched changes.

use crate::history::{History, HistoryEntry, UndoAction, MutationSource};
use std::collections::HashMap;

/// A report showing what changed since a history entry
#[derive(Debug, Clone)]
pub struct DiffReport {
    /// Entry ID we're diffing from (the "before" point)
    pub since_entry_id: u64,
    /// Label of the entry we're diffing from
    pub since_entry_label: String,
    /// Number of history entries included in this diff
    pub entries_spanned: usize,
    /// Value changes (non-formula cells)
    pub value_changes: Vec<DiffEntry>,
    /// Formula changes
    pub formula_changes: Vec<DiffEntry>,
    /// Structural changes (row/col insert/delete)
    pub structural_changes: Vec<StructuralChange>,
    /// Named range changes
    pub named_range_changes: Vec<NamedRangeChange>,
    /// Validation changes
    pub validation_changes: Vec<ValidationChange>,
    /// Format changes (count only, not individual cells)
    pub format_change_count: usize,
}

impl DiffReport {
    /// Total number of individual changes
    pub fn total_changes(&self) -> usize {
        self.value_changes.len()
            + self.formula_changes.len()
            + self.structural_changes.len()
            + self.named_range_changes.len()
            + self.validation_changes.len()
            + self.format_change_count
    }

    /// Number of AI-touched changes
    pub fn ai_touched_count(&self) -> usize {
        self.value_changes.iter().filter(|e| e.ai_touched).count()
            + self.formula_changes.iter().filter(|e| e.ai_touched).count()
    }

    /// Get value changes, optionally filtered to AI-touched only
    pub fn value_changes_filtered(&self, ai_only: bool) -> Vec<&DiffEntry> {
        self.value_changes.iter()
            .filter(|e| !ai_only || e.ai_touched)
            .collect()
    }

    /// Get formula changes, optionally filtered to AI-touched only
    pub fn formula_changes_filtered(&self, ai_only: bool) -> Vec<&DiffEntry> {
        self.formula_changes.iter()
            .filter(|e| !ai_only || e.ai_touched)
            .collect()
    }
}

/// A single cell's net change (before → after)
#[derive(Debug, Clone)]
pub struct DiffEntry {
    /// Sheet index
    pub sheet_index: usize,
    /// Cell row
    pub row: usize,
    /// Cell column
    pub col: usize,
    /// Value at the "before" point
    pub old_value: String,
    /// Current value
    pub new_value: String,
    /// Whether this was modified by AI
    pub ai_touched: bool,
    /// AI source label if touched by AI (e.g., "OpenAI gpt-4o")
    pub ai_source: Option<String>,
}

impl DiffEntry {
    /// Format cell address (e.g., "A1")
    pub fn cell_address(&self) -> String {
        format!("{}{}", col_to_letter(self.col), self.row + 1)
    }

    /// Format with sheet (e.g., "Sheet1!A1")
    pub fn full_address(&self, sheet_names: &[String]) -> String {
        let sheet_name = sheet_names.get(self.sheet_index)
            .map(|s| s.as_str())
            .unwrap_or("Sheet?");
        format!("{}!{}", sheet_name, self.cell_address())
    }
}

/// Structural change (row/column operations)
#[derive(Debug, Clone)]
pub enum StructuralChange {
    RowsInserted { sheet_index: usize, at_row: usize, count: usize },
    RowsDeleted { sheet_index: usize, at_row: usize, count: usize },
    ColsInserted { sheet_index: usize, at_col: usize, count: usize },
    ColsDeleted { sheet_index: usize, at_col: usize, count: usize },
    Sort { sheet_index: usize, column: usize, ascending: bool },
    SortCleared { sheet_index: usize },
}

impl StructuralChange {
    pub fn description(&self) -> String {
        match self {
            StructuralChange::RowsInserted { at_row, count, .. } => {
                if *count == 1 {
                    format!("Inserted row at {}", at_row + 1)
                } else {
                    format!("Inserted {} rows at {}", count, at_row + 1)
                }
            }
            StructuralChange::RowsDeleted { at_row, count, .. } => {
                if *count == 1 {
                    format!("Deleted row at {}", at_row + 1)
                } else {
                    format!("Deleted {} rows at {}", count, at_row + 1)
                }
            }
            StructuralChange::ColsInserted { at_col, count, .. } => {
                if *count == 1 {
                    format!("Inserted column at {}", col_to_letter(*at_col))
                } else {
                    format!("Inserted {} columns at {}", count, col_to_letter(*at_col))
                }
            }
            StructuralChange::ColsDeleted { at_col, count, .. } => {
                if *count == 1 {
                    format!("Deleted column at {}", col_to_letter(*at_col))
                } else {
                    format!("Deleted {} columns at {}", count, col_to_letter(*at_col))
                }
            }
            StructuralChange::Sort { column, ascending, .. } => {
                let dir = if *ascending { "ascending" } else { "descending" };
                format!("Sorted by column {} {}", col_to_letter(*column), dir)
            }
            StructuralChange::SortCleared { .. } => {
                "Cleared sort".to_string()
            }
        }
    }
}

/// Named range change
#[derive(Debug, Clone)]
pub enum NamedRangeChange {
    Created { name: String },
    Deleted { name: String },
    Renamed { old_name: String, new_name: String },
    DescriptionChanged { name: String },
}

impl NamedRangeChange {
    pub fn description(&self) -> String {
        match self {
            NamedRangeChange::Created { name } => format!("Created '{}'", name),
            NamedRangeChange::Deleted { name } => format!("Deleted '{}'", name),
            NamedRangeChange::Renamed { old_name, new_name } => {
                format!("Renamed '{}' → '{}'", old_name, new_name)
            }
            NamedRangeChange::DescriptionChanged { name } => {
                format!("Changed '{}' description", name)
            }
        }
    }
}

/// Validation change
#[derive(Debug, Clone)]
pub enum ValidationChange {
    Set { sheet_index: usize, range: String, rule_desc: String },
    Cleared { sheet_index: usize, range: String },
    Excluded { sheet_index: usize, range: String },
    ExclusionCleared { sheet_index: usize, range: String },
}

impl ValidationChange {
    pub fn description(&self) -> String {
        match self {
            ValidationChange::Set { range, rule_desc, .. } => {
                format!("Set {} on {}", rule_desc, range)
            }
            ValidationChange::Cleared { range, .. } => format!("Cleared validation on {}", range),
            ValidationChange::Excluded { range, .. } => format!("Excluded {} from validation", range),
            ValidationChange::ExclusionCleared { range, .. } => {
                format!("Cleared exclusion on {}", range)
            }
        }
    }
}

/// Convert column index to letter (0 = A, 25 = Z, 26 = AA)
fn col_to_letter(col: usize) -> String {
    if col < 26 {
        ((b'A' + col as u8) as char).to_string()
    } else {
        let first = (b'A' + (col / 26 - 1) as u8) as char;
        let second = (b'A' + (col % 26) as u8) as char;
        format!("{}{}", first, second)
    }
}

/// Format a cell range as "A1:B10" or "A1" for single cells
fn format_range(start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> String {
    let start = format!("{}{}", col_to_letter(start_col), start_row + 1);
    if start_row == end_row && start_col == end_col {
        start
    } else {
        let end = format!("{}{}", col_to_letter(end_col), end_row + 1);
        format!("{}:{}", start, end)
    }
}

/// Build a diff report showing changes since a specific history entry.
///
/// # Arguments
/// * `history` - The history stack
/// * `since_entry_id` - The entry ID to diff from (changes AFTER this entry)
///
/// Returns None if entry_id not found in undo stack.
pub fn build_diff_since(history: &History, since_entry_id: u64) -> Option<DiffReport> {
    // Find the entry's index in the undo stack
    let entry_index = history.global_index_for_id(since_entry_id)?;
    let since_entry = history.entry_at(entry_index)?;
    let since_entry_label = since_entry.action.label();

    // Collect all entries AFTER this one (index+1 to end)
    let entries_to_process: Vec<&HistoryEntry> = history
        .canonical_entries()
        .iter()
        .skip(entry_index + 1)
        .collect();

    let entries_spanned = entries_to_process.len();

    // Track net changes per cell: (sheet_index, row, col) -> (first_old_value, latest_new_value, ai_touched, ai_source)
    let mut cell_changes: HashMap<(usize, usize, usize), (String, String, bool, Option<String>)> = HashMap::new();

    // Collect structural changes (in order)
    let mut structural_changes = Vec::new();
    let mut named_range_changes = Vec::new();
    let mut validation_changes = Vec::new();
    let mut format_change_count = 0usize;

    // Process entries in chronological order
    for entry in entries_to_process {
        let ai_touched = matches!(&entry.source, MutationSource::Ai(_));
        let ai_source = match &entry.source {
            MutationSource::Ai(meta) => Some(meta.label()),
            MutationSource::Human => None,
        };

        process_action(
            &entry.action,
            ai_touched,
            ai_source,
            &mut cell_changes,
            &mut structural_changes,
            &mut named_range_changes,
            &mut validation_changes,
            &mut format_change_count,
        );
    }

    // Convert cell changes to DiffEntries, splitting into values vs formulas
    let mut value_changes = Vec::new();
    let mut formula_changes = Vec::new();

    for ((sheet_index, row, col), (old_value, new_value, ai_touched, ai_source)) in cell_changes {
        // Skip unchanged cells (can happen if value was changed then changed back)
        if old_value == new_value {
            continue;
        }

        let entry = DiffEntry {
            sheet_index,
            row,
            col,
            old_value: old_value.clone(),
            new_value: new_value.clone(),
            ai_touched,
            ai_source,
        };

        // Classify as formula or value
        if new_value.starts_with('=') || old_value.starts_with('=') {
            formula_changes.push(entry);
        } else {
            value_changes.push(entry);
        }
    }

    // Sort changes by cell address for consistent display
    value_changes.sort_by(|a, b| {
        (a.sheet_index, a.row, a.col).cmp(&(b.sheet_index, b.row, b.col))
    });
    formula_changes.sort_by(|a, b| {
        (a.sheet_index, a.row, a.col).cmp(&(b.sheet_index, b.row, b.col))
    });

    Some(DiffReport {
        since_entry_id,
        since_entry_label,
        entries_spanned,
        value_changes,
        formula_changes,
        structural_changes,
        named_range_changes,
        validation_changes,
        format_change_count,
    })
}

/// Process a single action, updating the change trackers
fn process_action(
    action: &UndoAction,
    ai_touched: bool,
    ai_source: Option<String>,
    cell_changes: &mut HashMap<(usize, usize, usize), (String, String, bool, Option<String>)>,
    structural_changes: &mut Vec<StructuralChange>,
    named_range_changes: &mut Vec<NamedRangeChange>,
    validation_changes: &mut Vec<ValidationChange>,
    format_change_count: &mut usize,
) {
    match action {
        UndoAction::Values { sheet_index, changes } => {
            for change in changes {
                let key = (*sheet_index, change.row, change.col);
                cell_changes
                    .entry(key)
                    .and_modify(|(_, latest_new, touched, source)| {
                        *latest_new = change.new_value.clone();
                        if ai_touched {
                            *touched = true;
                            *source = ai_source.clone();
                        }
                    })
                    .or_insert((
                        change.old_value.clone(),
                        change.new_value.clone(),
                        ai_touched,
                        ai_source.clone(),
                    ));
            }
        }

        UndoAction::Format { patches, .. } => {
            *format_change_count += patches.len();
        }

        UndoAction::RowsInserted { sheet_index, at_row, count } => {
            structural_changes.push(StructuralChange::RowsInserted {
                sheet_index: *sheet_index,
                at_row: *at_row,
                count: *count,
            });
        }

        UndoAction::RowsDeleted { sheet_index, at_row, count, .. } => {
            structural_changes.push(StructuralChange::RowsDeleted {
                sheet_index: *sheet_index,
                at_row: *at_row,
                count: *count,
            });
        }

        UndoAction::ColsInserted { sheet_index, at_col, count } => {
            structural_changes.push(StructuralChange::ColsInserted {
                sheet_index: *sheet_index,
                at_col: *at_col,
                count: *count,
            });
        }

        UndoAction::ColsDeleted { sheet_index, at_col, count, .. } => {
            structural_changes.push(StructuralChange::ColsDeleted {
                sheet_index: *sheet_index,
                at_col: *at_col,
                count: *count,
            });
        }

        UndoAction::SortApplied { sheet_index, new_sort_state, .. } => {
            let (col, ascending) = *new_sort_state;
            structural_changes.push(StructuralChange::Sort {
                sheet_index: *sheet_index,
                column: col,
                ascending,
            });
        }

        UndoAction::SortCleared { sheet_index, .. } => {
            structural_changes.push(StructuralChange::SortCleared {
                sheet_index: *sheet_index,
            });
        }

        UndoAction::NamedRangeCreated { named_range } => {
            named_range_changes.push(NamedRangeChange::Created {
                name: named_range.name.clone(),
            });
        }

        UndoAction::NamedRangeDeleted { named_range } => {
            named_range_changes.push(NamedRangeChange::Deleted {
                name: named_range.name.clone(),
            });
        }

        UndoAction::NamedRangeRenamed { old_name, new_name } => {
            named_range_changes.push(NamedRangeChange::Renamed {
                old_name: old_name.clone(),
                new_name: new_name.clone(),
            });
        }

        UndoAction::NamedRangeDescriptionChanged { name, .. } => {
            named_range_changes.push(NamedRangeChange::DescriptionChanged {
                name: name.clone(),
            });
        }

        UndoAction::ValidationSet { sheet_index, range, new_rule, .. } => {
            let range_str = format_range(range.start_row, range.start_col, range.end_row, range.end_col);
            let rule_desc = format!("{:?}", new_rule.rule_type); // Simple for now
            validation_changes.push(ValidationChange::Set {
                sheet_index: *sheet_index,
                range: range_str,
                rule_desc,
            });
        }

        UndoAction::ValidationCleared { sheet_index, range, .. } => {
            let range_str = format_range(range.start_row, range.start_col, range.end_row, range.end_col);
            validation_changes.push(ValidationChange::Cleared {
                sheet_index: *sheet_index,
                range: range_str,
            });
        }

        UndoAction::ValidationExcluded { sheet_index, range, .. } => {
            let range_str = format_range(range.start_row, range.start_col, range.end_row, range.end_col);
            validation_changes.push(ValidationChange::Excluded {
                sheet_index: *sheet_index,
                range: range_str,
            });
        }

        UndoAction::ValidationExclusionCleared { sheet_index, range, .. } => {
            let range_str = format_range(range.start_row, range.start_col, range.end_row, range.end_col);
            validation_changes.push(ValidationChange::ExclusionCleared {
                sheet_index: *sheet_index,
                range: range_str,
            });
        }

        UndoAction::Group { actions, .. } => {
            for sub_action in actions {
                process_action(
                    sub_action,
                    ai_touched,
                    ai_source.clone(),
                    cell_changes,
                    structural_changes,
                    named_range_changes,
                    validation_changes,
                    format_change_count,
                );
            }
        }

        // Skip column/row sizing changes (visual only)
        UndoAction::ColumnWidthSet { .. } | UndoAction::RowHeightSet { .. } => {}

        // Skip rewind entries (audit-only)
        UndoAction::Rewind { .. } => {}

        // Merge topology changes are structural-like, skip cell tracking
        UndoAction::SetMerges { .. } => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::history::{History, CellChange};

    #[test]
    fn test_build_diff_empty() {
        let history = History::new();
        let result = build_diff_since(&history, 999);
        assert!(result.is_none(), "Should return None for non-existent entry");
    }

    #[test]
    fn test_build_diff_single_change() {
        let mut history = History::new();

        // Add first entry
        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "".to_string(),
            new_value: "A".to_string(),
        }]);

        // Add second entry
        history.record_batch(0, vec![CellChange {
            row: 0, col: 1,
            old_value: "".to_string(),
            new_value: "B".to_string(),
        }]);

        // Get first entry's ID
        let first_id = history.entry_at(0).unwrap().id;

        // Build diff since first entry
        let report = build_diff_since(&history, first_id).unwrap();

        assert_eq!(report.entries_spanned, 1);
        assert_eq!(report.value_changes.len(), 1);
        assert_eq!(report.value_changes[0].col, 1);
        assert_eq!(report.value_changes[0].new_value, "B");
    }

    #[test]
    fn test_build_diff_net_change() {
        let mut history = History::new();

        // First entry
        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "".to_string(),
            new_value: "A".to_string(),
        }]);

        let first_id = history.entry_at(0).unwrap().id;

        // Second entry: change same cell
        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "A".to_string(),
            new_value: "B".to_string(),
        }]);

        // Third entry: change same cell again
        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "B".to_string(),
            new_value: "C".to_string(),
        }]);

        let report = build_diff_since(&history, first_id).unwrap();

        // Should show net change: A → C (skipping intermediate B)
        assert_eq!(report.entries_spanned, 2);
        assert_eq!(report.value_changes.len(), 1);
        assert_eq!(report.value_changes[0].old_value, "A");
        assert_eq!(report.value_changes[0].new_value, "C");
    }

    #[test]
    fn test_build_diff_unchanged_excluded() {
        let mut history = History::new();

        // First entry
        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "".to_string(),
            new_value: "A".to_string(),
        }]);

        let first_id = history.entry_at(0).unwrap().id;

        // Second entry: change cell
        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "A".to_string(),
            new_value: "B".to_string(),
        }]);

        // Third entry: change back to original
        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "B".to_string(),
            new_value: "A".to_string(),
        }]);

        let report = build_diff_since(&history, first_id).unwrap();

        // Net change is A → A, which should be excluded
        assert_eq!(report.entries_spanned, 2);
        assert_eq!(report.value_changes.len(), 0, "Unchanged cells should be excluded");
    }

    #[test]
    fn test_formula_vs_value_classification() {
        let mut history = History::new();

        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "".to_string(),
            new_value: "plain text".to_string(),
        }]);

        let first_id = history.entry_at(0).unwrap().id;

        // Add formula change
        history.record_batch(0, vec![CellChange {
            row: 0, col: 1,
            old_value: "".to_string(),
            new_value: "=SUM(A1:A10)".to_string(),
        }]);

        // Add value change
        history.record_batch(0, vec![CellChange {
            row: 0, col: 2,
            old_value: "".to_string(),
            new_value: "100".to_string(),
        }]);

        let report = build_diff_since(&history, first_id).unwrap();

        assert_eq!(report.formula_changes.len(), 1);
        assert_eq!(report.value_changes.len(), 1);
        assert!(report.formula_changes[0].new_value.starts_with('='));
    }
}
