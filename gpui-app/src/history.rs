/// Undo/Redo history system for spreadsheet operations

use visigrid_engine::cell::CellFormat;
use visigrid_engine::named_range::NamedRange;
use visigrid_engine::provenance::Provenance;
use visigrid_engine::sheet::SheetId;
use visigrid_engine::workbook::Workbook;
use std::time::Instant;

/// Cryptographic fingerprint of the history stack.
///
/// Used to detect concurrent modifications between preview and commit.
/// 128-bit blake3 hash ensures collision resistance (~2^64 birthday bound).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct HistoryFingerprint {
    /// Number of entries in the undo stack
    pub len: usize,
    /// High 64 bits of blake3 hash
    pub hash_hi: u64,
    /// Low 64 bits of blake3 hash
    pub hash_lo: u64,
}

impl std::fmt::Display for HistoryFingerprint {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}:{:016x}{:016x}", self.len, self.hash_hi, self.hash_lo)
    }
}

/// Display-ready entry for the History panel.
/// Pre-computed strings so render doesn't rebuild them.
#[derive(Clone, Debug)]
pub struct HistoryDisplayEntry {
    /// Stable ID for list keying (index in combined undo+redo view)
    pub id: u64,
    /// Primary label (e.g., "Paste", "Fill Down", "Edit cell")
    pub label: String,
    /// Scope description (e.g., "Sheet1!B2:D4", "47 cells")
    pub scope: String,
    /// Action-specific summary (e.g., "Whole number 1-100", "Column C ascending")
    pub summary: Option<String>,
    /// Location string for display (e.g., "A1:B10") - computed from affected_range
    pub location: Option<String>,
    /// When this action occurred
    pub timestamp: Instant,
    /// Lua snippet if provenance exists (from scripting)
    pub lua: Option<String>,
    /// Auto-generated Lua from action (Phase 9A provenance export)
    /// Available for all replayable actions, even without explicit provenance
    pub generated_lua: Option<String>,
    /// Whether this entry has provenance
    pub is_provenanced: bool,
    /// Whether this entry can be undone (vs already undone/redoable)
    pub is_undoable: bool,
    /// Sheet index for the action (if applicable)
    pub sheet_index: Option<usize>,
    /// Affected cells: (row, col, old_value, new_value)
    /// For value changes, shows before/after. For format changes, may be empty.
    pub affected_cells: Vec<(usize, usize, String, String)>,
    /// Bounding box of affected cells (start_row, start_col, end_row, end_col)
    pub affected_range: Option<(usize, usize, usize, usize)>,
    /// AI source label if this was an AI-generated mutation (e.g., "AI: OpenAI gpt-4o")
    pub ai_source: Option<String>,
}

#[derive(Clone, Debug)]
pub struct CellChange {
    pub row: usize,
    pub col: usize,
    pub old_value: String,
    pub new_value: String,
}

/// A patch for a single cell's format (before/after snapshot)
#[derive(Clone, Debug)]
pub struct CellFormatPatch {
    pub row: usize,
    pub col: usize,
    pub before: CellFormat,
    pub after: CellFormat,
}

/// Kind of format action (for coalescing)
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum FormatActionKind {
    Bold,
    Italic,
    Underline,
    Strikethrough,
    Font,
    Alignment,
    VerticalAlignment,
    TextOverflow,
    NumberFormat,
    DecimalPlaces,  // Special: coalesces rapidly
    BackgroundColor,
    Border,
    PasteFormats,  // Paste Special > Formats
    ClearFormatting,
}

/// An undoable action
#[derive(Clone, Debug)]
pub enum UndoAction {
    /// Cell value changes
    Values {
        sheet_index: usize,
        changes: Vec<CellChange>,
    },
    /// Cell format changes
    Format {
        sheet_index: usize,
        patches: Vec<CellFormatPatch>,
        kind: FormatActionKind,
        description: String,
    },
    /// Named range deleted (for undo)
    NamedRangeDeleted {
        named_range: NamedRange,
    },
    /// Named range created (for undo - delete it)
    NamedRangeCreated {
        /// Full named range payload (needed for forward replay)
        named_range: NamedRange,
    },
    /// Named range renamed (for undo)
    NamedRangeRenamed {
        old_name: String,
        new_name: String,
    },
    /// Named range description changed (for undo)
    NamedRangeDescriptionChanged {
        name: String,
        old_description: Option<String>,
        new_description: Option<String>,
    },
    /// Grouped actions that should be undone/redone together
    Group {
        actions: Vec<UndoAction>,
        description: String,
    },
    /// Rows inserted (for undo: delete the inserted rows)
    RowsInserted {
        sheet_index: usize,
        at_row: usize,
        count: usize,
    },
    /// Rows deleted (for undo: re-insert rows and restore cell data)
    RowsDeleted {
        sheet_index: usize,
        at_row: usize,
        count: usize,
        /// Deleted cell data: (row, col, value, format)
        deleted_cells: Vec<(usize, usize, String, CellFormat)>,
        /// Deleted row heights: (row, height)
        deleted_row_heights: Vec<(usize, f32)>,
    },
    /// Columns inserted (for undo: delete the inserted columns)
    ColsInserted {
        sheet_index: usize,
        at_col: usize,
        count: usize,
    },
    /// Columns deleted (for undo: re-insert columns and restore cell data)
    ColsDeleted {
        sheet_index: usize,
        at_col: usize,
        count: usize,
        /// Deleted cell data: (row, col, value, format)
        deleted_cells: Vec<(usize, usize, String, CellFormat)>,
        /// Deleted column widths: (col, width)
        deleted_col_widths: Vec<(usize, f32)>,
    },
    /// Column width changed (for undo: restore old width)
    ColumnWidthSet {
        /// Sheet ID (stable across reorder/delete, unlike index)
        sheet_id: SheetId,
        col: usize,
        /// Old width (None = was using default)
        old: Option<f32>,
        /// New width (None = reset to default)
        new: Option<f32>,
    },
    /// Row height changed (for undo: restore old height)
    RowHeightSet {
        /// Sheet ID (stable across reorder/delete, unlike index)
        sheet_id: SheetId,
        row: usize,
        /// Old height (None = was using default)
        old: Option<f32>,
        /// New height (None = reset to default)
        new: Option<f32>,
    },
    /// Sort applied (for undo: restore previous row order)
    SortApplied {
        /// Sheet where sort was applied (required for replay)
        sheet_index: usize,
        /// Previous row order before sorting
        previous_row_order: Vec<usize>,
        /// Previous sort state (column and direction)
        previous_sort_state: Option<(usize, bool)>, // (column, is_ascending)
        /// New row order after sorting (for redo)
        new_row_order: Vec<usize>,
        /// New sort state (column and direction) for redo
        new_sort_state: (usize, bool), // (column, is_ascending)
    },
    /// Sort cleared (for undo: restore previous sort state)
    SortCleared {
        /// Sheet where sort was cleared
        sheet_index: usize,
        /// Previous row order before clearing
        previous_row_order: Vec<usize>,
        /// Previous sort state (column, is_ascending)
        previous_sort_state: (usize, bool),
    },
    /// Validation rule set (for undo: restore previous rules)
    ValidationSet {
        sheet_index: usize,
        /// Target range where validation was applied
        range: visigrid_engine::validation::CellRange,
        /// Rules that were removed (for undo)
        previous_rules: Vec<(visigrid_engine::validation::CellRange, visigrid_engine::validation::ValidationRule)>,
        /// The new rule that was set (for redo)
        new_rule: visigrid_engine::validation::ValidationRule,
    },
    /// Validation rules cleared (for undo: restore the cleared rules)
    ValidationCleared {
        sheet_index: usize,
        /// Target range where validation was cleared
        range: visigrid_engine::validation::CellRange,
        /// Rules that were cleared (for undo: restore these)
        cleared_rules: Vec<(visigrid_engine::validation::CellRange, visigrid_engine::validation::ValidationRule)>,
    },
    /// Validation exclusion added (for undo: remove the exclusion)
    ValidationExcluded {
        sheet_index: usize,
        /// Range that was excluded from validation
        range: visigrid_engine::validation::CellRange,
    },
    /// Validation exclusion cleared (for undo: restore the exclusions)
    ValidationExclusionCleared {
        sheet_index: usize,
        /// Target range where exclusions were cleared
        range: visigrid_engine::validation::CellRange,
        /// Exclusions that were cleared (for undo: restore these)
        cleared_exclusions: Vec<visigrid_engine::validation::CellRange>,
    },
    /// Hard rewind: workbook reverted to historical state (audit-only, cannot undo)
    /// This action is for provenance tracking - it records that a rewind occurred
    /// Hard rewind: workbook reverted to historical state.
    /// This is an audit-only action - cannot be undone/redone.
    /// Contains full provenance for courtroom-grade explainability.
    Rewind {
        /// ID of the history entry we rewound "before"
        target_entry_id: u64,
        /// Index of the target entry in the original history stack
        target_index: usize,
        /// Summary of the target action (what we rewound "before")
        target_action_summary: String,
        /// How many history entries were discarded
        discarded_count: usize,
        /// History length before rewind
        old_history_len: usize,
        /// History length after rewind (should be old - discarded + 1 for this entry)
        new_history_len: usize,
        /// Wall-clock timestamp when rewind was committed (ISO 8601 format)
        timestamp_utc: String,
        /// Number of actions that were replayed to build the preview
        preview_replay_count: usize,
        /// Time spent building the preview snapshot (milliseconds)
        preview_build_ms: u64,
    },
}

impl UndoAction {
    /// Generate a human-readable label for this action.
    pub fn label(&self) -> String {
        match self {
            UndoAction::Values { changes, .. } => {
                if changes.len() == 1 {
                    "Edit cell".to_string()
                } else {
                    format!("Edit {} cells", changes.len())
                }
            }
            UndoAction::Format { description, patches, .. } => {
                if patches.len() == 1 {
                    description.clone()
                } else {
                    format!("{} ({} cells)", description, patches.len())
                }
            }
            UndoAction::NamedRangeDeleted { named_range } => {
                format!("Delete range '{}'", named_range.name)
            }
            UndoAction::NamedRangeCreated { named_range } => {
                format!("Create range '{}'", named_range.name)
            }
            UndoAction::NamedRangeRenamed { old_name, new_name } => {
                format!("Rename '{}' to '{}'", old_name, new_name)
            }
            UndoAction::NamedRangeDescriptionChanged { name, .. } => {
                format!("Change '{}' description", name)
            }
            UndoAction::Group { description, .. } => {
                description.clone()
            }
            UndoAction::RowsInserted { count, .. } => {
                if *count == 1 {
                    "Insert row".to_string()
                } else {
                    format!("Insert {} rows", count)
                }
            }
            UndoAction::RowsDeleted { count, .. } => {
                if *count == 1 {
                    "Delete row".to_string()
                } else {
                    format!("Delete {} rows", count)
                }
            }
            UndoAction::ColsInserted { count, .. } => {
                if *count == 1 {
                    "Insert column".to_string()
                } else {
                    format!("Insert {} columns", count)
                }
            }
            UndoAction::ColsDeleted { count, .. } => {
                if *count == 1 {
                    "Delete column".to_string()
                } else {
                    format!("Delete {} columns", count)
                }
            }
            UndoAction::ColumnWidthSet { .. } => {
                "Set column width".to_string()
            }
            UndoAction::RowHeightSet { .. } => {
                "Set row height".to_string()
            }
            UndoAction::SortApplied { .. } => {
                "Sort".to_string()
            }
            UndoAction::SortCleared { .. } => {
                "Clear sort".to_string()
            }
            UndoAction::ValidationSet { range, .. } => {
                let count = range.cell_count();
                if count == 1 {
                    "Set validation".to_string()
                } else {
                    format!("Set validation ({} cells)", count)
                }
            }
            UndoAction::ValidationCleared { range, .. } => {
                let count = range.cell_count();
                if count == 1 {
                    "Clear validation".to_string()
                } else {
                    format!("Clear validation ({} cells)", count)
                }
            }
            UndoAction::ValidationExcluded { range, .. } => {
                let count = range.cell_count();
                if count == 1 {
                    "Exclude from validation".to_string()
                } else {
                    format!("Exclude from validation ({} cells)", count)
                }
            }
            UndoAction::ValidationExclusionCleared { range, .. } => {
                let count = range.cell_count();
                if count == 1 {
                    "Clear exclusion".to_string()
                } else {
                    format!("Clear exclusions ({} cells)", count)
                }
            }
            UndoAction::Rewind { discarded_count, .. } => {
                format!("Rewind (discarded {} change{})", discarded_count, if *discarded_count == 1 { "" } else { "s" })
            }
        }
    }

    /// Generate an action-specific summary for the detail view.
    /// Returns None for simple actions where label is sufficient.
    pub fn summary(&self) -> Option<String> {
        match self {
            UndoAction::ValidationSet { range, new_rule, .. } => {
                let range_str = format_range(range.start_row, range.start_col, range.end_row, range.end_col);
                let rule_desc = format_validation_rule(new_rule);
                Some(format!("{} → {}", rule_desc, range_str))
            }
            UndoAction::ValidationCleared { range, cleared_rules, .. } => {
                let range_str = format_range(range.start_row, range.start_col, range.end_row, range.end_col);
                let count = cleared_rules.len();
                Some(format!("Cleared {} rule(s) from {}", count, range_str))
            }
            UndoAction::ValidationExcluded { range, .. } => {
                let range_str = format_range(range.start_row, range.start_col, range.end_row, range.end_col);
                Some(format!("+ {}", range_str))
            }
            UndoAction::ValidationExclusionCleared { range, .. } => {
                let range_str = format_range(range.start_row, range.start_col, range.end_row, range.end_col);
                Some(format!("- {}", range_str))
            }
            UndoAction::SortApplied { new_sort_state, .. } => {
                let (col, is_asc) = new_sort_state;
                let col_letter = col_to_letter(*col);
                let dir = if *is_asc { "ascending" } else { "descending" };
                Some(format!("Column {} {}", col_letter, dir))
            }
            UndoAction::SortCleared { previous_sort_state, .. } => {
                let (col, is_asc) = previous_sort_state;
                let col_letter = col_to_letter(*col);
                let dir = if *is_asc { "ascending" } else { "descending" };
                Some(format!("Was column {} {}", col_letter, dir))
            }
            UndoAction::RowsInserted { at_row, count, .. } => {
                Some(format!("{} row(s) at row {}", count, at_row + 1))
            }
            UndoAction::RowsDeleted { at_row, count, .. } => {
                Some(format!("{} row(s) at row {}", count, at_row + 1))
            }
            UndoAction::ColsInserted { at_col, count, .. } => {
                let col_letter = col_to_letter(*at_col);
                Some(format!("{} column(s) at {}", count, col_letter))
            }
            UndoAction::ColsDeleted { at_col, count, .. } => {
                let col_letter = col_to_letter(*at_col);
                Some(format!("{} column(s) at {}", count, col_letter))
            }
            UndoAction::ColumnWidthSet { col, old, new, .. } => {
                let col_letter = col_to_letter(*col);
                // Use unit-free numbers (internal units, not guaranteed to match Excel)
                let old_str = old.map(|w| format!("{:.0}", w)).unwrap_or_else(|| "default".to_string());
                let new_str = new.map(|w| format!("{:.0}", w)).unwrap_or_else(|| "default".to_string());
                Some(format!("Col {}: {} → {}", col_letter, old_str, new_str))
            }
            UndoAction::RowHeightSet { row, old, new, .. } => {
                // Use unit-free numbers (internal units, not guaranteed to match Excel)
                let old_str = old.map(|h| format!("{:.0}", h)).unwrap_or_else(|| "default".to_string());
                let new_str = new.map(|h| format!("{:.0}", h)).unwrap_or_else(|| "default".to_string());
                Some(format!("Row {}: {} → {}", row + 1, old_str, new_str))
            }
            // Simple actions - label is sufficient
            _ => None,
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

/// Format a validation rule for display
fn format_validation_rule(rule: &visigrid_engine::validation::ValidationRule) -> String {
    use visigrid_engine::validation::{ValidationType, ListSource};

    let blank_part = if rule.ignore_blank { " (allow blank)" } else { "" };

    match &rule.rule_type {
        ValidationType::WholeNumber(constraint) => {
            let op_str = format_numeric_constraint(constraint);
            format!("Whole number {}{}", op_str, blank_part)
        }
        ValidationType::Decimal(constraint) => {
            let op_str = format_numeric_constraint(constraint);
            format!("Decimal {}{}", op_str, blank_part)
        }
        ValidationType::List(source) => {
            let preview = match source {
                ListSource::Inline(items) => {
                    let display: Vec<_> = items.iter().take(3).cloned().collect();
                    if items.len() > 3 {
                        format!("{}, ... ({} items)", display.join(", "), items.len())
                    } else {
                        display.join(", ")
                    }
                }
                ListSource::Range(range_str) => range_str.clone(),
                ListSource::NamedRange(name) => name.clone(),
            };
            format!("List: {}{}", preview, blank_part)
        }
        ValidationType::Date(constraint) => {
            let op_str = format_numeric_constraint(constraint);
            format!("Date {}{}", op_str, blank_part)
        }
        ValidationType::Time(constraint) => {
            let op_str = format_numeric_constraint(constraint);
            format!("Time {}{}", op_str, blank_part)
        }
        ValidationType::TextLength(constraint) => {
            let op_str = format_numeric_constraint(constraint);
            format!("Text length {}{}", op_str, blank_part)
        }
        ValidationType::Custom(formula) => {
            format!("Custom: {}{}", formula, blank_part)
        }
    }
}

/// Format a numeric constraint for display
fn format_numeric_constraint(constraint: &visigrid_engine::validation::NumericConstraint) -> String {
    use visigrid_engine::validation::ComparisonOperator;

    let v1_str = format_constraint_value(&constraint.value1);
    let v2_str = constraint.value2.as_ref().map(format_constraint_value).unwrap_or_default();

    match constraint.operator {
        ComparisonOperator::Between => format!("between {} and {}", v1_str, v2_str),
        ComparisonOperator::NotBetween => format!("not between {} and {}", v1_str, v2_str),
        ComparisonOperator::EqualTo => format!("= {}", v1_str),
        ComparisonOperator::NotEqualTo => format!("≠ {}", v1_str),
        ComparisonOperator::GreaterThan => format!("> {}", v1_str),
        ComparisonOperator::LessThan => format!("< {}", v1_str),
        ComparisonOperator::GreaterThanOrEqual => format!("≥ {}", v1_str),
        ComparisonOperator::LessThanOrEqual => format!("≤ {}", v1_str),
    }
}

/// Format a constraint value for display
fn format_constraint_value(value: &visigrid_engine::validation::ConstraintValue) -> String {
    use visigrid_engine::validation::ConstraintValue;
    match value {
        ConstraintValue::Number(n) => {
            if n.fract() == 0.0 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        ConstraintValue::CellRef(r) => r.clone(),
        ConstraintValue::Formula(f) => f.clone(),
    }
}

/// Source of a mutation (for provenance tracking)
#[derive(Clone, Debug, Default)]
pub enum MutationSource {
    /// Human entered value manually (default)
    #[default]
    Human,
    /// AI-generated (Ask AI feature)
    Ai(AiMutationMeta),
}

/// Metadata for AI-generated mutations (minimal, no prompts/context stored)
#[derive(Clone, Debug)]
pub struct AiMutationMeta {
    /// Provider used (e.g., "openai")
    pub provider: String,
    /// Model used
    pub model: String,
    /// Whether privacy mode was enabled
    pub privacy_mode: bool,
    /// Request ID for correlation (optional)
    pub request_id: Option<String>,
    /// Context selection mode ("selection", "region", "used_range")
    pub context_mode: String,
    /// Truncation applied ("none", "rows", "cols", "both")
    pub truncation: String,
}

impl AiMutationMeta {
    /// Short label for display (e.g., "AI: OpenAI gpt-4o")
    pub fn label(&self) -> String {
        format!("AI: {} {}", self.provider, self.model)
    }
}

#[derive(Clone, Debug)]
pub struct HistoryEntry {
    /// Stable ID for this entry (monotonic, survives undo/redo moves)
    pub id: u64,
    pub action: UndoAction,
    pub timestamp: Instant,
    /// Lua provenance for multi-cell operations (Phase 4)
    pub provenance: Option<Provenance>,
    /// Source of mutation (human or AI) for provenance tracking
    pub source: MutationSource,
}

/// Coalescing window for rapid format changes (e.g., decimal +/-)
const COALESCE_WINDOW_MS: u128 = 500;

pub struct History {
    undo_stack: Vec<HistoryEntry>,
    redo_stack: Vec<HistoryEntry>,
    max_entries: usize,
    /// Save point for dirty detection: undo_stack length when document was saved
    save_point: usize,
    /// Monotonic counter for stable entry IDs
    next_id: u64,
}

impl Default for History {
    fn default() -> Self {
        Self::new()
    }
}

impl History {
    pub fn new() -> Self {
        Self {
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            max_entries: 100,
            save_point: 0,
            next_id: 1,
        }
    }

    fn next_entry_id(&mut self) -> u64 {
        let id = self.next_id;
        self.next_id += 1;
        id
    }

    /// Mark current position as the save point (document is now "clean")
    pub fn mark_saved(&mut self) {
        self.save_point = self.undo_stack.len();
    }

    /// Check if document has unsaved changes (dirty).
    /// Dirty = current history position differs from save point.
    pub fn is_dirty(&self) -> bool {
        self.undo_stack.len() != self.save_point
    }

    /// Get the save point (for debugging/testing)
    pub fn save_point(&self) -> usize {
        self.save_point
    }

    /// Record a single cell value change (human source)
    pub fn record_change(&mut self, sheet_index: usize, row: usize, col: usize, old_value: String, new_value: String) {
        self.record_change_with_source(sheet_index, row, col, old_value, new_value, MutationSource::Human);
    }

    /// Record a single cell value change with explicit source
    pub fn record_change_with_source(
        &mut self,
        sheet_index: usize,
        row: usize,
        col: usize,
        old_value: String,
        new_value: String,
        source: MutationSource,
    ) {
        if old_value == new_value {
            return;
        }

        let id = self.next_entry_id();
        let entry = HistoryEntry {
            id,
            action: UndoAction::Values {
                sheet_index,
                changes: vec![CellChange { row, col, old_value, new_value }],
            },
            timestamp: Instant::now(),
            provenance: None,  // Single cell edits don't need Lua provenance
            source,
        };
        self.push_entry(entry);
    }

    /// Record multiple cell value changes as a single undoable operation
    pub fn record_batch(&mut self, sheet_index: usize, changes: Vec<CellChange>) {
        self.record_batch_with_provenance(sheet_index, changes, None);
    }

    /// Record multiple cell value changes with optional Lua provenance
    pub fn record_batch_with_provenance(&mut self, sheet_index: usize, changes: Vec<CellChange>, provenance: Option<Provenance>) {
        if changes.is_empty() {
            return;
        }

        let id = self.next_entry_id();
        let entry = HistoryEntry {
            id,
            action: UndoAction::Values { sheet_index, changes },
            timestamp: Instant::now(),
            provenance,
            source: MutationSource::Human,
        };
        self.push_entry(entry);
    }

    /// Record format changes with coalescing support
    pub fn record_format(&mut self, sheet_index: usize, patches: Vec<CellFormatPatch>, kind: FormatActionKind, description: String) {
        if patches.is_empty() {
            return;
        }

        let now = Instant::now();

        // Try to coalesce with previous entry if:
        // 1. Same sheet
        // 2. Same kind (especially DecimalPlaces)
        // 3. Within time window
        // 4. Same cell positions
        if let Some(last) = self.undo_stack.last_mut() {
            if let UndoAction::Format { sheet_index: last_sheet, patches: last_patches, kind: last_kind, description: _ } = &mut last.action {
                if *last_sheet == sheet_index && *last_kind == kind && last.timestamp.elapsed().as_millis() < COALESCE_WINDOW_MS {
                    // Check if same cells
                    if Self::same_cell_positions(last_patches, &patches) {
                        // Coalesce: keep original 'before', update to new 'after'
                        for (old_patch, new_patch) in last_patches.iter_mut().zip(patches.iter()) {
                            old_patch.after = new_patch.after.clone();
                        }
                        last.timestamp = now;
                        // Clear redo stack since we modified history
                        self.redo_stack.clear();
                        return;
                    }
                }
            }
        }

        // No coalescing, create new entry
        let id = self.next_entry_id();
        let entry = HistoryEntry {
            id,
            action: UndoAction::Format { sheet_index, patches, kind, description },
            timestamp: now,
            provenance: None,  // Format changes don't need Lua provenance
            source: MutationSource::Human,
        };
        self.push_entry(entry);
    }

    /// Record format changes with optional Lua provenance (for Paste Formats)
    pub fn record_format_with_provenance(
        &mut self,
        sheet_index: usize,
        patches: Vec<CellFormatPatch>,
        kind: FormatActionKind,
        description: String,
        provenance: Option<Provenance>,
    ) {
        if patches.is_empty() {
            return;
        }

        let id = self.next_entry_id();
        let entry = HistoryEntry {
            id,
            action: UndoAction::Format { sheet_index, patches, kind, description },
            timestamp: Instant::now(),
            provenance,
            source: MutationSource::Human,
        };
        self.push_entry(entry);
    }

    /// Record a named range action (create, delete, rename)
    pub fn record_named_range_action(&mut self, action: UndoAction) {
        self.record_action_with_provenance(action, None);
    }

    /// Record any action with optional Lua provenance
    pub fn record_action_with_provenance(&mut self, action: UndoAction, provenance: Option<Provenance>) {
        let id = self.next_entry_id();
        let entry = HistoryEntry {
            id,
            action,
            timestamp: Instant::now(),
            provenance,
            source: MutationSource::Human,
        };
        self.push_entry(entry);
    }

    /// Check if two patch lists affect the same cells
    fn same_cell_positions(a: &[CellFormatPatch], b: &[CellFormatPatch]) -> bool {
        if a.len() != b.len() {
            return false;
        }
        // Patches are in same order if from same selection iteration
        a.iter().zip(b.iter()).all(|(pa, pb)| pa.row == pb.row && pa.col == pb.col)
    }

    fn push_entry(&mut self, entry: HistoryEntry) {
        self.undo_stack.push(entry);
        self.redo_stack.clear();
        // Note: save_point may now be unreachable (undo→edit diverges from saved state).
        // That's correct - is_dirty() will return true until next save.

        // Limit history size (remove oldest entries from front)
        if self.undo_stack.len() > self.max_entries {
            self.undo_stack.remove(0);
            // Indices shifted down - adjust save_point (saturating to 0)
            self.save_point = self.save_point.saturating_sub(1);
        }
    }

    /// Pop the last entry for undo
    pub fn undo(&mut self) -> Option<HistoryEntry> {
        if let Some(entry) = self.undo_stack.pop() {
            self.redo_stack.push(entry.clone());
            Some(entry)
        } else {
            None
        }
    }

    /// Pop from redo stack
    pub fn redo(&mut self) -> Option<HistoryEntry> {
        if let Some(entry) = self.redo_stack.pop() {
            self.undo_stack.push(entry.clone());
            Some(entry)
        } else {
            None
        }
    }

    pub fn can_undo(&self) -> bool {
        !self.undo_stack.is_empty()
    }

    pub fn can_redo(&self) -> bool {
        !self.redo_stack.is_empty()
    }

    /// Get entries for history panel display (most recent first).
    /// Returns (entry, is_undoable) tuples - undoable entries are in undo stack.
    pub fn entries_for_display(&self) -> Vec<(&HistoryEntry, bool)> {
        // Combine: redo stack (top = most recently undone) + undo stack (top = most recent)
        // Display order: most recent action first
        let mut entries: Vec<(&HistoryEntry, bool)> = Vec::new();

        // Undo stack entries (can be undone)
        for entry in self.undo_stack.iter().rev() {
            entries.push((entry, true));
        }

        // Redo stack entries (already undone, can be redone)
        for entry in self.redo_stack.iter().rev() {
            entries.push((entry, false));
        }

        entries
    }

    /// Get pre-computed display entries for History panel.
    /// Labels/scope/lua are computed once here, not in render.
    /// Uses stable entry.id for keying (survives undo/redo moves).
    pub fn display_entries(&self) -> Vec<HistoryDisplayEntry> {
        let mut entries = Vec::new();

        // Undo stack entries (most recent first) - these can be undone
        for entry in self.undo_stack.iter().rev() {
            entries.push(Self::to_display_entry(entry, true));
        }

        // Redo stack entries (already undone) - these can be redone
        for entry in self.redo_stack.iter().rev() {
            entries.push(Self::to_display_entry(entry, false));
        }

        // Invariant check: IDs should be unique
        #[cfg(debug_assertions)]
        {
            let mut seen_ids = std::collections::HashSet::new();
            for entry in &entries {
                debug_assert!(
                    seen_ids.insert(entry.id),
                    "Duplicate history entry ID: {}",
                    entry.id
                );
            }
        }

        entries
    }

    /// Convert a HistoryEntry to a HistoryDisplayEntry.
    /// Uses entry.id for stable keying across undo/redo operations.
    pub fn to_display_entry(entry: &HistoryEntry, is_undoable: bool) -> HistoryDisplayEntry {
        let (label, scope, lua, is_provenanced) = if let Some(ref prov) = entry.provenance {
            (prov.label.clone(), prov.scope.clone(), Some(prov.lua.clone()), true)
        } else {
            // Fallback to UndoAction::label()
            (entry.action.label(), String::new(), None, false)
        };

        // Extract affected cells and range from action
        let (sheet_index, affected_cells, affected_range) = Self::extract_action_details(&entry.action);

        // Generate action-specific summary
        let summary = entry.action.summary();

        // Generate location string from affected range
        let location = affected_range.map(|(sr, sc, er, ec)| format_range(sr, sc, er, ec));

        // Generate Lua from action (Phase 9A provenance export)
        let generated_lua = entry.action.to_lua();

        // Extract AI source label if applicable
        let ai_source = match &entry.source {
            MutationSource::Human => None,
            MutationSource::Ai(meta) => Some(meta.label()),
        };

        HistoryDisplayEntry {
            id: entry.id,  // Use stable entry ID, not position
            label,
            scope,
            summary,
            location,
            timestamp: entry.timestamp,
            lua,
            generated_lua,
            is_provenanced,
            is_undoable,
            sheet_index,
            affected_cells,
            affected_range,
            ai_source,
        }
    }

    /// Extract sheet index, affected cells, and bounding range from an action.
    fn extract_action_details(action: &UndoAction) -> (Option<usize>, Vec<(usize, usize, String, String)>, Option<(usize, usize, usize, usize)>) {
        match action {
            UndoAction::Values { sheet_index, changes } => {
                let cells: Vec<_> = changes.iter()
                    .map(|c| (c.row, c.col, c.old_value.clone(), c.new_value.clone()))
                    .collect();
                let range = Self::bounding_box(&cells);
                (Some(*sheet_index), cells, range)
            }
            UndoAction::Format { sheet_index, patches, .. } => {
                let cells: Vec<_> = patches.iter()
                    .map(|p| (p.row, p.col, String::new(), String::new()))
                    .collect();
                let range = Self::bounding_box(&cells);
                (Some(*sheet_index), cells, range)
            }
            UndoAction::ValidationSet { sheet_index, range, .. } => {
                let bbox = (range.start_row, range.start_col, range.end_row, range.end_col);
                (Some(*sheet_index), vec![], Some(bbox))
            }
            UndoAction::ValidationCleared { sheet_index, range, .. } => {
                let bbox = (range.start_row, range.start_col, range.end_row, range.end_col);
                (Some(*sheet_index), vec![], Some(bbox))
            }
            UndoAction::ValidationExcluded { sheet_index, range, .. } => {
                let bbox = (range.start_row, range.start_col, range.end_row, range.end_col);
                (Some(*sheet_index), vec![], Some(bbox))
            }
            UndoAction::ValidationExclusionCleared { sheet_index, range, .. } => {
                let bbox = (range.start_row, range.start_col, range.end_row, range.end_col);
                (Some(*sheet_index), vec![], Some(bbox))
            }
            UndoAction::RowsInserted { sheet_index, at_row, count } => {
                // Highlight the inserted rows (full width, arbitrary column span)
                let bbox = (*at_row, 0, at_row + count - 1, 25); // Show first 26 columns
                (Some(*sheet_index), vec![], Some(bbox))
            }
            UndoAction::RowsDeleted { sheet_index, at_row: _, deleted_cells, .. } => {
                if deleted_cells.is_empty() {
                    (Some(*sheet_index), vec![], None)
                } else {
                    let cells: Vec<_> = deleted_cells.iter()
                        .map(|(r, c, v, _)| (*r, *c, v.clone(), String::new()))
                        .collect();
                    let range = Self::bounding_box(&cells);
                    (Some(*sheet_index), cells, range)
                }
            }
            UndoAction::ColsInserted { sheet_index, at_col, count } => {
                // Highlight the inserted columns
                let bbox = (0, *at_col, 99, at_col + count - 1); // Show first 100 rows
                (Some(*sheet_index), vec![], Some(bbox))
            }
            UndoAction::ColsDeleted { sheet_index, at_col: _, deleted_cells, .. } => {
                if deleted_cells.is_empty() {
                    (Some(*sheet_index), vec![], None)
                } else {
                    let cells: Vec<_> = deleted_cells.iter()
                        .map(|(r, c, v, _)| (*r, *c, v.clone(), String::new()))
                        .collect();
                    let range = Self::bounding_box(&cells);
                    (Some(*sheet_index), cells, range)
                }
            }
            UndoAction::ColumnWidthSet { col, .. } => {
                // Highlight the affected column (first 100 rows)
                // Note: sheet_id not resolved to index here; caller must handle
                let bbox = (0, *col, 99, *col);
                (None, vec![], Some(bbox))
            }
            UndoAction::RowHeightSet { row, .. } => {
                // Highlight the affected row (first 26 columns)
                // Note: sheet_id not resolved to index here; caller must handle
                let bbox = (*row, 0, *row, 25);
                (None, vec![], Some(bbox))
            }
            UndoAction::Group { actions, .. } => {
                // For groups, combine all sub-action details
                let mut all_cells = Vec::new();
                let mut sheet = None;
                for sub in actions {
                    let (s, cells, _) = Self::extract_action_details(sub);
                    if sheet.is_none() { sheet = s; }
                    all_cells.extend(cells);
                }
                let range = Self::bounding_box(&all_cells);
                (sheet, all_cells, range)
            }
            _ => (None, vec![], None),
        }
    }

    /// Compute bounding box from a list of cells
    fn bounding_box(cells: &[(usize, usize, String, String)]) -> Option<(usize, usize, usize, usize)> {
        if cells.is_empty() {
            return None;
        }
        let mut min_row = usize::MAX;
        let mut min_col = usize::MAX;
        let mut max_row = 0;
        let mut max_col = 0;
        for (row, col, _, _) in cells {
            min_row = min_row.min(*row);
            min_col = min_col.min(*col);
            max_row = max_row.max(*row);
            max_col = max_col.max(*col);
        }
        Some((min_row, min_col, max_row, max_col))
    }

    /// Get the number of entries in the undo stack.
    pub fn undo_count(&self) -> usize {
        self.undo_stack.len()
    }

    /// Get the number of entries in the redo stack.
    pub fn redo_count(&self) -> usize {
        self.redo_stack.len()
    }

    pub fn clear(&mut self) {
        self.undo_stack.clear();
        self.redo_stack.clear();
        self.save_point = 0;
        self.next_id = 1;
    }

    // ========================================================================
    // Soft-Rewind Preview (Phase 8A)
    // ========================================================================

    /// Get the canonical history entries in chronological order (oldest first).
    /// This is the undo_stack in order, since we push newest to end.
    pub fn canonical_entries(&self) -> &[HistoryEntry] {
        &self.undo_stack
    }

    /// Find the global index for a given entry ID.
    /// Returns None if entry not found in undo stack.
    pub fn global_index_for_id(&self, entry_id: u64) -> Option<usize> {
        self.undo_stack.iter().position(|e| e.id == entry_id)
    }

    /// Get entry by global index.
    pub fn entry_at(&self, index: usize) -> Option<&HistoryEntry> {
        self.undo_stack.get(index)
    }

    /// Compute a fingerprint of current history state.
    /// Used to detect concurrent changes between preview start and commit.
    /// Format: (undo_stack_len, sum of entry IDs)
    /// Compute a cryptographic fingerprint of the history stack.
    ///
    /// Returns (len, hash_hi, hash_lo) where hash is a 128-bit blake3 digest
    /// of the ordered sequence of (entry_id, action_kind_tag).
    ///
    /// This is order-sensitive: different orderings produce different hashes.
    /// Collisions are astronomically unlikely (~2^64 birthday bound).
    pub fn fingerprint(&self) -> HistoryFingerprint {
        use blake3::Hasher;

        let len = self.undo_stack.len();
        let mut hasher = Hasher::new();

        // Hash length first (prevents length-extension issues)
        hasher.update(&(len as u64).to_le_bytes());

        // Hash each entry: (id, kind_tag) in order
        for entry in &self.undo_stack {
            hasher.update(&entry.id.to_le_bytes());
            let kind_tag = entry.action.kind().tag();
            hasher.update(&[kind_tag]);
        }

        let hash = hasher.finalize();
        let bytes = hash.as_bytes();

        // Extract first 128 bits as two u64s
        let hash_hi = u64::from_le_bytes(bytes[0..8].try_into().unwrap());
        let hash_lo = u64::from_le_bytes(bytes[8..16].try_into().unwrap());

        HistoryFingerprint { len, hash_hi, hash_lo }
    }

    /// Check if history fingerprint matches current state.
    /// Returns false if history changed since fingerprint was taken.
    pub fn fingerprint_matches(&self, fingerprint: &HistoryFingerprint) -> bool {
        self.fingerprint() == *fingerprint
    }

    /// Truncate history at index, keeping entries [0..index).
    /// Appends a new audit entry after truncation.
    /// Clears redo stack since truncated entries cannot be redone.
    ///
    /// # Arguments
    /// * `truncate_at` - Index to truncate at (entries [0..truncate_at) are kept)
    /// * `target_entry_id` - ID of the entry we're rewinding "before"
    /// * `target_index` - Original index of the target entry
    /// * `target_action_summary` - Summary of the target action
    /// * `preview_replay_count` - How many actions were replayed to build preview
    /// * `preview_build_ms` - Time spent building preview (milliseconds)
    pub fn truncate_and_append_rewind(
        &mut self,
        truncate_at: usize,
        target_entry_id: u64,
        target_index: usize,
        target_action_summary: String,
        preview_replay_count: usize,
        preview_build_ms: u64,
    ) {
        let old_len = self.undo_stack.len();
        let discarded_count = old_len.saturating_sub(truncate_at);

        // Truncate history
        self.undo_stack.truncate(truncate_at);

        // Clear redo stack - truncated entries cannot be redone
        self.redo_stack.clear();

        // Generate Unix timestamp (seconds since epoch) - simpler than ISO 8601 but sufficient for audit
        let timestamp_utc = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_secs().to_string())
            .unwrap_or_else(|_| "0".to_string());

        // Append rewind audit entry
        let new_len_after_audit = self.undo_stack.len() + 1;
        let rewind_action = UndoAction::Rewind {
            target_entry_id,
            target_index,
            target_action_summary,
            discarded_count,
            old_history_len: old_len,
            new_history_len: new_len_after_audit,
            timestamp_utc,
            preview_replay_count,
            preview_build_ms,
        };

        let entry = HistoryEntry {
            id: self.next_entry_id(),
            action: rewind_action,
            timestamp: std::time::Instant::now(),
            provenance: None,
            source: MutationSource::Human,  // Rewind is always user-initiated
        };
        self.undo_stack.push(entry);

        // Update save point if it was beyond truncation
        // (Document is now "dirty" relative to last save)
        if self.save_point > truncate_at {
            // Save point was in discarded region - document is now dirty
            // Set to impossible value to ensure is_dirty() returns true
            self.save_point = usize::MAX;
        }
    }

    /// Build workbook state + view state immediately BEFORE action at index `i`.
    /// This replays actions [0..i) on the base workbook.
    ///
    /// Returns (Workbook, PreviewViewState) or error if:
    /// - i > undo_stack.len()
    /// - timeout exceeded
    /// - too many actions to replay
    /// - unsupported action in replay prefix
    pub fn build_workbook_before(
        &self,
        i: usize,
        base: &Workbook,
        max_replay: usize,
        timeout_ms: u64,
    ) -> Result<PreviewBuildResult, PreviewBuildError> {
        use std::time::Instant;
        use crate::app::{PreviewViewState, PreviewSheetView};

        // Bounds check
        if i > self.undo_stack.len() {
            return Err(PreviewBuildError::InvalidIndex);
        }

        // Safety check: don't replay too many actions
        if i > max_replay {
            return Err(PreviewBuildError::TooManyActions(i));
        }

        // REPLAY GATE: Scan [0..i) for unsupported actions BEFORE starting replay.
        // This ensures deterministic failure - same history always fails the same way.
        for entry in self.undo_stack.iter().take(i) {
            if let Some(kind) = entry.action.first_unsupported_kind() {
                return Err(PreviewBuildError::UnsupportedAction(kind));
            }
        }

        let start = Instant::now();
        let timeout = std::time::Duration::from_millis(timeout_ms);

        // Start with a clone of base workbook
        let mut workbook = base.clone();

        // Initialize preview view state (one entry per sheet, identity order)
        let sheet_count = workbook.sheet_count();
        let mut view_state = PreviewViewState {
            per_sheet: vec![PreviewSheetView::default(); sheet_count],
        };

        // Apply actions [0..i)
        for (idx, entry) in self.undo_stack.iter().take(i).enumerate() {
            // Check timeout periodically
            if idx % 100 == 0 && start.elapsed() > timeout {
                return Err(PreviewBuildError::Timeout);
            }
            // Apply action with invariant checking - abort on violation
            Self::apply_action_forward(&mut workbook, &mut view_state, &entry.action)?;
        }

        let build_ms = start.elapsed().as_millis() as u64;

        Ok(PreviewBuildResult {
            workbook,
            view_state,
            replay_count: i,
            build_ms,
        })
    }

    /// Apply an action forward (redo direction) to workbook and view state.
    /// This uses the "new" values from each action.
    ///
    /// Returns Err(InvariantViolation) if:
    /// - sheet_index is out of bounds (sheet deleted or never existed)
    /// - row_order length mismatches sheet row count (structural corruption)
    ///
    /// INVARIANT: Preview must abort on violation - no partial previews allowed.
    fn apply_action_forward(
        workbook: &mut Workbook,
        view_state: &mut crate::app::PreviewViewState,
        action: &UndoAction,
    ) -> Result<(), PreviewBuildError> {
        match action {
            UndoAction::Values { sheet_index, changes } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("Values action references invalid sheet {}", sheet_index)
                    ))?;
                for change in changes {
                    sheet.set_value(change.row, change.col, &change.new_value);
                }
            }
            UndoAction::Format { sheet_index, patches, .. } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("Format action references invalid sheet {}", sheet_index)
                    ))?;
                for patch in patches {
                    sheet.set_format(patch.row, patch.col, patch.after.clone());
                }
            }
            UndoAction::NamedRangeCreated { named_range } => {
                // Forward replay: create the named range
                let _ = workbook.named_ranges_mut().set(named_range.clone());
            }
            UndoAction::NamedRangeDeleted { named_range } => {
                // Deleting means we had it before, so "forward" means delete it
                workbook.delete_named_range(&named_range.name);
            }
            UndoAction::NamedRangeRenamed { old_name, new_name } => {
                let _ = workbook.rename_named_range(old_name, new_name);
            }
            UndoAction::NamedRangeDescriptionChanged { name, new_description, .. } => {
                // Forward replay: apply description change
                let _ = workbook.named_ranges_mut().set_description(name, new_description.clone());
            }
            UndoAction::Group { actions, .. } => {
                for sub_action in actions {
                    Self::apply_action_forward(workbook, view_state, sub_action)?;
                }
            }
            UndoAction::RowsInserted { sheet_index, at_row, count } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("RowsInserted action references invalid sheet {}", sheet_index)
                    ))?;
                sheet.insert_rows(*at_row, *count);
                // STRUCTURAL CHANGE: Invalidate sort for this sheet (Option B)
                // Row structure changed, previous sort order is no longer valid
                if let Some(sheet_view) = view_state.per_sheet.get_mut(*sheet_index) {
                    sheet_view.row_order = None;
                    sheet_view.sort = None;
                }
            }
            UndoAction::RowsDeleted { sheet_index, at_row, count, .. } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("RowsDeleted action references invalid sheet {}", sheet_index)
                    ))?;
                sheet.delete_rows(*at_row, *count);
                // STRUCTURAL CHANGE: Invalidate sort for this sheet (Option B)
                if let Some(sheet_view) = view_state.per_sheet.get_mut(*sheet_index) {
                    sheet_view.row_order = None;
                    sheet_view.sort = None;
                }
            }
            UndoAction::ColsInserted { sheet_index, at_col, count } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("ColsInserted action references invalid sheet {}", sheet_index)
                    ))?;
                sheet.insert_cols(*at_col, *count);
                // Column changes don't invalidate row order, but may affect sort column
                // For safety, invalidate sort state (column index may have shifted)
                if let Some(sheet_view) = view_state.per_sheet.get_mut(*sheet_index) {
                    sheet_view.sort = None;
                }
            }
            UndoAction::ColsDeleted { sheet_index, at_col, count, .. } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("ColsDeleted action references invalid sheet {}", sheet_index)
                    ))?;
                sheet.delete_cols(*at_col, *count);
                // Column changes don't invalidate row order, but may affect sort column
                if let Some(sheet_view) = view_state.per_sheet.get_mut(*sheet_index) {
                    sheet_view.sort = None;
                }
            }
            UndoAction::ColumnWidthSet { .. } | UndoAction::RowHeightSet { .. } => {
                // Column/row sizing is stored at the app level (Spreadsheet), not in Workbook.
                // For preview purposes, we skip these - the preview shows correct data values
                // even if column widths differ from the historical state.
                // This is acceptable because sizing is visual-only, not computational.
            }
            UndoAction::SortApplied { sheet_index, new_row_order, new_sort_state, .. } => {
                // Validate sheet exists
                if *sheet_index >= view_state.per_sheet.len() {
                    return Err(PreviewBuildError::InvariantViolation(
                        format!("SortApplied action references invalid sheet {}", sheet_index)
                    ));
                }
                // Update preview view state with sort info
                let sheet_view = &mut view_state.per_sheet[*sheet_index];
                sheet_view.row_order = Some(new_row_order.clone());
                sheet_view.sort = Some(*new_sort_state);
            }
            UndoAction::SortCleared { sheet_index, .. } => {
                // Validate sheet exists
                if *sheet_index >= view_state.per_sheet.len() {
                    return Err(PreviewBuildError::InvariantViolation(
                        format!("SortCleared action references invalid sheet {}", sheet_index)
                    ));
                }
                // Clear sort in preview view state
                let sheet_view = &mut view_state.per_sheet[*sheet_index];
                sheet_view.row_order = None;
                sheet_view.sort = None;
            }
            UndoAction::ValidationSet { sheet_index, range, new_rule, .. } => {
                // Replace-in-range semantics: clear overlaps THEN set
                // This matches the live app behavior in dialogs.rs
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("ValidationSet action references invalid sheet {}", sheet_index)
                    ))?;
                sheet.validations.clear_range(range);
                sheet.validations.set(range.clone(), new_rule.clone());
            }
            UndoAction::ValidationCleared { sheet_index, range, .. } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("ValidationCleared action references invalid sheet {}", sheet_index)
                    ))?;
                sheet.validations.clear_range(range);
            }
            UndoAction::ValidationExcluded { sheet_index, range, .. } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("ValidationExcluded action references invalid sheet {}", sheet_index)
                    ))?;
                sheet.validations.exclude(range.clone());
            }
            UndoAction::ValidationExclusionCleared { sheet_index, range, .. } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("ValidationExclusionCleared action references invalid sheet {}", sheet_index)
                    ))?;
                sheet.validations.clear_exclusions_in_range(range);
            }
            UndoAction::Rewind { .. } => {
                // Rewind is audit-only - should never appear in replay paths
                // because rewind truncates history (nothing follows it to replay)
                return Err(PreviewBuildError::InvariantViolation(
                    "Rewind action found in replay path - this should be impossible".to_string()
                ));
            }
        }
        Ok(())
    }
}

/// Classification of undo action types for replay support checking
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UndoActionKind {
    Values,
    Format,
    NamedRangeCreated,
    NamedRangeDeleted,
    NamedRangeRenamed,
    NamedRangeDescriptionChanged,
    Group,
    RowsInserted,
    RowsDeleted,
    ColsInserted,
    ColsDeleted,
    ColumnWidthSet,
    RowHeightSet,
    SortApplied,
    SortCleared,
    ValidationSet,
    ValidationCleared,
    ValidationExcluded,
    ValidationExclusionCleared,
    /// Rewind is an audit-only action - it should never appear in replay paths
    /// because rewind truncates history (no actions after it to replay)
    Rewind,
}

impl UndoActionKind {
    /// Returns true if this action type is fully supported for forward replay
    pub fn is_replay_supported(&self) -> bool {
        match self {
            // Fully supported
            UndoActionKind::Values => true,
            UndoActionKind::Format => true,
            UndoActionKind::NamedRangeCreated => true,
            UndoActionKind::NamedRangeDeleted => true,
            UndoActionKind::NamedRangeRenamed => true,
            UndoActionKind::NamedRangeDescriptionChanged => true,
            UndoActionKind::Group => true,
            UndoActionKind::RowsInserted => true,
            UndoActionKind::RowsDeleted => true,
            UndoActionKind::ColsInserted => true,
            UndoActionKind::ColsDeleted => true,
            UndoActionKind::ColumnWidthSet => true,
            UndoActionKind::RowHeightSet => true,
            UndoActionKind::ValidationSet => true,
            UndoActionKind::ValidationCleared => true,
            UndoActionKind::ValidationExcluded => true,
            UndoActionKind::ValidationExclusionCleared => true,

            // Supported via PreviewViewState (Phase 8B)
            UndoActionKind::SortApplied => true,
            UndoActionKind::SortCleared => true,

            // Rewind is audit-only - should never appear in replay paths
            // (Rewind truncates history, so nothing follows it to replay)
            UndoActionKind::Rewind => false,
        }
    }

    /// Human-readable name for error messages
    pub fn display_name(&self) -> &'static str {
        match self {
            UndoActionKind::Values => "Edit",
            UndoActionKind::Format => "Format",
            UndoActionKind::NamedRangeCreated => "Create named range",
            UndoActionKind::NamedRangeDeleted => "Delete named range",
            UndoActionKind::NamedRangeRenamed => "Rename named range",
            UndoActionKind::NamedRangeDescriptionChanged => "Change description",
            UndoActionKind::Group => "Group",
            UndoActionKind::RowsInserted => "Insert rows",
            UndoActionKind::RowsDeleted => "Delete rows",
            UndoActionKind::ColsInserted => "Insert columns",
            UndoActionKind::ColsDeleted => "Delete columns",
            UndoActionKind::ColumnWidthSet => "Set column width",
            UndoActionKind::RowHeightSet => "Set row height",
            UndoActionKind::SortApplied => "Sort",
            UndoActionKind::SortCleared => "Clear sort",
            UndoActionKind::ValidationSet => "Set validation",
            UndoActionKind::ValidationCleared => "Clear validation",
            UndoActionKind::ValidationExcluded => "Exclude validation",
            UndoActionKind::ValidationExclusionCleared => "Clear exclusion",
            UndoActionKind::Rewind => "Rewind",
        }
    }

    /// Unique byte tag for each action kind (used in fingerprint hashing).
    /// IMPORTANT: These values must be stable across versions.
    /// Never reuse or change assigned values.
    pub fn tag(&self) -> u8 {
        match self {
            UndoActionKind::Values => 0x01,
            UndoActionKind::Format => 0x02,
            UndoActionKind::NamedRangeCreated => 0x03,
            UndoActionKind::NamedRangeDeleted => 0x04,
            UndoActionKind::NamedRangeRenamed => 0x05,
            UndoActionKind::NamedRangeDescriptionChanged => 0x06,
            UndoActionKind::Group => 0x07,
            UndoActionKind::RowsInserted => 0x08,
            UndoActionKind::RowsDeleted => 0x09,
            UndoActionKind::ColsInserted => 0x0A,
            UndoActionKind::ColsDeleted => 0x0B,
            UndoActionKind::ColumnWidthSet => 0x12,
            UndoActionKind::RowHeightSet => 0x13,
            UndoActionKind::SortApplied => 0x0C,
            UndoActionKind::SortCleared => 0x11,
            UndoActionKind::ValidationSet => 0x0D,
            UndoActionKind::ValidationCleared => 0x0E,
            UndoActionKind::ValidationExcluded => 0x0F,
            UndoActionKind::ValidationExclusionCleared => 0x10,
            UndoActionKind::Rewind => 0xFF, // Sentinel value for audit action
        }
    }
}

impl UndoAction {
    /// Get the kind of this action for replay support checking
    pub fn kind(&self) -> UndoActionKind {
        match self {
            UndoAction::Values { .. } => UndoActionKind::Values,
            UndoAction::Format { .. } => UndoActionKind::Format,
            UndoAction::NamedRangeCreated { .. } => UndoActionKind::NamedRangeCreated,
            UndoAction::NamedRangeDeleted { .. } => UndoActionKind::NamedRangeDeleted,
            UndoAction::NamedRangeRenamed { .. } => UndoActionKind::NamedRangeRenamed,
            UndoAction::NamedRangeDescriptionChanged { .. } => UndoActionKind::NamedRangeDescriptionChanged,
            UndoAction::Group { .. } => UndoActionKind::Group,
            UndoAction::RowsInserted { .. } => UndoActionKind::RowsInserted,
            UndoAction::RowsDeleted { .. } => UndoActionKind::RowsDeleted,
            UndoAction::ColsInserted { .. } => UndoActionKind::ColsInserted,
            UndoAction::ColsDeleted { .. } => UndoActionKind::ColsDeleted,
            UndoAction::ColumnWidthSet { .. } => UndoActionKind::ColumnWidthSet,
            UndoAction::RowHeightSet { .. } => UndoActionKind::RowHeightSet,
            UndoAction::SortApplied { .. } => UndoActionKind::SortApplied,
            UndoAction::SortCleared { .. } => UndoActionKind::SortCleared,
            UndoAction::ValidationSet { .. } => UndoActionKind::ValidationSet,
            UndoAction::ValidationCleared { .. } => UndoActionKind::ValidationCleared,
            UndoAction::ValidationExcluded { .. } => UndoActionKind::ValidationExcluded,
            UndoAction::ValidationExclusionCleared { .. } => UndoActionKind::ValidationExclusionCleared,
            UndoAction::Rewind { .. } => UndoActionKind::Rewind,
        }
    }

    /// Check if this action (and any nested actions in Group) are replay-supported
    pub fn is_replay_supported(&self) -> bool {
        match self {
            UndoAction::Group { actions, .. } => {
                actions.iter().all(|a| a.is_replay_supported())
            }
            _ => self.kind().is_replay_supported(),
        }
    }

    /// Find the first unsupported action kind in this action (including nested)
    pub fn first_unsupported_kind(&self) -> Option<UndoActionKind> {
        match self {
            UndoAction::Group { actions, .. } => {
                for action in actions {
                    if let Some(kind) = action.first_unsupported_kind() {
                        return Some(kind);
                    }
                }
                None
            }
            _ => {
                if self.kind().is_replay_supported() {
                    None
                } else {
                    Some(self.kind())
                }
            }
        }
    }
}

/// Successful result from preview build
pub struct PreviewBuildResult {
    /// The reconstructed workbook state
    pub workbook: Workbook,
    /// View state (row ordering per sheet)
    pub view_state: crate::app::PreviewViewState,
    /// Number of actions that were replayed
    pub replay_count: usize,
    /// Time spent building the preview (milliseconds)
    pub build_ms: u64,
}

/// Error type for preview build failures
#[derive(Debug)]
pub enum PreviewBuildError {
    InvalidIndex,
    TooManyActions(usize),
    Timeout,
    /// History contains an action type not supported for replay
    UnsupportedAction(UndoActionKind),
    /// Replay detected an invariant violation (data integrity failure)
    /// Preview must abort - no partial previews allowed
    InvariantViolation(String),
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Test 1: Fingerprint is order-sensitive - different action sequences produce different hashes
    #[test]
    fn fingerprint_is_order_sensitive() {
        use visigrid_engine::cell::CellFormat;

        let mut history1 = History::new();
        let mut history2 = History::new();

        // Create changes and format patches
        let changes = vec![CellChange {
            row: 0, col: 0,
            old_value: "a".to_string(),
            new_value: "b".to_string(),
        }];
        let patches = vec![CellFormatPatch {
            row: 0,
            col: 0,
            before: CellFormat::default(),
            after: CellFormat { bold: true, ..Default::default() },
        }];

        // History 1: Values then Format
        history1.record_batch(0, changes.clone());
        history1.record_format(0, patches.clone(), FormatActionKind::Bold, "Bold".into());

        // History 2: Format then Values
        history2.record_format(0, patches, FormatActionKind::Bold, "Bold".into());
        history2.record_batch(0, changes);

        // Fingerprints should be different because action kind order differs:
        // history1: (id=1, Values), (id=2, Format)
        // history2: (id=1, Format), (id=2, Values)
        let fp1 = history1.fingerprint();
        let fp2 = history2.fingerprint();

        assert_ne!(fp1, fp2, "Fingerprints should differ for different action kind orderings");
        assert_eq!(fp1.len, fp2.len, "Lengths should be the same");
    }

    /// Test 2: Same history produces same fingerprint
    #[test]
    fn fingerprint_is_deterministic() {
        let mut history = History::new();

        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "x".to_string(),
            new_value: "y".to_string(),
        }]);

        let fp1 = history.fingerprint();
        let fp2 = history.fingerprint();

        assert_eq!(fp1, fp2, "Same history should produce same fingerprint");
    }

    /// Test 3: Fingerprint changes when history changes
    #[test]
    fn fingerprint_changes_with_history() {
        let mut history = History::new();

        let fp_empty = history.fingerprint();

        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "".to_string(),
            new_value: "test".to_string(),
        }]);

        let fp_after = history.fingerprint();

        assert_ne!(fp_empty, fp_after, "Fingerprint should change after adding entry");
        assert_eq!(fp_after.len, 1, "Length should be 1 after adding entry");
    }

    /// Test 4: Truncate and append rewind creates correct audit entry
    #[test]
    fn truncate_creates_rewind_audit_entry() {
        let mut history = History::new();

        // Add some history entries
        for i in 0..5 {
            history.record_batch(0, vec![CellChange {
                row: i, col: 0,
                old_value: "".to_string(),
                new_value: format!("value{}", i),
            }]);
        }

        assert_eq!(history.undo_count(), 5);

        // Truncate at index 2 (keep entries 0, 1; discard 2, 3, 4)
        let target_id = history.entry_at(2).unwrap().id;
        history.truncate_and_append_rewind(
            2,
            target_id,
            2,
            "Test action".to_string(),
            2, // replay_count
            50, // build_ms
        );

        // Should have 3 entries now: 0, 1, and the rewind audit entry
        assert_eq!(history.undo_count(), 3);

        // Last entry should be a Rewind action
        let last = history.entry_at(2).unwrap();
        match &last.action {
            UndoAction::Rewind {
                target_entry_id,
                discarded_count,
                old_history_len,
                new_history_len,
                preview_replay_count,
                preview_build_ms,
                ..
            } => {
                assert_eq!(*target_entry_id, target_id);
                assert_eq!(*discarded_count, 3, "Should discard 3 entries (2, 3, 4)");
                assert_eq!(*old_history_len, 5);
                assert_eq!(*new_history_len, 3);
                assert_eq!(*preview_replay_count, 2);
                assert_eq!(*preview_build_ms, 50);
            }
            _ => panic!("Expected Rewind action"),
        }
    }

    /// Test 5: Fingerprint mismatch detection
    #[test]
    fn fingerprint_mismatch_detected() {
        let mut history = History::new();

        history.record_batch(0, vec![CellChange {
            row: 0, col: 0,
            old_value: "".to_string(),
            new_value: "initial".to_string(),
        }]);

        // Take fingerprint
        let fp_before = history.fingerprint();

        // Simulate concurrent change
        history.record_batch(0, vec![CellChange {
            row: 1, col: 0,
            old_value: "".to_string(),
            new_value: "concurrent".to_string(),
        }]);

        // Check fingerprint no longer matches
        assert!(!history.fingerprint_matches(&fp_before),
            "Fingerprint should not match after history changed");
    }

    /// Test 6: UndoActionKind tags are unique
    #[test]
    fn action_kind_tags_are_unique() {
        use std::collections::HashSet;

        let kinds = [
            UndoActionKind::Values,
            UndoActionKind::Format,
            UndoActionKind::NamedRangeCreated,
            UndoActionKind::NamedRangeDeleted,
            UndoActionKind::NamedRangeRenamed,
            UndoActionKind::NamedRangeDescriptionChanged,
            UndoActionKind::Group,
            UndoActionKind::RowsInserted,
            UndoActionKind::RowsDeleted,
            UndoActionKind::ColsInserted,
            UndoActionKind::ColsDeleted,
            UndoActionKind::SortApplied,
            UndoActionKind::SortCleared,
            UndoActionKind::ValidationSet,
            UndoActionKind::ValidationCleared,
            UndoActionKind::ValidationExcluded,
            UndoActionKind::ValidationExclusionCleared,
            UndoActionKind::Rewind,
        ];

        let tags: HashSet<u8> = kinds.iter().map(|k| k.tag()).collect();
        assert_eq!(tags.len(), kinds.len(), "All action kind tags must be unique");
    }
}
