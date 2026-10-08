/// Undo/Redo history system for spreadsheet operations

use visigrid_engine::cell::CellFormat;
use visigrid_engine::named_range::NamedRange;
use visigrid_engine::provenance::Provenance;
use visigrid_engine::sheet::{MergedRegion, SheetId};
use visigrid_engine::workbook::Workbook;
use std::collections::HashMap;
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

#[derive(Clone, Debug, serde::Serialize)]
pub struct CellChange {
    pub row: usize,
    pub col: usize,
    pub old_value: String,
    pub new_value: String,
}

#[derive(Clone, Debug, serde::Serialize)]
pub struct CommentPatch {
    /// The editor materialized an absent cell; undo must restore that absence.
    pub remove_cell_on_undo: bool,
    pub row: usize,
    pub col: usize,
    pub before: Option<visigrid_engine::cell::CellComment>,
    pub after: Option<visigrid_engine::cell::CellComment>,
}

/// Apply only the comment metadata; never replace cell values or formats.
pub(crate) fn apply_comment_patches(workbook: &mut Workbook, sheet_index: usize, patches: &[CommentPatch], forward: bool) {
    if let Some(sheet) = workbook.sheet_mut(sheet_index) {
        for p in patches {
            sheet.set_comment(p.row, p.col, if forward { p.after.clone() } else { p.before.clone() });
            if !forward && p.remove_cell_on_undo && p.before.is_none() {
                sheet.remove_empty_metadata_cell(p.row, p.col);
            }
        }
        workbook.bump_revision_for_structure();
    }
}

/// A patch for a single cell's format (before/after snapshot)
#[derive(Clone, Debug, serde::Serialize)]
pub struct CellFormatPatch {
    /// Restore an absent cell when undo removes its only authored metadata.
    pub remove_cell_on_undo: bool,
    pub row: usize,
    pub col: usize,
    pub before: CellFormat,
    pub after: CellFormat,
}

/// Atomic before/after workbook state for actions whose fidelity crosses
/// several sheet-owned subsystems. Revisions remain monotonic across undo and
/// redo even though the stored snapshots retain their original revisions.
#[derive(Clone, Debug)]
pub struct WorkbookSnapshotCommit {
    pub description: String,
    before: Workbook,
    after: Workbook,
}

impl WorkbookSnapshotCommit {
    pub fn new(description: impl Into<String>, before: Workbook, after: Workbook) -> Self {
        Self {
            description: description.into(),
            before,
            after,
        }
    }

    pub fn undo_into(&self, workbook: &mut Workbook) {
        workbook.restore_snapshot_monotonic(&self.before);
    }

    pub fn redo_into(&self, workbook: &mut Workbook) {
        workbook.restore_snapshot_monotonic(&self.after);
    }

    fn replay_into(&self, workbook: &mut Workbook) {
        *workbook = self.after.clone();
    }
}

/// Kind of format action (for coalescing)
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize)]
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
    FontSize,
    FontColor,
    Border,
    PasteFormats,  // Paste Special > Formats
    ClearFormatting,
    CellStyle,
}

/// An undoable action
#[derive(Clone, Debug, serde::Serialize)]
pub enum UndoAction {
    PrintSetupChanged {
        sheet_id: SheetId,
        before: visigrid_engine::print_setup::PrintSetup,
        after: visigrid_engine::print_setup::PrintSetup,
    },
    Comments { sheet_index: usize, patches: Vec<CommentPatch>, description: String },
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
    /// Conditional format rule added (undo removes it)
    CondFormatAdded {
        sheet_index: usize,
        rule: visigrid_engine::cond_format::CondFormatRule,
    },
    /// Conditional format rules removed (undo re-adds them)
    CondFormatsCleared {
        sheet_index: usize,
        rules: Vec<visigrid_engine::cond_format::CondFormatRule>,
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
    /// Atomic Review Mode commit. Workbook and GUI-owned row state are kept
    /// as before/after snapshots for the first implementation.
    #[serde(skip)]
    PlanCommit {
        commit: Box<visigrid_engine::operation_plan::PlanCommit>,
        sheet_id: SheetId,
        before_row_view: visigrid_engine::filter::RowView,
        after_row_view: visigrid_engine::filter::RowView,
        before_row_heights: HashMap<usize, f32>,
        after_row_heights: HashMap<usize, f32>,
    },
    /// Atomic non-plan workbook mutation. Used when a cell/format patch cannot
    /// faithfully represent the action, such as cloning a complete sheet.
    #[serde(skip)]
    WorkbookSnapshot {
        commit: Box<WorkbookSnapshotCommit>,
        before_row_view: visigrid_engine::filter::RowView,
        after_row_view: visigrid_engine::filter::RowView,
    },
    #[serde(skip)]
    TableBatchChanged { sheet_index: usize, commit: Box<visigrid_engine::workbook::GuardedStructureCommit>, description: String },
    #[serde(skip)]
    TableStructureChanged {
        sheet_index: usize,
        history: Box<crate::table_structure::TableStructureHistory>,
        description: String,
    },
    TableCellsChanged {
        sheet_index: usize,
        commit: Box<crate::table_cell_history::TableCellsCommit>,
        description: String,
    },
    TableViewChanged {
        sheet_index: usize,
        commit: Box<visigrid_engine::workbook::TableViewCommit>,
        description: String,
    },
    #[serde(skip)]
    ReviewCopy { history: Box<crate::review_copy::ReviewCopyHistory> },
    #[serde(skip)]
    TableAppend {
        sheet_index: usize,
        history: Box<crate::table_append::TableAppendHistory>,
        description: String,
    },
    #[serde(skip)]
    TableCommit {
        header_layout: Option<Box<crate::table_create::HeaderLayout>>,
        sheet_index: usize,
        commit: Box<visigrid_engine::workbook::TableCommit>,
        description: String,
    },
    /// Pivot table action (create, apply fields, refresh, delete). Scoped to
    /// the pivot object and its output cells; never a workbook snapshot.
    #[serde(skip)]
    PivotCommit {
        commit: Box<visigrid_engine::workbook::PivotCommit>,
        /// When the action created the pivot's output sheet: its index and the
        /// sheet as created (empty, named), for undo to remove and redo to
        /// restore with the same id.
        created_sheet: Option<(usize, Box<visigrid_engine::sheet::Sheet>)>,
        description: String,
    },
    /// Rows inserted (for undo: delete the inserted rows)
    RowsInserted {
        sheet_index: usize,
        table_rows: Option<visigrid_engine::workbook::TableRowHistory>,
        row_layout: Option<Box<crate::table_structure::RowLayoutHistory>>,
        print_setup_before: visigrid_engine::print_setup::PrintSetup,
        at_row: usize,
        count: usize,
        /// Formulas rewritten by the edit, anywhere in the workbook:
        /// (sheet_index, row, col, text_before). A #REF! rewrite is
        /// destructive text — undo must restore it in the same step, or the
        /// cells come back while the formulas stay mangled.
        #[allow(dead_code)]
        formula_rewrites: Vec<(usize, usize, usize, String)>,
    },
    /// Rows deleted (for undo: re-insert rows and restore cell data)
    RowsDeleted {
        sheet_index: usize,
        table_rows: Option<visigrid_engine::workbook::TableRowHistory>,
        row_layout: Option<Box<crate::table_structure::RowLayoutHistory>>,
        print_setup_before: visigrid_engine::print_setup::PrintSetup,
        at_row: usize,
        count: usize,
        /// Deleted cell data: (row, col, value, format)
        deleted_cells: Vec<(usize, usize, String, CellFormat)>,
        deleted_comments: Vec<(usize, usize, visigrid_engine::cell::CellComment)>,
        /// Deleted row heights: (row, height)
        deleted_row_heights: Vec<(usize, f32)>,
        /// See RowsInserted::formula_rewrites.
        #[allow(dead_code)]
        formula_rewrites: Vec<(usize, usize, usize, String)>,
    },
    /// Columns inserted (for undo: delete the inserted columns)
    ColsInserted {
        sheet_index: usize,
        table_columns: Option<visigrid_engine::workbook::TableColumnHistory>,
        print_setup_before: visigrid_engine::print_setup::PrintSetup,
        at_col: usize,
        count: usize,
        /// See RowsInserted::formula_rewrites.
        #[allow(dead_code)]
        formula_rewrites: Vec<(usize, usize, usize, String)>,
    },
    /// Columns deleted (for undo: re-insert columns and restore cell data)
    ColsDeleted {
        sheet_index: usize,
        table_columns: Option<visigrid_engine::workbook::TableColumnHistory>,
        print_setup_before: visigrid_engine::print_setup::PrintSetup,
        at_col: usize,
        count: usize,
        /// Deleted cell data: (row, col, value, format)
        deleted_cells: Vec<(usize, usize, String, CellFormat)>,
        deleted_comments: Vec<(usize, usize, visigrid_engine::cell::CellComment)>,
        /// Deleted column widths: (col, width)
        deleted_col_widths: Vec<(usize, f32)>,
        /// See RowsInserted::formula_rewrites.
        #[allow(dead_code)]
        formula_rewrites: Vec<(usize, usize, usize, String)>,
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
    /// Rows hidden/unhidden (for undo: toggle visibility)
    RowVisibilityChanged {
        sheet_id: SheetId,
        rows: Vec<usize>,
        /// true = rows were hidden, false = rows were unhidden
        hidden: bool,
    },
    /// Columns hidden/unhidden (for undo: toggle visibility)
    ColVisibilityChanged {
        sheet_id: SheetId,
        cols: Vec<usize>,
        /// true = cols were hidden, false = cols were unhidden
        hidden: bool,
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
    /// Exact canonical metadata edit; legacy variants below retain their old semantics.
    ValidationChanged {
        sheet_index: usize,
        commit: Box<crate::validation_ui::plan::Commit>,
        description: String,
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
    /// Freeze panes changed (for undo: restore previous freeze state)
    FreezePanesChanged {
        sheet_id: visigrid_engine::sheet::SheetId,
        old_frozen_rows: usize,
        old_frozen_cols: usize,
        new_frozen_rows: usize,
        new_frozen_cols: usize,
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
    /// Merge regions changed (merge or unmerge operation).
    /// Captures both topology (merge list) and any cell values cleared during merge.
    SetMerges {
        sheet_index: usize,
        /// Merge regions before the operation
        before: Vec<MergedRegion>,
        /// Merge regions after the operation
        after: Vec<MergedRegion>,
        /// Cell values cleared during merge (row, col, old_value).
        /// Empty for unmerge operations.
        cleared_values: Vec<(usize, usize, String)>,
        description: String,
    },
}

impl UndoAction {
    fn estimated_history_bytes(&self) -> usize {
        use visigrid_engine::history_size::{serialized_bytes, workbook_bytes, sheet_bytes};
        let payload = match self {
            Self::TableCommit { commit, header_layout, description, .. } => commit.estimated_history_bytes() + serialized_bytes(header_layout) + description.capacity(),
            Self::TableBatchChanged { commit, description, .. } => commit.estimated_history_bytes() + description.capacity(),
            Self::TableAppend { history, description, .. } => history.estimated_history_bytes() + description.capacity(),
            Self::TableStructureChanged { history, description, .. } => history.estimated_history_bytes() + description.capacity(),
            Self::Group { actions, description } => actions.iter().map(Self::estimated_history_bytes).sum::<usize>() + description.capacity(),
            Self::PlanCommit { commit, before_row_view, after_row_view, before_row_heights, after_row_heights, .. } => workbook_bytes(&commit.source) + workbook_bytes(&commit.applied) + commit.affected_cells.capacity() * std::mem::size_of::<visigrid_engine::cell_id::CellId>() + serialized_bytes(&commit.verification) + before_row_view.retained_bytes() + after_row_view.retained_bytes() + serialized_bytes(&(before_row_heights, after_row_heights)),
            Self::WorkbookSnapshot { commit, before_row_view, after_row_view } => workbook_bytes(&commit.before) + workbook_bytes(&commit.after) + commit.description.capacity() + before_row_view.retained_bytes() + after_row_view.retained_bytes(),
            Self::ReviewCopy { history } => history.estimated_history_bytes(),
            Self::PivotCommit { commit, created_sheet, description } => commit.approx_bytes() + created_sheet.as_ref().map_or(0, |(_, s)| sheet_bytes(s)) + description.capacity(),
            _ => serialized_bytes(self),
        };
        std::mem::size_of_val(self).saturating_add(payload)
    }

    /// Generate a human-readable label for this action.
    pub fn label(&self) -> String {
        match self {
            UndoAction::Comments { description, .. } | UndoAction::ValidationChanged { description, .. } => description.clone(),
            UndoAction::CondFormatAdded { .. } => "Add conditional format".to_string(),
            UndoAction::CondFormatsCleared { rules, .. } => {
                format!("Clear {} conditional format{}", rules.len(), if rules.len() == 1 { "" } else { "s" })
            }
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
            UndoAction::PlanCommit { commit, .. } => {
                format!("Apply reviewed plan {}", commit.plan_id.0)
            }
            UndoAction::PrintSetupChanged { .. } => "Save print setup".into(),
            UndoAction::WorkbookSnapshot { commit, .. } => commit.description.clone(),
            UndoAction::ReviewCopy { history } => format!("Copy reviewed result to {}", history.sheet.name),
            UndoAction::TableBatchChanged { description, .. } | UndoAction::TableStructureChanged { description, .. }
            | UndoAction::TableViewChanged { description, .. }
            | UndoAction::TableCellsChanged { description, .. } => description.clone(),
            UndoAction::TableCommit { description, .. } | UndoAction::TableAppend { description, .. } => description.clone(),
            UndoAction::PivotCommit { description, .. } => description.clone(),
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
            UndoAction::RowVisibilityChanged { rows, hidden, .. } => {
                let action = if *hidden { "Hide" } else { "Unhide" };
                format!("{} {} row(s)", action, rows.len())
            }
            UndoAction::ColVisibilityChanged { cols, hidden, .. } => {
                let action = if *hidden { "Hide" } else { "Unhide" };
                format!("{} {} column(s)", action, cols.len())
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
            UndoAction::FreezePanesChanged { new_frozen_rows, new_frozen_cols, .. } => {
                if *new_frozen_rows == 0 && *new_frozen_cols == 0 {
                    "Unfreeze panes".to_string()
                } else {
                    "Freeze panes".to_string()
                }
            }
            UndoAction::Rewind { discarded_count, .. } => {
                format!("Rewind (discarded {} change{})", discarded_count, if *discarded_count == 1 { "" } else { "s" })
            }
            UndoAction::SetMerges { description, .. } => {
                description.clone()
            }
        }
    }

    /// Generate an action-specific summary for the detail view.
    /// Returns None for simple actions where label is sufficient.
    pub fn summary(&self) -> Option<String> {
        match self {
            UndoAction::ValidationChanged { commit, .. } => Some(crate::validation_ui::plan::range_summary(&commit.ranges)),
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
#[derive(Clone, Debug, Default, serde::Serialize)]
pub enum MutationSource {
    /// Human entered value manually (default)
    #[default]
    Human,
    /// AI-generated (Ask AI feature)
    Ai(AiMutationMeta),
    /// Applied by a connected client over the session protocol (agent, CLI).
    /// Distinct from Ai: that is the in-app Ask AI flow; this is an external
    /// client acting through MCP or `vgrid apply`.
    Agent {
        /// Authenticated client name, e.g. "Claude Code".
        client: String,
    },
}

/// Metadata for AI-generated mutations (minimal, no prompts/context stored)
#[derive(Clone, Debug, serde::Serialize)]
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
    max_bytes: usize,
    entry_bytes: HashMap<u64, usize>,
    last_record_too_large: bool,
    base_invalidated: bool,
    rewind_base: Option<(Workbook, crate::app::PreviewViewState)>,
    /// Cell-walk size of `rewind_base`, cached so a keystroke does not rescan it.
    rewind_base_bytes: usize,
    /// Evicted edits were stored without calculation. Rewind recomputes once on open.
    rewind_base_needs_calc: bool,
    pending_notice: Option<String>,
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
            max_bytes: 1024 * 1024 * 1024,
            entry_bytes: HashMap::new(),
            last_record_too_large: false,
            base_invalidated: false,
            rewind_base: None,
            rewind_base_bytes: 0,
            rewind_base_needs_calc: false,
            pending_notice: None,
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
    pub fn record_change(&mut self, live: &Workbook, sheet_index: usize, row: usize, col: usize, old_value: String, new_value: String) {
        self.record_change_with_source(live, sheet_index, row, col, old_value, new_value, MutationSource::Human);
    }

    /// Record a single cell value change with explicit source
    pub fn record_change_with_source(
        &mut self,
        live: &Workbook,
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
        self.push_entry(live, entry);
    }

    /// Record multiple cell value changes as a single undoable operation
    pub fn record_batch(&mut self, live: &Workbook, sheet_index: usize, changes: Vec<CellChange>) {
        self.record_batch_with_provenance(live, sheet_index, changes, None);
    }

    /// Re-tag the most recent undo entry's source. Used after routing a
    /// session client's edit through a normal GUI mutation path.
    pub fn retag_last_source(&mut self, live: &Workbook, source: MutationSource) {
        if let Some(entry) = self.undo_stack.last_mut() {
            entry.source = source;
            self.entry_bytes.remove(&entry.id);
            self.enforce_byte_budget(live);
        }
    }

    /// Record a value batch attributed to a specific source (session clients).
    pub fn record_batch_from(&mut self, live: &Workbook, sheet_index: usize, changes: Vec<CellChange>, source: MutationSource) {
        self.record_batch_with_provenance(live, sheet_index, changes, None);
        self.retag_last_source(live, source);
    }

    /// Record a format batch attributed to a specific source (session clients).
    pub fn record_format_from(
        &mut self,
        live: &Workbook,
        sheet_index: usize,
        patches: Vec<CellFormatPatch>,
        kind: FormatActionKind,
        description: String,
        source: MutationSource,
    ) {
        self.record_format(live, sheet_index, patches, kind, description);
        self.retag_last_source(live, source);
    }

    /// Source of the entry that `undo()` would revert next.
    pub fn peek_undo_source(&self) -> Option<&MutationSource> {
        self.undo_stack.last().map(|e| &e.source)
    }

    /// Description of the entry that `undo()` would revert next.
    pub fn peek_undo_description(&self) -> Option<String> {
        self.undo_stack.last().map(|e| e.action.label())
    }

    /// Record multiple cell value changes with optional Lua provenance
    pub fn record_batch_with_provenance(&mut self, live: &Workbook, sheet_index: usize, changes: Vec<CellChange>, provenance: Option<Provenance>) {
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
        self.push_entry(live, entry);
    }

    /// Record format changes with coalescing support
    pub fn record_format(&mut self, live: &Workbook, sheet_index: usize, patches: Vec<CellFormatPatch>, kind: FormatActionKind, description: String) {
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
                        self.entry_bytes.remove(&last.id);
                        self.redo_stack.clear();
                        self.enforce_byte_budget(live);
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
        self.push_entry(live, entry);
    }

    /// Record format changes with optional Lua provenance (for Paste Formats)
    pub fn record_format_with_provenance(
        &mut self,
        live: &Workbook,
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
        self.push_entry(live, entry);
    }

    /// Record a named range action (create, delete, rename)
    pub fn record_named_range_action(&mut self, live: &Workbook, action: UndoAction) {
        self.record_action_with_provenance(live, action, None);
    }

    /// Record any action with optional Lua provenance
    pub fn record_action_with_provenance(&mut self, live: &Workbook, action: UndoAction, provenance: Option<Provenance>) {
        let id = self.next_entry_id();
        let entry = HistoryEntry {
            id,
            action,
            timestamp: Instant::now(),
            provenance,
            source: MutationSource::Human,
        };
        self.push_entry(live, entry);
    }

    /// Check if two patch lists affect the same cells
    fn same_cell_positions(a: &[CellFormatPatch], b: &[CellFormatPatch]) -> bool {
        if a.len() != b.len() {
            return false;
        }
        // Patches are in same order if from same selection iteration
        a.iter().zip(b.iter()).all(|(pa, pb)| pa.row == pb.row && pa.col == pb.col)
    }

    pub fn has_rewind_base(&self) -> bool {
        self.rewind_base.is_some()
    }

    pub fn set_rewind_base(&mut self, workbook: &Workbook) {
        // The clone shares cells, pools and the dependency graph until an edit
        // diverges, so it costs about nothing at any workbook size. Copies
        // made for a previous baseline are not this baseline's bytes.
        let cloned = workbook.clone_sharing_cell_cow();
        cloned.reset_shared_cell_cow();
        self.rewind_base = Some((cloned, crate::app::PreviewViewState {
            per_sheet: vec![Default::default(); workbook.sheet_count()],
        }));
        self.rewind_base_needs_calc = false;
        self.base_invalidated = false;
        self.refresh_rewind_charge(workbook);
        self.drop_rewind_if_over_budget();
    }

    pub fn take_notice(&mut self) -> Option<String> { self.pending_notice.take() }

    fn note_rewind_unavailable(&mut self) {
        const SENTENCE: &str = "Rewind is unavailable; ordinary undo still works for recent changes.";
        match self.pending_notice.as_deref() {
            Some(existing) if existing.contains("Rewind is unavailable") => {}
            Some(existing) => self.pending_notice = Some(format!("{existing}. {SENTENCE}")),
            None => self.pending_notice = Some(SENTENCE.into()),
        }
    }

    /// Drop a captured baseline and say so. A baseline that was never captured
    /// stays silent: eviction already marks rewind invalid.
    fn fail_rewind(&mut self) {
        let had_base = self.rewind_base.take().is_some();
        self.rewind_base_bytes = 0;
        self.rewind_base_needs_calc = false;
        self.base_invalidated = true;
        if had_base {
            self.note_rewind_unavailable();
        }
    }

    fn drop_rewind_if_over_budget(&mut self) {
        if self.rewind_base.is_some() && self.rewind_base_bytes > self.max_bytes {
            self.fail_rewind();
        }
    }

    fn charged_bytes(&self) -> usize {
        let entries = self.entry_bytes.values().copied().sum::<usize>();
        if self.rewind_base.is_some() { entries.saturating_add(self.rewind_base_bytes) } else { entries }
    }

    /// Bytes of baseline allocations the live workbook no longer shares.
    fn refresh_rewind_charge(&mut self, live: &Workbook) {
        self.rewind_base_bytes = self.rewind_base.as_ref().map(|(workbook, _)| workbook.unshared_cow_bytes(live)).unwrap_or(0);
    }

    /// Recompute the retained-copy charge against the current workbook and drop
    /// rewind when it no longer fits. Call this after a preview settles the
    /// baseline, which can copy pages the last edit did not.
    pub fn observe_live(&mut self, live: &Workbook) {
        self.refresh_rewind_charge(live);
        self.drop_rewind_if_over_budget();
    }

    fn evicted_baseline_error() -> PreviewBuildError {
        PreviewBuildError::InvariantViolation("Older undo history was evicted to stay within the history limit. Rewind from the original snapshot is unavailable; ordinary undo is still available for retained changes.".into())
    }

    /// Evicted edits are stored without calculation. The first preview calculates
    /// the baseline in place and clears the flag, so scrubbing does not do it again.
    fn settle_rewind_base(&mut self) {
        if !self.rewind_base_needs_calc {
            return;
        }
        if let Some((workbook, _)) = self.rewind_base.as_mut() {
            workbook.rebuild_dep_graph();
            workbook.recompute_full_ordered();
        }
        self.rewind_base_needs_calc = false;
        self.drop_rewind_if_over_budget();
    }

    fn advance_rewind_base(&mut self, live: &Workbook, action: &UndoAction) {
        if self.rewind_base.is_none() {
            self.base_invalidated = true;
            return;
        }
        // Store the evicted edit. Calculating here repeated a workbook recalc
        // on every keystroke once history held max_entries.
        let applied = self.rewind_base.as_mut().is_some_and(|(workbook, view)| {
            Self::apply_action_forward(workbook, view, action, false).is_ok()
        });
        if applied {
            self.rewind_base_needs_calc = true;
            self.refresh_rewind_charge(live);
            return;
        }
        self.fail_rewind();
    }

    fn evict_oldest_undo(&mut self, live: &Workbook) {
        let removed = self.undo_stack.remove(0);
        self.advance_rewind_base(live, &removed.action);
        self.entry_bytes.remove(&removed.id);
        self.save_point = if self.save_point == 0 { usize::MAX } else { self.save_point - 1 };
    }

    pub fn last_record_too_large(&self) -> bool { self.last_record_too_large }

    #[cfg(test)]
    pub(crate) fn set_byte_budget_for_test(&mut self, bytes: usize) { self.max_bytes = bytes; }


    fn enforce_byte_budget(&mut self, live: &Workbook) {
        let ids: std::collections::HashSet<_> = self.undo_stack.iter().chain(&self.redo_stack).map(|e| e.id).collect();
        self.entry_bytes.retain(|id, _| ids.contains(id));
        for entry in self.undo_stack.iter().chain(&self.redo_stack) {
            self.entry_bytes.entry(entry.id).or_insert_with(|| entry.action.estimated_history_bytes()
                .saturating_add(visigrid_engine::history_size::serialized_bytes(&(&entry.provenance, &entry.source))));
        }
        self.last_record_too_large = self.undo_stack.last().is_some_and(|e| self.entry_bytes[&e.id] > self.max_bytes);
        if self.last_record_too_large {
            let conversion = matches!(self.undo_stack.last().map(|e| &e.action), Some(UndoAction::TableCommit { commit, .. }) if commit.is_conversion());
            self.pending_notice = Some(if conversion { "Converted; this change is too large to undo" } else { "This change is too large to undo; earlier undo history was cleared" }.into());
            // Advance across every discarded action, so the next retained
            // change can still rewind to the state after this barrier.
            for entry in std::mem::take(&mut self.undo_stack) { self.advance_rewind_base(live, &entry.action); }
            // There is no safe undo path across an unrecorded mutation.
            self.undo_stack.clear(); self.redo_stack.clear(); self.entry_bytes.clear();
            self.save_point = usize::MAX;
            self.drop_rewind_if_over_budget();
            return;
        }
        // The baseline shares the budget with retained entries.
        while self.charged_bytes() > self.max_bytes && !self.redo_stack.is_empty() {
            let removed = self.redo_stack.remove(0);
            self.entry_bytes.remove(&removed.id);
        }
        while self.undo_stack.len() > self.max_entries {
            self.evict_oldest_undo(live);
        }
        let mut guard = self.undo_stack.len().saturating_add(1);
        while self.charged_bytes() > self.max_bytes && !self.undo_stack.is_empty() && guard > 0 {
            guard -= 1;
            let before = self.charged_bytes();
            self.evict_oldest_undo(live);
            // Replaying into the baseline can move the bytes instead of freeing
            // them. Drop rewind and keep discarding entries without that copy.
            if self.charged_bytes() >= before && self.rewind_base.is_some() {
                self.fail_rewind();
            }
        }
        self.drop_rewind_if_over_budget();
    }

    fn push_entry(&mut self, live: &Workbook, entry: HistoryEntry) {
        if self.save_point > self.undo_stack.len() { self.save_point = usize::MAX; }
        // A live edit may already have copied a shared chunk. Count it before
        // the new entry joins the budget.
        self.refresh_rewind_charge(live);
        self.undo_stack.push(entry);
        self.redo_stack.clear();
        self.enforce_byte_budget(live);
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
            MutationSource::Agent { client } => Some(client.clone()),
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
            UndoAction::ValidationChanged { sheet_index, .. } => (Some(*sheet_index), vec![], None),
            UndoAction::TableCellsChanged { sheet_index, commit, .. } => {
                let cells = commit.changes();
                let range = Self::bounding_box(&cells);
                (Some(*sheet_index), cells, range)
            }
            UndoAction::TableBatchChanged { sheet_index, .. } | UndoAction::TableStructureChanged { sheet_index, .. } => {
                (Some(*sheet_index), vec![], None)
            }
            UndoAction::TableViewChanged { sheet_index, .. } => (Some(*sheet_index), vec![], None),
            UndoAction::ReviewCopy { history } => (Some(history.index), vec![], None),
            UndoAction::TableAppend { sheet_index, history, .. } => {
                let r = history.table.after_table().unwrap().full_range();
                (Some(*sheet_index), vec![], Some((r.start_row, r.start_col, r.end_row, r.end_col)))
            }
            UndoAction::TableCommit { sheet_index, commit, .. } => {
                let range=commit.after_table().or_else(||commit.before_table()).map(|t| {
                    let mut range = t.full_range();
                    if let Some(before) = commit.before_table() {
                        range.end_row = range.end_row.max(before.full_range().end_row);
                        range.end_col = range.end_col.max(before.full_range().end_col);
                    }
                    if commit.is_totals_change() { range.end_row = t.range.end_row + 1; }
                    (range.start_row,range.start_col,range.end_row,range.end_col)
                });
                (Some(*sheet_index),vec![],range)
            }
            UndoAction::Comments { sheet_index, patches, .. } => {
                let cells: Vec<_> = patches.iter().map(|p| (p.row, p.col, String::new(), String::new())).collect();
                (Some(*sheet_index), vec![], Self::bounding_box(&cells))
            }
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
            UndoAction::RowsInserted { sheet_index, at_row, count, .. } => {
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
            UndoAction::ColsInserted { sheet_index, at_col, count, .. } => {
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
            UndoAction::SetMerges { sheet_index, after, before, .. } => {
                // Use the larger set (before or after) to compute affected range
                let regions = if after.len() >= before.len() { after } else { before };
                if let Some(first) = regions.first() {
                    let mut bbox = (first.start.0, first.start.1, first.end.0, first.end.1);
                    for m in regions.iter().skip(1) {
                        bbox.0 = bbox.0.min(m.start.0);
                        bbox.1 = bbox.1.min(m.start.1);
                        bbox.2 = bbox.2.max(m.end.0);
                        bbox.3 = bbox.3.max(m.end.1);
                    }
                    (Some(*sheet_index), vec![], Some(bbox))
                } else {
                    (Some(*sheet_index), vec![], None)
                }
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
        self.entry_bytes.clear();
        self.last_record_too_large = false;
        self.base_invalidated = false;
        // The base stays, so a later eviction can still rewind. Load paths
        // clear and then capture: capture may notice that the baseline does
        // not fit, and a clear after that capture would erase the notice.
        self.pending_notice = None;
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
        live: &Workbook,
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
        self.enforce_byte_budget(live);

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
        &mut self,
        i: usize,
        base: Option<&Workbook>,
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

        if self.base_invalidated {
            return Err(Self::evicted_baseline_error());
        }
        // REPLAY GATE: Scan [0..i) for unsupported actions BEFORE starting replay.
        // This ensures deterministic failure - same history always fails the same way.
        for entry in self.undo_stack.iter().take(i) {
            if let Some(kind) = entry.action.first_unsupported_kind() {
                return Err(PreviewBuildError::UnsupportedAction(kind));
            }
        }

        // Evicted edits were stored without calculation. Settle that prefix
        // once, on the stored baseline, so the next scrub clones the result.
        self.settle_rewind_base();
        if self.base_invalidated {
            return Err(Self::evicted_baseline_error());
        }
        // No snapshot has been captured to replay from.
        let base = self.rewind_base.as_ref().map(|(w, _)| w).or(base).ok_or(PreviewBuildError::NoBaseSnapshot)?;

        let start = Instant::now();
        let timeout = std::time::Duration::from_millis(timeout_ms);

        let mut workbook = base.clone();

        // Initialize preview view state (one entry per sheet, identity order)
        let sheet_count = workbook.sheet_count();
        let mut view_state = self.rewind_base.as_ref().map(|(_, v)| v.clone()).unwrap_or_else(|| PreviewViewState {
            per_sheet: vec![PreviewSheetView::default(); sheet_count],
        });

        // Apply actions [0..i)
        for (idx, entry) in self.undo_stack.iter().take(i).enumerate() {
            // Check timeout periodically
            if idx % 100 == 0 && start.elapsed() > timeout {
                return Err(PreviewBuildError::Timeout);
            }
            // Apply action with invariant checking - abort on violation
            Self::apply_action_forward(&mut workbook, &mut view_state, &entry.action, true)?;
        }

        view_state.per_sheet.resize_with(workbook.sheet_count(), PreviewSheetView::default);
        for (index, entry) in self.undo_stack.iter().enumerate().skip(i) {
            let source = match &entry.action {
                UndoAction::TableStructureChanged { history, .. } => Some((history.commit.sheet, &history.before, history.source_frozen)),
                UndoAction::RowsInserted { row_layout: Some(layout), .. }
                | UndoAction::RowsDeleted { row_layout: Some(layout), .. } => Some((layout.sheet, &layout.before, None)),
                UndoAction::TableCommit { commit, header_layout: Some(layout), .. } => Some((commit.sheet_id(), &layout.before, Some(layout.frozen_before))),
                _ => None,
            };
            if let Some((target_sheet_id, before, source_frozen)) = source {
                if let Some(sheet_index) = workbook.sheet_index_by_id(target_sheet_id) {
                    let view = &mut view_state.per_sheet[sheet_index];
                    if view.structure_layout.is_none() {
                        let mut layout = before.clone();
                        let mut frozen = source_frozen;
                        for earlier in self.undo_stack[i..index].iter().rev() {
                            match &earlier.action {
                                UndoAction::ColumnWidthSet {
                                    sheet_id, col, old, ..
                                } if *sheet_id == target_sheet_id => {
                                    if let Some(v) = old {
                                        layout.widths.insert(*col, *v);
                                    } else {
                                        layout.widths.remove(col);
                                    }
                                }
                                UndoAction::RowHeightSet {
                                    sheet_id, row, old, ..
                                } if *sheet_id == target_sheet_id => {
                                    if let Some(v) = old {
                                        layout.heights.insert(*row, *v);
                                    } else {
                                        layout.heights.remove(row);
                                    }
                                }
                                UndoAction::RowVisibilityChanged { sheet_id, rows, hidden }
                                    if *sheet_id == target_sheet_id => {
                                    for row in rows {
                                        if *hidden { layout.hidden_rows.remove(row); }
                                        else { layout.hidden_rows.insert(*row); }
                                    }
                                }
                                UndoAction::ColVisibilityChanged { sheet_id, cols, hidden }
                                    if *sheet_id == target_sheet_id => {
                                    for col in cols {
                                        if *hidden { layout.hidden_cols.remove(col); }
                                        else { layout.hidden_cols.insert(*col); }
                                    }
                                }
                                UndoAction::FreezePanesChanged { sheet_id, old_frozen_rows, old_frozen_cols, .. }
                                    if *sheet_id == target_sheet_id => {
                                    frozen = Some((*old_frozen_rows, *old_frozen_cols));
                                }
                                UndoAction::RowsInserted { .. }
                                | UndoAction::RowsDeleted { .. }
                                | UndoAction::ColsInserted { .. }
                                | UndoAction::ColsDeleted { .. }
                                | UndoAction::WorkbookSnapshot { .. }
                                | UndoAction::Group { .. } => {
                                    return Err(PreviewBuildError::InvariantViolation("Cannot reconstruct layout across older structural history.".into()));
                                }
                                _ => {}
                            }
                        }
                        view.structure_layout = Some(layout);
                        if let Some(frozen) = frozen {
                            workbook.sheet_mut(sheet_index).unwrap().frozen_panes = frozen;
                        }
                    }
                }
            }
        }

        // Rebuild projections from the reconstructed cells, including views
        // already present in the base snapshot (no history action required).
        view_state.per_sheet.resize_with(workbook.sheet_count(), PreviewSheetView::default);
        for (sheet, view) in workbook.sheets().iter().zip(&mut view_state.per_sheet) {
            let projection = sheet.build_saved_table_view(crate::app::NUM_ROWS.min(sheet.rows))
                .map_err(PreviewBuildError::InvariantViolation)?;
            if let Some(table_view) = &projection {
                view.sort = table_view.filters().sort.as_ref().map(|s| (s.column, s.direction == visigrid_engine::filter::SortDirection::Ascending));
            }
            view.table_rows = projection.map(|v| v.rows().clone());
        }
        if start.elapsed() > timeout { return Err(PreviewBuildError::Timeout); }

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
        settle: bool,
    ) -> Result<(), PreviewBuildError> {
        // Nested groups each hold a guard. Depth stays above zero until the
        // outermost evicted action returns, including inside a cloned replay.
        let _defer_recalc = (!settle).then(visigrid_engine::workbook::DeferRecalcGuard::enter);
        crate::validation_ui::plan::validate_history(workbook, action, true)
            .map_err(PreviewBuildError::InvariantViolation)?;
        if view_state.per_sheet.iter().any(|v| v.structure_layout.is_some())
            && matches!(action, UndoAction::RowsInserted { row_layout: None, .. } | UndoAction::RowsDeleted { row_layout: None, .. }
                | UndoAction::ColsInserted { .. } | UndoAction::ColsDeleted { .. }
                | UndoAction::WorkbookSnapshot { .. })
        {
            return Err(PreviewBuildError::InvariantViolation(
                "Cannot reconstruct layout across older structural history.".into()));
        }
        match action {
            UndoAction::ValidationChanged { commit, .. } => {
                commit.apply(workbook, true).map_err(PreviewBuildError::InvariantViolation)?;
            }
            UndoAction::Comments { sheet_index, patches, .. } => {
                crate::comments::plan::validate_history(workbook, action, true)
                    .map_err(PreviewBuildError::InvariantViolation)?;
                let sheet = workbook.sheet_mut(*sheet_index).ok_or_else(|| PreviewBuildError::InvariantViolation("Missing comment sheet".into()))?;
                for patch in patches { sheet.set_comment(patch.row, patch.col, patch.after.clone()); }
            }
            UndoAction::CondFormatAdded { sheet_index, rule } => {
                crate::cond_format_ui::plan::validate_history(workbook, action, true)
                    .map_err(PreviewBuildError::InvariantViolation)?;
                crate::cond_format_ui::plan::apply(workbook, *sheet_index, std::slice::from_ref(rule), true);
            }
            UndoAction::CondFormatsCleared { sheet_index, rules } => {
                crate::cond_format_ui::plan::validate_history(workbook, action, true)
                    .map_err(PreviewBuildError::InvariantViolation)?;
                crate::cond_format_ui::plan::apply(workbook, *sheet_index, rules, false);
            }
            UndoAction::Values { sheet_index, changes } => {
                if workbook.sheet(*sheet_index).is_none() {
                    return Err(PreviewBuildError::InvariantViolation(format!("Values action references invalid sheet {}", sheet_index)));
                }
                if settle {
                    workbook.begin_batch();
                    for change in changes { workbook.set_cell_value_tracked(*sheet_index, change.row, change.col, &change.new_value); }
                    workbook.end_batch();
                } else {
                    let sheet = workbook.sheet_mut(*sheet_index).unwrap();
                    for change in changes { sheet.set_value_deferred(change.row, change.col, &change.new_value); }
                }
            }
            UndoAction::Format { sheet_index, patches, .. } => {
                crate::formatting::plan::validate_history(workbook, action, true)
                    .map_err(PreviewBuildError::InvariantViolation)?;
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
                    Self::apply_action_forward(workbook, view_state, sub_action, settle)?;
                }
            }
            UndoAction::PrintSetupChanged { sheet_id, after, .. } => {
                workbook.set_print_setup(*sheet_id, after.clone()).map_err(PreviewBuildError::InvariantViolation)?;
            }
            UndoAction::PlanCommit { commit, .. } => {
                *workbook = commit.applied.clone();
            }
            UndoAction::WorkbookSnapshot {
                commit,
                after_row_view,
                ..
            } => {
                commit.replay_into(workbook);
                view_state.per_sheet = vec![
                    crate::app::PreviewSheetView::default();
                    workbook.sheet_count()
                ];
                if after_row_view.is_sorted() {
                    let active_sheet = workbook.active_sheet_index();
                    if let Some(sheet_view) = view_state.per_sheet.get_mut(active_sheet) {
                        sheet_view.row_order = Some(after_row_view.row_order().to_vec());
                    }
                }
            }
            UndoAction::TableViewChanged { commit, .. } => {
                if settle {
                    workbook.rebuild_dep_graph();
                    workbook.recompute_full_ordered();
                }
                workbook.apply_table_view_commit(commit, false).map_err(PreviewBuildError::InvariantViolation)?;
            }
            UndoAction::TableBatchChanged { commit, .. } => {
                let ids: Vec<_> = workbook.sheets().iter().map(|s| s.id).collect();
                commit.replay(workbook, false).map_err(PreviewBuildError::InvariantViolation)?;
                let mut previous: std::collections::HashMap<_, _> = ids.into_iter()
                    .zip(std::mem::take(&mut view_state.per_sheet)).collect();
                view_state.per_sheet = workbook.sheets().iter()
                    .map(|s| previous.remove(&s.id).unwrap_or_default()).collect();
            }
            UndoAction::TableStructureChanged {
                sheet_index,
                history,
                ..
            } => {
                if let Some(frozen) = history.source_frozen {
                    workbook.sheet_by_id_mut(history.commit.sheet)
                        .ok_or_else(|| PreviewBuildError::InvariantViolation("Review source sheet no longer exists.".into()))?
                        .frozen_panes = frozen;
                }
                history
                    .commit
                    .replay(workbook, false)
                    .map_err(PreviewBuildError::InvariantViolation)?;
                if let Some(view) = view_state.per_sheet.get_mut(*sheet_index) {
                    // A visibility-only commit keeps worksheet sort history.
                    // Structural row/column edits still invalidate that mapping.
                    if !history.commit.steps.is_empty() {
                        view.row_order = None;
                        view.sort = None;
                    }
                    view.structure_layout = Some(history.after.clone());
                }
            }
            UndoAction::TableCellsChanged { commit, .. } => {
                if settle { workbook.rebuild_dep_graph(); }
                commit.replay(workbook, false).map_err(PreviewBuildError::InvariantViolation)?;
            }
            UndoAction::ReviewCopy { history } => {
                let candidate = history.replay(workbook, false).map_err(PreviewBuildError::InvariantViolation)?;
                workbook.restore_snapshot_monotonic(&candidate);
                view_state.per_sheet.resize_with(workbook.sheet_count(), crate::app::PreviewSheetView::default);
                view_state.per_sheet[history.index].structure_layout = Some(history.layout.clone());
            }
            UndoAction::TableAppend { history, .. } => {
                let candidate = history.replay(workbook, false)
                    .map_err(PreviewBuildError::InvariantViolation)?;
                workbook.restore_snapshot_monotonic(&candidate);
            }
            UndoAction::TableCommit { sheet_index, commit, header_layout, .. } => {
                if crate::table_create::is_creation(commit) {
                    if let Some(layout) = header_layout {
                        workbook.sheet_by_id_mut(commit.sheet_id())
                            .ok_or_else(|| PreviewBuildError::InvariantViolation("Missing creation sheet".into()))?
                            .frozen_panes = layout.frozen_before;
                    }
                    let candidate = crate::table_create::prepare_creation_replay(workbook, commit, header_layout.as_deref(), false)
                        .map_err(PreviewBuildError::InvariantViolation)?;
                    workbook.restore_snapshot_monotonic(&candidate);
                    if let (Some(layout), Some(view)) = (header_layout, view_state.per_sheet.get_mut(*sheet_index)) {
                        view.structure_layout = Some(layout.after.clone());
                    }
                } else if commit.is_calculated_change() {
                    let candidate = crate::table_calculated::prepare_replay(workbook, commit, false)
                        .map_err(PreviewBuildError::InvariantViolation)?;
                    workbook.restore_snapshot_monotonic(&candidate);
                } else if commit.is_totals_change() {
                    let candidate = crate::table_totals::prepare_replay(workbook, commit, false)
                        .map_err(PreviewBuildError::InvariantViolation)?;
                    workbook.restore_snapshot_monotonic(&candidate);
                } else if crate::table_resize::is_resize(commit) {
                    let candidate = crate::table_resize::prepare_resize_replay(workbook, commit, false)
                        .map_err(PreviewBuildError::InvariantViolation)?;
                    workbook.restore_snapshot_monotonic(&candidate);
                } else if commit.is_name_change() {
                    let candidate = crate::table_header_paste::prepare_header_replay(workbook, commit, false)
                        .map_err(PreviewBuildError::InvariantViolation)?;
                    workbook.restore_snapshot_monotonic(&candidate);
                } else {
                    workbook.apply_table_commit(commit, false).map_err(PreviewBuildError::InvariantViolation)?;
                }
                if commit.inserted_header_row().is_some() {
                    if let Some(view) = view_state.per_sheet.get_mut(*sheet_index) {
                        view.row_order = None;
                        view.sort = None;
                    }
                }
            }
            UndoAction::PivotCommit { commit, created_sheet, .. } => {
                if let Some((index, sheet)) = created_sheet {
                    if workbook.sheet_index_by_id(sheet.id).is_none() {
                        workbook.restore_sheet(*index, (**sheet).clone());
                    }
                }
                workbook
                    .apply_pivot_state(&commit.after)
                    .map_err(|e| PreviewBuildError::InvariantViolation(e.to_string()))?;
            }
            UndoAction::RowsInserted { sheet_index, at_row, count, table_rows, row_layout, .. } => {
                if let Some(history) = table_rows {
                    workbook.apply_table_row_history(history, false).map_err(PreviewBuildError::InvariantViolation)?;
                } else if row_layout.is_some() {
                    workbook.structural_edit(*sheet_index, visigrid_engine::structural::Axis::Row, *at_row, *count, false)
                        .map_err(PreviewBuildError::InvariantViolation)?;
                } else {
                    let sheet = workbook.sheet_mut(*sheet_index).ok_or_else(|| PreviewBuildError::InvariantViolation("Sheet no longer exists".into()))?;
                    sheet.insert_rows(*at_row, *count);
                }
                if let Some(history) = row_layout {
                    let layout = history.apply(workbook, *sheet_index, false)
                        .map_err(PreviewBuildError::InvariantViolation)?;
                    view_state.per_sheet.resize_with(workbook.sheet_count(), crate::app::PreviewSheetView::default);
                    view_state.per_sheet[*sheet_index].structure_layout = Some(layout.clone());
                }
                // STRUCTURAL CHANGE: Invalidate sort for this sheet (Option B)
                // Row structure changed, previous sort order is no longer valid
                if let Some(sheet_view) = view_state.per_sheet.get_mut(*sheet_index) {
                    sheet_view.row_order = None;
                    sheet_view.sort = None;
                }
            }
            UndoAction::RowsDeleted { sheet_index, at_row, count, table_rows, row_layout, .. } => {
                if let Some(history) = table_rows {
                    workbook.apply_table_row_history(history, false).map_err(PreviewBuildError::InvariantViolation)?;
                } else if row_layout.is_some() {
                    workbook.structural_edit(*sheet_index, visigrid_engine::structural::Axis::Row, *at_row, *count, true)
                        .map_err(PreviewBuildError::InvariantViolation)?;
                } else {
                    let sheet = workbook.sheet_mut(*sheet_index).ok_or_else(|| PreviewBuildError::InvariantViolation("Sheet no longer exists".into()))?;
                    sheet.delete_rows(*at_row, *count);
                }
                if let Some(history) = row_layout {
                    let layout = history.apply(workbook, *sheet_index, false)
                        .map_err(PreviewBuildError::InvariantViolation)?;
                    view_state.per_sheet.resize_with(workbook.sheet_count(), crate::app::PreviewSheetView::default);
                    view_state.per_sheet[*sheet_index].structure_layout = Some(layout.clone());
                }
                // STRUCTURAL CHANGE: Invalidate sort for this sheet (Option B)
                if let Some(sheet_view) = view_state.per_sheet.get_mut(*sheet_index) {
                    sheet_view.row_order = None;
                    sheet_view.sort = None;
                }
            }
            UndoAction::ColsInserted { sheet_index, at_col, count, table_columns, .. } => {
                if let Some(history) = table_columns {
                    workbook.apply_table_column_history(history, false).map_err(PreviewBuildError::InvariantViolation)?;
                } else {
                    let sheet = workbook.sheet_mut(*sheet_index).ok_or_else(|| PreviewBuildError::InvariantViolation("Sheet no longer exists".into()))?;
                    sheet.insert_cols(*at_col, *count);
                }
                // Column changes don't invalidate row order, but may affect sort column
                // For safety, invalidate sort state (column index may have shifted)
                if let Some(sheet_view) = view_state.per_sheet.get_mut(*sheet_index) {
                    sheet_view.sort = None;
                }
            }
            UndoAction::ColsDeleted { sheet_index, at_col, count, table_columns, .. } => {
                if let Some(history) = table_columns {
                    workbook.apply_table_column_history(history, false).map_err(PreviewBuildError::InvariantViolation)?;
                } else {
                    let sheet = workbook.sheet_mut(*sheet_index).ok_or_else(|| PreviewBuildError::InvariantViolation("Sheet no longer exists".into()))?;
                    sheet.delete_cols(*at_col, *count);
                }
                // Column changes don't invalidate row order, but may affect sort column
                if let Some(sheet_view) = view_state.per_sheet.get_mut(*sheet_index) {
                    sheet_view.sort = None;
                }
            }
            UndoAction::ColumnWidthSet {
                sheet_id, col, new, ..
            } => {
                if let Some(layout) = workbook
                    .sheet_index_by_id(*sheet_id)
                    .and_then(|i| view_state.per_sheet.get_mut(i))
                    .and_then(|v| v.structure_layout.as_mut())
                {
                    if let Some(v) = new {
                        layout.widths.insert(*col, *v);
                    } else {
                        layout.widths.remove(col);
                    }
                }
            }
            UndoAction::RowHeightSet {
                sheet_id, row, new, ..
            } => {
                if let Some(layout) = workbook
                    .sheet_index_by_id(*sheet_id)
                    .and_then(|i| view_state.per_sheet.get_mut(i))
                    .and_then(|v| v.structure_layout.as_mut())
                {
                    if let Some(v) = new {
                        layout.heights.insert(*row, *v);
                    } else {
                        layout.heights.remove(row);
                    }
                }
            }
            UndoAction::RowVisibilityChanged { sheet_id, rows, hidden } => {
                let sheet = workbook.sheet_by_id_mut(*sheet_id).ok_or_else(|| PreviewBuildError::InvariantViolation("Missing visibility sheet".into()))?;
                let mut canonical = sheet.manual_hidden_rows();
                for row in rows {
                    if *hidden { canonical.insert(*row); } else { canonical.remove(row); }
                }
                sheet.set_manual_hidden_rows(canonical).map_err(PreviewBuildError::InvariantViolation)?;
                if let Some(layout) = workbook.sheet_index_by_id(*sheet_id)
                    .and_then(|i| view_state.per_sheet.get_mut(i))
                    .and_then(|v| v.structure_layout.as_mut()) {
                    for row in rows {
                        if *hidden { layout.hidden_rows.insert(*row); }
                        else { layout.hidden_rows.remove(row); }
                    }
                }
            }
            UndoAction::ColVisibilityChanged { sheet_id, cols, hidden } => {
                if let Some(layout) = workbook.sheet_index_by_id(*sheet_id)
                    .and_then(|i| view_state.per_sheet.get_mut(i))
                    .and_then(|v| v.structure_layout.as_mut()) {
                    for col in cols {
                        if *hidden { layout.hidden_cols.insert(*col); }
                        else { layout.hidden_cols.remove(col); }
                    }
                }
            }
            UndoAction::FreezePanesChanged { sheet_id, new_frozen_rows, new_frozen_cols, .. } => {
                crate::table_command_scope::restore_freeze_panes(workbook, *sheet_id, (*new_frozen_rows, *new_frozen_cols))
                    .map_err(PreviewBuildError::InvariantViolation)?;
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
            // The retained prefix already describes this state. Audit entries
            // have no workbook effect, including during a later rewind.
            UndoAction::Rewind { .. } => {}
            UndoAction::SetMerges { sheet_index, after, cleared_values, .. } => {
                let sheet = workbook.sheet_mut(*sheet_index)
                    .ok_or_else(|| PreviewBuildError::InvariantViolation(
                        format!("SetMerges action references invalid sheet {}", sheet_index)
                    ))?;
                // Clear values that the merge hides
                for (row, col, _) in cleared_values {
                    sheet.set_value(*row, *col, "");
                }
                // Apply merge topology
                sheet.set_merges(after.clone());
            }
        }
        Ok(())
    }
}

/// Classification of undo action types for replay support checking
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UndoActionKind {
    ValidationChanged,
    PrintSetupChanged,
    Comments,
    CondFormatAdded,
    CondFormatsCleared,
    Values,
    Format,
    NamedRangeCreated,
    NamedRangeDeleted,
    NamedRangeRenamed,
    NamedRangeDescriptionChanged,
    Group,
    PlanCommit,
    WorkbookSnapshot,
    PivotCommit,
    TableCommit,
    TableAppend,
    ReviewCopy,
    TableViewChanged,
    TableCellsChanged,
    TableStructureChanged,
    TableBatchChanged,
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
    RowVisibilityChanged,
    ColVisibilityChanged,
    FreezePanesChanged,
    /// Rewind is an audit-only action - it should never appear in replay paths
    /// because rewind truncates history (no actions after it to replay)
    Rewind,
    SetMerges,
}

impl UndoActionKind {
    /// Returns true if this action type is fully supported for forward replay
    pub fn is_replay_supported(&self) -> bool {
        match self {
            // Fully supported
            UndoActionKind::ValidationChanged => true,
            UndoActionKind::Values => true,
            UndoActionKind::CondFormatAdded => true,
            UndoActionKind::CondFormatsCleared => true,
            UndoActionKind::Format => true,
            UndoActionKind::NamedRangeCreated => true,
            UndoActionKind::NamedRangeDeleted => true,
            UndoActionKind::NamedRangeRenamed => true,
            UndoActionKind::NamedRangeDescriptionChanged => true,
            UndoActionKind::Group => true,
            UndoActionKind::PlanCommit => true,
            UndoActionKind::PrintSetupChanged => true,
            UndoActionKind::Comments => true,
            UndoActionKind::WorkbookSnapshot => true,
            UndoActionKind::TableCommit => true,
            UndoActionKind::TableAppend => true,
            UndoActionKind::ReviewCopy => true,
            UndoActionKind::TableViewChanged => true,
            UndoActionKind::TableCellsChanged => true,
            UndoActionKind::TableStructureChanged => true,
            UndoActionKind::TableBatchChanged => true,
            UndoActionKind::PivotCommit => true,
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

            // View-only changes (not serialized to file) - skip for replay
            UndoActionKind::RowVisibilityChanged => false,
            UndoActionKind::ColVisibilityChanged => false,
            // Freeze panes now carry a stable sheet identity and replay into Workbook.
            UndoActionKind::FreezePanesChanged => true,

            // Merge topology changes are replay-supported
            UndoActionKind::SetMerges => true,

            // Rewind is audit-only - should never appear in replay paths
            // (Rewind truncates history, so nothing follows it to replay)
            UndoActionKind::Rewind => true,
        }
    }

    /// Human-readable name for error messages
    pub fn display_name(&self) -> &'static str {
        match self {
            UndoActionKind::ValidationChanged => "Validation",
            UndoActionKind::Values => "Edit",
            UndoActionKind::CondFormatAdded => "Add conditional format",
            UndoActionKind::CondFormatsCleared => "Clear conditional formats",
            UndoActionKind::Format => "Format",
            UndoActionKind::NamedRangeCreated => "Create named range",
            UndoActionKind::NamedRangeDeleted => "Delete named range",
            UndoActionKind::NamedRangeRenamed => "Rename named range",
            UndoActionKind::NamedRangeDescriptionChanged => "Change description",
            UndoActionKind::Group => "Group",
            UndoActionKind::PlanCommit => "Reviewed plan",
            UndoActionKind::PrintSetupChanged => "Print setup",
            UndoActionKind::Comments => "Comment",
            UndoActionKind::WorkbookSnapshot => "Workbook snapshot",
            UndoActionKind::TableCommit => "Table",
            UndoActionKind::TableAppend => "Append Table row",
            UndoActionKind::ReviewCopy => "Copy reviewed sheet",
            UndoActionKind::TableViewChanged => "Table view",
            UndoActionKind::TableCellsChanged => "Table cells",
            UndoActionKind::TableStructureChanged => "Table structure",
            UndoActionKind::TableBatchChanged => "Table batch",
            UndoActionKind::PivotCommit => "Pivot table",
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
            UndoActionKind::RowVisibilityChanged => "Hide/unhide rows",
            UndoActionKind::ColVisibilityChanged => "Hide/unhide columns",
            UndoActionKind::FreezePanesChanged => "Freeze panes",
            UndoActionKind::Rewind => "Rewind",
            UndoActionKind::SetMerges => "Merge cells",
        }
    }

    /// Unique byte tag for each action kind (used in fingerprint hashing).
    /// IMPORTANT: These values must be stable across versions.
    /// Never reuse or change assigned values.
    pub fn tag(&self) -> u8 {
        match self {
            UndoActionKind::CondFormatAdded => 0x18,
            UndoActionKind::CondFormatsCleared => 0x19,
            UndoActionKind::ValidationChanged => 0x29,
            UndoActionKind::Values => 0x01,
            UndoActionKind::Format => 0x02,
            UndoActionKind::NamedRangeCreated => 0x03,
            UndoActionKind::NamedRangeDeleted => 0x04,
            UndoActionKind::NamedRangeRenamed => 0x05,
            UndoActionKind::NamedRangeDescriptionChanged => 0x06,
            UndoActionKind::Group => 0x07,
            UndoActionKind::PlanCommit => 0x1A,
            UndoActionKind::PrintSetupChanged => 0x1D,
            UndoActionKind::Comments => 0x1E,
            UndoActionKind::WorkbookSnapshot => 0x1B,
            UndoActionKind::TableCommit => 0x1F,
            UndoActionKind::TableAppend => 0x27,
            UndoActionKind::ReviewCopy => 0x28,
            UndoActionKind::TableViewChanged => 0x20,
            UndoActionKind::TableCellsChanged => 0x21,
            UndoActionKind::TableStructureChanged => 0x22,
            UndoActionKind::TableBatchChanged => 0x23,
            UndoActionKind::PivotCommit => 0x1C,
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
            UndoActionKind::RowVisibilityChanged => 0x15,
            UndoActionKind::ColVisibilityChanged => 0x16,
            UndoActionKind::FreezePanesChanged => 0x17,
            UndoActionKind::Rewind => 0xFF, // Sentinel value for audit action
            UndoActionKind::SetMerges => 0x14,
        }
    }
}

impl UndoAction {
    /// Get the kind of this action for replay support checking
    pub fn kind(&self) -> UndoActionKind {
        match self {
            UndoAction::CondFormatAdded { .. } => UndoActionKind::CondFormatAdded,
            UndoAction::CondFormatsCleared { .. } => UndoActionKind::CondFormatsCleared,
            UndoAction::ValidationChanged { .. } => UndoActionKind::ValidationChanged,
            UndoAction::Values { .. } => UndoActionKind::Values,
            UndoAction::Format { .. } => UndoActionKind::Format,
            UndoAction::NamedRangeCreated { .. } => UndoActionKind::NamedRangeCreated,
            UndoAction::NamedRangeDeleted { .. } => UndoActionKind::NamedRangeDeleted,
            UndoAction::NamedRangeRenamed { .. } => UndoActionKind::NamedRangeRenamed,
            UndoAction::NamedRangeDescriptionChanged { .. } => UndoActionKind::NamedRangeDescriptionChanged,
            UndoAction::Group { .. } => UndoActionKind::Group,
            UndoAction::PlanCommit { .. } => UndoActionKind::PlanCommit,
            UndoAction::PrintSetupChanged { .. } => UndoActionKind::PrintSetupChanged,
            UndoAction::Comments { .. } => UndoActionKind::Comments,
            UndoAction::WorkbookSnapshot { .. } => UndoActionKind::WorkbookSnapshot,
            UndoAction::TableCommit { .. } => UndoActionKind::TableCommit,
            UndoAction::TableAppend { .. } => UndoActionKind::TableAppend,
            UndoAction::ReviewCopy { .. } => UndoActionKind::ReviewCopy,
            UndoAction::TableViewChanged { .. } => UndoActionKind::TableViewChanged,
            UndoAction::TableCellsChanged { .. } => UndoActionKind::TableCellsChanged,
            UndoAction::TableStructureChanged { .. } => UndoActionKind::TableStructureChanged,
            UndoAction::TableBatchChanged { .. } => UndoActionKind::TableBatchChanged,
            UndoAction::PivotCommit { .. } => UndoActionKind::PivotCommit,
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
            UndoAction::RowVisibilityChanged { .. } => UndoActionKind::RowVisibilityChanged,
            UndoAction::ColVisibilityChanged { .. } => UndoActionKind::ColVisibilityChanged,
            UndoAction::FreezePanesChanged { .. } => UndoActionKind::FreezePanesChanged,
            UndoAction::Rewind { .. } => UndoActionKind::Rewind,
            UndoAction::SetMerges { .. } => UndoActionKind::SetMerges,
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
    /// No load-time snapshot was captured to replay from
    NoBaseSnapshot,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ordinary_row_rewind_preserves_hidden_rows_and_layout_before_and_after_edits() {
        use crate::table_structure::{RowLayoutHistory, StructureLayout};
        use visigrid_engine::{structural::Axis, workbook::StructureStep};
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_manual_hidden_rows([2].into()).unwrap();
        let base = wb.clone();
        let mut layout = StructureLayout {
            hidden_rows: [2].into(), heights: [(2, 40.0)].into(),
            widths: [(4, 120.0)].into(), hidden_cols: [5].into(),
        };
        let mut layouts = vec![layout.clone()];
        let mut history = History::new();
        for (at, delete) in [(1, false), (3, true), (1, false)] {
            let rows = RowLayoutHistory::capture(layout, wb.active_sheet(), StructureStep {
                axis: Axis::Row, at, count: 1, delete,
            }).unwrap();
            let print_setup_before = wb.active_sheet().print_setup.clone();
            wb.structural_edit(0, Axis::Row, at, 1, delete).unwrap();
            layout = rows.after.clone();
            layouts.push(layout.clone());
            let action = if delete {
                UndoAction::RowsDeleted {
                    sheet_index: 0, table_rows: None, row_layout: Some(Box::new(rows)),
                    print_setup_before, at_row: at, count: 1, formula_rewrites: vec![],
                    deleted_cells: vec![], deleted_comments: vec![], deleted_row_heights: vec![(3, 40.0)],
                }
            } else {
                UndoAction::RowsInserted {
                    sheet_index: 0, table_rows: None, row_layout: Some(Box::new(rows)),
                    print_setup_before, at_row: at, count: 1, formula_rewrites: vec![],
                }
            };
            history.record_named_range_action(&visigrid_engine::workbook::Workbook::new(), action);
        }
        for (index, expected) in layouts.iter().enumerate() {
            let preview = history.build_workbook_before(index, Some(&base), 100, 10_000).unwrap();
            assert_eq!(preview.workbook.active_sheet().manual_hidden_rows(), expected.hidden_rows);
            assert_eq!(preview.view_state.per_sheet[0].structure_layout.as_ref(), Some(expected));
        }
    }

    #[test]
    fn table_history_replays_schema_formulas_and_style_without_body_snapshots() {
        use visigrid_engine::table::TableRange;
        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_id();
        workbook.set_cell_value_tracked(0, 0, 0, "Amount");
        workbook.set_cell_value_tracked(0, 1, 0, "12");
        workbook.set_cell_value_tracked(0, 0, 3, "=SUM(Sales[Amount])");
        let base = workbook.clone();
        let create = workbook
            .create_table(
                sheet,
                TableRange {
                    start_row: 0,
                    start_col: 0,
                    end_row: 1,
                    end_col: 0,
                },
                "Sales",
            )
            .unwrap();
        let id = create.table_id();
        let rename = workbook.rename_table(id, "Orders").unwrap();
        let style = workbook
            .set_table_style(
                id,
                visigrid_engine::table::TableStyle { banded_rows: false, ..Default::default() },
            )
            .unwrap();
        let mut replay = base;
        let mut view = crate::app::PreviewViewState::default();
        for commit in [&create, &rename, &style] {
            let action = UndoAction::TableCommit { header_layout: None,
                sheet_index: 0,
                commit: Box::new(commit.clone()),
                description: "Table change".into(),
            };
            History::apply_action_forward(&mut replay, &mut view, &action, true).unwrap();
            assert!(UndoActionKind::TableCommit.is_replay_supported());
        }
        assert_eq!(
            replay.saved_tables().sheets[0].tables,
            workbook.saved_tables().sheets[0].tables
        );
        assert_eq!(replay.active_sheet().get_raw(1, 0), "12");
        assert_eq!(replay.active_sheet().get_display(0, 3), "12");
        assert!(replay
            .active_sheet()
            .get_raw(0, 3)
            .contains("Orders[Amount]"));
        for commit in [&style, &rename, &create] {
            replay.apply_table_commit(commit, true).unwrap();
        }
        assert_eq!(replay.tables().count(), 0);
        assert_eq!(replay.active_sheet().get_raw(1, 0), "12");
    }

    #[test]
    fn table_growth_and_whole_row_history_rewind_matches_live_workbook() {
        use visigrid_engine::table::TableRange;
        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_id();
        let id = workbook.create_table(sheet, TableRange {
            start_row: 0, start_col: 0, end_row: 0, end_col: 1,
        }, "Sales").unwrap().table_id();
        let mut replay = workbook.clone();
        let mut view = crate::app::PreviewViewState::default();
        let append = workbook.append_table_rows(id, 2, &[
            (1, 0, "3".into()), (1, 1, "=[@Column1]*10".into()),
            (2, 0, "4".into()), (2, 1, "=[@Column1]*10".into()),
        ]).unwrap();
        let action = UndoAction::TableCommit { header_layout: None,
            sheet_index: 0, commit: Box::new(append), description: "Append Table rows".into(),
        };
        History::apply_action_forward(&mut replay, &mut view, &action, true).unwrap();
        for delete in [false, true] {
            let count = if delete { 3 } else { 1 };
            let history = workbook.prepare_table_row_history(0, 1, count, delete).unwrap().unwrap();
            let print_setup_before = workbook.active_sheet().print_setup.clone();
            workbook.apply_table_row_history(&history, false).unwrap();
            let action = if delete {
                UndoAction::RowsDeleted {
                    sheet_index: 0, at_row: 1, count, table_rows: Some(history), row_layout: None,
                    print_setup_before, formula_rewrites: vec![], deleted_cells: vec![],
                    deleted_comments: vec![], deleted_row_heights: vec![],
                }
            } else {
                UndoAction::RowsInserted {
                    sheet_index: 0, at_row: 1, count, table_rows: Some(history), row_layout: None,
                    print_setup_before, formula_rewrites: vec![],
                }
            };
            History::apply_action_forward(&mut replay, &mut view, &action, true).unwrap();
            assert_eq!(replay.table(id).unwrap().1, workbook.table(id).unwrap().1);
            for row in 0..5 {
                assert_eq!(replay.active_sheet().get_raw(row, 0), workbook.active_sheet().get_raw(row, 0));
                assert_eq!(replay.active_sheet().get_display(row, 1), workbook.active_sheet().get_display(row, 1));
            }
        }
        assert_eq!(replay.table(id).unwrap().1.range.data_rows(), 0);
    }

    #[test]
    fn headerless_table_rewind_matches_live_and_preserves_records() {
        use visigrid_engine::table::TableRange;
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 0, 0, "42");
        workbook.set_cell_value_tracked(0, 0, 1, "=A1*2");
        workbook.set_cell_value_tracked(0, 1, 0, "17");
        let mut replay = workbook.clone();
        let mut view = crate::app::PreviewViewState::default();
        let commit = workbook.create_table_without_headers(workbook.active_sheet_id(), TableRange {
            start_row: 0, start_col: 0, end_row: 1, end_col: 1,
        }, "Sales").unwrap();
        let id = commit.table_id();
        let action = UndoAction::TableCommit { header_layout: None, sheet_index: 0, commit: Box::new(commit.clone()), description: "Create Table: Sales".into() };
        History::apply_action_forward(&mut replay, &mut view, &action, true).unwrap();
        assert_eq!(replay.table(id).unwrap().1, workbook.table(id).unwrap().1);
        assert_eq!(replay.active_sheet().get_display(1,1), "84");
        assert_eq!(replay.active_sheet().get_raw(2,0), "17");
        replay.apply_table_commit(&commit, true).unwrap();
        assert!(replay.table(id).is_none());
        assert_eq!(replay.active_sheet().get_display(0,1), "84");
        History::apply_action_forward(&mut replay, &mut view, &action, true).unwrap();
        assert_eq!(replay.table(id).unwrap().1, workbook.table(id).unwrap().1);
    }

    #[test]
    fn table_column_rewind_preserves_schema_rules_overrides_and_dependencies() {
        use visigrid_engine::table::TableRange;
        let mut workbook = Workbook::new();
        let id = workbook.create_table(workbook.active_sheet_id(), TableRange {
            start_row: 0, start_col: 0, end_row: 3, end_col: 2,
        }, "Sales").unwrap().table_id();
        workbook.set_cell_value_tracked(0, 1, 0, "2");
        workbook.set_cell_value_tracked(0, 1, 1, "10");
        workbook.set_calculated_column(id, 2, 1, "=[@Column1]*B2", true).unwrap();
        workbook.set_cell_value_tracked(0, 2, 2, "999");
        workbook.clear_cell_tracked(0, 3, 2);
        let summary = workbook.add_sheet_named("Summary").unwrap();
        workbook.set_cell_value_tracked(summary, 0, 0, "=SUM(Sales[Column3])");
        let mut replay = workbook.clone();
        let mut view = crate::app::PreviewViewState::default();
        // Insert internally, remove an input, then remove the calculated field.
        for (at, delete) in [(1, false), (0, true), (2, true)] {
            let table_columns = workbook.prepare_table_column_history(0, at, 1, delete).unwrap();
            let print_setup_before = workbook.sheet(0).unwrap().print_setup.clone();
            let deleted_cells = workbook.sheet(0).unwrap().occupied_cells_in_cols(at, 1);
            let rewrites = workbook.apply_table_column_history(table_columns.as_ref().unwrap(), false).unwrap();
            let formula_rewrites = rewrites.into_iter().map(|(s,r,c,old,_)| (s,r,c,old)).collect();
            let action = if delete {
                UndoAction::ColsDeleted { sheet_index: 0, at_col: at, count: 1, table_columns,
                    print_setup_before, deleted_cells, deleted_comments: vec![], deleted_col_widths: vec![], formula_rewrites }
            } else {
                UndoAction::ColsInserted { sheet_index: 0, at_col: at, count: 1, table_columns,
                    print_setup_before, formula_rewrites }
            };
            History::apply_action_forward(&mut replay, &mut view, &action, true).unwrap();
            assert_eq!(replay.table(id).unwrap().1, workbook.table(id).unwrap().1);
            for r in 0..4 { for c in 0..5 {
                assert_eq!(replay.sheet(0).unwrap().get_raw(r,c), workbook.sheet(0).unwrap().get_raw(r,c));
                assert_eq!(replay.sheet(0).unwrap().get_display(r,c), workbook.sheet(0).unwrap().get_display(r,c));
            }}
            assert_eq!(replay.sheet(summary).unwrap().get_raw(0,0), workbook.sheet(summary).unwrap().get_raw(0,0));
        }
    }

    #[test]
    fn native_totals_rewind_with_active_filters_keeps_hidden_footer_and_criteria() {
        use visigrid_engine::{table::TableRange, table_view::{TableViewSpec, TableFilter}, filter::{ColumnFilter, NormalizedFilterKey}};
        let mut wb = Workbook::new();
        for (row, value) in ["Amount", "10", "20"].iter().enumerate() { wb.set_cell_value_tracked(0, row, 0, value); }
        let id = wb.create_table(wb.active_sheet_id(), TableRange { start_row:0, start_col:0, end_row:2, end_col:0 }, "Sales").unwrap().table_id();
        let mut spec = TableViewSpec::new(id);
        spec.filters.push(TableFilter { column: wb.table(id).unwrap().1.columns[0].id,
            criteria: ColumnFilter { selected: Some([NormalizedFilterKey::Number(10.0.into())].into()), text_filter: None } });
        wb.set_table_view_spec(wb.active_sheet_id(), Some(spec.clone())).unwrap();
        let (mut wb, _) = wb.prepare_table_row_visibility(wb.active_sheet_id(), [2, 3].into()).unwrap();
        let mut replay = wb.clone();
        let show = wb.set_table_totals_visible(id, true, [2, 3].into()).unwrap();
        let edit = wb.set_table_total(id, 0, crate::table_totals::total_setting("countNums", "").unwrap()).unwrap();
        let hide = wb.set_table_totals_visible(id, false, [2, 3].into()).unwrap();
        let mut view = crate::app::PreviewViewState::default();
        for (commit, expected) in [(show, "10"), (edit, "1"), (hide, "")] {
            let action = UndoAction::TableCommit { sheet_index: 0, commit: Box::new(commit.clone()), header_layout: None, description: "Totals".into() };
            History::apply_action_forward(&mut replay, &mut view, &action, true).unwrap();
            assert_eq!(replay.sheet(0).unwrap().get_display(3, 0), expected);
            assert_eq!(replay.sheet(0).unwrap().table_view_spec(), Some(&spec));
            assert_eq!(replay.sheet(0).unwrap().manual_hidden_rows(), [2, 3].into());
            let before = crate::table_totals::prepare_replay(&replay, &commit, true).unwrap();
            let after = crate::table_totals::prepare_replay(&before, &commit, false).unwrap();
            assert_eq!(after.sheet(0).unwrap().get_display(3, 0), expected);
            assert_eq!(after.sheet(0).unwrap().manual_hidden_rows(), [2, 3].into());
        }
    }

    #[test]
    fn calculated_column_rewind_preserves_cleared_exceptions_and_rule_updates() {
        use visigrid_engine::table::TableRange;
        let mut wb = Workbook::new();
        let id = wb.create_table(wb.active_sheet_id(), TableRange { start_row:0,start_col:0,end_row:3,end_col:1 }, "Sales").unwrap().table_id();
        let mut replay = wb.clone();
        let mut view = crate::app::PreviewViewState::default();
        let rule = wb.set_calculated_column(id,1,1,"=A2*2",true).unwrap();
        History::apply_action_forward(&mut replay,&mut view,&UndoAction::TableCommit { header_layout: None,sheet_index:0,commit:Box::new(rule),description:"Formula rule".into()}, true).unwrap();
        let before = wb.active_sheet().get_raw(2,1);
        wb.clear_cell_tracked(0,2,1);
        History::apply_action_forward(&mut replay,&mut view,&UndoAction::Values {sheet_index:0,changes:vec![CellChange {row:2,col:1,old_value:before,new_value:String::new()}]}, true).unwrap();
        let update = wb.set_calculated_column(id,1,1,"=A2*3",false).unwrap();
        History::apply_action_forward(&mut replay,&mut view,&UndoAction::TableCommit { header_layout: None,sheet_index:0,commit:Box::new(update),description:"Update rule".into()}, true).unwrap();
        assert!(replay.active_sheet().is_calculated_exception(2,1));
        assert_eq!(replay.active_sheet().get_raw(2,1),"");
        assert_eq!(replay.active_sheet().get_raw(3,1),"=A4*3");
        assert_eq!(replay.table(id).unwrap().1,wb.table(id).unwrap().1);
    }

    #[test]
    fn table_history_replay_reports_stale_state_without_overwriting() {
        use visigrid_engine::table::TableRange;
        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_id();
        workbook.set_cell_value_tracked(0, 0, 0, "Amount");
        let create = workbook
            .create_table(
                sheet,
                TableRange {
                    start_row: 0,
                    start_col: 0,
                    end_row: 1,
                    end_col: 0,
                },
                "Sales",
            )
            .unwrap();
        let rename = workbook.rename_table(create.table_id(), "Orders").unwrap();
        workbook.apply_table_commit(&rename, true).unwrap();
        workbook.rename_table(create.table_id(), "Changed").unwrap();
        let action = UndoAction::TableCommit { header_layout: None,
            sheet_index: 0,
            commit: Box::new(rename),
            description: "Rename Table".into(),
        };
        assert!(History::apply_action_forward(
            &mut workbook,
            &mut crate::app::PreviewViewState::default(),
            &action,
            true
        )
        .is_err());
        assert!(workbook.table_by_name("Changed").is_some());
    }

    #[test]
    fn print_setup_replay_targets_sheet_identity_and_preserves_values() {
        use visigrid_engine::print_setup::{PrintRows, PrintSetup};
        let mut workbook = Workbook::new();
        let index = workbook.add_sheet_named("Report").unwrap();
        let sheet = workbook.sheet_mut(index).unwrap();
        sheet.set_value(0, 0, "Title");
        let sheet_id = sheet.id;
        let after = PrintSetup {
            gridlines: true,
            repeat_rows: Some(PrintRows { start: 0, end: 2 }),
            ..Default::default()
        };
        let action = UndoAction::PrintSetupChanged {
            sheet_id,
            before: PrintSetup::default(),
            after: after.clone(),
        };
        let sheet = workbook.take_sheet(index).unwrap();
        assert!(workbook.restore_sheet(0, sheet));
        let mut view = crate::app::PreviewViewState::default();
        History::apply_action_forward(&mut workbook, &mut view, &action, true).unwrap();
        assert_eq!(workbook.sheet(0).unwrap().print_setup, after);
        assert!(workbook.sheet(1).unwrap().print_setup.is_default());
        assert_eq!(workbook.sheet(0).unwrap().get_raw(0, 0), "Title");
        workbook.take_sheet(0).unwrap();
        assert!(History::apply_action_forward(&mut workbook, &mut view, &action, true).is_err());
    }

    /// Large workbooks keep no load-time copy (#18); preview must say so
    /// rather than replay from something that isn't there.
    #[test]
    fn preview_without_a_base_snapshot_is_refused() {
        let mut history = History::new();
        let result = history.build_workbook_before(0, None, 100, 1_000);
        assert!(matches!(result, Err(PreviewBuildError::NoBaseSnapshot)));

        let base = Workbook::new();
        assert!(history.build_workbook_before(0, Some(&base), 100, 1_000).is_ok());
    }

    #[test]
    fn workbook_snapshot_undo_redo_restores_complete_sheet_and_monotonic_revisions() {
        use visigrid_engine::sheet::MergedRegion;

        let before = Workbook::new();
        let mut after = before.clone();
        let copied_index = after.add_sheet_named("Copied preview").unwrap();
        let copied = after.sheet_mut(copied_index).unwrap();
        copied.set_value(0, 0, "kept");
        copied.add_merge(MergedRegion::new(0, 0, 0, 1)).unwrap();
        assert!(after.set_active_sheet(copied_index));

        let commit = WorkbookSnapshotCommit::new("Copy reviewed result", before, after.clone());
        let mut current = after;
        let applied_revision = current.revision();

        commit.undo_into(&mut current);
        assert_eq!(current.sheet_count(), 1);
        assert_eq!(current.revision(), applied_revision + 1);

        commit.redo_into(&mut current);
        assert_eq!(current.sheet_count(), 2);
        assert_eq!(current.active_sheet_index(), copied_index);
        assert_eq!(current.sheet(copied_index).unwrap().get_raw(0, 0), "kept");
        assert_eq!(current.sheet(copied_index).unwrap().merged_regions.len(), 1);
        assert_eq!(current.revision(), applied_revision + 2);
    }

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
        let patches = vec![CellFormatPatch { remove_cell_on_undo: false,
            row: 0,
            col: 0,
            before: CellFormat::default(),
            after: CellFormat { bold: true, ..Default::default() },
        }];

        // History 1: Values then Format
        history1.record_batch(&visigrid_engine::workbook::Workbook::new(), 0, changes.clone());
        history1.record_format(&visigrid_engine::workbook::Workbook::new(), 0, patches.clone(), FormatActionKind::Bold, "Bold".into());

        // History 2: Format then Values
        history2.record_format(&visigrid_engine::workbook::Workbook::new(), 0, patches, FormatActionKind::Bold, "Bold".into());
        history2.record_batch(&visigrid_engine::workbook::Workbook::new(), 0, changes);

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

        history.record_batch(&visigrid_engine::workbook::Workbook::new(), 0, vec![CellChange {
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

        history.record_batch(&visigrid_engine::workbook::Workbook::new(), 0, vec![CellChange {
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
            history.record_batch(&visigrid_engine::workbook::Workbook::new(), 0, vec![CellChange {
                row: i, col: 0,
                old_value: "".to_string(),
                new_value: format!("value{}", i),
            }]);
        }

        assert_eq!(history.undo_count(), 5);

        // Truncate at index 2 (keep entries 0, 1; discard 2, 3, 4)
        let target_id = history.entry_at(2).unwrap().id;
        history.truncate_and_append_rewind(
            &visigrid_engine::workbook::Workbook::new(),
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

        history.record_batch(&visigrid_engine::workbook::Workbook::new(), 0, vec![CellChange {
            row: 0, col: 0,
            old_value: "".to_string(),
            new_value: "initial".to_string(),
        }]);

        // Take fingerprint
        let fp_before = history.fingerprint();

        // Simulate concurrent change
        history.record_batch(&visigrid_engine::workbook::Workbook::new(), 0, vec![CellChange {
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
            UndoActionKind::Comments,
            UndoActionKind::Values,
            UndoActionKind::Format,
            UndoActionKind::NamedRangeCreated,
            UndoActionKind::NamedRangeDeleted,
            UndoActionKind::NamedRangeRenamed,
            UndoActionKind::NamedRangeDescriptionChanged,
            UndoActionKind::Group,
            UndoActionKind::PlanCommit,
            UndoActionKind::PrintSetupChanged,
            UndoActionKind::WorkbookSnapshot,
            UndoActionKind::TableCommit,
            UndoActionKind::TableAppend,
            UndoActionKind::ReviewCopy,
            UndoActionKind::TableViewChanged,
            UndoActionKind::TableCellsChanged,
            UndoActionKind::TableStructureChanged,
            UndoActionKind::TableBatchChanged,
            UndoActionKind::PivotCommit,
            UndoActionKind::RowsInserted,
            UndoActionKind::RowsDeleted,
            UndoActionKind::ColsInserted,
            UndoActionKind::ColsDeleted,
            UndoActionKind::SortApplied,
            UndoActionKind::SortCleared,
            UndoActionKind::ValidationChanged,
            UndoActionKind::ValidationSet,
            UndoActionKind::ValidationCleared,
            UndoActionKind::ValidationExcluded,
            UndoActionKind::ValidationExclusionCleared,
            UndoActionKind::ColumnWidthSet,
            UndoActionKind::RowHeightSet,
            UndoActionKind::RowVisibilityChanged,
            UndoActionKind::ColVisibilityChanged,
            UndoActionKind::FreezePanesChanged,
            UndoActionKind::SetMerges,
            UndoActionKind::Rewind,
        ];

        let tags: HashSet<u8> = kinds.iter().map(|k| k.tag()).collect();
        assert_eq!(tags.len(), kinds.len(), "All action kind tags must be unique");
    }
}

#[cfg(test)]
mod comment_tests {
    use super::*;
    use visigrid_engine::cell::CellComment;
    #[test]
    fn comment_undo_redo_and_history_replay_preserve_values() {
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_text(0, 0, "00123");
        let before = CellComment { text: "Original".into(), author: "Alice".into() };
        let after = CellComment { text: "Updated\n日本語".into(), author: "Alice".into() };
        let patches = vec![CommentPatch { remove_cell_on_undo: false, row: 0, col: 0, before: Some(before.clone()), after: Some(after.clone()) }];
        apply_comment_patches(&mut wb, 0, &patches, true);
        assert_eq!(wb.active_sheet().comment(0,0), Some(&after));
        let rev = wb.revision();
        apply_comment_patches(&mut wb, 0, &patches, false);
        assert_eq!(wb.active_sheet().comment(0,0), Some(&before));
        assert!(wb.revision() > rev);
        let action = UndoAction::Comments { sheet_index: 0, patches, description: "Edit comment".into() };
        assert!(action.is_replay_supported());
        History::apply_action_forward(&mut wb, &mut Default::default(), &action, true).unwrap();
        assert_eq!(wb.active_sheet().comment(0,0), Some(&after));
        assert_eq!(wb.active_sheet().get_raw(0,0), "00123");
        let delete = vec![CommentPatch {remove_cell_on_undo: false,row:0,col:0,before:Some(after.clone()),after:None}];
        apply_comment_patches(&mut wb,0,&delete,true);
        assert!(wb.active_sheet().comment(0,0).is_none());
        apply_comment_patches(&mut wb,0,&delete,false);
        assert_eq!(wb.active_sheet().comment(0,0), Some(&after));
    }
}

#[cfg(test)]
mod byte_budget_tests {
    use super::*;
    static REWIND_SPIKE: std::sync::Mutex<()> = std::sync::Mutex::new(());

    #[test]
    fn retagging_a_source_reaccounts_its_retained_payload() {
        let mut history = History::new();
        history.max_bytes = 32_000;
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 0, 0, String::new(), "a".into());
        history.retag_last_source(&visigrid_engine::workbook::Workbook::new(), MutationSource::Agent { client: "x".repeat(40_000) });
        assert!(history.last_record_too_large());
        assert!(!history.can_undo());
        assert!(history.is_dirty());
    }

    #[test]
    fn oldest_entries_are_evicted_by_bytes_and_redo_transfers_keep_the_budget() {
        let mut history = History::new();
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 0, 0, String::new(), "a".repeat(4_000));
        let first = history.undo_stack.last().map(|e| e.id);
        let one = history.entry_bytes.values().sum::<usize>();
        assert!(one > 4_000, "history bytes should count the stored text, got {one}");
        // Two measured entries do not fit; the oldest one is dropped.
        history.max_bytes = one + one / 2;
        history.mark_saved();
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 1, 0, String::new(), "b".repeat(4_000));
        assert!(history.undo_stack.iter().all(|e| Some(e.id) != first));
        assert!(history.is_dirty());
        assert!(matches!(history.build_workbook_before(0, Some(&Workbook::new()), 100, 1_000), Err(PreviewBuildError::InvariantViolation(_))));
        let bytes = history.entry_bytes.values().sum::<usize>();
        assert!(bytes <= history.max_bytes);
        history.undo().unwrap();
        assert_eq!(history.entry_bytes.values().sum::<usize>(), bytes);
        history.redo().unwrap();
        assert_eq!(history.entry_bytes.values().sum::<usize>(), bytes);
        assert!(history.is_dirty());
    }

    #[test]
    fn oversized_mutation_clears_both_stacks_and_cannot_jump_back_over_it() {
        let mut history = History::new();
        history.max_bytes = 32_000;
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 0, 0, String::new(), "a".into());
        history.mark_saved();
        history.undo().unwrap();
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 1, 0, String::new(), "b".repeat(40_000));
        assert!(history.last_record_too_large());
        assert!(!history.can_undo());
        assert!(!history.can_redo());
        assert!(history.is_dirty());
        assert!(history.entry_bytes.is_empty());
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 2, 0, String::new(), "c".into());
        assert!(!history.last_record_too_large());
        assert!(history.can_undo());
    }

    #[test]
    fn count_eviction_keeps_rewind_to_the_oldest_retained_entry() {
        assert_eq!(History::new().max_entries, 100);
        let base = Workbook::new();
        let mut history = History::new();
        history.max_entries = 3;
        history.set_rewind_base(&base);
        for i in 0..8 {
            history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, i, 0, String::new(), format!("v{i}"));
        }
        assert!(!history.base_invalidated);
        assert!(history.rewind_base.is_some());
        assert_eq!(history.undo_stack.len(), 3);
        assert!(!history.last_record_too_large());
        let oldest = history.build_workbook_before(0, None, 20, 10_000).expect("rewind after count eviction");
        assert_eq!(oldest.workbook.active_sheet().get_raw(4, 0), "v4");
        assert_eq!(oldest.workbook.active_sheet().get_raw(5, 0), "");
        let current = history.build_workbook_before(history.undo_stack.len(), None, 20, 10_000).unwrap();
        assert_eq!(current.workbook.active_sheet().get_raw(4, 0), "v4");
        assert_eq!(current.workbook.active_sheet().get_raw(7, 0), "v7");
    }

    #[test]
    fn clear_keeps_the_captured_base_so_later_eviction_still_rewinds() {
        let base = Workbook::new();
        let mut history = History::new();
        history.set_rewind_base(&base);
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 0, 0, String::new(), "stale".into());
        history.clear();
        assert!(history.rewind_base.is_some());
        assert!(!history.base_invalidated);
        assert!(!history.can_undo());
        history.max_entries = 2;
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 0, 0, String::new(), "b".into());
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 1, 0, String::new(), "c".into());
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 2, 0, String::new(), "d".into());
        let oldest = history.build_workbook_before(0, None, 20, 10_000).expect("rewind after clear and eviction");
        assert_eq!(oldest.workbook.active_sheet().get_raw(0, 0), "b");
        assert_eq!(oldest.workbook.active_sheet().get_raw(1, 0), "");
    }

    #[test]
    fn eviction_stores_values_without_recalculating_until_rewind_opens() {
        use std::sync::atomic::{AtomicUsize, Ordering};
        use visigrid_engine::custom_fns;
        use visigrid_engine::formula::eval::{EvalArg, EvalResult};
        static CALLS: AtomicUsize = AtomicUsize::new(0);
        fn spike(name: &str, _: &[EvalArg]) -> Option<EvalResult> {
            (name == "REWINDSPIKE").then(|| {
                CALLS.fetch_add(1, Ordering::SeqCst);
                EvalResult::Number(1.0)
            })
        }
        struct Reset(Option<custom_fns::CustomFnHandler>);
        impl Drop for Reset {
            fn drop(&mut self) { custom_fns::set_default_custom_fn_handler(self.0); }
        }
        let _lock = REWIND_SPIKE.lock().unwrap();
        let _reset = Reset(custom_fns::default_custom_fn_handler());
        custom_fns::set_default_custom_fn_handler(Some(spike));
        let mut base = Workbook::new();
        base.set_cell_value_tracked(0, 0, 1, "=REWINDSPIKE()+A1");
        CALLS.store(0, Ordering::SeqCst);
        let mut history = History::new();
        history.max_entries = 1;
        history.set_rewind_base(&base);
        // A formula, not a literal: set_value would evaluate REWINDSPIKE here.
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 0, 0, String::new(), "=REWINDSPIKE()".into());
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 2, 0, String::new(), "z".into());
        assert_eq!(CALLS.load(Ordering::SeqCst), 0, "eviction must not calculate");
        let opened = history.build_workbook_before(0, None, 20, 10_000).expect("rewind");
        assert!(CALLS.load(Ordering::SeqCst) > 0, "opening rewind calculates the stored prefix");
        let calls_after_open = CALLS.load(Ordering::SeqCst);
        assert_eq!(opened.workbook.active_sheet().get_raw(0, 0), "=REWINDSPIKE()");
        assert_eq!(opened.workbook.active_sheet().get_display(0, 0), "1");
        assert_eq!(opened.workbook.active_sheet().get_display(0, 1), "2");
        let again = history.build_workbook_before(0, None, 20, 10_000).expect("second preview");
        assert_eq!(CALLS.load(Ordering::SeqCst), calls_after_open, "a second preview must not recalculate the settled baseline");
        assert_eq!(again.workbook.active_sheet().get_display(0, 1), "2");
        let later = history.build_workbook_before(1, None, 20, 10_000).expect("scrub");
        assert_eq!(CALLS.load(Ordering::SeqCst), calls_after_open, "scrubbing past the settled prefix must not recalculate it");
        assert_eq!(later.workbook.active_sheet().get_raw(2, 0), "z");
        assert_eq!(later.workbook.active_sheet().get_display(0, 1), "2");
    }

    #[test]
    fn evicted_table_edits_wait_for_rewind_and_then_match_the_live_sheet() {
        use std::sync::atomic::{AtomicUsize, Ordering};
        use visigrid_engine::custom_fns;
        use visigrid_engine::filter::{ColumnFilter, NormalizedFilterKey};
        use visigrid_engine::formula::eval::{EvalArg, EvalResult};
        use visigrid_engine::table::TableRange;
        use visigrid_engine::table_view::{TableFilter, TableViewSpec};
        static CALLS: AtomicUsize = AtomicUsize::new(0);
        fn spike(name: &str, _: &[EvalArg]) -> Option<EvalResult> {
            (name == "TABLESPIKE").then(|| {
                CALLS.fetch_add(1, Ordering::SeqCst);
                EvalResult::Number(1.0)
            })
        }
        struct Reset(Option<custom_fns::CustomFnHandler>);
        impl Drop for Reset {
            fn drop(&mut self) { custom_fns::set_default_custom_fn_handler(self.0); }
        }
        let _lock = REWIND_SPIKE.lock().unwrap();
        let _reset = Reset(custom_fns::default_custom_fn_handler());
        custom_fns::set_default_custom_fn_handler(Some(spike));

        let mut base = Workbook::new();
        base.set_cell_value_tracked(0, 0, 0, "Amount");
        base.set_cell_value_tracked(0, 1, 0, "10");
        base.set_cell_value_tracked(0, 2, 0, "20");
        let id = base.create_table(base.active_sheet_id(), TableRange {
            start_row: 0, start_col: 0, end_row: 2, end_col: 0,
        }, "Sales").unwrap().table_id();
        base.set_cell_value_tracked(0, 0, 2, "=TABLESPIKE()+SUBTOTAL(109,Sales[Amount])");
        CALLS.store(0, Ordering::SeqCst);

        let mut live = base.clone();
        let mut spec = TableViewSpec::new(id);
        spec.filters.push(TableFilter {
            column: live.table(id).unwrap().1.columns[0].id,
            criteria: ColumnFilter { selected: Some([NormalizedFilterKey::Number(10.0.into())].into()), text_filter: None },
        });
        let view = live.set_table_view_spec(live.active_sheet_id(), Some(spec)).unwrap();
        let style = live.set_table_style(id, visigrid_engine::table::TableStyle { banded_rows: false, ..Default::default() }).unwrap();
        let mut edited = live.clone();
        edited.set_cell_value_tracked(0, 1, 0, "99");
        let cells = crate::table_cell_history::TableCellsCommit::capture(live.active_sheet(), edited.active_sheet(), [(1, 0)]);
        edited.recompute_full_ordered();
        CALLS.store(0, Ordering::SeqCst);

        let mut history = History::new();
        history.max_entries = 1;
        history.set_rewind_base(&base);
        history.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableViewChanged {
            sheet_index: 0, commit: Box::new(view), description: "Filter".into(),
        }, None);
        history.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableCommit {
            header_layout: None, sheet_index: 0, commit: Box::new(style), description: "Banding".into(),
        }, None);
        history.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), UndoAction::TableCellsChanged {
            sheet_index: 0, commit: Box::new(cells), description: "Cell".into(),
        }, None);
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 4, 0, String::new(), "kept".into());
        assert_eq!(CALLS.load(Ordering::SeqCst), 0, "evicting Table edits must not calculate");
        let opened = history.build_workbook_before(0, None, 20, 10_000).expect("rewind");
        assert!(CALLS.load(Ordering::SeqCst) > 0, "opening rewind calculates the stored prefix");
        assert_eq!(opened.workbook.active_sheet().get_raw(1, 0), edited.active_sheet().get_raw(1, 0));
        assert_eq!(opened.workbook.active_sheet().get_display(0, 2), edited.active_sheet().get_display(0, 2));
        assert_eq!(opened.workbook.active_sheet().table_view_spec(), edited.active_sheet().table_view_spec());
        assert_eq!(opened.workbook.table(id).unwrap().1.style, edited.table(id).unwrap().1.style);
        assert_eq!(opened.workbook.active_sheet().get_raw(4, 0), "");
    }

    #[test]
    fn evicted_row_insert_waits_for_rewind_and_then_matches_the_live_sheet() {
        use std::sync::atomic::{AtomicUsize, Ordering};
        use visigrid_engine::custom_fns;
        use visigrid_engine::formula::eval::{EvalArg, EvalResult};
        use visigrid_engine::table::TableRange;
        static CALLS: AtomicUsize = AtomicUsize::new(0);
        fn spike(name: &str, _: &[EvalArg]) -> Option<EvalResult> {
            (name == "TABLESPIKE").then(|| {
                CALLS.fetch_add(1, Ordering::SeqCst);
                EvalResult::Number(1.0)
            })
        }
        struct Reset(Option<custom_fns::CustomFnHandler>);
        impl Drop for Reset {
            fn drop(&mut self) { custom_fns::set_default_custom_fn_handler(self.0); }
        }
        let _lock = REWIND_SPIKE.lock().unwrap();
        let _reset = Reset(custom_fns::default_custom_fn_handler());
        custom_fns::set_default_custom_fn_handler(Some(spike));

        let mut base = Workbook::new();
        base.set_cell_value_tracked(0, 0, 0, "Amount");
        base.set_cell_value_tracked(0, 1, 0, "10");
        base.set_cell_value_tracked(0, 2, 0, "20");
        let id = base.create_table(base.active_sheet_id(), TableRange {
            start_row: 0, start_col: 0, end_row: 2, end_col: 0,
        }, "Sales").unwrap().table_id();
        // No cell references, so inserting a row does not rewrite this formula.
        // Only a full recalculation evaluates it.
        base.set_cell_value_tracked(0, 6, 0, "=TABLESPIKE()");
        let rows = base.prepare_table_row_history(0, 2, 1, false).unwrap().expect("table row history");
        let mut live = base.clone();
        live.apply_table_row_history(&rows, false).unwrap();
        CALLS.store(0, Ordering::SeqCst);
        assert_eq!(live.active_sheet().get_display(7, 0), "1");
        assert!(live.table(id).unwrap().1.range.end_row > 2);

        let mut history = History::new();
        history.max_entries = 1;
        history.set_rewind_base(&base);
        history.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(), UndoAction::RowsInserted {
            sheet_index: 0,
            table_rows: Some(rows),
            row_layout: None,
            print_setup_before: visigrid_engine::print_setup::PrintSetup::default(),
            at_row: 2,
            count: 1,
            formula_rewrites: vec![],
        }, None);
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 8, 0, String::new(), "kept".into());
        assert_eq!(CALLS.load(Ordering::SeqCst), 0, "evicting a row insert must not calculate");
        let opened = history.build_workbook_before(0, None, 20, 10_000).expect("rewind");
        assert!(CALLS.load(Ordering::SeqCst) > 0, "opening rewind calculates the stored insert");
        assert_eq!(opened.workbook.table(id).unwrap().1.range, live.table(id).unwrap().1.range);
        assert_eq!(opened.workbook.active_sheet().get_display(7, 0), live.active_sheet().get_display(7, 0));
        assert_eq!(opened.workbook.active_sheet().get_raw(6, 0), "");
        assert_eq!(opened.workbook.active_sheet().get_raw(8, 0), "");
    }

    #[test]
    fn failed_eviction_replay_reports_that_rewind_is_unavailable() {
        let mut history = History::new();
        history.set_rewind_base(&Workbook::new());
        history.max_entries = 1;
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 5, 0, 0, String::new(), "a".into());
        history.record_change(&visigrid_engine::workbook::Workbook::new(), 0, 0, 0, String::new(), "b".into());
        assert!(history.base_invalidated);
        assert!(history.rewind_base.is_none());
        assert!(history.can_undo());
        let notice = history.take_notice().unwrap_or_default();
        assert!(notice.contains("Rewind is unavailable"), "{notice}");
    }

    #[test]
    fn rewind_base_larger_than_the_budget_is_dropped_with_a_notice() {
        let mut workbook = Workbook::new();
        for row in 0..1024 {
            workbook.set_cell_value_tracked(0, row, 0, "xxxxxxxxxx");
        }
        let mut history = History::new();
        history.max_bytes = 1_000_000;
        history.set_rewind_base(&workbook);
        assert!(history.rewind_base_bytes < 1_000, "capture charges copied maps, not the shared column, got {}", history.rewind_base_bytes);
        workbook.set_cell_value_tracked(0, 0, 0, "y");
        history.record_change(&workbook, 0, 0, 0, "xxxxxxxxxx".into(), "y".into());
        assert!(history.rewind_base_bytes > 2_000, "a one-character edit copies the whole chunk, got {}", history.rewind_base_bytes);
        history.max_bytes = history.rewind_base_bytes - 1;
        history.record_change(&workbook, 0, 1, 0, String::new(), "z".into());
        assert!(history.base_invalidated);
        assert!(history.rewind_base.is_none());
        let notice = history.take_notice().unwrap_or_default();
        assert!(notice.contains("Rewind is unavailable"), "{notice}");
    }

    #[test]
    fn shared_baseline_keeps_rewind_for_a_workbook_larger_than_the_budget() {
        let mut workbook = Workbook::new();
        for row in 0..200 {
            workbook.set_cell_value_tracked(0, row, 0, "xxxxxxxxxx");
        }
        let bytes = visigrid_engine::history_size::workbook_bytes(&workbook);
        assert!(bytes > 1_000, "the fixture must be larger than the budget, got {bytes}");
        let mut history = History::new();
        history.max_bytes = 1_000;
        history.max_entries = 1;
        history.set_rewind_base(&workbook);
        assert!(history.rewind_base_bytes < 1_000, "capture charges copied maps, not the shared cells, got {}", history.rewind_base_bytes);
        assert!(history.rewind_base_bytes < bytes / 2);
        history.record_change(&workbook, 0, 0, 1, String::new(), "a".into());
        history.record_change(&workbook, 0, 0, 2, String::new(), "b".into());
        assert!(history.rewind_base.is_some(), "a one-character divergence stays inside the budget");
        assert!(!history.base_invalidated);
        assert!(history.rewind_base_bytes < history.max_bytes);
        let opened = history.build_workbook_before(0, None, 20, 10_000).expect("rewind");
        assert_eq!(opened.workbook.active_sheet().get_raw(0, 1), "a");
        assert_eq!(opened.workbook.active_sheet().get_raw(0, 0), "xxxxxxxxxx");
    }

    #[test]
    fn too_large_edit_also_says_rewind_is_unavailable() {
        let live = Workbook::new();
        let mut history = History::new();
        history.set_rewind_base(&live);
        history.max_bytes = 32_000;
        history.record_change(&live, 0, 0, 0, String::new(), "a".into());
        history.record_change(&live, 0, 1, 0, String::new(), "b".repeat(40_000));
        let notice = history.take_notice().unwrap_or_default();
        assert!(notice.contains("too large to undo"), "{notice}");
        assert!(notice.contains("Rewind is unavailable"), "{notice}");
        assert!(history.rewind_base.is_none());
    }

    #[test]
    fn table_swap_keeps_the_copied_chunks_and_a_later_edit_on_the_charge() {
        let mut workbook = Workbook::new();
        for row in 0..2048 {
            workbook.set_cell_value_tracked(0, row, 0, "xxxxxxxxxx");
        }
        let second = workbook.add_sheet();
        for row in 0..1024 {
            workbook.set_cell_value_tracked(second, row, 0, "yyyyyyyyyy");
        }
        let mut history = History::new();
        history.set_rewind_base(&workbook);
        assert!(history.rewind_base_bytes < 1_000, "capture charges copied maps, not the shared columns, got {}", history.rewind_base_bytes);
        workbook.create_table_without_headers(workbook.active_sheet_id(), visigrid_engine::table::TableRange {
            start_row: 0, start_col: 0, end_row: 2, end_col: 0,
        }, "Sales").unwrap();
        history.record_change(&workbook, 0, 3, 0, String::new(), "n".into());
        let after_insert = history.rewind_base_bytes;
        assert!(after_insert > 8_000, "creating a table copies both chunks of the column, got {after_insert}");
        workbook.set_cell_value_tracked(second, 0, 0, "z");
        history.record_change(&workbook, second, 0, 0, "yyyyyyyyyy".into(), "z".into());
        assert!(history.rewind_base_bytes > after_insert, "a later edit stays charged after the workbook is replaced, got {} then {after_insert}", history.rewind_base_bytes);
    }
}
