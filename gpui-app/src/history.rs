/// Undo/Redo history system for spreadsheet operations

use visigrid_engine::cell::CellFormat;
use visigrid_engine::named_range::NamedRange;
use visigrid_engine::provenance::Provenance;
use std::time::Instant;

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
    /// When this action occurred
    pub timestamp: Instant,
    /// Lua snippet if provenance exists
    pub lua: Option<String>,
    /// Whether this entry has provenance
    pub is_provenanced: bool,
    /// Whether this entry can be undone (vs already undone/redoable)
    pub is_undoable: bool,
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
    Font,
    Alignment,
    VerticalAlignment,
    TextOverflow,
    NumberFormat,
    DecimalPlaces,  // Special: coalesces rapidly
    BackgroundColor,
    Border,
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
        name: String,
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
    /// Sort applied (for undo: restore previous row order)
    SortApplied {
        /// Previous row order before sorting
        previous_row_order: Vec<usize>,
        /// Previous sort state (column and direction)
        previous_sort_state: Option<(usize, bool)>, // (column, is_ascending)
        /// New row order after sorting (for redo)
        new_row_order: Vec<usize>,
        /// New sort state (column and direction) for redo
        new_sort_state: (usize, bool), // (column, is_ascending)
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
            UndoAction::NamedRangeCreated { name } => {
                format!("Create range '{}'", name)
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
            UndoAction::SortApplied { .. } => {
                "Sort".to_string()
            }
        }
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

    /// Record a single cell value change
    pub fn record_change(&mut self, sheet_index: usize, row: usize, col: usize, old_value: String, new_value: String) {
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
    fn to_display_entry(entry: &HistoryEntry, is_undoable: bool) -> HistoryDisplayEntry {
        let (label, scope, lua, is_provenanced) = if let Some(ref prov) = entry.provenance {
            (prov.label.clone(), prov.scope.clone(), Some(prov.lua.clone()), true)
        } else {
            // Fallback to UndoAction::label()
            (entry.action.label(), String::new(), None, false)
        };

        HistoryDisplayEntry {
            id: entry.id,  // Use stable entry ID, not position
            label,
            scope,
            timestamp: entry.timestamp,
            lua,
            is_provenanced,
            is_undoable,
        }
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
}
