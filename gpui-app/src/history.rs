/// Undo/Redo history system for spreadsheet operations

use visigrid_engine::cell::CellFormat;
use visigrid_engine::named_range::NamedRange;
use std::time::Instant;

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
}

#[derive(Clone, Debug)]
pub struct HistoryEntry {
    pub action: UndoAction,
    pub timestamp: Instant,
}

pub struct History {
    undo_stack: Vec<HistoryEntry>,
    redo_stack: Vec<HistoryEntry>,
    max_entries: usize,
}

/// Coalescing window for rapid format changes (e.g., decimal +/-)
const COALESCE_WINDOW_MS: u128 = 500;

impl History {
    pub fn new() -> Self {
        Self {
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            max_entries: 100,
        }
    }

    /// Record a single cell value change
    pub fn record_change(&mut self, sheet_index: usize, row: usize, col: usize, old_value: String, new_value: String) {
        if old_value == new_value {
            return;
        }

        let entry = HistoryEntry {
            action: UndoAction::Values {
                sheet_index,
                changes: vec![CellChange { row, col, old_value, new_value }],
            },
            timestamp: Instant::now(),
        };
        self.push_entry(entry);
    }

    /// Record multiple cell value changes as a single undoable operation
    pub fn record_batch(&mut self, sheet_index: usize, changes: Vec<CellChange>) {
        if changes.is_empty() {
            return;
        }

        let entry = HistoryEntry {
            action: UndoAction::Values { sheet_index, changes },
            timestamp: Instant::now(),
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
        let entry = HistoryEntry {
            action: UndoAction::Format { sheet_index, patches, kind, description },
            timestamp: now,
        };
        self.push_entry(entry);
    }

    /// Record a named range action (create, delete, rename)
    pub fn record_named_range_action(&mut self, action: UndoAction) {
        let entry = HistoryEntry {
            action,
            timestamp: Instant::now(),
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

        // Limit history size
        if self.undo_stack.len() > self.max_entries {
            self.undo_stack.remove(0);
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

    pub fn clear(&mut self) {
        self.undo_stack.clear();
        self.redo_stack.clear();
    }
}
