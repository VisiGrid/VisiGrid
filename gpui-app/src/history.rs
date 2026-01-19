/// Undo/Redo history system for spreadsheet operations

#[derive(Clone, Debug)]
pub struct CellChange {
    pub row: usize,
    pub col: usize,
    pub old_value: String,
    pub new_value: String,
}

#[derive(Clone, Debug)]
pub struct HistoryEntry {
    pub changes: Vec<CellChange>,
}

pub struct History {
    undo_stack: Vec<HistoryEntry>,
    redo_stack: Vec<HistoryEntry>,
    max_entries: usize,
}

impl History {
    pub fn new() -> Self {
        Self {
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            max_entries: 100,
        }
    }

    /// Record a single cell change
    pub fn record_change(&mut self, row: usize, col: usize, old_value: String, new_value: String) {
        if old_value == new_value {
            return;
        }

        let entry = HistoryEntry {
            changes: vec![CellChange { row, col, old_value, new_value }],
        };
        self.push_entry(entry);
    }

    /// Record multiple cell changes as a single undoable operation
    pub fn record_batch(&mut self, changes: Vec<CellChange>) {
        if changes.is_empty() {
            return;
        }

        let entry = HistoryEntry { changes };
        self.push_entry(entry);
    }

    fn push_entry(&mut self, entry: HistoryEntry) {
        self.undo_stack.push(entry);
        self.redo_stack.clear();

        // Limit history size
        if self.undo_stack.len() > self.max_entries {
            self.undo_stack.remove(0);
        }
    }

    /// Pop the last entry for undo, returns the changes to apply
    pub fn undo(&mut self) -> Option<HistoryEntry> {
        if let Some(entry) = self.undo_stack.pop() {
            self.redo_stack.push(entry.clone());
            Some(entry)
        } else {
            None
        }
    }

    /// Pop from redo stack, returns the changes to apply
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
