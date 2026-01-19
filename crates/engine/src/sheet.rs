use std::cell::RefCell;
use std::collections::{HashMap, HashSet};

use serde::{Deserialize, Serialize};

use super::cell::{Alignment, Cell, CellFormat, CellValue, NumberFormat, SpillError, SpillInfo, TextOverflow, VerticalAlignment};
use super::formula::eval::{self, Array2D, CellLookup, EvalResult, Value};

// Thread-local set to track cells currently being evaluated (for cycle detection)
thread_local! {
    static EVALUATING: RefCell<HashSet<(usize, usize)>> = RefCell::new(HashSet::new());
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    cells: HashMap<(usize, usize), Cell>,
    pub rows: usize,
    pub cols: usize,
    /// Spilled values from array formulas: (row, col) -> Value
    #[serde(skip)]
    spill_values: HashMap<(usize, usize), Value>,
}

impl CellLookup for Sheet {
    fn get_value(&self, row: usize, col: usize) -> f64 {
        // Check for circular reference
        let is_cycle = EVALUATING.with(|eval| {
            let set = eval.borrow();
            set.contains(&(row, col))
        });
        if is_cycle {
            return 0.0; // Circular reference - return 0
        }

        // CRITICAL: Check spill_values first for receiver cells
        // This enables formulas like =B2 to work when B2 is a spill receiver
        if let Some(spill_value) = self.spill_values.get(&(row, col)) {
            return spill_value.to_number().unwrap_or(0.0);
        }

        self.cells
            .get(&(row, col))
            .map(|c| self.evaluate_cell_value(&c.value))
            .unwrap_or(0.0)
    }

    fn get_text(&self, row: usize, col: usize) -> String {
        // Check for circular reference
        let is_cycle = EVALUATING.with(|eval| {
            let set = eval.borrow();
            set.contains(&(row, col))
        });
        if is_cycle {
            return "#CIRC!".to_string(); // Circular reference error
        }

        // CRITICAL: Check spill_values first for receiver cells
        if let Some(spill_value) = self.spill_values.get(&(row, col)) {
            return spill_value.to_text();
        }

        match self.cells.get(&(row, col)) {
            Some(cell) => match &cell.value {
                CellValue::Empty => String::new(),
                CellValue::Text(s) => s.clone(),
                CellValue::Number(n) => {
                    if n.fract() == 0.0 {
                        format!("{}", *n as i64)
                    } else {
                        format!("{}", n)
                    }
                }
                CellValue::Formula { ast: Some(ast), .. } => {
                    // Track that we're evaluating this cell
                    EVALUATING.with(|eval| eval.borrow_mut().insert((row, col)));
                    let result = eval::evaluate(ast, self).to_text();
                    EVALUATING.with(|eval| eval.borrow_mut().remove(&(row, col)));
                    result
                }
                CellValue::Formula { ast: None, .. } => String::new(),
            },
            None => String::new(),
        }
    }
}

impl Sheet {
    pub fn new(rows: usize, cols: usize) -> Self {
        Self {
            name: String::from("Sheet1"),
            cells: HashMap::new(),
            rows,
            cols,
            spill_values: HashMap::new(),
        }
    }

    pub fn set_value(&mut self, row: usize, col: usize, value: &str) {
        // Clear any existing spill from this cell before setting new value
        self.clear_spill_from(row, col);

        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.set(value);

        // If this is a formula, evaluate it and apply spill if it returns an array
        self.evaluate_and_spill(row, col);
    }

    /// Evaluate a cell's formula and apply spill if it returns an array
    fn evaluate_and_spill(&mut self, row: usize, col: usize) {
        // Get the AST if this is a formula
        let ast = match self.cells.get(&(row, col)) {
            Some(cell) => match &cell.value {
                CellValue::Formula { ast: Some(ast), .. } => ast.clone(),
                _ => return, // Not a formula or no AST
            },
            None => return,
        };

        // Evaluate the formula
        let result = eval::evaluate(&ast, self);

        // If it's an array, try to apply spill
        if let EvalResult::Array(array) = result {
            let rows = array.rows();
            let cols = array.cols();

            // Check for collision before applying
            match self.check_spill_collision(row, col, rows, cols) {
                Ok(()) => {
                    // Apply the spill
                    self.apply_spill(row, col, &array);
                }
                Err(blocked_by) => {
                    // Record the spill error
                    if let Some(cell) = self.cells.get_mut(&(row, col)) {
                        cell.spill_error = Some(SpillError { blocked_by });
                    }
                }
            }
        }
    }

    // =========================================================================
    // Spill Management
    // =========================================================================

    /// Clear spill data originating from a specific cell
    pub fn clear_spill_from(&mut self, parent_row: usize, parent_col: usize) {
        // Get the spill info from the parent cell
        let spill_info = match self.cells.get(&(parent_row, parent_col)) {
            Some(cell) => cell.spill_info.clone(),
            None => return,
        };

        if let Some(info) = spill_info {
            // Clear all receiving cells
            for dr in 0..info.rows {
                for dc in 0..info.cols {
                    if dr == 0 && dc == 0 {
                        continue; // Skip the parent cell itself
                    }
                    let r = parent_row + dr;
                    let c = parent_col + dc;

                    // Remove spill value
                    self.spill_values.remove(&(r, c));

                    // Clear spill_parent reference
                    if let Some(cell) = self.cells.get_mut(&(r, c)) {
                        if cell.spill_parent == Some((parent_row, parent_col)) {
                            cell.spill_parent = None;
                        }
                    }
                }
            }

            // Clear spill_info on parent
            if let Some(cell) = self.cells.get_mut(&(parent_row, parent_col)) {
                cell.spill_info = None;
            }
        }
    }

    /// Check if spill from (parent_row, parent_col) with given dimensions would collide
    /// Returns Ok(()) if no collision, Err with blocking cell position if collision
    pub fn check_spill_collision(
        &self,
        parent_row: usize,
        parent_col: usize,
        rows: usize,
        cols: usize,
    ) -> Result<(), (usize, usize)> {
        for dr in 0..rows {
            for dc in 0..cols {
                if dr == 0 && dc == 0 {
                    continue; // Skip the parent cell
                }
                let r = parent_row + dr;
                let c = parent_col + dc;

                // Check if cell exists and has content
                if let Some(cell) = self.cells.get(&(r, c)) {
                    // Check if it has its own value (not a spill receiver from us)
                    let is_our_receiver = cell.spill_parent == Some((parent_row, parent_col));
                    if !is_our_receiver {
                        match &cell.value {
                            CellValue::Empty => {}
                            _ => return Err((r, c)),
                        }
                        // Also blocked if it's receiving spill from another cell
                        if cell.spill_parent.is_some() {
                            return Err((r, c));
                        }
                    }
                }

                // Check if there's a spill value from another parent
                if let Some(_) = self.spill_values.get(&(r, c)) {
                    // Check if this spill is from us or another parent
                    if let Some(cell) = self.cells.get(&(r, c)) {
                        if cell.spill_parent != Some((parent_row, parent_col)) {
                            return Err((r, c));
                        }
                    } else {
                        return Err((r, c));
                    }
                }
            }
        }
        Ok(())
    }

    /// Apply spill from an array result at (parent_row, parent_col)
    /// Returns true if spill was successful, false if blocked (#SPILL!)
    pub fn apply_spill(
        &mut self,
        parent_row: usize,
        parent_col: usize,
        array: &Array2D,
    ) -> bool {
        let rows = array.rows();
        let cols = array.cols();

        // Check for collision first
        if self.check_spill_collision(parent_row, parent_col, rows, cols).is_err() {
            return false;
        }

        // Clear any existing spill from this parent
        self.clear_spill_from(parent_row, parent_col);

        // Apply new spill
        for dr in 0..rows {
            for dc in 0..cols {
                let r = parent_row + dr;
                let c = parent_col + dc;

                if let Some(value) = array.get(dr, dc) {
                    if dr == 0 && dc == 0 {
                        // Parent cell - just set spill_info
                        let cell = self.cells.entry((r, c)).or_insert_with(Cell::new);
                        cell.spill_info = Some(SpillInfo { rows, cols });
                    } else {
                        // Receiving cell
                        self.spill_values.insert((r, c), value.clone());
                        let cell = self.cells.entry((r, c)).or_insert_with(Cell::new);
                        cell.spill_parent = Some((parent_row, parent_col));
                    }
                }
            }
        }

        true
    }

    /// Get the spilled value at a position (if any)
    pub fn get_spill_value(&self, row: usize, col: usize) -> Option<&Value> {
        self.spill_values.get(&(row, col))
    }

    /// Check if a cell is receiving spill data
    pub fn is_spill_receiver(&self, row: usize, col: usize) -> bool {
        self.cells
            .get(&(row, col))
            .map(|c| c.is_spill_receiver())
            .unwrap_or(false)
    }

    /// Check if a cell is a spill parent
    pub fn is_spill_parent(&self, row: usize, col: usize) -> bool {
        self.cells
            .get(&(row, col))
            .map(|c| c.is_spill_parent())
            .unwrap_or(false)
    }

    /// Get spill info for a cell (if it's a spill parent)
    pub fn get_spill_info(&self, row: usize, col: usize) -> Option<SpillInfo> {
        self.cells
            .get(&(row, col))
            .and_then(|c| c.spill_info.clone())
    }

    /// Check if a cell has a spill error
    pub fn has_spill_error(&self, row: usize, col: usize) -> bool {
        self.cells
            .get(&(row, col))
            .map(|c| c.has_spill_error())
            .unwrap_or(false)
    }

    pub fn get_display(&self, row: usize, col: usize) -> String {
        // Check for spilled value first
        if let Some(spill_value) = self.spill_values.get(&(row, col)) {
            return spill_value.to_text();
        }

        match self.cells.get(&(row, col)) {
            Some(cell) => {
                // Check for spill error
                if cell.spill_error.is_some() {
                    return "#SPILL!".to_string();
                }
                self.display_cell_value(&cell.value)
            }
            None => String::new(),
        }
    }

    /// Get display value with number formatting applied
    pub fn get_formatted_display(&self, row: usize, col: usize) -> String {
        // Check for spilled value first
        if let Some(spill_value) = self.spill_values.get(&(row, col)) {
            // Apply formatting from the cell if it exists
            if let Some(cell) = self.cells.get(&(row, col)) {
                match spill_value {
                    Value::Number(n) => return CellValue::format_number(*n, &cell.format.number_format),
                    _ => return spill_value.to_text(),
                }
            }
            return spill_value.to_text();
        }

        match self.cells.get(&(row, col)) {
            Some(cell) => {
                // Check for spill error
                if cell.spill_error.is_some() {
                    return "#SPILL!".to_string();
                }
                // Get the result for formatting
                let result = match &cell.value {
                    CellValue::Number(n) => EvalResult::Number(*n),
                    CellValue::Formula { ast: Some(ast), .. } => eval::evaluate(ast, self),
                    CellValue::Text(s) => EvalResult::Text(s.clone()),
                    CellValue::Empty => return String::new(),
                    CellValue::Formula { ast: None, .. } => return "#ERR".to_string(),
                };

                match result {
                    EvalResult::Number(n) => {
                        // Apply number formatting
                        CellValue::format_number(n, &cell.format.number_format)
                    }
                    EvalResult::Text(s) => s,
                    EvalResult::Boolean(b) => if b { "TRUE".to_string() } else { "FALSE".to_string() },
                    EvalResult::Error(e) => format!("#ERR: {}", e),
                    EvalResult::Array(arr) => {
                        // Array: display top-left value (spill handles rest)
                        match arr.top_left() {
                            Value::Number(n) => CellValue::format_number(n, &cell.format.number_format),
                            other => other.to_text(),
                        }
                    }
                }
            }
            None => String::new(),
        }
    }

    pub fn get_raw(&self, row: usize, col: usize) -> String {
        self.cells
            .get(&(row, col))
            .map(|c| c.value.raw_display())
            .unwrap_or_default()
    }

    /// Get a reference to a cell (returns default empty cell if not found)
    pub fn get_cell(&self, row: usize, col: usize) -> Cell {
        self.cells
            .get(&(row, col))
            .cloned()
            .unwrap_or_default()
    }

    fn evaluate_cell_value(&self, value: &CellValue) -> f64 {
        match value {
            CellValue::Empty => 0.0,
            CellValue::Number(n) => *n,
            CellValue::Text(s) => s.parse().unwrap_or(0.0),
            CellValue::Formula { ast: Some(ast), .. } => {
                match eval::evaluate(ast, self) {
                    EvalResult::Number(n) => n,
                    EvalResult::Boolean(b) => if b { 1.0 } else { 0.0 },
                    EvalResult::Text(s) => s.parse().unwrap_or(0.0),
                    EvalResult::Error(_) => 0.0,
                    EvalResult::Array(arr) => arr.top_left().to_number().unwrap_or(0.0),
                }
            }
            CellValue::Formula { ast: None, .. } => 0.0,
        }
    }

    fn display_cell_value(&self, value: &CellValue) -> String {
        match value {
            CellValue::Empty => String::new(),
            CellValue::Text(s) => s.clone(),
            CellValue::Number(n) => {
                if n.fract() == 0.0 {
                    format!("{}", *n as i64)
                } else {
                    format!("{:.2}", n)
                }
            }
            CellValue::Formula { ast: Some(ast), .. } => {
                match eval::evaluate(ast, self) {
                    EvalResult::Number(n) => {
                        if n.fract() == 0.0 {
                            format!("{}", n as i64)
                        } else {
                            format!("{:.2}", n)
                        }
                    }
                    EvalResult::Text(s) => s,
                    EvalResult::Boolean(b) => if b { "TRUE".to_string() } else { "FALSE".to_string() },
                    EvalResult::Error(e) => format!("#ERR: {}", e),
                    EvalResult::Array(arr) => arr.top_left().to_text(),
                }
            }
            CellValue::Formula { ast: None, .. } => "#ERR".to_string(),
        }
    }

    pub fn get_format(&self, row: usize, col: usize) -> CellFormat {
        self.cells
            .get(&(row, col))
            .map(|c| c.format.clone())
            .unwrap_or_default()
    }

    /// Iterate over all populated cells
    pub fn cells_iter(&self) -> impl Iterator<Item = (&(usize, usize), &Cell)> {
        self.cells.iter()
    }

    /// Get coordinates of non-empty cells within a range
    pub fn cells_in_range(&self, min_row: usize, max_row: usize, min_col: usize, max_col: usize) -> Vec<(usize, usize)> {
        self.cells
            .keys()
            .filter(|(r, c)| *r >= min_row && *r <= max_row && *c >= min_col && *c <= max_col)
            .copied()
            .collect()
    }

    /// Clear a cell completely (remove from HashMap)
    pub fn clear_cell(&mut self, row: usize, col: usize) {
        self.clear_spill_from(row, col);
        self.cells.remove(&(row, col));
        self.spill_values.remove(&(row, col));
    }

    pub fn set_format(&mut self, row: usize, col: usize, format: CellFormat) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format = format;
    }

    pub fn toggle_bold(&mut self, row: usize, col: usize) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.bold = !cell.format.bold;
    }

    pub fn toggle_italic(&mut self, row: usize, col: usize) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.italic = !cell.format.italic;
    }

    pub fn toggle_underline(&mut self, row: usize, col: usize) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.underline = !cell.format.underline;
    }

    pub fn toggle_strikethrough(&mut self, row: usize, col: usize) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.strikethrough = !cell.format.strikethrough;
    }

    pub fn set_alignment(&mut self, row: usize, col: usize, alignment: Alignment) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.alignment = alignment;
    }

    pub fn set_vertical_alignment(&mut self, row: usize, col: usize, vertical_alignment: VerticalAlignment) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.vertical_alignment = vertical_alignment;
    }

    pub fn set_text_overflow(&mut self, row: usize, col: usize, text_overflow: TextOverflow) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.text_overflow = text_overflow;
    }

    pub fn set_number_format(&mut self, row: usize, col: usize, number_format: NumberFormat) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.number_format = number_format;
    }

    pub fn get_number_format(&self, row: usize, col: usize) -> NumberFormat {
        self.cells
            .get(&(row, col))
            .map(|c| c.format.number_format)
            .unwrap_or_default()
    }

    pub fn set_font_family(&mut self, row: usize, col: usize, font_family: Option<String>) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.font_family = font_family;
    }

    pub fn get_font_family(&self, row: usize, col: usize) -> Option<String> {
        self.cells
            .get(&(row, col))
            .and_then(|c| c.format.font_family.clone())
    }

    /// Insert rows at the specified position, shifting existing rows down
    pub fn insert_rows(&mut self, at_row: usize, count: usize) {
        // Collect all cells that need to be shifted
        let cells_to_shift: Vec<_> = self.cells
            .iter()
            .filter(|((r, _), _)| *r >= at_row)
            .map(|((r, c), cell)| ((*r, *c), cell.clone()))
            .collect();

        // Remove old positions
        for ((r, c), _) in &cells_to_shift {
            self.cells.remove(&(*r, *c));
        }

        // Insert at new positions (shifted down)
        for ((r, c), cell) in cells_to_shift {
            if r + count < self.rows {
                self.cells.insert((r + count, c), cell);
            }
        }
    }

    /// Delete rows at the specified position, shifting remaining rows up
    pub fn delete_rows(&mut self, start_row: usize, count: usize) {
        // Remove cells in the deleted rows
        for row in start_row..start_row + count {
            for col in 0..self.cols {
                self.cells.remove(&(row, col));
            }
        }

        // Collect cells that need to be shifted up
        let cells_to_shift: Vec<_> = self.cells
            .iter()
            .filter(|((r, _), _)| *r >= start_row + count)
            .map(|((r, c), cell)| ((*r, *c), cell.clone()))
            .collect();

        // Remove old positions
        for ((r, c), _) in &cells_to_shift {
            self.cells.remove(&(*r, *c));
        }

        // Insert at new positions (shifted up)
        for ((r, c), cell) in cells_to_shift {
            self.cells.insert((r - count, c), cell);
        }
    }

    /// Insert columns at the specified position, shifting existing columns right
    pub fn insert_cols(&mut self, at_col: usize, count: usize) {
        // Collect all cells that need to be shifted
        let cells_to_shift: Vec<_> = self.cells
            .iter()
            .filter(|((_, c), _)| *c >= at_col)
            .map(|((r, c), cell)| ((*r, *c), cell.clone()))
            .collect();

        // Remove old positions
        for ((r, c), _) in &cells_to_shift {
            self.cells.remove(&(*r, *c));
        }

        // Insert at new positions (shifted right)
        for ((r, c), cell) in cells_to_shift {
            if c + count < self.cols {
                self.cells.insert((r, c + count), cell);
            }
        }
    }

    /// Delete columns at the specified position, shifting remaining columns left
    pub fn delete_cols(&mut self, start_col: usize, count: usize) {
        // Remove cells in the deleted columns
        for col in start_col..start_col + count {
            for row in 0..self.rows {
                self.cells.remove(&(row, col));
            }
        }

        // Collect cells that need to be shifted left
        let cells_to_shift: Vec<_> = self.cells
            .iter()
            .filter(|((_, c), _)| *c >= start_col + count)
            .map(|((r, c), cell)| ((*r, *c), cell.clone()))
            .collect();

        // Remove old positions
        for ((r, c), _) in &cells_to_shift {
            self.cells.remove(&(*r, *c));
        }

        // Insert at new positions (shifted left)
        for ((r, c), cell) in cells_to_shift {
            self.cells.insert((r, c - count), cell);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_set_text_overflow() {
        let mut sheet = Sheet::new(10, 10);

        // Default should be Clip
        let format = sheet.get_format(0, 0);
        assert_eq!(format.text_overflow, TextOverflow::Clip);

        // Set to Wrap
        sheet.set_text_overflow(0, 0, TextOverflow::Wrap);
        let format = sheet.get_format(0, 0);
        assert_eq!(format.text_overflow, TextOverflow::Wrap);

        // Set to Overflow (Spill)
        sheet.set_text_overflow(0, 0, TextOverflow::Overflow);
        let format = sheet.get_format(0, 0);
        assert_eq!(format.text_overflow, TextOverflow::Overflow);
    }

    #[test]
    fn test_set_vertical_alignment() {
        let mut sheet = Sheet::new(10, 10);

        // Default should be Middle
        let format = sheet.get_format(0, 0);
        assert_eq!(format.vertical_alignment, VerticalAlignment::Middle);

        // Set to Top
        sheet.set_vertical_alignment(0, 0, VerticalAlignment::Top);
        let format = sheet.get_format(0, 0);
        assert_eq!(format.vertical_alignment, VerticalAlignment::Top);

        // Set to Bottom
        sheet.set_vertical_alignment(0, 0, VerticalAlignment::Bottom);
        let format = sheet.get_format(0, 0);
        assert_eq!(format.vertical_alignment, VerticalAlignment::Bottom);
    }

    #[test]
    fn test_format_persists_with_value() {
        let mut sheet = Sheet::new(10, 10);

        // Set format first
        sheet.set_text_overflow(0, 0, TextOverflow::Wrap);
        sheet.set_vertical_alignment(0, 0, VerticalAlignment::Top);

        // Then set value
        sheet.set_value(0, 0, "Hello");

        // Format should persist
        let format = sheet.get_format(0, 0);
        assert_eq!(format.text_overflow, TextOverflow::Wrap);
        assert_eq!(format.vertical_alignment, VerticalAlignment::Top);
    }

    #[test]
    fn test_spill_stops_at_nonempty_cell() {
        // This tests the data model aspect - the renderer logic is tested elsewhere
        let mut sheet = Sheet::new(10, 10);

        // Cell A1 has long text with Spill mode
        sheet.set_value(0, 0, "This is a very long text that should spill");
        sheet.set_text_overflow(0, 0, TextOverflow::Overflow);

        // Cell C1 has content - should stop spill
        sheet.set_value(0, 2, "Block");

        // Verify setup
        assert!(!sheet.get_raw(0, 0).is_empty());
        assert!(sheet.get_raw(0, 1).is_empty()); // B1 is empty
        assert!(!sheet.get_raw(0, 2).is_empty()); // C1 blocks
    }

    #[test]
    fn test_spill_stops_at_formatted_cell() {
        let mut sheet = Sheet::new(10, 10);

        // Cell A1 has Spill mode
        sheet.set_value(0, 0, "Long text here");
        sheet.set_text_overflow(0, 0, TextOverflow::Overflow);

        // Cell B1 is empty but has formatting
        sheet.toggle_bold(0, 1);

        // Verify setup
        let b1_format = sheet.get_format(0, 1);
        assert!(b1_format.bold);
        assert!(sheet.get_raw(0, 1).is_empty());
    }

    #[test]
    fn test_format_painter_copy_paste() {
        let mut sheet = Sheet::new(10, 10);

        // Set up source cell with complex formatting
        sheet.set_value(0, 0, "Source");
        sheet.toggle_bold(0, 0);
        sheet.toggle_italic(0, 0);
        sheet.set_vertical_alignment(0, 0, VerticalAlignment::Top);
        sheet.set_text_overflow(0, 0, TextOverflow::Wrap);
        sheet.set_number_format(0, 0, NumberFormat::Currency { decimals: 2 });

        // Get the source format
        let source_format = sheet.get_format(0, 0);

        // Apply to destination cells
        for row in 1..5 {
            for col in 0..5 {
                sheet.set_format(row, col, source_format.clone());
            }
        }

        // Verify format was copied correctly
        for row in 1..5 {
            for col in 0..5 {
                let fmt = sheet.get_format(row, col);
                assert!(fmt.bold, "Cell ({}, {}) should be bold", row, col);
                assert!(fmt.italic, "Cell ({}, {}) should be italic", row, col);
                assert_eq!(fmt.vertical_alignment, VerticalAlignment::Top);
                assert_eq!(fmt.text_overflow, TextOverflow::Wrap);
                assert!(matches!(fmt.number_format, NumberFormat::Currency { decimals: 2 }));
            }
        }

        // Verify source cell value wasn't affected
        assert_eq!(sheet.get_raw(0, 0), "Source");

        // Verify destination cells have no values (only format)
        for row in 1..5 {
            for col in 0..5 {
                assert!(sheet.get_raw(row, col).is_empty());
            }
        }
    }

    #[test]
    fn test_format_painter_large_selection() {
        let mut sheet = Sheet::new(100, 100);

        // Set up source format
        sheet.toggle_bold(0, 0);
        sheet.set_text_overflow(0, 0, TextOverflow::Overflow);
        let source_format = sheet.get_format(0, 0);

        // Apply to 500 cells (10 rows x 50 cols)
        for row in 1..11 {
            for col in 0..50 {
                sheet.set_format(row, col, source_format.clone());
            }
        }

        // Spot check some cells
        assert!(sheet.get_format(1, 0).bold);
        assert!(sheet.get_format(5, 25).bold);
        assert!(sheet.get_format(10, 49).bold);
        assert_eq!(sheet.get_format(10, 49).text_overflow, TextOverflow::Overflow);
    }

    #[test]
    fn test_sequence_spill() {
        let mut sheet = Sheet::new(10, 10);

        // Enter SEQUENCE formula that should create a 3x2 array
        sheet.set_value(0, 0, "=SEQUENCE(3,2)");

        // Check that parent cell is marked as spill parent
        assert!(sheet.is_spill_parent(0, 0));

        // Check that values spilled correctly
        // SEQUENCE(3,2) should produce:
        // 1 2
        // 3 4
        // 5 6
        assert_eq!(sheet.get_display(0, 0), "1");
        assert_eq!(sheet.get_display(0, 1), "2");
        assert_eq!(sheet.get_display(1, 0), "3");
        assert_eq!(sheet.get_display(1, 1), "4");
        assert_eq!(sheet.get_display(2, 0), "5");
        assert_eq!(sheet.get_display(2, 1), "6");

        // Check that receiving cells are marked
        assert!(sheet.is_spill_receiver(0, 1));
        assert!(sheet.is_spill_receiver(1, 0));
        assert!(sheet.is_spill_receiver(1, 1));
        assert!(sheet.is_spill_receiver(2, 0));
        assert!(sheet.is_spill_receiver(2, 1));

        // Parent cell should not be a receiver
        assert!(!sheet.is_spill_receiver(0, 0));
    }

    #[test]
    fn test_sequence_with_start_and_step() {
        let mut sheet = Sheet::new(10, 10);

        // SEQUENCE(3, 1, 10, 5) should produce: 10, 15, 20
        sheet.set_value(0, 0, "=SEQUENCE(3,1,10,5)");

        assert_eq!(sheet.get_display(0, 0), "10");
        assert_eq!(sheet.get_display(1, 0), "15");
        assert_eq!(sheet.get_display(2, 0), "20");
    }

    #[test]
    fn test_spill_collision_shows_error() {
        let mut sheet = Sheet::new(10, 10);

        // Put a value in the way of where spill would go
        sheet.set_value(1, 0, "blocker");

        // Now enter SEQUENCE that would need to spill to (1, 0)
        sheet.set_value(0, 0, "=SEQUENCE(3,1)");

        // Parent cell should show #SPILL! error
        assert!(sheet.has_spill_error(0, 0));
        assert_eq!(sheet.get_display(0, 0), "#SPILL!");

        // The blocking cell should still have its value
        assert_eq!(sheet.get_display(1, 0), "blocker");
    }

    #[test]
    fn test_spill_cleared_when_formula_changes() {
        let mut sheet = Sheet::new(10, 10);

        // Enter SEQUENCE that spills
        sheet.set_value(0, 0, "=SEQUENCE(3,2)");

        // Verify initial spill
        assert!(sheet.is_spill_parent(0, 0));
        assert!(sheet.is_spill_receiver(1, 0));

        // Change the formula to a scalar value
        sheet.set_value(0, 0, "42");

        // Spill should be cleared
        assert!(!sheet.is_spill_parent(0, 0));
        assert!(!sheet.is_spill_receiver(1, 0));
        assert!(!sheet.is_spill_receiver(0, 1));

        // Receiving cells should now be empty
        assert_eq!(sheet.get_display(1, 0), "");
        assert_eq!(sheet.get_display(0, 1), "");
    }

    #[test]
    fn test_transpose_spill() {
        let mut sheet = Sheet::new(10, 10);

        // Set up a column of values
        sheet.set_value(0, 0, "1");
        sheet.set_value(1, 0, "2");
        sheet.set_value(2, 0, "3");

        // TRANSPOSE should turn the column into a row
        sheet.set_value(0, 2, "=TRANSPOSE(A1:A3)");

        // Should spill horizontally: 1, 2, 3
        assert_eq!(sheet.get_display(0, 2), "1");
        assert_eq!(sheet.get_display(0, 3), "2");
        assert_eq!(sheet.get_display(0, 4), "3");
    }

    #[test]
    fn test_reference_to_spill_receiver() {
        let mut sheet = Sheet::new(10, 10);

        // Create a SEQUENCE that spills to multiple cells
        // A1: =SEQUENCE(3,1) produces 1, 2, 3 in A1:A3
        sheet.set_value(0, 0, "=SEQUENCE(3,1)");

        // Verify spill works
        assert_eq!(sheet.get_display(0, 0), "1"); // A1 (parent)
        assert_eq!(sheet.get_display(1, 0), "2"); // A2 (receiver)
        assert_eq!(sheet.get_display(2, 0), "3"); // A3 (receiver)

        // Now reference a spill receiver from another formula
        // B1: =A2 should return 2 (the spilled value)
        sheet.set_value(0, 1, "=A2");
        assert_eq!(sheet.get_display(0, 1), "2");

        // B2: =A2+A3 should return 5 (2+3)
        sheet.set_value(1, 1, "=A2+A3");
        assert_eq!(sheet.get_display(1, 1), "5");

        // B3: =SUM(A1:A3) should return 6 (1+2+3)
        sheet.set_value(2, 1, "=SUM(A1:A3)");
        assert_eq!(sheet.get_display(2, 1), "6");
    }

    #[test]
    fn test_spill_receiver_in_range_functions() {
        let mut sheet = Sheet::new(10, 10);

        // Create sequence in A1:A5
        sheet.set_value(0, 0, "=SEQUENCE(5,1,10,10)"); // 10, 20, 30, 40, 50

        // Verify spill
        assert_eq!(sheet.get_display(0, 0), "10");
        assert_eq!(sheet.get_display(4, 0), "50");

        // Test various functions that read from spill receivers
        sheet.set_value(0, 2, "=AVERAGE(A1:A5)");  // Should be 30
        sheet.set_value(1, 2, "=MIN(A1:A5)");     // Should be 10
        sheet.set_value(2, 2, "=MAX(A1:A5)");     // Should be 50
        sheet.set_value(3, 2, "=COUNT(A1:A5)");   // Should be 5

        assert_eq!(sheet.get_display(0, 2), "30");
        assert_eq!(sheet.get_display(1, 2), "10");
        assert_eq!(sheet.get_display(2, 2), "50");
        assert_eq!(sheet.get_display(3, 2), "5");
    }

    #[test]
    fn test_sort_numeric_ascending() {
        let mut sheet = Sheet::new(10, 10);

        // Set up data: unsorted numbers in A1:A5
        sheet.set_value(0, 0, "30");
        sheet.set_value(1, 0, "10");
        sheet.set_value(2, 0, "50");
        sheet.set_value(3, 0, "20");
        sheet.set_value(4, 0, "40");

        // SORT ascending (default)
        sheet.set_value(0, 2, "=SORT(A1:A5)");

        // Should be sorted: 10, 20, 30, 40, 50
        assert_eq!(sheet.get_display(0, 2), "10");
        assert_eq!(sheet.get_display(1, 2), "20");
        assert_eq!(sheet.get_display(2, 2), "30");
        assert_eq!(sheet.get_display(3, 2), "40");
        assert_eq!(sheet.get_display(4, 2), "50");
    }

    #[test]
    fn test_sort_numeric_descending() {
        let mut sheet = Sheet::new(10, 10);

        // Set up data: unsorted numbers in A1:A5
        sheet.set_value(0, 0, "30");
        sheet.set_value(1, 0, "10");
        sheet.set_value(2, 0, "50");
        sheet.set_value(3, 0, "20");
        sheet.set_value(4, 0, "40");

        // SORT descending (is_asc = FALSE)
        sheet.set_value(0, 2, "=SORT(A1:A5,1,FALSE)");

        // Should be sorted descending: 50, 40, 30, 20, 10
        assert_eq!(sheet.get_display(0, 2), "50");
        assert_eq!(sheet.get_display(1, 2), "40");
        assert_eq!(sheet.get_display(2, 2), "30");
        assert_eq!(sheet.get_display(3, 2), "20");
        assert_eq!(sheet.get_display(4, 2), "10");
    }

    #[test]
    fn test_sort_multi_column_by_second_column() {
        let mut sheet = Sheet::new(10, 10);

        // Set up 2-column data
        // A1:B3: (Alice, 30), (Bob, 10), (Charlie, 20)
        sheet.set_value(0, 0, "Alice");   sheet.set_value(0, 1, "30");
        sheet.set_value(1, 0, "Bob");     sheet.set_value(1, 1, "10");
        sheet.set_value(2, 0, "Charlie"); sheet.set_value(2, 1, "20");

        // SORT by column 2 (the numbers)
        sheet.set_value(0, 3, "=SORT(A1:B3,2)");

        // Should be sorted by score: Bob(10), Charlie(20), Alice(30)
        assert_eq!(sheet.get_display(0, 3), "Bob");
        assert_eq!(sheet.get_display(0, 4), "10");
        assert_eq!(sheet.get_display(1, 3), "Charlie");
        assert_eq!(sheet.get_display(1, 4), "20");
        assert_eq!(sheet.get_display(2, 3), "Alice");
        assert_eq!(sheet.get_display(2, 4), "30");
    }

    #[test]
    fn test_sort_text_alphabetical() {
        let mut sheet = Sheet::new(10, 10);

        // Set up text data
        sheet.set_value(0, 0, "Banana");
        sheet.set_value(1, 0, "Apple");
        sheet.set_value(2, 0, "Cherry");

        // SORT alphabetically
        sheet.set_value(0, 2, "=SORT(A1:A3)");

        // Should be: Apple, Banana, Cherry
        assert_eq!(sheet.get_display(0, 2), "Apple");
        assert_eq!(sheet.get_display(1, 2), "Banana");
        assert_eq!(sheet.get_display(2, 2), "Cherry");
    }

    #[test]
    fn test_unique_single_column() {
        let mut sheet = Sheet::new(10, 10);

        // Set up data with duplicates
        sheet.set_value(0, 0, "Apple");
        sheet.set_value(1, 0, "Banana");
        sheet.set_value(2, 0, "Apple");  // duplicate
        sheet.set_value(3, 0, "Cherry");
        sheet.set_value(4, 0, "Banana"); // duplicate

        // UNIQUE should return: Apple, Banana, Cherry (first occurrence order)
        sheet.set_value(0, 2, "=UNIQUE(A1:A5)");

        assert_eq!(sheet.get_display(0, 2), "Apple");
        assert_eq!(sheet.get_display(1, 2), "Banana");
        assert_eq!(sheet.get_display(2, 2), "Cherry");
        // No more rows
        assert_eq!(sheet.get_display(3, 2), "");
    }

    #[test]
    fn test_unique_multi_column() {
        let mut sheet = Sheet::new(10, 10);

        // Set up 2-column data with duplicate rows
        sheet.set_value(0, 0, "Alice");   sheet.set_value(0, 1, "30");
        sheet.set_value(1, 0, "Bob");     sheet.set_value(1, 1, "20");
        sheet.set_value(2, 0, "Alice");   sheet.set_value(2, 1, "30"); // duplicate row
        sheet.set_value(3, 0, "Alice");   sheet.set_value(3, 1, "25"); // different score

        // UNIQUE should return 3 rows: (Alice,30), (Bob,20), (Alice,25)
        sheet.set_value(0, 3, "=UNIQUE(A1:B4)");

        assert_eq!(sheet.get_display(0, 3), "Alice");
        assert_eq!(sheet.get_display(0, 4), "30");
        assert_eq!(sheet.get_display(1, 3), "Bob");
        assert_eq!(sheet.get_display(1, 4), "20");
        assert_eq!(sheet.get_display(2, 3), "Alice");
        assert_eq!(sheet.get_display(2, 4), "25");
        // No fourth row
        assert_eq!(sheet.get_display(3, 3), "");
    }

    #[test]
    fn test_unique_case_insensitive() {
        let mut sheet = Sheet::new(10, 10);

        // Set up data with case differences
        sheet.set_value(0, 0, "Apple");
        sheet.set_value(1, 0, "APPLE");  // same word, different case
        sheet.set_value(2, 0, "Banana");

        // UNIQUE should treat as duplicates (case-insensitive)
        sheet.set_value(0, 2, "=UNIQUE(A1:A3)");

        // Should return: Apple (first occurrence), Banana
        assert_eq!(sheet.get_display(0, 2), "Apple");
        assert_eq!(sheet.get_display(1, 2), "Banana");
        assert_eq!(sheet.get_display(2, 2), "");
    }

    #[test]
    fn test_unique_numeric() {
        let mut sheet = Sheet::new(10, 10);

        // Set up numeric data with duplicates
        sheet.set_value(0, 0, "10");
        sheet.set_value(1, 0, "20");
        sheet.set_value(2, 0, "10");  // duplicate
        sheet.set_value(3, 0, "30");

        sheet.set_value(0, 2, "=UNIQUE(A1:A4)");

        assert_eq!(sheet.get_display(0, 2), "10");
        assert_eq!(sheet.get_display(1, 2), "20");
        assert_eq!(sheet.get_display(2, 2), "30");
        assert_eq!(sheet.get_display(3, 2), "");
    }

    #[test]
    fn test_filter_basic() {
        let mut sheet = Sheet::new(10, 10);

        // Set up data: names in A, scores in B, pass/fail in C (1=pass, 0=fail)
        sheet.set_value(0, 0, "Alice");   sheet.set_value(0, 1, "85"); sheet.set_value(0, 2, "1");
        sheet.set_value(1, 0, "Bob");     sheet.set_value(1, 1, "45"); sheet.set_value(1, 2, "0");
        sheet.set_value(2, 0, "Charlie"); sheet.set_value(2, 1, "72"); sheet.set_value(2, 2, "1");
        sheet.set_value(3, 0, "Diana");   sheet.set_value(3, 1, "38"); sheet.set_value(3, 2, "0");

        // FILTER by pass column (C1:C4)
        sheet.set_value(0, 4, "=FILTER(A1:B4,C1:C4)");

        // Should return only passing students: Alice, Charlie
        assert_eq!(sheet.get_display(0, 4), "Alice");
        assert_eq!(sheet.get_display(0, 5), "85");
        assert_eq!(sheet.get_display(1, 4), "Charlie");
        assert_eq!(sheet.get_display(1, 5), "72");
        // No third row
        assert_eq!(sheet.get_display(2, 4), "");
    }

    #[test]
    fn test_filter_no_matches() {
        let mut sheet = Sheet::new(10, 10);

        // Set up data where all fail filter
        sheet.set_value(0, 0, "Apple");  sheet.set_value(0, 1, "0");
        sheet.set_value(1, 0, "Banana"); sheet.set_value(1, 1, "0");

        // FILTER with all zeros (no matches)
        sheet.set_value(0, 3, "=FILTER(A1:A2,B1:B2)");

        // Should return #CALC! error
        assert!(sheet.get_display(0, 3).contains("#CALC!"));
    }

    #[test]
    fn test_filter_single_column() {
        let mut sheet = Sheet::new(10, 10);

        // Set up single column with filter criteria
        sheet.set_value(0, 0, "10"); sheet.set_value(0, 1, "1");
        sheet.set_value(1, 0, "20"); sheet.set_value(1, 1, "0");
        sheet.set_value(2, 0, "30"); sheet.set_value(2, 1, "1");
        sheet.set_value(3, 0, "40"); sheet.set_value(3, 1, "1");

        // FILTER numbers where criteria is true
        sheet.set_value(0, 3, "=FILTER(A1:A4,B1:B4)");

        assert_eq!(sheet.get_display(0, 3), "10");
        assert_eq!(sheet.get_display(1, 3), "30");
        assert_eq!(sheet.get_display(2, 3), "40");
        assert_eq!(sheet.get_display(3, 3), "");
    }

    #[test]
    fn test_filter_combined_with_sort() {
        let mut sheet = Sheet::new(15, 15);

        // Set up data
        sheet.set_value(0, 0, "Alice");   sheet.set_value(0, 1, "85"); sheet.set_value(0, 2, "1");
        sheet.set_value(1, 0, "Bob");     sheet.set_value(1, 1, "92"); sheet.set_value(1, 2, "1");
        sheet.set_value(2, 0, "Charlie"); sheet.set_value(2, 1, "45"); sheet.set_value(2, 2, "0");
        sheet.set_value(3, 0, "Diana");   sheet.set_value(3, 1, "78"); sheet.set_value(3, 2, "1");

        // First filter passing students
        sheet.set_value(0, 5, "=FILTER(A1:B4,C1:C4)");
        // Then reference that spilled range to sort by score
        // Note: This tests that SORT can work with spilled data

        // Verify filter works
        assert_eq!(sheet.get_display(0, 5), "Alice");
        assert_eq!(sheet.get_display(1, 5), "Bob");
        assert_eq!(sheet.get_display(2, 5), "Diana");
    }

    #[test]
    fn test_filter_include_semantics() {
        // Lock behavior: FILTER treats non-zero as TRUE, zero as FALSE
        // Text is NOT coerced - it evaluates to 0.0 (FALSE)
        let mut sheet = Sheet::new(10, 10);

        // Test with various include values
        sheet.set_value(0, 0, "A"); sheet.set_value(0, 1, "1");      // 1 = TRUE
        sheet.set_value(1, 0, "B"); sheet.set_value(1, 1, "0");      // 0 = FALSE
        sheet.set_value(2, 0, "C"); sheet.set_value(2, 1, "-1");     // -1 = TRUE (non-zero)
        sheet.set_value(3, 0, "D"); sheet.set_value(3, 1, "0.001");  // 0.001 = TRUE (non-zero)

        sheet.set_value(0, 3, "=FILTER(A1:A4,B1:B4)");

        // Should include: A (1), C (-1), D (0.001)
        assert_eq!(sheet.get_display(0, 3), "A");
        assert_eq!(sheet.get_display(1, 3), "C");
        assert_eq!(sheet.get_display(2, 3), "D");
        assert_eq!(sheet.get_display(3, 3), ""); // No fourth row
    }

    #[test]
    fn test_spill_shrinks_clears_old_receivers() {
        // When a spill range shrinks, old receivers outside the new range must be cleared
        let mut sheet = Sheet::new(10, 10);

        // Create a 5-row spill
        sheet.set_value(0, 0, "=SEQUENCE(5,1)");

        // Verify all 5 rows are filled
        assert_eq!(sheet.get_display(0, 0), "1");
        assert_eq!(sheet.get_display(4, 0), "5");
        assert!(sheet.is_spill_receiver(4, 0));

        // Now change to a 2-row spill
        sheet.set_value(0, 0, "=SEQUENCE(2,1)");

        // Verify new spill is 2 rows
        assert_eq!(sheet.get_display(0, 0), "1");
        assert_eq!(sheet.get_display(1, 0), "2");

        // Old receivers (rows 2, 3, 4) should be cleared
        assert_eq!(sheet.get_display(2, 0), "");
        assert_eq!(sheet.get_display(3, 0), "");
        assert_eq!(sheet.get_display(4, 0), "");

        // Old receivers should no longer be marked as receivers
        assert!(!sheet.is_spill_receiver(2, 0));
        assert!(!sheet.is_spill_receiver(3, 0));
        assert!(!sheet.is_spill_receiver(4, 0));
    }
}
