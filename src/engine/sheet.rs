use std::collections::HashMap;

use serde::{Deserialize, Serialize};

use super::cell::{Alignment, Cell, CellFormat, CellValue, NumberFormat};
use super::formula::eval::{self, CellLookup, EvalResult};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    cells: HashMap<(usize, usize), Cell>,
    pub rows: usize,
    pub cols: usize,
}

impl CellLookup for Sheet {
    fn get_value(&self, row: usize, col: usize) -> f64 {
        self.cells
            .get(&(row, col))
            .map(|c| self.evaluate_cell_value(&c.value))
            .unwrap_or(0.0)
    }
}

impl Sheet {
    pub fn new(rows: usize, cols: usize) -> Self {
        Self {
            name: String::from("Sheet1"),
            cells: HashMap::new(),
            rows,
            cols,
        }
    }

    pub fn set_value(&mut self, row: usize, col: usize, value: &str) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.set(value);
    }

    pub fn get_display(&self, row: usize, col: usize) -> String {
        match self.cells.get(&(row, col)) {
            Some(cell) => self.display_cell_value(&cell.value),
            None => String::new(),
        }
    }

    /// Get display value with number formatting applied
    pub fn get_formatted_display(&self, row: usize, col: usize) -> String {
        match self.cells.get(&(row, col)) {
            Some(cell) => {
                // Get the numeric value for formatting
                let num_value = match &cell.value {
                    CellValue::Number(n) => Some(*n),
                    CellValue::Formula { ast: Some(ast), .. } => {
                        match eval::evaluate(ast, self) {
                            EvalResult::Number(n) => Some(n),
                            EvalResult::Error(_) => None,
                        }
                    }
                    _ => None,
                };

                if let Some(n) = num_value {
                    // Apply number formatting
                    CellValue::format_number(n, &cell.format.number_format)
                } else {
                    self.display_cell_value(&cell.value)
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
                    EvalResult::Error(_) => 0.0,
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
            CellValue::Formula { ast: Some(ast), source } => {
                match eval::evaluate(ast, self) {
                    EvalResult::Number(n) => {
                        if n.fract() == 0.0 {
                            format!("{}", n as i64)
                        } else {
                            format!("{:.2}", n)
                        }
                    }
                    EvalResult::Error(_) => format!("#ERR"),
                }
            }
            CellValue::Formula { ast: None, source: _ } => format!("#ERR"),
        }
    }

    pub fn get_format(&self, row: usize, col: usize) -> CellFormat {
        self.cells
            .get(&(row, col))
            .map(|c| c.format.clone())
            .unwrap_or_default()
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

    pub fn set_alignment(&mut self, row: usize, col: usize, alignment: Alignment) {
        let cell = self.cells.entry((row, col)).or_insert_with(Cell::new);
        cell.format.alignment = alignment;
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
