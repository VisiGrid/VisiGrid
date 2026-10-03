//! Workbook-level pivot operations: placement, atomic apply, scoped undo
//! state and stale tracking. A child module of `workbook` so it can reach the
//! sheets and the batch/recalc machinery directly.
//!
//! A pivot action never snapshots the workbook. It produces a
//! [`PivotCommit`]: the pivot object and the values of the cells it touches,
//! before and after. Applying `after` performs the action; applying `before`
//! undoes it. Both are bounded by the output budget in [`crate::pivot`].

use crate::cell::NumberFormat;
use crate::cell_id::CellId;
use crate::formula::eval::Value;
use crate::pivot::{self, PivotError, PivotOutput, PivotSnapshot, PivotTable, RefreshRecord};
use crate::sheet::{SheetId, NUM_COLS, NUM_ROWS};

use super::Workbook;

fn now_secs() -> i64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0)
}

/// The pivot object and the output cells it covers, at one moment.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotState {
    /// Sheet showing the output.
    pub sheet_id: SheetId,
    pub pivot_id: u64,
    /// `None` = no such pivot (before a create, after a delete).
    pub table: Option<PivotTable>,
    /// Cells over the union of the old and new regions: value and number
    /// format. `Value::Empty` clears the value but keeps the cell's other
    /// formatting. Other style attributes are never touched by a pivot.
    pub cells: Vec<PivotCell>,
}

/// One cell of a pivot state.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCell {
    pub row: u32,
    pub col: u32,
    pub value: Value,
    pub number_format: NumberFormat,
}

impl PivotState {
    /// Rough retained size, for the history budget.
    pub fn approx_bytes(&self) -> usize {
        let cell_bytes: usize = self
            .cells
            .iter()
            .map(|c| {
                64 + match &c.value {
                    Value::Text(s) | Value::Error(s) => s.len(),
                    _ => 0,
                }
            })
            .sum();
        256 + cell_bytes
    }
}

/// One pivot action as its before/after states.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCommit {
    pub before: PivotState,
    pub after: PivotState,
}

impl PivotCommit {
    pub fn approx_bytes(&self) -> usize {
        self.before.approx_bytes() + self.after.approx_bytes()
    }
}

/// A pivot as saved in a file. Files identify sheets by position, and loading
/// assigns fresh sheet ids, so the source sheet is stored by index.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct SavedPivot {
    /// Index of the source sheet in the saved workbook.
    pub source_sheet: usize,
    pub table: PivotTable,
}

/// Why a pivot action cannot be performed. The workbook is unchanged.
#[derive(Debug, Clone, PartialEq)]
pub enum PivotOpError {
    Pivot(PivotError),
    NotFound(u64),
    SheetMissing,
    SourceSheetMissing,
    TableSource(String),
    /// The output would leave the grid.
    OffGrid { rows: usize, cols: usize },
    /// A cell in the new output area is in use.
    Blocked { row: usize, col: usize, reason: String },
}

impl std::fmt::Display for PivotOpError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            PivotOpError::TableSource(message) => write!(f, "{message}"),
            PivotOpError::Pivot(e) => write!(f, "{e}"),
            PivotOpError::NotFound(id) => write!(f, "Pivot table {id} no longer exists."),
            PivotOpError::SheetMissing => write!(f, "The pivot's sheet no longer exists."),
            PivotOpError::SourceSheetMissing => write!(f, "The pivot's source sheet no longer exists."),
            PivotOpError::OffGrid { rows, cols } => {
                write!(f, "The pivot output ({rows} × {cols}) would run past the edge of the sheet.")
            }
            PivotOpError::Blocked { row, col, reason } => {
                write!(f, "Cannot place the pivot: {}{} is {reason}.", crate::formula::parser::col_to_letters(*col), row + 1)
            }
        }
    }
}

impl From<PivotError> for PivotOpError {
    fn from(e: PivotError) -> Self {
        PivotOpError::Pivot(e)
    }
}

impl Workbook {
    /// A pivot id not used by any sheet.
    pub fn next_pivot_id(&self) -> u64 {
        self.sheets.iter().flat_map(|s| s.pivots.iter().map(|p| p.id)).max().map_or(1, |m| m + 1)
    }

    /// A display name not used by any pivot: "PivotTable1", "PivotTable2", …
    pub fn next_pivot_name(&self) -> String {
        let taken: Vec<String> = self
            .sheets
            .iter()
            .flat_map(|s| s.pivots.iter().map(|p| p.name.to_lowercase()))
            .collect();
        (1..).map(|n| format!("PivotTable{n}")).find(|n| !taken.contains(&n.to_lowercase())).unwrap()
    }

    /// (sheet index, pivot) for an id.
    pub fn find_pivot(&self, id: u64) -> Option<(usize, &PivotTable)> {
        self.sheets
            .iter()
            .enumerate()
            .find_map(|(i, s)| s.pivots.iter().find(|p| p.id == id).map(|p| (i, p)))
    }

    /// Every pivot in the workbook, with the index of the sheet showing it.
    pub fn pivots(&self) -> Vec<(usize, &PivotTable)> {
        self.sheets.iter().enumerate().flat_map(|(i, s)| s.pivots.iter().map(move |p| (i, p))).collect()
    }

    /// Select a Table by stable identity. Bounds are a cache, resolved afresh
    /// for each field-list or refresh operation.
    pub fn table_pivot_source(&self, id: crate::table::TableId) -> Result<pivot::PivotSource, PivotOpError> {
        let (sheet_id, t) = self.table(id).ok_or_else(|| PivotOpError::TableSource(
            "The source Table no longer exists. The last pivot result has been kept.".into()))?;
        Ok(pivot::PivotSource { table_id: Some(id), sheet_id,
            start_row: t.range.start_row as u32, start_col: t.range.start_col as u32,
            end_row: t.range.end_row as u32, end_col: t.range.end_col as u32 })
    }

    /// Resolve stable field IDs, never names or saved offsets, on refresh.
    /// Missing fields refuse the whole operation without changing output.
    pub fn resolve_pivot_source(&self, source: &mut pivot::PivotSource, definition: &mut pivot::PivotDefinition) -> Result<(), PivotOpError> {
        let Some(id) = source.table_id else { return Ok(()) };
        let current = self.table_pivot_source(id)?;
        let (_, table) = self.table(id).unwrap();
        let mut resolved = definition.clone();
        for field in resolved.rows.iter_mut().chain(resolved.column.iter_mut()).chain(resolved.values.iter_mut().map(|v| &mut v.field)) {
            let found = field.column_id.and_then(|id| table.columns.iter().enumerate().find(|(_, c)| c.id == id));
            let Some((offset, column)) = found else {
                return Err(PivotOpError::TableSource(format!("The source column '{}' is missing from {}. Edit the pivot fields to replace it. The last result has been kept.", field.header, table.name)));
            };
            field.offset = offset as u32;
            field.header = column.name.clone();
        }
        *source = current;
        *definition = resolved;
        Ok(())
    }

    /// Read a resolved definition and canonical source snapshot together.
    /// Callers must aggregate and commit the returned definition, since Table
    /// columns may have moved or been renamed since the previous refresh.
    pub fn pivot_snapshot(&self, table: &PivotTable) -> Result<(PivotTable, PivotSnapshot, u64), PivotOpError> {
        let mut resolved = table.clone();
        self.resolve_pivot_source(&mut resolved.source, &mut resolved.definition)?;
        let sheet = self.sheet_by_id(resolved.source.sheet_id).ok_or(PivotOpError::SourceSheetMissing)?;
        let snapshot = pivot::capture_snapshot(sheet, &resolved.source, &resolved.definition);
        let generation = self.pivot_source_generation(&resolved).ok_or(PivotOpError::SourceSheetMissing)?;
        Ok((resolved, snapshot, generation))
    }

    /// Rows appended directly below a pivot's source: the last row of the
    /// contiguous block of rows under `end_row` that hold data in any of the
    /// source's columns. `None` if the next row is empty. The caller offers to
    /// extend the source; it is never extended silently (a totals row under the
    /// data would otherwise be summed).
    pub fn pivot_source_growth(&self, table: &PivotTable) -> Option<u32> {
        if table.source.table_id.is_some() { return None; }
        let sheet = self.sheet_by_id(table.source.sheet_id)?;
        let (c0, c1) = (table.source.start_col as usize, table.source.end_col as usize);
        let row_has_data = |r: usize| (c0..=c1).any(|c| !sheet.get_raw(r, c).is_empty());
        let mut r = table.source.end_row as usize + 1;
        let limit = sheet.rows.min(NUM_ROWS);
        let mut last = None;
        while r < limit && row_has_data(r) {
            last = Some(r as u32);
            r += 1;
        }
        last
    }

    /// The source sheet's current edit generation (for stale checks before apply).
    pub fn pivot_source_generation(&self, table: &PivotTable) -> Option<u64> {
        let generation = self.sheet_by_id(table.source.sheet_id)?.edit_generation();
        let Some(id) = table.source.table_id else { return Some(generation) };
        let (_, source) = self.table(id)?;
        // Metadata-only resize/rename must invalidate a background job too.
        use std::hash::{Hash, Hasher};
        let mut hash = std::collections::hash_map::DefaultHasher::new();
        generation.hash(&mut hash);
        id.hash(&mut hash);
        source.name.hash(&mut hash);
        (source.range.start_row, source.range.start_col, source.range.end_row, source.range.end_col).hash(&mut hash);
        for column in &source.columns { column.id.hash(&mut hash); column.name.hash(&mut hash); }
        Some(hash.finish())
    }

    /// Prepare placing `output` for `table` on the sheet `sheet_id`. `table`
    /// carries the new definition/source/anchor; any pivot with the same id is
    /// the one being replaced. Validates the whole new area first; nothing is
    /// changed.
    pub fn prepare_pivot_commit(
        &self,
        sheet_id: SheetId,
        mut table: PivotTable,
        output: &PivotOutput,
        source_generation: u64,
        refreshed_at: i64,
    ) -> Result<PivotCommit, PivotOpError> {
        let sheet = self.sheet_by_id(sheet_id).ok_or(PivotOpError::SheetMissing)?;
        let old = sheet.pivots.iter().find(|p| p.id == table.id).cloned();
        let (h, w) = (output.height(), output.width());
        let (r0, c0) = (table.anchor_row as usize, table.anchor_col as usize);
        if h == 0 || w == 0 || r0 + h > NUM_ROWS.min(sheet.rows) || c0 + w > NUM_COLS.min(sheet.cols) {
            return Err(PivotOpError::OffGrid { rows: h, cols: w });
        }
        let (r1, c1) = (r0 + h - 1, c0 + w - 1);

        // Every cell of the new area must be free: owned by this pivot, or
        // empty, unmerged, not a spill, not another pivot's.
        if let Some(t) = sheet.tables().iter().find(|t| t.full_range().intersects(crate::table::TableRange {
            start_row: r0, start_col: c0, end_row: r1, end_col: c1,
        })) {
            return Err(PivotOpError::Blocked { row: r0.max(t.range.start_row), col: c0.max(t.range.start_col), reason: format!("part of table {}", t.name) });
        }
        for other in sheet.pivots.iter().filter(|p| p.id != table.id) {
            if other.intersects(r0, c0, r1, c1) {
                let (ar, ac, _, _) = other.region().unwrap();
                return Err(PivotOpError::Blocked {
                    row: ar.max(r0),
                    col: ac.max(c0),
                    reason: format!("part of {}", other.name),
                });
            }
        }
        if let Some(m) = sheet
            .merged_regions
            .iter()
            .find(|m| m.start.0 <= r1 && r0 <= m.end.0 && m.start.1 <= c1 && c0 <= m.end.1)
        {
            return Err(PivotOpError::Blocked { row: m.start.0, col: m.start.1, reason: "a merged cell".into() });
        }
        let in_old = |r: usize, c: usize| old.as_ref().is_some_and(|o| o.contains(r, c));
        // Scan only occupied cells, not the whole rectangle.
        let mut blockers: Vec<(usize, usize)> = sheet
            .cells_in_range(r0, r1, c0, c1)
            .into_iter()
            .filter(|&(r, c)| !in_old(r, c) && !sheet.get_raw(r, c).is_empty())
            .collect();
        blockers.sort_unstable();
        if let Some(&(row, col)) = blockers.first() {
            return Err(PivotOpError::Blocked { row, col, reason: "not empty".into() });
        }
        for r in r0..=r1 {
            for c in c0..=c1 {
                if !in_old(r, c) && sheet.is_spill_receiver(r, c) {
                    return Err(PivotOpError::Blocked { row: r, col: c, reason: "part of a spilled array".into() });
                }
            }
        }

        // Union of old and new regions, row-major.
        let mut coords: Vec<(usize, usize)> = Vec::with_capacity(h * w);
        for r in r0..=r1 {
            for c in c0..=c1 {
                coords.push((r, c));
            }
        }
        if let Some((or0, oc0, or1, oc1)) = old.as_ref().and_then(|o| o.region()) {
            for r in or0..=or1 {
                for c in oc0..=oc1 {
                    if !(r >= r0 && r <= r1 && c >= c0 && c <= c1) {
                        coords.push((r, c));
                    }
                }
            }
        }

        let current_format = |r: usize, c: usize| sheet.get_format(r, c).number_format.clone();
        let before_cells: Vec<PivotCell> = coords
            .iter()
            .map(|&(r, c)| PivotCell {
                row: r as u32,
                col: c as u32,
                value: if in_old(r, c) { owned_value(sheet, r, c) } else { Value::Empty },
                number_format: current_format(r, c),
            })
            .collect();
        let after_cells: Vec<PivotCell> = coords
            .iter()
            .map(|&(r, c)| {
                let inside = r >= r0 && r <= r1 && c >= c0 && c <= c1;
                let value = if inside { output.cells[r - r0][c - c0].clone() } else { Value::Empty };
                // Value cells take their field's format; everything else keeps
                // whatever the cell has, so user formatting survives refresh.
                let field_format = inside
                    .then(|| r - r0 >= output.header_rows)
                    .filter(|&is_data| is_data)
                    .and_then(|_| output.value_columns.get(c - c0).copied().flatten())
                    .and_then(|vi| table.definition.values.get(vi))
                    .and_then(|v| v.number_format.clone());
                PivotCell {
                    row: r as u32,
                    col: c as u32,
                    value,
                    number_format: field_format.unwrap_or_else(|| current_format(r, c)),
                }
            })
            .collect();

        table.extent = Some((h as u32, w as u32));
        table.stale = false;
        table.source_generation = Some(source_generation);
        table.last_refresh = Some(RefreshRecord { source_rows: output.source_rows as u64, refreshed_at });

        Ok(PivotCommit {
            before: PivotState { sheet_id, pivot_id: table.id, table: old, cells: before_cells },
            after: PivotState { sheet_id, pivot_id: table.id, table: Some(table), cells: after_cells },
        })
    }

    /// Create a pivot on a new sheet without an undo stack (headless hosts).
    /// Same output and styling as the desktop's create; the new sheet is
    /// named "Pivot" (or "Pivot (2)", …) and does not become active. On any
    /// failure the workbook is unchanged. Returns the pivot id and the new
    /// sheet's index.
    pub fn create_pivot(&mut self, source: pivot::PivotSource, definition: pivot::PivotDefinition) -> Result<(u64, usize), String> {
        if self.sheets.iter().any(|s| s.table_view_spec().is_some()) {
            let mut candidate = self.clone();
            let result = candidate.create_pivot_inner(source, definition)?;
            self.restore_snapshot_monotonic(&candidate);
            return Ok(result);
        }
        self.create_pivot_inner(source, definition)
    }

    fn create_pivot_inner(&mut self, source: pivot::PivotSource, definition: pivot::PivotDefinition) -> Result<(u64, usize), String> {
        let table = PivotTable {
            id: self.next_pivot_id(),
            name: self.next_pivot_name(),
            source,
            definition,
            anchor_row: 0,
            anchor_col: 0,
            extent: None,
            last_refresh: None,
            stale: false,
            source_generation: None,
        };
        let (mut table, snapshot, generation) = self.pivot_snapshot(&table).map_err(|e| e.to_string())?;
        let output = pivot::aggregate(&table.definition, &snapshot).map_err(|e| e.to_string())?;
        pivot::format_new_pivot_values(&mut table.definition, &output);

        let mut name = "Pivot".to_string();
        let mut n = 2;
        while self.sheet_name_exists(&name) {
            name = format!("Pivot ({n})");
            n += 1;
        }
        let idx = self.add_sheet_named(&name).unwrap_or_else(|| self.add_sheet());
        let sheet_id = self.sheets[idx].id;
        pivot::style_new_pivot(&mut self.sheets[idx], &table, &output);
        let commit = match self.prepare_pivot_commit(sheet_id, table.clone(), &output, generation, now_secs()) {
            Ok(c) => c,
            Err(e) => {
                self.take_sheet(idx);
                return Err(e.to_string());
            }
        };
        self.apply_pivot_state(&commit.after).map_err(|e| e.to_string())?;
        self.bump_revision_for_structure();
        Ok((table.id, idx))
    }

    /// Refresh one pivot in place without an undo stack (headless hosts).
    /// Returns the new output size (rows, cols).
    pub fn refresh_pivot(&mut self, pivot_id: u64) -> Result<(usize, usize), String> {
        let (idx, table) = self.find_pivot(pivot_id).ok_or_else(|| PivotOpError::NotFound(pivot_id).to_string())?;
        let table = table.clone();
        let sheet_id = self.sheets[idx].id;
        let (table, snapshot, generation) = self.pivot_snapshot(&table).map_err(|e| e.to_string())?;
        let output = pivot::aggregate(&table.definition, &snapshot).map_err(|e| e.to_string())?;
        let commit = self
            .prepare_pivot_commit(sheet_id, table, &output, generation, now_secs())
            .map_err(|e| e.to_string())?;
        self.apply_pivot_state(&commit.after).map_err(|e| e.to_string())?;
        Ok((output.height(), output.width()))
    }

    /// Find a pivot by name (case-insensitive) or by numeric id.
    pub fn find_pivot_by_name(&self, name: &str) -> Option<(usize, &PivotTable)> {
        let want = name.trim();
        self.pivots()
            .into_iter()
            .find(|(_, t)| t.name.eq_ignore_ascii_case(want))
            .or_else(|| want.parse::<u64>().ok().and_then(|id| self.find_pivot(id)))
    }

    /// Prepare deleting a pivot and clearing its output.
    pub fn prepare_pivot_delete(&self, pivot_id: u64) -> Result<PivotCommit, PivotOpError> {
        let (idx, table) = self.find_pivot(pivot_id).ok_or(PivotOpError::NotFound(pivot_id))?;
        let sheet = &self.sheets[idx];
        let mut before_cells = Vec::new();
        let mut after_cells = Vec::new();
        if let Some((r0, c0, r1, c1)) = table.region() {
            for r in r0..=r1 {
                for c in c0..=c1 {
                    let nf = sheet.get_format(r, c).number_format.clone();
                    before_cells.push(PivotCell { row: r as u32, col: c as u32, value: owned_value(sheet, r, c), number_format: nf.clone() });
                    after_cells.push(PivotCell { row: r as u32, col: c as u32, value: Value::Empty, number_format: nf });
                }
            }
        }
        Ok(PivotCommit {
            before: PivotState { sheet_id: sheet.id, pivot_id, table: Some(table.clone()), cells: before_cells },
            after: PivotState { sheet_id: sheet.id, pivot_id, table: None, cells: after_cells },
        })
    }

    /// Apply a pivot state atomically: set (or remove) the pivot object and
    /// write every listed cell, then recalculate dependents once. Used for the
    /// action itself and for its undo/redo.
    pub fn apply_pivot_state(&mut self, state: &PivotState) -> Result<Vec<CellId>, PivotOpError> {
        if self.sheets.iter().any(|s| s.table_view_spec().is_some()) {
            let mut candidate = self.clone();
            let changed = candidate.apply_pivot_state_unchecked(state)?;
            for sheet in &candidate.sheets {
                sheet.build_saved_table_view(NUM_ROWS.min(sheet.rows)).map_err(PivotOpError::TableSource)?;
            }
            if let Some(error) = candidate.take_incremental_errors().first() {
                return Err(PivotOpError::TableSource(format!("The pivot could not be recalculated: {error:?}")));
            }
            self.restore_snapshot_monotonic(&candidate);
            return Ok(changed);
        }
        self.apply_pivot_state_unchecked(state)
    }

    fn apply_pivot_state_unchecked(&mut self, state: &PivotState) -> Result<Vec<CellId>, PivotOpError> {
        let idx = self.sheet_index_by_id(state.sheet_id).ok_or(PivotOpError::SheetMissing)?;
        if let Some((r0, c0, r1, c1)) = state.table.as_ref().and_then(|p| p.region()) {
            if let Some(t) = self.sheets[idx].tables().iter().find(|t| t.full_range().intersects(crate::table::TableRange {
                start_row: r0, start_col: c0, end_row: r1, end_col: c1,
            })) {
                return Err(PivotOpError::Blocked { row: r0, col: c0, reason: format!("part of table {}", t.name) });
            }
        }
        for cell in &state.cells {
            if let Some(t) = self.sheets[idx].table_at(cell.row as usize, cell.col as usize) {
                return Err(PivotOpError::Blocked { row: cell.row as usize, col: cell.col as usize, reason: format!("part of table {}", t.name) });
            }
        }
        self.begin_batch();
        {
            let sheet = &mut self.sheets[idx];
            // Remove the object first, so the writer is the only thing that
            // touches its cells.
            sheet.pivots.retain(|p| p.id != state.pivot_id);
            for cell in &state.cells {
                let (r, c) = (cell.row as usize, cell.col as usize);
                sheet.write_pivot_cell(r, c, &cell.value);
                if sheet.get_format(r, c).number_format != cell.number_format {
                    sheet.set_number_format(r, c, cell.number_format.clone());
                }
            }
            if let Some(t) = &state.table {
                sheet.pivots.push(t.clone());
                sheet.pivots.sort_by_key(|p| p.id);
            }
        }
        for cell in &state.cells {
            let (r, c) = (cell.row as usize, cell.col as usize);
            self.update_cell_deps(state.sheet_id, r, c);
            self.note_cell_changed(CellId::new(state.sheet_id, r, c));
        }
        Ok(self.end_batch())
    }

    /// Is this pivot stale right now? Read-only: the saved flag, or a source
    /// edit since the last refresh, or a missing source sheet.
    pub fn is_pivot_stale(&self, p: &PivotTable) -> bool {
        if p.stale { return true; }
        if p.source.table_id.is_some() {
            let mut source = p.source;
            let mut definition = p.definition.clone();
            if self.resolve_pivot_source(&mut source, &mut definition).is_err()
                || source != p.source || definition != p.definition { return true; }
        }
        match (p.source_generation, self.pivot_source_generation(p)) {
            (Some(recorded), Some(now)) => recorded != now,
            (_, None) => true,
            (None, Some(_)) => false,
        }
    }

    /// Mark pivots stale whose source sheet has been edited since their last
    /// refresh. Cheap; call after edits. Returns true if any flag changed.
    ///
    /// Conservative: any value write on the source sheet counts, even outside
    /// the source range. A source whose formulas depend on other sheets may
    /// change without being flagged; Refresh always reads current values.
    pub fn update_pivot_staleness(&mut self) -> bool {
        let gens: Vec<(u64, Option<u64>, bool)> = self.pivots().into_iter().map(|(_, p)| (p.id, self.pivot_source_generation(p), self.is_pivot_stale(p))).collect();
        let mut changed = false;
        for sheet in &mut self.sheets {
            for p in &mut sheet.pivots {
                if p.stale {
                    continue;
                }
                let (_, current, changed_source) = gens.iter().find(|(id, _, _)| *id == p.id).copied().unwrap();
                let is_stale = match (p.source_generation, current) {
                    (Some(recorded), Some(now)) => recorded != now,
                    // Never refreshed this session (e.g. just loaded): trust the
                    // saved flag until an edit happens; record the baseline.
                    (None, Some(now)) => {
                        p.source_generation = Some(now);
                        false
                    }
                    (_, None) => true, // source sheet gone
                };
                if is_stale || changed_source {
                    p.stale = true;
                    changed = true;
                }
            }
        }
        changed
    }

    /// The pivots on sheet `sheet_idx`, in their saved form.
    pub fn saved_pivots(&self, sheet_idx: usize) -> Vec<SavedPivot> {
        let Some(sheet) = self.sheets.get(sheet_idx) else { return Vec::new() };
        sheet
            .pivots
            .iter()
            .filter_map(|p| {
                let source_sheet = self.sheet_index_by_id(p.source.sheet_id)?;
                Some(SavedPivot { source_sheet, table: p.clone() })
            })
            .collect()
    }

    /// Restore saved pivots onto sheet `sheet_idx` after its cells are loaded.
    /// Each is validated: its source sheet must exist, its output must fit the
    /// grid and not overlap another pivot or a merge, and its id must be unique
    /// in the workbook. An invalid pivot is dropped (its output stays as plain
    /// values) and a warning is returned. Loaded pivots keep their saved stale
    /// flag; nothing is refreshed.
    pub fn restore_pivots(&mut self, sheet_idx: usize, saved: Vec<SavedPivot>) -> Vec<String> {
        let mut warnings = Vec::new();
        if sheet_idx >= self.sheets.len() {
            return warnings;
        }
        for sp in saved {
            let mut t = sp.table;
            let Some(source_id) = self.sheets.get(sp.source_sheet).map(|s| s.id) else {
                warnings.push(format!("{}: its source sheet is missing; kept as values.", t.name));
                continue;
            };
            t.source.sheet_id = source_id;
            t.source_generation = None;
            if t.source.end_row < t.source.start_row || t.source.end_col < t.source.start_col {
                warnings.push(format!("{}: invalid source range; kept as values.", t.name));
                continue;
            }
            if self.find_pivot(t.id).is_some() {
                warnings.push(format!("{}: duplicate pivot id; kept as values.", t.name));
                continue;
            }
            let sheet = &self.sheets[sheet_idx];
            if let Some((r0, c0, r1, c1)) = t.region() {
                if r1 >= sheet.rows.min(NUM_ROWS) || c1 >= sheet.cols.min(NUM_COLS) {
                    warnings.push(format!("{}: output runs off the sheet; kept as values.", t.name));
                    continue;
                }
                if let Some(other) = sheet.pivots.iter().find(|p| p.intersects(r0, c0, r1, c1)) {
                    warnings.push(format!("{}: output overlaps {}; kept as values.", t.name, other.name));
                    continue;
                }
                if sheet
                    .merged_regions
                    .iter()
                    .any(|m| m.start.0 <= r1 && r0 <= m.end.0 && m.start.1 <= c1 && c0 <= m.end.1)
                {
                    warnings.push(format!("{}: output overlaps a merged cell; kept as values.", t.name));
                    continue;
                }
            }
            let sheet = &mut self.sheets[sheet_idx];
            sheet.pivots.push(t);
            sheet.pivots.sort_by_key(|p| p.id);
        }
        warnings
    }

    /// Structural-edit support: would this edit cut through a pivot's output
    /// on the edited sheet? Returns the pivot's name if so.
    pub(crate) fn pivot_cut_by_structural(
        &self,
        sheet_index: usize,
        is_row: bool,
        at: usize,
        count: usize,
        delete: bool,
    ) -> Option<String> {
        let sheet = self.sheets.get(sheet_index)?;
        sheet.pivots.iter().find_map(|p| {
            let (r0, c0, r1, c1) = p.region()?;
            let (s, e) = if is_row { (r0, r1) } else { (c0, c1) };
            let cut = if delete { at <= e && at + count > s } else { at > s && at <= e };
            cut.then(|| p.name.clone())
        })
    }

    /// Structural-edit support: move pivot outputs on the edited sheet and
    /// pivot sources that live on it. Call after cells have moved.
    pub(crate) fn shift_pivots_for_structural(
        &mut self,
        sheet_index: usize,
        is_row: bool,
        at: usize,
        count: usize,
        delete: bool,
    ) {
        let edited_id = self.sheets[sheet_index].id;
        for (idx, sheet) in self.sheets.iter_mut().enumerate() {
            for p in &mut sheet.pivots {
                // Output anchor (only on the edited sheet). Cuts were refused
                // beforehand, so the region moves as a whole or not at all.
                if idx == sheet_index {
                    let anchor = if is_row { p.anchor_row as usize } else { p.anchor_col as usize };
                    let moved = if delete {
                        if at + count <= anchor { anchor - count } else { anchor }
                    } else if at <= anchor {
                        anchor + count
                    } else {
                        anchor
                    };
                    if is_row {
                        p.anchor_row = moved as u32;
                    } else {
                        p.anchor_col = moved as u32;
                    }
                }
                // Source range (any sheet's pivot whose source is here).
                if p.source.sheet_id == edited_id {
                    let (s, e) = if is_row {
                        (p.source.start_row as usize, p.source.end_row as usize)
                    } else {
                        (p.source.start_col as usize, p.source.end_col as usize)
                    };
                    match crate::structural::shift_span(s, e, at, count, delete) {
                        Some((ns, ne)) => {
                            if is_row {
                                p.source.start_row = ns as u32;
                                p.source.end_row = ne as u32;
                            } else {
                                p.source.start_col = ns as u32;
                                p.source.end_col = ne as u32;
                            }
                        }
                        None => {} // source deleted; refresh will report it
                    }
                    p.stale = true;
                }
            }
        }
    }
}

/// The value a pivot wrote to an owned cell, read back for undo.
fn owned_value(sheet: &crate::sheet::Sheet, r: usize, c: usize) -> Value {
    sheet.get_computed_value(r, c)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pivot::{aggregate, Aggregation, PivotDefinition, PivotField, PivotSource, PivotValueField};

    fn book() -> (Workbook, SheetId, SheetId) {
        let mut wb = Workbook::new();
        let data = wb.sheet(0).unwrap().id;
        {
            let s = wb.sheet_mut(0).unwrap();
            for (c, h) in ["Region", "Amount"].iter().enumerate() {
                s.set_value(0, c, h);
            }
            for (r, (reg, amt)) in [("West", "10"), ("East", "5"), ("West", "2")].iter().enumerate() {
                s.set_value(r + 1, 0, reg);
                s.set_value(r + 1, 1, amt);
            }
        }
        let out_idx = wb.add_sheet();
        let out = wb.sheet(out_idx).unwrap().id;
        (wb, data, out)
    }

    fn table(wb: &Workbook, data: SheetId) -> PivotTable {
        PivotTable {
            id: wb.next_pivot_id(),
            name: wb.next_pivot_name(),
            source: PivotSource { table_id: None, sheet_id: data, start_row: 0, start_col: 0, end_row: 3, end_col: 1 },
            definition: PivotDefinition {
                rows: vec![PivotField { column_id: None, offset: 0, header: "Region".into() }],
                column: None,
                values: vec![PivotValueField { field: PivotField { column_id: None, offset: 1, header: "Amount".into() }, aggregation: Aggregation::Sum, number_format: None }],
            },
            anchor_row: 0,
            anchor_col: 0,
            extent: None,
            last_refresh: None,
            stale: false,
            source_generation: None,
        }
    }

    #[test]
    fn editable_table_and_pivot_output_cannot_overlap_even_in_blank_cells() {
        let (mut wb, data, out) = book();
        let editable = wb.create_table(out, crate::table::TableRange { start_row: 0, start_col: 0, end_row: 10, end_col: 3 }, "Records").unwrap();
        let mut pivot = table(&wb, data);
        pivot.anchor_row = 2;
        let (pivot, snap, gen) = wb.pivot_snapshot(&pivot).unwrap();
        let output = aggregate(&pivot.definition, &snap).unwrap();
        assert!(wb.prepare_pivot_commit(out, pivot.clone(), &output, gen, 0).is_err());
        wb.remove_table(editable.table_id()).unwrap();
        let commit = wb.prepare_pivot_commit(out, pivot, &output, gen, 0).unwrap();
        wb.apply_pivot_state(&commit.after).unwrap();
        assert!(wb.create_table(out, crate::table::TableRange { start_row: 2, start_col: 0, end_row: 10, end_col: 3 }, "Records").is_err());
        wb.apply_pivot_state(&commit.before).unwrap();
        wb.create_table(out, crate::table::TableRange { start_row: 0, start_col: 0, end_row: 10, end_col: 3 }, "Records").unwrap();
        let mut state = commit.after;
        state.cells.clear();
        assert!(wb.apply_pivot_state(&state).is_err());
    }

    fn refresh(wb: &mut Workbook, out: SheetId, t: PivotTable) -> PivotCommit {
        let (t, snap, gen) = wb.pivot_snapshot(&t).unwrap();
        let output = aggregate(&t.definition, &snap).unwrap();
        let commit = wb.prepare_pivot_commit(out, t, &output, gen, 0).unwrap();
        wb.apply_pivot_state(&commit.after).unwrap();
        commit
    }

    #[test]
    fn create_refresh_protect_undo_redo() {
        let (mut wb, data, out) = book();
        let t = table(&wb, data);
        let id = t.id;
        let create = refresh(&mut wb, out, t);
        let oi = wb.sheet_index_by_id(out).unwrap();
        assert_eq!(wb.sheet(oi).unwrap().get_display(1, 0), "East");
        assert_eq!(wb.sheet(oi).unwrap().get_display(2, 1), "12");
        assert_eq!(wb.sheet(oi).unwrap().get_display(3, 0), "Grand Total");

        // Protected: ordinary writes and clears inside the region are refused.
        wb.set_cell_value_tracked(oi, 2, 1, "999");
        wb.clear_cell_tracked(oi, 1, 0);
        assert_eq!(wb.sheet(oi).unwrap().get_display(2, 1), "12");
        assert_eq!(wb.sheet(oi).unwrap().get_display(1, 0), "East");

        // Source edit → stale; refresh picks it up.
        let di = wb.sheet_index_by_id(data).unwrap();
        wb.set_cell_value_tracked(di, 3, 1, "20");
        assert!(wb.update_pivot_staleness());
        assert!(wb.find_pivot(id).unwrap().1.stale);
        let t2 = wb.find_pivot(id).unwrap().1.clone();
        let r2 = refresh(&mut wb, out, t2);
        assert_eq!(wb.sheet(oi).unwrap().get_display(2, 1), "30");
        assert!(!wb.find_pivot(id).unwrap().1.stale);

        // Undo the refresh, then the create: back to an empty sheet with no pivot.
        wb.apply_pivot_state(&r2.before).unwrap();
        assert_eq!(wb.sheet(oi).unwrap().get_display(2, 1), "12");
        wb.apply_pivot_state(&create.before).unwrap();
        assert!(wb.find_pivot(id).is_none());
        assert_eq!(wb.sheet(oi).unwrap().get_display(1, 0), "");
        // Redo.
        wb.apply_pivot_state(&create.after).unwrap();
        assert_eq!(wb.sheet(oi).unwrap().get_display(1, 0), "East");
    }

    #[test]
    fn formulas_can_read_owned_output_and_recalc_on_refresh() {
        let (mut wb, data, out) = book();
        let t = table(&wb, data);
        refresh(&mut wb, out, t.clone());
        let oi = wb.sheet_index_by_id(out).unwrap();
        wb.set_cell_value_tracked(oi, 10, 0, "=B3*2");
        assert_eq!(wb.sheet(oi).unwrap().get_display(10, 0), "24");
        let di = wb.sheet_index_by_id(data).unwrap();
        wb.set_cell_value_tracked(di, 1, 1, "100");
        let t2 = wb.find_pivot(t.id).unwrap().1.clone();
        refresh(&mut wb, out, t2);
        assert_eq!(wb.sheet(oi).unwrap().get_display(10, 0), "204"); // (100+2)*2
    }

    #[test]
    fn placement_refuses_occupied_cells_and_keeps_old_output() {
        let (mut wb, data, out) = book();
        let t = table(&wb, data);
        refresh(&mut wb, out, t.clone());
        let oi = wb.sheet_index_by_id(out).unwrap();
        // Something next to the output blocks growth to a wider layout.
        wb.set_cell_value_tracked(oi, 0, 3, "mine");
        let mut wider = wb.find_pivot(t.id).unwrap().1.clone();
        wider.definition.values.push(PivotValueField {
            field: PivotField { column_id: None, offset: 1, header: "Amount".into() },
            aggregation: Aggregation::Count,
            number_format: None,
        });
        wider.definition.values.push(PivotValueField {
            field: PivotField { column_id: None, offset: 1, header: "Amount".into() },
            aggregation: Aggregation::Max,
            number_format: None,
        });
        let (wider, snap, gen) = wb.pivot_snapshot(&wider).unwrap();
        let output = aggregate(&wider.definition, &snap).unwrap();
        let err = wb.prepare_pivot_commit(out, wider, &output, gen, 0).unwrap_err();
        assert!(matches!(err, PivotOpError::Blocked { row: 0, col: 3, .. }), "{err:?}");
        assert_eq!(wb.sheet(oi).unwrap().get_display(0, 3), "mine");
        assert_eq!(wb.sheet(oi).unwrap().get_display(2, 1), "12");
    }

    #[test]
    fn shrinking_output_clears_trailing_cells_and_releases_them() {
        let (mut wb, data, out) = book();
        let mut t = table(&wb, data);
        t.definition.values.push(PivotValueField {
            field: PivotField { column_id: None, offset: 1, header: "Amount".into() },
            aggregation: Aggregation::Count,
            number_format: None,
        });
        refresh(&mut wb, out, t.clone());
        let oi = wb.sheet_index_by_id(out).unwrap();
        assert_eq!(wb.sheet(oi).unwrap().get_display(0, 2), "Count of Amount");
        let mut narrow = wb.find_pivot(t.id).unwrap().1.clone();
        narrow.definition.values.pop();
        refresh(&mut wb, out, narrow);
        assert_eq!(wb.sheet(oi).unwrap().get_display(0, 2), "");
        wb.set_cell_value_tracked(oi, 0, 2, "free again");
        assert_eq!(wb.sheet(oi).unwrap().get_display(0, 2), "free again");
    }

    #[test]
    fn spill_is_blocked_by_owned_blank_cells() {
        let (mut wb, data, out) = book();
        let t = table(&wb, data);
        refresh(&mut wb, out, t);
        let oi = wb.sheet_index_by_id(out).unwrap();
        // A SEQUENCE spilling down column B from above would run into B2..B4.
        wb.set_cell_value_tracked(oi, 5, 5, "=SEQUENCE(3)"); // clear area: fine
        assert_eq!(wb.sheet(oi).unwrap().get_display(6, 5), "2");
        // The grand-total row's empty cell (row 3 has values; row 4 is blank
        // below it is not owned). Spill into the owned region instead:
        wb.sheet_mut(oi).unwrap().set_value(0, 5, "");
        wb.set_cell_value_tracked(oi, 0, 1, "x"); // refused: owned
        assert_eq!(wb.sheet(oi).unwrap().get_display(0, 1), "Sum of Amount");
    }

    #[test]
    fn structural_edits_shift_or_refuse() {
        let (mut wb, data, out) = book();
        let t = table(&wb, data);
        let id = t.id;
        let mut t_shifted = t.clone();
        t_shifted.anchor_row = 5;
        refresh(&mut wb, out, t_shifted);
        let oi = wb.sheet_index_by_id(out).unwrap();
        // Cutting through the output is refused.
        assert!(wb.pivot_cut_by_structural(oi, true, 6, 1, false).is_some());
        assert!(wb.pivot_cut_by_structural(oi, true, 5, 1, true).is_some());
        // Inserting above it moves it.
        assert!(wb.pivot_cut_by_structural(oi, true, 2, 2, false).is_none());
        wb.shift_pivots_for_structural(oi, true, 2, 2, false);
        assert_eq!(wb.find_pivot(id).unwrap().1.anchor_row, 7);
        // Inserting rows inside the source range grows it and marks stale.
        let di = wb.sheet_index_by_id(data).unwrap();
        wb.shift_pivots_for_structural(di, true, 2, 1, false);
        let p = wb.find_pivot(id).unwrap().1;
        assert_eq!(p.source.end_row, 4);
        assert!(p.stale);
    }

    #[test]
    fn formats_follow_their_field_and_user_styling_survives_refresh() {
        let (mut wb, data, out) = book();
        let mut t = table(&wb, data);
        let money = NumberFormat::Currency { decimals: 2, thousands: true, negative: Default::default(), symbol: None };
        t.definition.values[0].number_format = Some(money.clone());
        let id = t.id;
        refresh(&mut wb, out, t);
        let oi = wb.sheet_index_by_id(out).unwrap();
        // Data and grand-total cells in the value column get the field's format;
        // header and label cells keep General.
        assert_eq!(wb.sheet(oi).unwrap().get_format(1, 1).number_format, money);
        assert_eq!(wb.sheet(oi).unwrap().get_format(3, 1).number_format, money);
        assert_eq!(wb.sheet(oi).unwrap().get_format(0, 1).number_format, NumberFormat::General);
        assert_eq!(wb.sheet(oi).unwrap().get_format(1, 0).number_format, NumberFormat::General);

        // The user bolds a label cell; a refresh leaves it bold.
        wb.sheet_mut(oi).unwrap().set_bold(1, 0, true);
        let t2 = wb.find_pivot(id).unwrap().1.clone();
        refresh(&mut wb, out, t2);
        assert!(wb.sheet(oi).unwrap().get_format(1, 0).bold);

        // Rearranging (adding a Count field before Sum) moves the currency
        // format with the Sum field to its new column; the Count column is plain.
        let mut t3 = wb.find_pivot(id).unwrap().1.clone();
        t3.definition.values.insert(
            0,
            PivotValueField {
                field: PivotField { column_id: None, offset: 1, header: "Amount".into() },
                aggregation: Aggregation::Count,
                number_format: crate::pivot::default_number_format(Aggregation::Count, &money),
            },
        );
        refresh(&mut wb, out, t3.clone());
        assert_eq!(wb.sheet(oi).unwrap().get_format(1, 2).number_format, money);
        assert!(matches!(wb.sheet(oi).unwrap().get_format(1, 1).number_format, NumberFormat::Number { decimals: 0, .. }));

        // Undo restores the previous formats exactly.
        let (t3, snap, gen) = wb.pivot_snapshot(&t3).unwrap();
        let output = crate::pivot::aggregate(&t3.definition, &snap).unwrap();
        let again = wb.prepare_pivot_commit(out, t3, &output, gen, 0).unwrap();
        wb.apply_pivot_state(&again.after).unwrap();
        wb.apply_pivot_state(&again.before).unwrap();
        assert_eq!(wb.sheet(oi).unwrap().get_format(1, 2).number_format, money);
    }

    #[test]
    fn blank_result_cells_keep_their_formatting() {
        let (mut wb, data, out) = book();
        let di = wb.sheet_index_by_id(data).unwrap();
        // Cross-tab with a missing intersection: East has no row with Month=Q2.
        wb.sheet_mut(di).unwrap().set_value(0, 2, "Q");
        wb.sheet_mut(di).unwrap().set_value(1, 2, "Q1");
        wb.sheet_mut(di).unwrap().set_value(2, 2, "Q1");
        wb.sheet_mut(di).unwrap().set_value(3, 2, "Q2");
        let mut t = table(&wb, data);
        t.source.end_col = 2;
        t.definition.column = Some(PivotField { column_id: None, offset: 2, header: "Q".into() });
        let id = t.id;
        refresh(&mut wb, out, t);
        let oi = wb.sheet_index_by_id(out).unwrap();
        // Rows: header(2), East, West, Grand. East × Q2 is blank at (2, 2).
        assert_eq!(wb.sheet(oi).unwrap().get_display(2, 2), "");
        wb.sheet_mut(oi).unwrap().set_background_color(2, 2, Some([255, 255, 0, 255]));
        let t2 = wb.find_pivot(id).unwrap().1.clone();
        refresh(&mut wb, out, t2);
        assert_eq!(wb.sheet(oi).unwrap().get_format(2, 2).background_color, Some([255, 255, 0, 255]));
    }

    #[test]
    fn appended_rows_below_the_source_are_detected_not_absorbed() {
        let (mut wb, data, out) = book();
        let t = table(&wb, data);
        let id = t.id;
        refresh(&mut wb, out, t);
        let p = wb.find_pivot(id).unwrap().1.clone();
        assert_eq!(wb.pivot_source_growth(&p), None);
        let di = wb.sheet_index_by_id(data).unwrap();
        wb.set_cell_value_tracked(di, 4, 0, "North");
        wb.set_cell_value_tracked(di, 4, 1, "7");
        wb.set_cell_value_tracked(di, 5, 1, "1");
        // A gap, then something else: not part of the appended block.
        wb.set_cell_value_tracked(di, 7, 0, "Note");
        assert_eq!(wb.pivot_source_growth(&p), Some(5));
        // The source itself is unchanged until the user accepts.
        assert_eq!(wb.find_pivot(id).unwrap().1.source.end_row, 3);
    }

    #[test]
    fn saved_form_round_trips_and_invalid_pivots_are_dropped() {
        let (mut wb, data, out) = book();
        let t = table(&wb, data);
        refresh(&mut wb, out, t);
        let oi = wb.sheet_index_by_id(out).unwrap();
        let saved = wb.saved_pivots(oi);
        assert_eq!(saved.len(), 1);
        assert_eq!(saved[0].source_sheet, wb.sheet_index_by_id(data).unwrap());
        let json = serde_json::to_string(&saved).unwrap();

        // A fresh workbook with different sheet ids: restore maps the index.
        let mut wb2 = Workbook::new();
        wb2.add_sheet();
        let back: Vec<SavedPivot> = serde_json::from_str(&json).unwrap();
        let warnings = wb2.restore_pivots(1, back.clone());
        assert!(warnings.is_empty(), "{warnings:?}");
        let (_, p) = wb2.find_pivot(back[0].table.id).unwrap();
        assert_eq!(p.source.sheet_id, wb2.sheet(0).unwrap().id);
        // Restoring the same pivot again is a duplicate id: dropped.
        let w = wb2.restore_pivots(1, back.clone());
        assert_eq!(w.len(), 1);
        // A missing source sheet: dropped.
        let mut bad = back.clone();
        bad[0].source_sheet = 9;
        bad[0].table.id = 99;
        let w = wb2.restore_pivots(1, bad);
        assert!(w[0].contains("source sheet is missing"));
    }

    #[test]
    fn take_and_restore_sheet_keeps_id_position_and_contents() {
        let (mut wb, data, out) = book();
        let oi = wb.sheet_index_by_id(out).unwrap();
        wb.sheet_mut(oi).unwrap().set_value(0, 0, "kept");
        let taken = wb.take_sheet(oi).unwrap();
        assert!(wb.sheet_index_by_id(out).is_none());
        assert!(wb.restore_sheet(oi, taken));
        assert_eq!(wb.sheet_index_by_id(out), Some(oi));
        assert_eq!(wb.sheet(oi).unwrap().get_display(0, 0), "kept");
        // A new sheet never reuses the restored id.
        let n = wb.add_sheet();
        assert_ne!(wb.sheet(n).unwrap().id, out);
        assert!(wb.sheet_index_by_id(data).is_some());
    }

    #[test]
    fn is_pivot_stale_is_read_only() {
        let (mut wb, data, out) = book();
        let t = table(&wb, data);
        let id = t.id;
        refresh(&mut wb, out, t);
        let p = wb.find_pivot(id).unwrap().1.clone();
        assert!(!wb.is_pivot_stale(&p));
        let di = wb.sheet_index_by_id(data).unwrap();
        wb.set_cell_value_tracked(di, 1, 1, "1");
        assert!(wb.is_pivot_stale(&p));
        assert!(!wb.find_pivot(id).unwrap().1.stale, "flag itself unchanged");
    }

    #[test]
    fn commit_is_bounded_to_the_regions_not_the_workbook() {
        let (mut wb, data, out) = book();
        let t = table(&wb, data);
        let c = refresh(&mut wb, out, t);
        // 4 rows × 2 cols output; before has the same 8 coords, all empty.
        assert_eq!(c.after.cells.len(), 8);
        assert_eq!(c.before.cells.len(), 8);
        assert!(c.before.cells.iter().all(|cell| cell.value == Value::Empty));
        assert!(c.approx_bytes() < 4096);
    }

    #[test]
    fn headless_create_names_fields_and_refresh_follows_edits() {
        let (mut wb, data, _) = book();
        let headers = vec!["Region".to_string(), "Amount".to_string()];
        let def = PivotDefinition::from_names(&headers, &["region".into()], None, &[(Some(Aggregation::Sum), "AMOUNT".into())], &[]).unwrap();
        let source = PivotSource { table_id: None, sheet_id: data, start_row: 0, start_col: 0, end_row: 3, end_col: 1 };
        let active = wb.active_sheet_index();
        let (id, idx) = wb.create_pivot(source, def).unwrap();
        assert_eq!(wb.active_sheet_index(), active, "creation does not steal the active sheet");
        assert_eq!(wb.sheet(idx).unwrap().name, "Pivot");
        let out = wb.sheet(idx).unwrap();
        assert_eq!(out.get_display(0, 1), "Sum of Amount");
        assert_eq!(out.get_display(1, 0), "East");
        assert_eq!(out.get_display(2, 1), "12");
        assert_eq!(out.get_display(3, 0), "Grand Total");

        wb.sheet_mut(0).unwrap().set_value(1, 1, "100");
        assert_eq!(wb.refresh_pivot(id).unwrap(), (4, 2));
        assert_eq!(wb.sheet(idx).unwrap().get_display(2, 1), "102");
        assert_eq!(wb.find_pivot_by_name("pivottable1").map(|(_, t)| t.id), Some(id));

        // A second pivot gets the next free sheet name.
        let def = PivotDefinition::from_names(&headers, &[], None, &[(Some(Aggregation::Count), "Region".into())], &[]).unwrap();
        let (_, idx2) = wb.create_pivot(source, def).unwrap();
        assert_eq!(wb.sheet(idx2).unwrap().name, "Pivot (2)");
    }

    #[test]
    fn headless_create_refusals_leave_the_workbook_unchanged() {
        let (mut wb, data, _) = book();
        let headers = vec!["Region".to_string(), "Amount".to_string()];
        let err = PivotDefinition::from_names(&headers, &["Month".into()], None, &[], &[]).unwrap_err();
        assert!(err.contains("no column headed \"Month\"") && err.contains("Region, Amount"), "{err}");
        assert!(PivotDefinition::from_names(&headers, &[], None, &[], &[]).is_err());
        assert_eq!(Aggregation::parse("Distinct-Count"), Some(Aggregation::DistinctCount));
        assert_eq!(Aggregation::parse("mean"), Some(Aggregation::Average));
        assert_eq!(Aggregation::parse("median"), None);

        // Source headers changed under a stored field: refused, no sheet added.
        let sheets = wb.sheets().len();
        let def = PivotDefinition::from_names(&headers, &["Region".into()], None, &[], &[]).unwrap();
        wb.sheet_mut(0).unwrap().set_value(0, 0, "Area");
        let source = PivotSource { table_id: None, sheet_id: data, start_row: 0, start_col: 0, end_row: 3, end_col: 1 };
        assert!(wb.create_pivot(source, def).is_err());
        assert_eq!(wb.sheets().len(), sheets);
    }
}
