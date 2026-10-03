//! Workbook Table operations. A commit retains schema, headers and dependent
//! formula changes, never a body or workbook snapshot. Public operations
//! validate before writing.

#[path = "workbook_calculated.rs"]
mod calculated;
#[path = "workbook_table_columns.rs"]
mod columns;
pub use columns::TableColumnHistory;
#[path = "workbook_table_create.rs"]
mod create;

use super::table_refs::TableFormulaChange;
use super::Workbook;
use crate::cell::{CellValue, ValueRef};
use crate::sheet::SheetId;
use crate::table::{self, DataTable, TableColumn, TableColumnId, TableId, TableRange, TableSource, TableStyle};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, HashSet};

#[derive(Debug, Clone)]
struct HeaderCell {
    row: usize,
    col: usize,
    value: CellValue,
}

#[derive(Debug, Clone)]
struct TableState {
    table: Option<DataTable>,
    headers: Vec<HeaderCell>,
    checks: Vec<HeaderCell>,
}

/// Opaque, validated change. Replay through `apply_table_commit`, which checks
/// both schema and header preconditions before touching the workbook.
#[derive(Debug, Clone)]
pub struct TableCommit {
    sheet_id: SheetId,
    id: TableId,
    before: TableState,
    after: TableState,
    formulas: Vec<TableFormulaChange>,
    creation_references: Option<Vec<(crate::cell_id::CellId, String)>>,
    cells: Vec<(HeaderCell, HeaderCell)>,
    append_region: Option<TableRange>,
    header_insertion: Option<Box<create::HeaderInsertion>>,
    rules: Vec<calculated::RuleChange>,
}

impl TableCommit {
    pub fn inserted_header_row(&self) -> Option<usize> {
        self.header_insertion.as_ref().map(|h| h.at)
    }
    pub fn table_id(&self) -> TableId {
        self.id
    }
    pub fn sheet_id(&self) -> SheetId {
        self.sheet_id
    }
    pub fn before_table(&self) -> Option<&DataTable> {
        self.before.table.as_ref()
    }
    pub fn after_table(&self) -> Option<&DataTable> {
        self.after.table.as_ref()
    }
    pub fn header_cell_count(&self) -> usize {
        self.before.headers.len()
    }
}

/// Only Table metadata for a whole-row edit; ordinary row history owns the
/// deleted cells and presentation. Bounds are explicit because undoing a last
/// body-row deletion inserts at the new bottom edge (normally outside a Table).
#[derive(Debug, Clone)]
pub struct TableRowHistory {
    sheet: SheetId,
    at: usize,
    count: usize,
    delete: bool,
    before: Vec<DataTable>,
    after: Vec<DataTable>,
    rules: Vec<calculated::RuleChange>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct SavedTableSheet {
    pub sheet: usize,
    pub tables: Vec<DataTable>,
    #[serde(default)]
    pub column_allocators: BTreeMap<u64, u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view: Option<crate::table_view::TableViewSpec>,
}

/// Versioned, workbook-wide envelope. Sheet IDs are remapped on load; Table
/// and column IDs survive. Keep the allocator even after the final deletion.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct SavedTableCatalog {
    pub version: u32,
    pub next_table_id: u64,
    pub sheets: Vec<SavedTableSheet>,
}

impl Workbook {
    pub fn read_only_reason(&self) -> Option<&str> {
        self.sheets().iter().find_map(|s| s.read_only_reason.as_deref())
    }

    pub fn ensure_writable(&self) -> Result<(), String> {
        match self.read_only_reason() {
            Some(reason) => Err(format!("Read-only recovery: {reason} Saving, including Save As, is disabled.")),
            None => Ok(()),
        }
    }

    pub fn prepare_table_row_history(
        &self,
        sheet_index: usize,
        at: usize,
        count: usize,
        delete: bool,
    ) -> Result<Option<TableRowHistory>, String> {
        self.validate_structural_edit(
            sheet_index,
            crate::structural::Axis::Row,
            at,
            count,
            delete,
        )?;
        let sheet = &self.sheets[sheet_index];
        let rules = self.structural_rule_changes(
            sheet_index,
            crate::structural::Axis::Row,
            at,
            count,
            delete,
        );
        if sheet.tables().is_empty() && rules.is_empty() {
            return Ok(None);
        }
        let before = sheet.tables().to_vec();
        let mut after = before.clone();
        for table in &mut after {
            let (start, end) = crate::structural::shift_span(
                table.range.start_row,
                table.range.end_row,
                at,
                count,
                delete,
            )
            .ok_or("Cannot delete a Table header. Convert to a range first.")?;
            table.range.start_row = start;
            table.range.end_row = end;
            for change in rules.iter().filter(|r| r.table == table.id) {
                table
                    .columns
                    .iter_mut()
                    .find(|c| c.id == change.column)
                    .unwrap()
                    .formula = Some(change.after.clone());
                table
                    .columns
                    .iter_mut()
                    .find(|c| c.id == change.column)
                    .unwrap()
                    .formula_origin = change.after_origin;
            }
        }
        Ok(Some(TableRowHistory {
            rules,
            sheet: sheet.id,
            at,
            count,
            delete,
            before,
            after,
        }))
    }

    pub fn validate_table_row_history(
        &self,
        history: &TableRowHistory,
        undo: bool,
    ) -> Result<(), String> {
        let index = self
            .sheet_index_by_id(history.sheet)
            .ok_or("Table sheet no longer exists.")?;
        let expected = if undo {
            &history.after
        } else {
            &history.before
        };
        let current = self.sheets[index].tables();
        if current.len() != expected.len()
            || expected
                .iter()
                .any(|table| !current.iter().any(|t| same_schema(t, table)))
        {
            return Err("Tables changed since the row operation was prepared.".into());
        }
        self.validate_rule_changes(&history.rules, undo)?;
        self.validate_structural_edit(
            index,
            crate::structural::Axis::Row,
            history.at,
            history.count,
            history.delete != undo,
        )
    }

    pub fn apply_table_row_history(
        &mut self,
        history: &TableRowHistory,
        undo: bool,
    ) -> Result<Vec<(usize, usize, usize, String, String)>, String> {
        self.validate_table_row_history(history, undo)?;
        let index = self.sheet_index_by_id(history.sheet).unwrap();
        let rewrites = self.structural_edit_with_rules(
            index,
            crate::structural::Axis::Row,
            history.at,
            history.count,
            history.delete != undo,
            !undo,
        )?;
        self.apply_rule_changes(&history.rules, undo);
        let target = if undo {
            &history.before
        } else {
            &history.after
        };
        // Restore exact bounds and rules, preserving allocator high-water marks and IDs.
        let mut corrected = false;
        for table in &mut self.sheets[index].data_tables {
            let range = target.iter().find(|t| t.id == table.id).unwrap().range;
            corrected |= table.range != range;
            table.range = range;
            table.columns = target
                .iter()
                .find(|t| t.id == table.id)
                .unwrap()
                .columns
                .clone();
        }
        if corrected || !history.rules.is_empty() {
            self.rebuild_dep_graph();
            self.recompute_full_ordered();
        }
        Ok(rewrites)
    }

    pub fn tables(&self) -> impl Iterator<Item = (SheetId, &DataTable)> {
        self.sheets
            .iter()
            .flat_map(|s| s.tables().iter().map(move |t| (s.id, t)))
    }

    pub fn table(&self, id: TableId) -> Option<(SheetId, &DataTable)> {
        self.tables().find(|(_, t)| t.id == id)
    }

    pub fn table_by_name(&self, name: &str) -> Option<(SheetId, &DataTable)> {
        self.tables()
            .find(|(_, t)| t.name.eq_ignore_ascii_case(name))
    }

    pub fn next_table_name(&self) -> String {
        (1u64..)
            .map(|i| format!("Table{i}"))
            .find(|name| self.table_by_name(name).is_none() && self.get_named_range(name).is_none())
            .expect("table name space exhausted")
    }

    pub fn has_table_history(&self) -> bool {
        self.next_table_id > 1 || self.tables().next().is_some()
    }

    fn validate_table_name_available(
        &self,
        name: &str,
        except: Option<TableId>,
    ) -> Result<(), String> {
        table::validate_table_name(name)?;
        if self.get_named_range(name).is_some()
            || self
                .tables()
                .any(|(_, t)| Some(t.id) != except && t.name.eq_ignore_ascii_case(name))
        {
            return Err(format!(
                "The name '{name}' is already used by a table or named range."
            ));
        }
        Ok(())
    }

    fn validate_table_region(
        &self,
        sheet_id: SheetId,
        range: TableRange,
        except: Option<TableId>,
    ) -> Result<(), String> {
        let sheet = self
            .sheet_by_id(sheet_id)
            .ok_or("Table sheet no longer exists.")?;
        range.validate(sheet.rows, sheet.cols)?;
        if let Some(t) = sheet
            .tables()
            .iter()
            .find(|t| Some(t.id) != except && t.range.intersects(range))
        {
            return Err(format!("Table range overlaps '{}'.", t.name));
        }
        if sheet.merged_regions.iter().any(|m| {
            range.intersects(TableRange {
                start_row: m.start.0,
                start_col: m.start.1,
                end_row: m.end.0,
                end_col: m.end.1,
            })
        }) {
            return Err("Table range contains merged cells. Unmerge them first.".into());
        }
        if let Some(p) = sheet.pivot_in_rect(
            range.start_row,
            range.start_col,
            range.end_row,
            range.end_col,
        ) {
            return Err(format!("Table range overlaps {}'s pivot output.", p.name));
        }
        // Sparse iteration: even a header-only table on a large empty range
        // never allocates one entry per body cell.
        for ((row, col), cell) in sheet.cells_iter() {
            if range.contains(row, col)
                && (cell.spill_parent().is_some() || cell.spill_info().is_some())
            {
                return Err(format!(
                    "Table range contains an array spill at row {}, column {}.",
                    row + 1,
                    col + 1
                ));
            }
        }
        Ok(())
    }

    /// Preview header normalization without changing values or allocating IDs.
    pub fn preview_table_headers(
        &self,
        sheet_id: SheetId,
        range: TableRange,
    ) -> Result<Vec<String>, String> {
        self.validate_table_region(sheet_id, range, None)?;
        let sheet = self.sheet_by_id(sheet_id).unwrap();
        Ok(table::normalize_headers(
            &(range.start_col..=range.end_col)
                .map(|col| sheet.get_display(range.start_row, col))
                .collect::<Vec<_>>(),
        ))
    }

    /// Create from an explicit range whose first row supplies headers.
    /// Use `create_table_without_headers` to preserve every selected data row.
    pub fn create_table(
        &mut self,
        sheet_id: SheetId,
        range: TableRange,
        name: &str,
    ) -> Result<TableCommit, String> {
        self.validate_table_name_available(name, None)?;
        let names = self.preview_table_headers(sheet_id, range)?;
        let id = TableId(self.next_table_id);
        let next_id = id.0.checked_add(1).ok_or("Table IDs exhausted.")?;
        let columns: Vec<_> = names
            .into_iter()
            .enumerate()
            .map(|(i, name)| TableColumn {
                formula: None,
                formula_origin: 1,
                id: TableColumnId(i as u64 + 1),
                name,
            })
            .collect();
        let table = DataTable {
            id,
            name: name.into(),
            range,
            next_column_id: columns.len() as u64 + 1,
            columns,
            style: TableStyle::default(),
            source: None,
        };
        let commit = self.table_commit(sheet_id, id, None, Some(table))?;
        self.apply_table_commit(&commit, false)?;
        self.next_table_id = self.next_table_id.max(next_id);
        Ok(commit)
    }

    pub fn rename_table(&mut self, id: TableId, name: &str) -> Result<TableCommit, String> {
        self.validate_table_name_available(name, Some(id))?;
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        let mut new = old.clone();
        new.name = name.into();
        let commit = self.table_commit(sheet_id, id, Some(old.clone()), Some(new))?;
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    /// Rename multiple headers together, so swapping two names is atomic.
    pub fn rename_table_columns(
        &mut self,
        id: TableId,
        names: &[String],
    ) -> Result<TableCommit, String> {
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        if names.len() != old.columns.len() {
            return Err("Supply one name for every table column.".into());
        }
        let mut new = old.clone();
        for (col, name) in new.columns.iter_mut().zip(names) {
            col.name = name.clone();
        }
        let sheet = self.sheet_by_id(sheet_id).unwrap();
        new.validate(sheet.rows, sheet.cols)?;
        let commit = self.table_commit(sheet_id, id, Some(old.clone()), Some(new))?;
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    /// Explicitly include/release existing cells. Only bottom/right edges may
    /// move. Body cells and cells released by shrinking are never rewritten.
    pub fn resize_table(&mut self, id: TableId, range: TableRange) -> Result<TableCommit, String> {
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        if (range.start_row, range.start_col) != (old.range.start_row, old.range.start_col) {
            return Err("Resize keeps the table's header and first column fixed.".into());
        }
        self.validate_table_region(sheet_id, range, Some(id))?;
        let mut new = old.clone();
        new.range = range;
        new.columns.truncate(range.width());
        if range.width() > old.columns.len() {
            let sheet = self.sheet_by_id(sheet_id).unwrap();
            let mut names: Vec<_> = old.columns.iter().map(|c| c.name.clone()).collect();
            names.extend(
                (old.range.end_col + 1..=range.end_col)
                    .map(|c| sheet.get_display(range.start_row, c)),
            );
            let names = table::normalize_headers(&names);
            // Existing names win over new headers. Normalize each incoming
            // header against those names, without changing existing columns.
            let mut used: HashSet<_> = old.columns.iter().map(|c| c.name.to_lowercase()).collect();
            for name in names.into_iter().skip(old.columns.len()) {
                let base = name.clone();
                let mut name = name;
                let mut suffix = 2;
                while used.contains(&name.to_lowercase()) {
                    name = format!("{base}{suffix}");
                    suffix += 1;
                }
                used.insert(name.to_lowercase());
                let col_id = new.next_column_id;
                new.next_column_id = col_id.checked_add(1).ok_or("Column IDs exhausted.")?;
                new.columns.push(TableColumn {
                    formula: None,
                    formula_origin: 1,
                    id: TableColumnId(col_id),
                    name,
                });
            }
        }
        let commit = self.table_commit(sheet_id, id, Some(old.clone()), Some(new))?;
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    /// Detect desktop append intent without making ordinary loaders grow Tables.
    /// A side-crossing paste must be explicitly resized before it can be applied.
    pub fn table_append_target(
        &self,
        sheet_id: SheetId,
        range: TableRange,
    ) -> Result<Option<TableId>, String> {
        let sheet = self
            .sheet_by_id(sheet_id)
            .ok_or("Sheet no longer exists.")?;
        range.validate(sheet.rows, sheet.cols)?;
        let mut target = None;
        for table in sheet.tables() {
            let r = table.range;
            if range.start_row > r.start_row
                && range.start_row <= r.end_row.saturating_add(1)
                && range.end_row > r.end_row
                && range.start_col <= r.end_col
                && range.end_col >= r.start_col
            {
                if range.start_col < r.start_col || range.end_col > r.end_col || target.is_some() {
                    return Err("Resize the Table first: this paste crosses its side and bottom boundaries.".into());
                }
                target = Some(table.id);
            }
        }
        Ok(target)
    }

    /// Add empty records, optionally writing a rectangular paste or a typed
    /// value in the same guarded history commit. Never moves adjacent cells.
    pub fn append_table_rows(
        &mut self,
        id: TableId,
        count: usize,
        writes: &[(usize, usize, String)],
    ) -> Result<TableCommit, String> {
        self.append_table_rows_impl(id, count, writes, false)
    }

    pub fn append_table_rows_with_edit(
        &mut self,
        id: TableId,
        count: usize,
        row: usize,
        col: usize,
        value: &str,
    ) -> Result<TableCommit, String> {
        self.append_table_rows_impl(id, count, &[(row, col, value.into())], true)
    }

    fn append_table_rows_impl(
        &mut self,
        id: TableId,
        count: usize,
        writes: &[(usize, usize, String)],
        infer_rule: bool,
    ) -> Result<TableCommit, String> {
        if count == 0 {
            return Err("Append at least one row.".into());
        }
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        let mut new = old.clone();
        new.range.end_row = new
            .range
            .end_row
            .checked_add(count)
            .ok_or("Append exceeds the sheet boundary.")?;
        self.validate_table_region(sheet_id, new.range, Some(id))?;
        let region = TableRange {
            start_row: old.range.end_row + 1,
            ..new.range
        };
        self.validate_empty_table_append(sheet_id, region)?;
        let mut inferred = None;
        let sheet = self.sheet_by_id(sheet_id).unwrap();
        if infer_rule && writes.len() == 1 {
            let (row, col, source) = &writes[0];
            if new.range.contains(*row, *col)
                && *row > new.range.start_row
                && source.starts_with('=')
                && new.columns[*col - new.range.start_col].formula.is_none()
                && (old.range.start_row + 1..=old.range.end_row)
                    .all(|r| r == *row || sheet.get_raw(r, *col).is_empty())
            {
                crate::formula::parser::parse(source)
                    .map_err(|e| format!("Invalid column formula: {e}"))?;
                new.columns[*col - new.range.start_col].formula = Some(source.clone());
                new.columns[*col - new.range.start_col].formula_origin = *row - new.range.start_row;
                inferred = Some(*col);
            }
        }
        if inferred.is_some() {
            let (row, col, source) = &writes[0];
            self.validate_calculated_formula(sheet_id, *row, *col, source)?;
        }
        let mut seen = HashSet::new();
        let mut cells = Vec::with_capacity(writes.len());
        for (row, col, text) in writes {
            if *row <= old.range.start_row
                || !new.range.contains(*row, *col)
                || !seen.insert((*row, *col))
            {
                return Err("Append writes must be unique cells within the Table body.".into());
            }
            cells.push((
                HeaderCell {
                    row: *row,
                    col: *col,
                    value: sheet.get_cell(*row, *col).value,
                },
                HeaderCell {
                    row: *row,
                    col: *col,
                    value: CellValue::from_input(text),
                },
            ));
        }
        let fill_start = if inferred.is_some() {
            new.range.start_row + 1
        } else {
            region.start_row
        };
        for row in fill_start..=new.range.end_row {
            for col in region.start_col..=region.end_col {
                if (row >= region.start_row || inferred == Some(col)) && !seen.contains(&(row, col))
                {
                    if let Some(formula) = new.formula_at(row, col) {
                        self.validate_calculated_formula(sheet_id, row, col, &formula)?;
                        cells.push((
                            HeaderCell {
                                row,
                                col,
                                value: sheet.get_cell(row, col).value,
                            },
                            HeaderCell {
                                row,
                                col,
                                value: CellValue::from_input(&formula),
                            },
                        ));
                    }
                }
            }
        }
        let mut commit = self.table_commit(sheet_id, id, Some(old.clone()), Some(new))?;
        commit.cells = cells;
        commit.append_region = Some(region);
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    fn validate_empty_table_append(
        &self,
        sheet_id: SheetId,
        region: TableRange,
    ) -> Result<(), String> {
        let sheet = self
            .sheet_by_id(sheet_id)
            .ok_or("Sheet no longer exists.")?;
        for ((row, col), cell) in sheet.cells_iter() {
            if region.contains(row, col)
                && (!matches!(cell.value(), ValueRef::Empty) || cell.comment().is_some())
            {
                return Err(format!("Cannot append: {} already contains data. Resize the Table to include existing records.", crate::cell_id::CellId::new(sheet_id, row, col)));
            }
        }
        Ok(())
    }

    pub fn set_table_style(
        &mut self,
        id: TableId,
        style: TableStyle,
    ) -> Result<TableCommit, String> {
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        let mut new = old.clone();
        new.style = style;
        let commit = self.table_commit(sheet_id, id, Some(old.clone()), Some(new))?;
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    /// Link a Table to the recipe it is loaded from, record a refresh, or
    /// (None) unlink it. Cells are not touched.
    pub fn set_table_source(
        &mut self,
        id: TableId,
        source: Option<TableSource>,
    ) -> Result<TableCommit, String> {
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        let mut new = old.clone();
        new.source = source;
        let commit = self.table_commit(sheet_id, id, Some(old.clone()), Some(new))?;
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    /// Convert to a range: remove metadata and rewrite dependent structured
    /// formulas to absolute A1 references, preserving explicit formatting.
    pub fn remove_table(&mut self, id: TableId) -> Result<TableCommit, String> {
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        let commit = self.table_commit(sheet_id, id, Some(old.clone()), None)?;
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    fn table_commit(
        &self,
        sheet_id: SheetId,
        id: TableId,
        before: Option<DataTable>,
        mut after: Option<DataTable>,
    ) -> Result<TableCommit, String> {
        let sheet = self
            .sheet_by_id(sheet_id)
            .ok_or("Table sheet no longer exists.")?;
        let mut headers = Vec::new();
        let mut changed = Vec::new();
        let mut before_checks = Vec::new();
        let mut after_checks = Vec::new();
        if let Some(table) = &after {
            for (offset, column) in table.columns.iter().enumerate() {
                let (row, col) = (table.range.start_row, table.range.start_col + offset);
                let value = sheet.get_cell(row, col).value;
                before_checks.push(HeaderCell {
                    row,
                    col,
                    value: value.clone(),
                });
                after_checks.push(HeaderCell {
                    row,
                    col,
                    value: CellValue::Text(column.name.clone()),
                });
                if !matches!(&value, CellValue::Text(t) if t == &column.name) {
                    headers.push(HeaderCell { row, col, value });
                    changed.push(HeaderCell {
                        row,
                        col,
                        value: CellValue::Text(column.name.clone()),
                    });
                }
            }
        }
        if let Some(table) = &before {
            for offset in 0..table.columns.len() {
                let (row, col) = (table.range.start_row, table.range.start_col + offset);
                if offset >= after.as_ref().map_or(0, |t| t.columns.len()) {
                    let cell = HeaderCell {
                        row,
                        col,
                        value: sheet.get_cell(row, col).value,
                    };
                    before_checks.push(cell.clone());
                    after_checks.push(cell);
                }
            }
        }
        let creation_references = if before.is_none() {
            after
                .as_ref()
                .map(|table| self.table_reference_sources(sheet_id, table))
        } else {
            None
        };
        let formulas = self.table_formula_changes(sheet_id, before.as_ref(), after.as_ref())?;
        let rules = self.schema_rule_changes(sheet_id, before.as_ref(), after.as_ref())?;
        // The operation's own rule changes are part of its primary schema state.
        if let Some(table) = &mut after {
            for change in rules.iter().filter(|r| r.table == id) {
                if let Some(column) = table.columns.iter_mut().find(|c| c.id == change.column) {
                    column.formula = Some(change.after.clone());
                    column.formula_origin = change.after_origin;
                }
            }
        }
        let rules = rules.into_iter().filter(|r| r.table != id).collect();
        Ok(TableCommit {
            rules,
            cells: Vec::new(),
            append_region: None,
            header_insertion: None,
            creation_references,
            formulas,
            sheet_id,
            id,
            before: TableState {
                table: before,
                headers,
                checks: before_checks,
            },
            after: TableState {
                table: after,
                headers: changed,
                checks: after_checks,
            },
        })
    }

    /// `undo = true` restores the before state; false reapplies the after
    /// state. A stale commit fails atomically instead of overwriting edits.
    pub fn apply_table_commit(&mut self, commit: &TableCommit, undo: bool) -> Result<(), String> {
        if commit.header_insertion.is_some() {
            return self.apply_headerless_table_commit(commit, undo);
        }
        self.apply_table_commit_inner(commit, undo)
    }

    fn apply_table_commit_inner(&mut self, commit: &TableCommit, undo: bool) -> Result<(), String> {
        let (expected, target) = if undo {
            (&commit.after, &commit.before)
        } else {
            (&commit.before, &commit.after)
        };
        let sheet = self
            .sheet_by_id(commit.sheet_id)
            .ok_or("Table sheet no longer exists.")?;
        let current = sheet.tables().iter().find(|t| t.id == commit.id);
        if let Some(spec) = sheet.table_view_spec().filter(|spec| spec.table == commit.id) {
            let table = target.table.as_ref()
                .ok_or("Clear the saved Table view before removing its Table.")?;
            spec.validate_schema(table)?;
        }
        let same_schema = match (current, expected.table.as_ref()) {
            (Some(a), Some(b)) => {
                let mut a = a.clone();
                a.next_column_id = b.next_column_id;
                a == *b
            }
            (None, None) => true,
            _ => false,
        };
        if !same_schema {
            return Err("Table changed since this operation was prepared.".into());
        }
        if let Some(region) = commit.append_region {
            if undo {
                let owned: HashSet<_> = commit
                    .cells
                    .iter()
                    .map(|(cell, _)| (cell.row, cell.col))
                    .collect();
                for ((row, col), cell) in sheet.cells_iter() {
                    if region.contains(row, col)
                        && (cell.comment().is_some()
                            || (!owned.contains(&(row, col))
                                && !matches!(cell.value(), ValueRef::Empty)))
                    {
                        return Err(
                            "An appended row changed since this operation was prepared.".into()
                        );
                    }
                }
            } else {
                self.validate_empty_table_append(commit.sheet_id, region)?;
            }
        }
        for (before, after) in &commit.cells {
            let cell = if undo { after } else { before };
            if !same_value(&sheet.get_cell(cell.row, cell.col).value, &cell.value) {
                return Err("An appended cell changed since this operation was prepared.".into());
            }
        }
        for cell in &expected.checks {
            let value = sheet.get_cell(cell.row, cell.col).value;
            let same = match (&value, &cell.value) {
                (CellValue::Empty, CellValue::Empty) => true,
                (CellValue::Number(a), CellValue::Number(b)) => a.to_bits() == b.to_bits(),
                (CellValue::Text(a), CellValue::Text(b)) => a == b,
                (CellValue::Formula { source: a, .. }, CellValue::Formula { source: b, .. }) => {
                    a == b
                }
                _ => false,
            };
            if !same {
                return Err("Table header changed since this operation was prepared.".into());
            }
        }
        if let Some(table) = &target.table {
            table.validate(sheet.rows, sheet.cols)?;
            self.validate_table_name_available(&table.name, Some(commit.id))?;
            self.validate_table_region(commit.sheet_id, table.range, Some(commit.id))?;
            if self
                .tables()
                .any(|(s, t)| t.id == table.id && s != commit.sheet_id)
            {
                return Err("Table identity is already used on another sheet.".into());
            }
        }
        // Headers share one row. Validate their bounding span once, avoiding
        // a full sparse-cell scan for every changed column on a wide table.
        if let (Some(first), Some(last)) = (target.headers.first(), target.headers.last()) {
            self.validate_table_region(
                commit.sheet_id,
                TableRange {
                    start_row: first.row,
                    end_row: first.row,
                    start_col: first.col,
                    end_col: last.col,
                },
                Some(commit.id),
            )?;
        }
        for change in &commit.formulas {
            let expected_source = if undo { &change.after } else { &change.before };
            let formula_sheet = self
                .sheet_by_id(change.cell.sheet)
                .ok_or("Formula sheet no longer exists.")?;
            if formula_sheet.get_raw(change.cell.row, change.cell.col) != *expected_source
                || formula_sheet
                    .table_header_at(change.cell.row, change.cell.col)
                    .is_some()
                || formula_sheet.is_pivot_owned(change.cell.row, change.cell.col)
            {
                return Err(
                    "A dependent formula changed since the table operation was prepared.".into(),
                );
            }
        }
        // Refuse replay if newer formulas would also require a rewrite. Existing
        // destructive rewrites (#REF!/A1 conversion) are restored from the commit.
        if let Some(sources) = &commit.creation_references {
            // Undo creation restores the originally unbound formulas, rather
            // than converting them to A1. New references make that undo stale.
            let table = commit.after.table.as_ref().unwrap();
            if self.table_reference_sources(commit.sheet_id, table) != *sources {
                return Err("Table references changed since creation.".into());
            }
        } else {
            for change in self.table_formula_changes(
                commit.sheet_id,
                expected.table.as_ref(),
                target.table.as_ref(),
            )? {
                if !commit
                    .formulas
                    .iter()
                    .any(|saved| saved.cell == change.cell)
                    && !commit.cells.iter().any(|(cell, _)| {
                        change.cell.sheet == commit.sheet_id
                            && change.cell.row == cell.row
                            && change.cell.col == cell.col
                    })
                {
                    return Err("New dependent formulas require a fresh table operation.".into());
                }
            }
        }
        self.validate_rule_changes(&commit.rules, undo)?;
        let fresh_rules = self.schema_rule_changes(
            commit.sheet_id,
            expected.table.as_ref(),
            target.table.as_ref(),
        )?;
        if fresh_rules.iter().any(|r| {
            r.table != commit.id
                && !commit
                    .rules
                    .iter()
                    .any(|saved| saved.table == r.table && saved.column == r.column)
        }) {
            return Err("New calculated-column references require a fresh Table operation.".into());
        }
        let sheet = self.sheet_by_id_mut(commit.sheet_id).unwrap();
        sheet.data_tables.retain(|t| t.id != commit.id);
        for cell in &target.headers {
            sheet.write_table_header(cell.row, cell.col, cell.value.clone());
        }
        let allocator = sheet
            .table_column_allocators
            .entry(commit.id.0)
            .or_insert(1);
        for table in [commit.before.table.as_ref(), commit.after.table.as_ref()]
            .into_iter()
            .flatten()
        {
            *allocator = (*allocator).max(table.next_column_id);
        }
        if let Some(table) = &target.table {
            let mut table = table.clone();
            table.next_column_id = table.next_column_id.max(*allocator);
            sheet.data_tables.push(table);
        }
        sheet.table_id_high_water = sheet.table_id_high_water.max(commit.id.0);
        sheet.mark_table_changed();
        self.next_table_id = self.next_table_id.max(commit.id.0 + 1);
        self.refresh_table_name_reservations();
        for change in &commit.formulas {
            let source = if undo { &change.before } else { &change.after };
            self.sheet_by_id_mut(change.cell.sheet).unwrap().set_value(
                change.cell.row,
                change.cell.col,
                source,
            );
        }
        for (before, after) in &commit.cells {
            let cell = if undo { before } else { after };
            self.sheet_by_id_mut(commit.sheet_id)
                .unwrap()
                .write_table_header(cell.row, cell.col, cell.value.clone());
        }
        self.apply_rule_changes(&commit.rules, undo);
        // Membership changes affect symbolic shape dependencies even when no
        // cell was written (including empty -> nonempty bodies).
        self.rebuild_dep_graph();
        self.recompute_full_ordered();
        self.bump_revision_for_structure();
        Ok(())
    }

    pub(crate) fn refresh_table_name_reservations(&mut self) {
        self.named_ranges.table_names = self.tables().map(|(_, t)| t.name.to_lowercase()).collect();
    }

    pub fn saved_tables(&self) -> SavedTableCatalog {
        SavedTableCatalog {
            // Older readers must refuse a recipe link rather than drop it
            version: if self.tables().any(|(_, t)| t.source.is_some()) {
                4
            } else if self.sheets.iter().any(|s| {
                s.table_view_spec().is_some_and(|v| v.requires_persistence())
            }) {
                3
            } else if self
                .tables()
                .any(|(_, t)| t.columns.iter().any(|c| c.formula.is_some()))
            {
                2
            } else {
                1
            },
            next_table_id: self.next_table_id,
            sheets: self
                .sheets
                .iter()
                .enumerate()
                .filter(|(_, s)| !s.tables().is_empty() || !s.table_column_allocators.is_empty())
                .map(|(i, s)| SavedTableSheet {
                    sheet: i,
                    tables: s.tables().to_vec(),
                    column_allocators: s.table_column_allocators.clone(),
                    view: s.table_view_spec.clone().filter(|v| v.requires_persistence()),
                })
                .collect(),
        }
    }

    /// Strict, atomic restore after sheets/cells/merges/pivots are loaded and
    /// before recalculation. Reject corrupt metadata, never discard silently.
    pub fn restore_tables(&mut self, saved: SavedTableCatalog) -> Result<(), String> {
        if ![1, 2, 3, 4].contains(&saved.version) || saved.next_table_id == 0 {
            return Err("Unsupported or invalid Tables metadata version/allocator.".into());
        }
        let mut names = HashSet::new();
        let mut ids = HashSet::new();
        let mut sheet_indices = HashSet::new();
        for entry in &saved.sheets {
            if !sheet_indices.insert(entry.sheet) {
                return Err("Duplicate table sheet entry.".into());
            }
            let sheet = self
                .sheet(entry.sheet)
                .ok_or("Table references a missing sheet.")?;
            if let Some(spec) = &entry.view {
                if saved.version < 3 {
                    return Err("Table views require Tables metadata version 3.".into());
                }
                let table = entry.tables.iter().find(|t| t.id == spec.table)
                    .ok_or("Saved Table view refers to a missing Table on this sheet.")?;
                spec.validate_schema(table)?;
            }
            if entry
                .column_allocators
                .iter()
                .any(|(&id, &next)| id == 0 || id >= saved.next_table_id || next == 0)
            {
                return Err("Invalid saved column allocator.".into());
            }
            for (i, table) in entry.tables.iter().enumerate() {
                table.validate(sheet.rows, sheet.cols)?;
                if saved.version == 1 && table.columns.iter().any(|c| c.formula.is_some()) {
                    return Err("Calculated columns require Tables metadata version 2.".into());
                }
                if saved.version < 4 && table.source.is_some() {
                    return Err("Recipe-backed Tables require Tables metadata version 4.".into());
                }
                if table.id.0 >= saved.next_table_id
                    || !ids.insert(table.id)
                    || !names.insert(table.name.to_lowercase())
                    || self.get_named_range(&table.name).is_some()
                {
                    return Err("Duplicate table name/identity or invalid table allocator.".into());
                }
                if entry.tables[..i]
                    .iter()
                    .any(|t| t.range.intersects(table.range))
                {
                    return Err("Saved tables overlap.".into());
                }
                // Existing tables are replaced only after validating the whole
                // envelope; collisions against their old bounds are irrelevant.
                if sheet.merged_regions.iter().any(|m| {
                    table.range.intersects(TableRange {
                        start_row: m.start.0,
                        start_col: m.start.1,
                        end_row: m.end.0,
                        end_col: m.end.1,
                    })
                }) || sheet
                    .pivot_in_rect(
                        table.range.start_row,
                        table.range.start_col,
                        table.range.end_row,
                        table.range.end_col,
                    )
                    .is_some()
                {
                    return Err("Saved table overlaps merged cells or pivot output.".into());
                }
                for (offset, col) in table.columns.iter().enumerate() {
                    if !matches!(sheet.get_cell_opt(table.range.start_row, table.range.start_col + offset).map(|c| c.value()), Some(ValueRef::Text(t)) if t == col.name)
                    {
                        return Err(format!(
                            "Saved table '{}' does not match its header cells.",
                            table.name
                        ));
                    }
                }
                for ((r, c), cell) in sheet.cells_iter() {
                    if table.range.contains(r, c)
                        && (cell.spill_info().is_some() || cell.spill_parent().is_some())
                    {
                        return Err("Saved table overlaps an array spill.".into());
                    }
                }
            }
        }
        for sheet in &mut self.sheets {
            sheet.data_tables.clear();
            sheet.table_view_spec = None;
        }
        for entry in saved.sheets {
            let sheet = &mut self.sheets[entry.sheet];
            for (id, next) in entry.column_allocators {
                let current = sheet.table_column_allocators.entry(id).or_insert(1);
                *current = (*current).max(next);
            }
            sheet.data_tables = entry.tables;
            sheet.table_view_spec = entry.view;
            for table in &mut sheet.data_tables {
                let next = sheet.table_column_allocators.entry(table.id.0).or_insert(1);
                *next = (*next).max(table.next_column_id);
                table.next_column_id = *next;
            }
        }
        self.next_table_id = self.next_table_id.max(saved.next_table_id);
        // Retain the allocator when a caller later extracts a single Sheet.
        for sheet in &mut self.sheets {
            sheet.table_id_high_water = self.next_table_id - 1;
        }
        self.refresh_table_name_reservations();
        // Pivots load before Tables; establish their metadata-aware baseline
        // only once the complete source catalog is available.
        self.update_pivot_staleness();
        Ok(())
    }
}

fn same_value(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Empty, CellValue::Empty) => true,
        (CellValue::Number(a), CellValue::Number(b)) => a.to_bits() == b.to_bits(),
        (CellValue::Text(a), CellValue::Text(b)) => a == b,
        (CellValue::Formula { source: a, .. }, CellValue::Formula { source: b, .. }) => a == b,
        _ => false,
    }
}

fn same_schema(a: &DataTable, b: &DataTable) -> bool {
    let mut a = a.clone();
    a.next_column_id = b.next_column_id;
    a == *b
}
