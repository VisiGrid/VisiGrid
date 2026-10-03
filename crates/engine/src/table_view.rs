//! Bounded Table sort/filter projections. Cells and formula coordinates never move.
//!
//! This is the Phase 2 engine foundation, not an installed desktop view. Hosts must
//! retain one owner per sheet, use the workbook's saved spec, rebuild after calculation, and
//! preflight every mutation before exposing this through editing UI. The desktop
//! Table sort/filter refusal remains until that integration is complete.

use crate::{
    cell::{CellFormat, ValueRef},
    filter::{ColumnFilter, FilterKey, FilterState, RowView, SortDirection, SortKey, SortState},
    sheet::{Sheet, SheetId},
    table::{DataTable, TableColumnId, TableId, TableRange},
    validation::CellRange,
};
use serde::{Deserialize, Serialize};
use std::collections::HashSet;

/// The existing owner must be cleared explicitly before switching datasets.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ViewOwner {
    Range,
    Table(TableId),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct TableSort {
    pub column: TableColumnId,
    pub direction: SortDirection,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct TableFilter {
    pub column: TableColumnId,
    pub criteria: ColumnFilter,
}

/// Persist intent, never column offsets or a cached permutation. IDs are
/// resolved on every build; the Table catalog stores this in format version 3.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct TableViewSpec {
    pub table: TableId,
    pub sort: Option<TableSort>,
    pub filters: Vec<TableFilter>,
    pub show_filter_buttons: bool,
}

impl TableViewSpec {
    pub fn new(table: TableId) -> Self {
        Self {
            table,
            sort: None,
            filters: Vec::new(),
            show_filter_buttons: true,
        }
    }

    pub fn has_criteria(&self) -> bool {
        self.sort.is_some() || !self.filters.is_empty()
    }

    /// Default controls with no criteria need no persisted view or format bump.
    pub(crate) fn requires_persistence(&self) -> bool {
        self.has_criteria() || !self.show_filter_buttons
    }

    pub fn clear_sort(&mut self) {
        self.sort = None;
    }
    pub fn clear_filters(&mut self) {
        self.filters.clear();
    }

    /// Validate saved identities without evaluating cells or activating a view.
    /// Empty bodies retain intent until records are added again.
    pub(crate) fn validate_schema(&self, table: &DataTable) -> Result<(), String> {
        // Catalogs are untrusted. Check bounds/width before resolving offsets,
        // including a catalog restored into an existing workbook.
        table.validate(crate::sheet::NUM_ROWS, crate::sheet::NUM_COLS)?;
        if self.table != table.id {
            return Err("Table view refers to a different Table.".into());
        }
        self.resolve(table).map(|_| ())
    }

    fn resolve(&self, table: &DataTable) -> Result<FilterState, String> {
        let column = |id| {
            table
                .columns
                .iter()
                .position(|c| c.id == id)
                .map(|offset| table.range.start_col + offset)
                .ok_or_else(|| {
                    format!(
                        "Table column {} no longer exists. Clear its sort/filter criterion.",
                        id.0
                    )
                })
        };
        let range = table.range;
        let mut state = FilterState {
            filter_range: Some((
                range.start_row,
                range.start_col,
                range.end_row,
                range.end_col,
            )),
            ..Default::default()
        };
        if let Some(sort) = &self.sort {
            state.sort = Some(SortState {
                column: column(sort.column)?,
                direction: sort.direction,
            });
        }
        let mut seen = HashSet::new();
        for filter in &self.filters {
            if !seen.insert(filter.column) {
                return Err("A Table column can have only one filter criterion.".into());
            }
            if filter.criteria.selected.as_ref().is_some_and(|keys| {
                keys.iter().any(|key| {
                matches!(key, crate::filter::NormalizedFilterKey::Number(n) if !n.0.is_finite())
            })
            }) {
                return Err(
                    "Table filter numbers must be finite to save and reopen reliably.".into(),
                );
            }
            state
                .column_filters
                .insert(column(filter.column)?, filter.criteria.clone());
        }
        Ok(state)
    }
}

/// Immutable projection of one computed sheet snapshot. Rebuild after any edit,
/// schema change or recalculation, including changes to cross-sheet precedents.
/// Build failures leave both the source and any previous projection untouched.
#[derive(Debug, Clone)]
pub struct TableView {
    spec: TableViewSpec,
    sheet: SheetId,
    range: TableRange,
    rows: RowView,
    filters: FilterState,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableViewFocus {
    pub view_row: usize,
    pub data_row: usize,
    /// Hosts should explain when the edited record has been filtered out.
    pub record_hidden: bool,
}

impl TableView {
    /// `row_count` is the host's represented worksheet extent, not a scan limit.
    /// It must include the entire Table; no truncated sort/filter is permitted.
    pub fn build(
        sheet: &Sheet,
        spec: TableViewSpec,
        row_count: usize,
        current_owner: Option<ViewOwner>,
    ) -> Result<Self, String> {
        if current_owner.is_some_and(|owner| owner != ViewOwner::Table(spec.table)) {
            return Err(
                "Clear the current sort/filter view before activating another Table.".into(),
            );
        }
        let table = find_table(sheet, spec.table)?;
        let range = table.range;
        validate_table_view_layout(sheet, spec.table)?;
        if row_count <= range.end_row || row_count > sheet.rows {
            return Err(
                "The row view must include the entire Table and fit within the worksheet.".into(),
            );
        }
        let filters = spec.resolve(table)?;
        let mut rows = RowView::new(row_count);
        if let Some(sort) = &filters.sort {
            let mut keys: Vec<_> = (range.start_row + 1..=range.end_row)
                .map(|row| {
                    let key = FilterKey::from_value(&sheet.get_computed_value(row, sort.column));
                    (row, SortKey::from_filter_key(&key, row))
                })
                .collect();
            // Keep type groups and blanks-last in both directions. Equal keys
            // retain canonical record order, so rebuild and reopen agree.
            keys.sort_by(|a, b| {
                a.1.type_rank.cmp(&b.1.type_rank).then_with(|| {
                    let value = a.1.value.cmp(&b.1.value);
                    if sort.direction == SortDirection::Descending {
                        value.reverse()
                    } else {
                        value
                    }
                })
            });
            let mut order: Vec<_> = (0..row_count).collect();
            for (slot, (data_row, _)) in (range.start_row + 1..=range.end_row).zip(keys) {
                order[slot] = data_row;
            }
            rows.apply_sort(order);
        }
        let mut visible = vec![true; row_count];
        for (row, show) in visible
            .iter_mut()
            .enumerate()
            .take(range.end_row + 1)
            .skip(range.start_row + 1)
        {
            *show = filters.column_filters.iter().all(|(column, filter)| {
                filter.passes(&FilterKey::from_value(
                    &sheet.get_computed_value(row, *column),
                ))
            });
        }
        rows.apply_filter(visible);
        Ok(Self {
            spec,
            sheet: sheet.id,
            range,
            rows,
            filters,
        })
    }

    pub fn spec(&self) -> &TableViewSpec {
        &self.spec
    }
    pub fn rows(&self) -> &RowView {
        &self.rows
    }
    pub fn filters(&self) -> &FilterState {
        &self.filters
    }
    pub fn range(&self) -> TableRange {
        self.range
    }
    pub fn owner(&self) -> ViewOwner {
        ViewOwner::Table(self.spec.table)
    }

    pub fn rebuild(&self, sheet: &Sheet) -> Result<Self, String> {
        self.check_sheet(sheet)?;
        let table = find_table(sheet, self.spec.table)?;
        Self::build(
            sheet,
            self.spec.clone(),
            self.rows.row_count().max(table.range.end_row + 1),
            Some(self.owner()),
        )
    }

    fn check_sheet(&self, sheet: &Sheet) -> Result<(), String> {
        if sheet.id != self.sheet {
            return Err("The Table view belongs to a different sheet.".into());
        }
        Ok(())
    }

    /// Preflight canonical rectangles for a whole mutation batch, including
    /// format/comment/validation writes. This does not apply writes or replace
    /// other protection checks. Hosts must use this on *every* mutation path.
    /// Clear even an empty adjacent target only after clearing the active view.
    pub fn validate_mutation_ranges(
        &self,
        sheet: &Sheet,
        ranges: &[CellRange],
    ) -> Result<(), String> {
        self.check_sheet(sheet)?;
        let table = find_table(sheet, self.spec.table)?;
        if table.range != self.range {
            return Err("The Table changed shape. Rebuild its view before editing.".into());
        }
        self.spec.resolve(table)?;
        validate_table_view_layout(sheet, self.spec.table)?;
        for range in ranges {
            if range.start_row > range.end_row
                || range.start_col > range.end_col
                || range.end_row >= sheet.rows
                || range.end_col >= sheet.cols
            {
                return Err("Mutation range is outside the worksheet.".into());
            }
            if crosses_adjacent_body(self.range, range) {
                return Err("Clear the Table sort/filter view before changing adjacent cells in its body rows.".into());
            }
        }
        Ok(())
    }

    /// Map a visible-record paste once, before writing anything. Hidden records
    /// are skipped; crossing the Table's body boundary rejects the entire paste.
    /// This is opt-in: existing worksheet paste semantics are not changed here.
    pub fn visible_body_rows(
        &self,
        start_view_row: usize,
        count: usize,
    ) -> Result<Vec<usize>, String> {
        if count == 0 {
            return Ok(Vec::new());
        }
        if start_view_row <= self.range.start_row
            || start_view_row > self.range.end_row
            || !self.rows.is_view_row_visible(start_view_row)
        {
            return Err("Select a visible Table body cell before pasting records.".into());
        }
        let start = self
            .rows
            .visible_index_of(start_view_row)
            .ok_or("Select a visible Table body cell before pasting records.")?;
        let targets: Vec<_> = self.rows.visible_rows()[start..]
            .iter()
            .copied()
            .take_while(|row| *row <= self.range.end_row)
            .take(count)
            .map(|row| self.rows.view_to_data(row))
            .collect();
        if targets.len() != count {
            return Err("The paste exceeds the visible Table body. Clear the view before growing the Table.".into());
        }
        Ok(targets)
    }

    /// Retain the same canonical record after rebuilding. If filtered out,
    /// select the nearest visible position (the following row wins a tie).
    pub fn focus_record(&self, data_row: usize) -> Result<TableViewFocus, String> {
        if data_row >= self.rows.row_count() {
            return Err("Record is outside the row view.".into());
        }
        if let Some(view_row) = self.rows.data_to_view(data_row) {
            return Ok(TableViewFocus {
                view_row,
                data_row,
                record_hidden: false,
            });
        }
        let old_slot = self.rows.data_to_view_unchecked(data_row);
        let visible = self.rows.visible_rows();
        let next = visible.partition_point(|row| *row < old_slot);
        let after = visible.get(next).copied();
        let before = next.checked_sub(1).map(|i| visible[i]);
        let view_row = match (before, after) {
            (Some(a), Some(b)) => {
                if old_slot - a < b - old_slot {
                    a
                } else {
                    b
                }
            }
            (Some(a), None) => a,
            (None, Some(b)) => b,
            (None, None) => return Err("No visible worksheet rows.".into()),
        };
        Ok(TableViewFocus {
            view_row,
            data_row: self.rows.view_to_data(view_row),
            record_hidden: true,
        })
    }
}

fn find_table(sheet: &Sheet, id: TableId) -> Result<&DataTable, String> {
    sheet
        .tables()
        .iter()
        .find(|table| table.id == id)
        .ok_or_else(|| format!("Table {} no longer exists on this sheet.", id.0))
}

fn body_rows_overlap(table: TableRange, start: usize, end: usize) -> bool {
    start <= table.end_row && end > table.start_row
}

fn crosses_adjacent_body(table: TableRange, range: &CellRange) -> bool {
    body_rows_overlap(table, range.start_row, range.end_row)
        && (range.start_col < table.start_col || range.end_col > table.end_col)
}

/// Conservative whole-row-view eligibility. Scan sparse cells and metadata,
/// never the worksheet rectangle. Titles and notes above/below the body are safe.
/// Re-run for activation, rebuild and mutation preflight: activation alone is
/// insufficient because a later write could introduce neighboring content.
pub fn validate_table_view_layout(sheet: &Sheet, id: TableId) -> Result<(), String> {
    let table = find_table(sheet, id)?;
    let range = table.range;
    range.validate(sheet.rows, sheet.cols)?;
    if range.data_rows() == 0 {
        return Err("This Table has no body rows to sort or filter.".into());
    }
    let refuse = |kind| {
        Err(format!("Table view would move or hide {kind}. Use a dedicated data sheet or clear the adjacent content first."))
    };
    for ((row, col), cell) in sheet.cells_iter() {
        if row <= range.start_row || row > range.end_row {
            continue;
        }
        if cell.spill_parent().is_some() || cell.spill_info().is_some() {
            return refuse("an array spill");
        }
        if (col < range.start_col || col > range.end_col)
            && (!matches!(cell.value(), ValueRef::Empty)
                || cell.format() != &CellFormat::default()
                || cell.comment().is_some()
                || cell.style_id().is_some()
                || cell.frozen_formula().is_some())
        {
            return refuse("adjacent cell content, formatting or comments");
        }
    }
    if sheet
        .spill_receiver_coords()
        .any(|(row, _)| body_rows_overlap(range, row, row))
    {
        return refuse("an array spill");
    }
    if sheet
        .merged_regions
        .iter()
        .any(|m| body_rows_overlap(range, m.start.0, m.end.0))
    {
        return refuse("merged cells in the Table's body rows");
    }
    if sheet.tables().iter().any(|other| {
        other.id != id && body_rows_overlap(range, other.range.start_row, other.full_range().end_row)
    }) {
        return refuse("another Table");
    }
    if sheet.pivots.iter().any(|pivot| {
        pivot
            .region()
            .is_some_and(|(start, _, end, _)| body_rows_overlap(range, start, end))
    }) {
        return refuse("pivot output");
    }
    if sheet
        .validations
        .iter()
        .any(|(r, _)| crosses_adjacent_body(range, r))
        || sheet
            .validations
            .exclusions_iter()
            .any(|r| crosses_adjacent_body(range, r))
    {
        return refuse("adjacent validation metadata");
    }
    if sheet
        .cond_formats
        .iter()
        .any(|rule| rule.ranges.iter().any(|r| crosses_adjacent_body(range, r)))
    {
        return refuse("adjacent conditional formatting");
    }
    let has_adjacent_columns = range.start_col > 0 || range.end_col + 1 < sheet.cols;
    if (has_adjacent_columns
        && sheet.row_formats.iter().any(|(row, format)| {
            body_rows_overlap(range, *row, *row) && format != &CellFormat::default()
        }))
        || sheet.col_formats.iter().any(|(col, format)| {
            (*col < range.start_col || *col > range.end_col) && format != &CellFormat::default()
        })
    {
        return refuse("inherited formatting outside the Table");
    }
    Ok(())
}
