//! Atomic structural edits with sparse authored-cell and metadata history.
//! Temporary candidates are discarded. Lifecycle history retains only the
//! added/deleted sheet; survivors use sparse cell and metadata patches.
use super::Workbook;
#[path = "workbook_history_cells.rs"]
mod cell_history;
use cell_history::{same_cells, CellPatch, SignatureEncoder};
#[cfg(test)]
use crate::cell::Cell;
use crate::{
    cell::{CellFormat, ValueRef},
    cond_format::CondFormatStore,
    named_range::NamedRangeStore,
    pivot::PivotTable,
    print_setup::PrintSetup,
    sheet::{MergedRegion, Sheet, SheetId},
    structural::Axis,
    table::DataTable,
    table_view::TableViewSpec,
    validation::ValidationStore,
};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::collections::{BTreeSet, HashMap};

#[derive(Clone, Copy, Debug, serde::Serialize)]
pub struct StructureStep {
    pub axis: Axis,
    pub at: usize,
    pub count: usize,
    pub delete: bool,
}

#[derive(Clone, Debug, Serialize)]
struct Metadata {
    tables: Vec<DataTable>,
    #[serde(skip_serializing_if = "BTreeSet::is_empty")]
    manual_hidden_rows: BTreeSet<usize>,
    view: Option<TableViewSpec>,
    #[serde(skip)]
    validations: ValidationStore,
    cond_formats: CondFormatStore,
    merges: Vec<MergedRegion>,
    row_formats: HashMap<usize, CellFormat>,
    col_formats: HashMap<usize, CellFormat>,
    frozen: (usize, usize),
    print: PrintSetup,
    pivots: Vec<PivotTable>,
}
impl Metadata {
    fn capture(s: &Sheet) -> Self {
        Self {
            tables: s.tables().to_vec(),
            manual_hidden_rows: s.manual_hidden_rows.clone(),
            view: s.table_view_spec().cloned(),
            validations: s.validations.clone(),
            cond_formats: s.cond_formats.clone(),
            merges: s.merged_regions.clone(),
            row_formats: s.row_formats.clone(),
            col_formats: s.col_formats.clone(),
            frozen: s.frozen_panes,
            print: s.print_setup.clone(),
            pivots: s.pivots.clone(),
        }
    }
    fn signature(&self) -> serde_json::Value {
        let mut copy = self.clone();
        copy.tables.sort_by_key(|t| t.id.0);
        for table in &mut copy.tables {
            table.next_column_id = 0;
        }
        for pivot in &mut copy.pivots {
            pivot.stale = false;
        }
        let mut value = serde_json::to_value(&copy).expect("structural metadata is serializable");
        let mut rules: Vec<_> = copy.validations.iter().collect();
        rules.sort_by_key(|(r, _)| (r.start_row, r.start_col, r.end_row, r.end_col));
        let mut exclusions: Vec<_> = copy.validations.exclusions_iter().collect();
        exclusions.sort_by_key(|r| (r.start_row, r.start_col, r.end_row, r.end_col));
        value["validations"] = serde_json::json!([rules, exclusions]);
        value
    }
    fn install(&self, s: &mut Sheet) {
        s.install_column_tables(self.tables.clone()); // preserves allocation high-water marks
        s.table_view_spec = self.view.clone();
        s.manual_hidden_rows = self.manual_hidden_rows.clone();
        s.validations = self.validations.clone();
        s.cond_formats = self.cond_formats.clone();
        s.merged_regions = self.merges.clone();
        s.rebuild_merge_index();
        s.row_formats = self.row_formats.clone();
        s.col_formats = self.col_formats.clone();
        s.frozen_panes = self.frozen;
        s.print_setup = self.print.clone();
        s.pivots = self.pivots.clone();
        for pivot in &mut s.pivots {
            pivot.stale = true;
            pivot.source_generation = None;
        }
    }
}
#[derive(Clone, Debug, serde::Serialize)]
struct MetaPatch {
    sheet: SheetId,
    before: Metadata,
    after: Metadata,
}

#[derive(Clone, Debug)]
struct SheetChange {
    index: usize,
    sheet: Box<Sheet>,
    added: bool,
}

#[derive(Clone, Debug)]
pub struct GuardedStructureCommit {
    pub sheet: SheetId,
    pub steps: Vec<StructureStep>,
    cells: Vec<CellPatch>,
    metadata: Vec<MetaPatch>,
    names: Option<(NamedRangeStore, NamedRangeStore)>,
    fingerprint_sheets: Option<Vec<SheetId>>,
    append_table: Option<crate::table::TableId>,
    before_fingerprint: [u8; 32],
    after_fingerprint: [u8; 32],
    // Replay restores authored states, including their existing cycles. Keep
    // each side's identities because structural edits can move or delete them.
    before_cycles: rustc_hash::FxHashSet<crate::cell_id::CellId>,
    after_cycles: rustc_hash::FxHashSet<crate::cell_id::CellId>,
    sheets: Vec<(SheetId, String, usize, usize)>,
    renamed_sheets: Vec<(SheetId, String, String)>,
    sheet_change: Option<SheetChange>,
    views: Vec<(SheetId, Option<TableViewSpec>)>,
}

#[cfg(test)]
fn image(s: &Sheet, r: usize, c: usize) -> Option<Cell> {
    s.get_cell_opt(r, c).map(|cell| {
        let mut cell = cell.to_cell();
        cell.clear_spill_state();
        cell
    })
}
#[cfg(test)]
fn signature(cell: &Option<Cell>) -> serde_json::Value {
    let Some(cell) = cell else {
        return serde_json::Value::Null;
    };
    let value = match cell.value() {
        ValueRef::Empty => ("empty", String::new()),
        ValueRef::Text(s) => ("text", s.to_owned()),
        ValueRef::Number(n) => ("number", n.to_bits().to_string()),
        ValueRef::Formula { source, .. } => ("formula", source.to_owned()),
    };
    serde_json::json!([
        value,
        cell.format,
        cell.comment(),
        cell.style_id(),
        cell.frozen_formula()
    ])
}
fn fingerprint(s: &Sheet) -> [u8; 32] {
    // Retain only coordinates while imposing canonical row/column order.
    // Materializing every cell's JSON tree here used over a gigabyte for a
    // 300k-cell sheet. Serialize one signature at a time, preserving the exact
    // fingerprint encoding and all stale-history checks.
    let mut positions: Vec<_> = s.cells_iter().map(|(p, _)| p).collect();
    positions.sort_unstable();
    let mut h = Sha256::new();
    let mut encoder = SignatureEncoder::default();
    for (r, c) in positions {
        h.update((r as u64).to_le_bytes());
        h.update((c as u64).to_le_bytes());
        let bytes = encoder.encode(s.get_cell_opt(r, c).unwrap());
        h.update((bytes.len() as u64).to_le_bytes());
        h.update(bytes);
    }
    h.finalize().into()
}

#[cfg(test)]
mod fingerprint_tests {
    use super::*;
    use crate::cell::{CellComment, CellValue};

    // The previous materialized encoder is an independent compatibility
    // oracle: optimizing the scan must not weaken what history detects.
    fn materialized(s: &Sheet) -> [u8; 32] {
        let cells: std::collections::BTreeMap<_, _> = s.cells_iter()
            .map(|(p, _)| (p, signature(&image(s, p.0, p.1)))).collect();
        let mut h = Sha256::new();
        for ((r, c), value) in cells {
            h.update((r as u64).to_le_bytes());
            h.update((c as u64).to_le_bytes());
            let bytes = serde_json::to_vec(&value).unwrap();
            h.update((bytes.len() as u64).to_le_bytes());
            h.update(bytes);
        }
        h.finalize().into()
    }

    #[test]
    fn streamed_fingerprint_preserves_authored_cells_and_order_independence() {
        let mut a = Sheet::new(SheetId(1), 200, 20);
        let mut b = a.clone();
        let cells: Vec<_> = (0..1000).map(|i| {
            let (row, col) = (i % 200, i / 200);
            let mut cell = Cell::new();
            cell.value = match i % 6 {
                0 => CellValue::Empty,
                1 => CellValue::Number(-0.0),
                2 => CellValue::Number(i as f64 / 7.0),
                3 => CellValue::Text(format!("Unicode é / \"quotes\" / {i}\n")),
                4 => CellValue::from_input("=(A1+1)*2"),
                _ => CellValue::Number(f64::NAN),
            };
            if i % 3 == 0 {
                std::sync::Arc::make_mut(&mut cell.format).bold = true;
                cell.set_comment(Some(CellComment { text: "note".into(), author: "tester".into() }));
                cell.set_style_id(Some(3));
                cell.set_frozen_formula(Some("=A1".into()));
            }
            (row, col, cell)
        }).collect();
        for (r, c, cell) in &cells {
            a.restore_history_cell(*r, *c, Some(cell.clone()));
        }
        for (r, c, cell) in cells.iter().rev() {
            let mut cell = cell.clone();
            cell.clear_spill_state();
            b.restore_history_cell(*r, *c, Some(cell));
        }
        assert_eq!(fingerprint(&a), materialized(&a));
        assert_eq!(fingerprint(&a), fingerprint(&b));
        let mut derived = cells[1].2.clone();
        let authored = signature(&Some(derived.clone()));
        derived.set_spill_parent(Some((0, 0)));
        assert_eq!(signature(&Some(derived)), authored);
        let baseline = fingerprint(&b);
        for change in 0..8 {
            let mut altered = b.clone();
            let row = match change { 6 => 0, 7 => 4, _ => 1 };
            let mut cell = altered.get_cell(row, 0);
            match change {
                0 => cell.value = CellValue::Number(0.0), // distinguish signed zero
                1 => std::sync::Arc::make_mut(&mut cell.format).italic = true,
                2 => cell.set_comment(Some(CellComment { text: "new".into(), author: String::new() })),
                3 => cell.set_style_id(Some(4)),
                4 => cell.set_frozen_formula(Some("=B1".into())),
                7 => cell.value = CellValue::from_input("=A1+1*2"),
                _ => {},
            }
            // Case 6 specifically distinguishes stored empty from absence.
            altered.restore_history_cell(row, 0, if change == 5 || change == 6 { None } else { Some(cell) });
            assert_ne!(fingerprint(&altered), baseline, "missed authored change {change}");
            assert_eq!(fingerprint(&altered), materialized(&altered));
        }
    }
}
fn scoped_fingerprint(wb: &Workbook, sheets: Option<&[SheetId]>) -> [u8; 32] {
    let mut h = Sha256::new();
    for s in wb.sheets() {
        if sheets.is_none_or(|ids| ids.contains(&s.id)) { h.update(fingerprint(s)); }
        h.update(serde_json::to_vec(&Metadata::capture(s).signature()).unwrap());
    }
    h.update(serde_json::to_vec(&serde_json::to_value(wb.named_ranges()).unwrap()).unwrap());
    h.finalize().into()
}
fn identity(wb: &Workbook) -> Vec<(SheetId, String, usize, usize)> {
    wb.sheets()
        .iter()
        .map(|s| (s.id, s.name.clone(), s.rows, s.cols))
        .collect()
}
fn views(wb: &Workbook) -> Vec<(SheetId, Option<TableViewSpec>)> {
    wb.sheets()
        .iter()
        .map(|s| (s.id, s.table_view_spec().cloned()))
        .collect()
}
pub(crate) fn validate_views(wb: &Workbook) -> Result<(), String> {
    for s in wb.sheets() {
        s.build_saved_table_view(s.rows)?;
    }
    Ok(())
}
/// Move an index or discard it when deleted. Used for desktop sizing/visibility too.
pub fn shift_structure_index(index: usize, step: StructureStep, limit: usize) -> Option<usize> {
    if index < step.at {
        Some(index)
    } else if step.delete {
        if index < step.at + step.count {
            None
        } else {
            Some(index - step.count)
        }
    } else {
        index.checked_add(step.count).filter(|i| *i < limit)
    }
}
fn boundary(index: usize, step: StructureStep) -> usize {
    if step.at >= index {
        index
    } else if step.delete {
        index - step.count.min(index - step.at)
    } else {
        index + step.count
    }
}
impl Workbook {
    /// Steps use canonical coordinates at each step. A host deleting a projected
    /// selection must resolve once and submit descending, coalesced spans.
    pub fn prepare_guarded_structure(
        &self,
        index: usize,
        steps: Vec<StructureStep>,
    ) -> Result<(Workbook, GuardedStructureCommit), String> {
        self.ensure_writable()?;
        if steps.is_empty() || steps.len() > 1000 {
            return Err("Choose between 1 and 1,000 structural spans.".into());
        }
        let before = self
            .sheets
            .get(index)
            .ok_or("The sheet no longer exists.")?;
        let mut candidate = self.clone();
        for step in &steps {
            let s = &candidate.sheets[index];
            let limit = if step.axis == Axis::Row {
                s.rows
            } else {
                s.cols
            };
            candidate.validate_structural_edit(
                index,
                step.axis,
                step.at,
                step.count,
                step.delete,
            )?;
            if !step.delete {
                // The ordinary engine preflight checks values. Explicit blank
                // cells, comments, styles and merges must not be dropped either.
                let edge = limit - step.count;
                if s.cells_iter().any(|((r, c), _)| {
                    if step.axis == Axis::Row {
                        r >= edge
                    } else {
                        c >= edge
                    }
                }) || (if step.axis == Axis::Row {
                    s.row_formats.keys()
                } else {
                    s.col_formats.keys()
                })
                .any(|i| *i >= edge)
                    || s.merged_regions.iter().any(|m| {
                        let end = if step.axis == Axis::Row { m.end.0 } else { m.end.1 };
                        // Already on the sheet edge: the insert keeps that end
                        // there instead of truncating the merge.
                        end != limit - 1 && end >= edge
                    })
                {
                    return Err(
                        "The insertion would push cell data or formatting off the worksheet."
                            .into(),
                    );
                }
            }
            candidate.structural_edit(index, step.axis, step.at, step.count, step.delete)?;
            let frozen = &mut candidate.sheets[index].frozen_panes;
            if step.axis == Axis::Row {
                frozen.0 = boundary(frozen.0, *step).min(limit);
            } else {
                frozen.1 = boundary(frozen.1, *step).min(limit);
            }
        }
        if let Some(error) = candidate.take_incremental_errors().first() {
            return Err(format!("Structural recalculation failed: {error:?}"));
        }
        validate_views(&candidate)?;
        let mut commit = self.capture_guarded_batch(&candidate)?;
        commit.sheet = before.id;
        commit.steps = steps;
        Ok((candidate, commit))
    }
}
impl Workbook {
    /// Capture an already materialized, validated batch for atomic history.
    /// The host must preflight targets and validate its own presentation state.
    pub fn capture_guarded_batch(
        &self,
        candidate: &Workbook,
    ) -> Result<GuardedStructureCommit, String> {
        self.capture_guarded_batch_inner(candidate, false, None, Some(100_000), None)
    }

    pub(super) fn capture_footer_append(&self, candidate: &Workbook, id: crate::table::TableId) -> Result<GuardedStructureCommit, String> {
        self.capture_guarded_batch_inner(candidate, false, None, Some(100_000), Some(id))
    }

    /// Conversion can rewrite every calculated cell in an existing Table.
    /// Its source-only patches are compact; the ordinary edit limit must not
    /// make a previously supported Table impossible to convert.
    pub(super) fn capture_table_conversion(&self, candidate: &Workbook) -> Result<GuardedStructureCommit, String> {
        self.capture_guarded_batch_inner(candidate, false, None, None, None)
    }

    pub(super) fn capture_sheet_rename(&self, candidate: &Workbook) -> Result<GuardedStructureCommit, String> {
        self.capture_guarded_batch_inner(candidate, true, None, Some(100_000), None)
    }

    pub(super) fn capture_sheet_change(&self, candidate: &Workbook, index: usize, added: bool) -> Result<GuardedStructureCommit, String> {
        let source = if added { candidate } else { self };
        let sheet = source.sheet(index).ok_or("The changed sheet no longer exists.")?;
        self.capture_guarded_batch_inner(candidate, false, Some(SheetChange {
            index, sheet: Box::new(sheet.clone()), added,
        }), Some(100_000), None)
    }

    fn capture_guarded_batch_inner(&self, candidate: &Workbook, allow_rename: bool, sheet_change: Option<SheetChange>, cell_limit: Option<usize>, append_table: Option<crate::table::TableId>) -> Result<GuardedStructureCommit, String> {
        self.ensure_writable()?;
        let mut before = identity(self);
        let mut after = identity(candidate);
        let mut before_views = views(self);
        let mut after_views = views(candidate);
        if let Some(change) = &sheet_change {
            let (identities, specs) = if change.added { (&mut after, &mut after_views) } else { (&mut before, &mut before_views) };
            identities.remove(change.index);
            specs.remove(change.index);
        }
        let same_shape = before.len() == after.len() && before.iter().zip(&after)
            .all(|(b, a)| b.0 == a.0 && b.2 == a.2 && b.3 == a.3);
        if !same_shape || (!allow_rename && before != after) || before_views != after_views {
            return Err("A guarded batch cannot change sheets or saved Table criteria.".into());
        }
        let renamed_sheets = before.iter().zip(&after).filter(|(b, a)| b.1 != a.1)
            .map(|(b, a)| (b.0, b.1.clone(), a.1.clone())).collect();
        // Static footer appends own sparse sheets. Dynamic addresses and name
        // changes retain a full-workbook guard: their reference scope can move.
        let fingerprint_sheets = append_table.filter(|_| self.volatile_cells.is_empty() && candidate.volatile_cells.is_empty()
            && serde_json::to_value(&self.named_ranges).unwrap() == serde_json::to_value(&candidate.named_ranges).unwrap())
            .map(|_| self.sheets.iter().filter(|b| candidate.sheet_by_id(b.id).is_some_and(|a| a.edit_generation() != b.edit_generation() || !a.exceptional_reference_sources().is_empty() || !b.exceptional_reference_sources().is_empty())).map(|s| s.id).collect::<Vec<_>>());
        for sheet in candidate.sheets() {
            if fingerprint_sheets.as_ref().is_none_or(|ids| ids.contains(&sheet.id)) { sheet.build_saved_table_view(sheet.rows)?; }
        }
        let mut cells = Vec::new();
        let mut metadata = Vec::new();
        for b in &self.sheets {
            if fingerprint_sheets.as_ref().is_some_and(|ids| !ids.contains(&b.id)) { continue; }
            let Some(a) = candidate.sheet_by_id(b.id) else { continue; };
            let positions: BTreeSet<_> = b
                .cells_iter()
                .map(|(p, _)| p)
                .chain(a.cells_iter().map(|(p, _)| p))
                .collect();
            for (row, col) in positions {
                let before = b.get_cell_opt(row, col);
                let after = a.get_cell_opt(row, col);
                if !same_cells(before, after) {
                    if cell_limit.is_some_and(|limit| cells.len() >= limit) {
                        if allow_rename {
                            return Err("Renaming would change more than 100,000 stored cells. Nothing was renamed.".into());
                        }
                        if sheet_change.is_some() {
                            return Err("This sheet change would rewrite more than 100,000 stored cells. Nothing was changed.".into());
                        }
                        return Err("This transaction changes more than 100,000 stored cells. Use a smaller selection or clear Table views first.".into());
                    }
                    cells.push(CellPatch::capture(b.id, row, col, before, after));
                }
            }
            let bm = Metadata::capture(b);
            let am = Metadata::capture(a);
            if bm.signature() != am.signature() {
                metadata.push(MetaPatch {
                    sheet: b.id,
                    before: bm,
                    after: am,
                });
            }
        }
        let names = (serde_json::to_value(&self.named_ranges).unwrap()
            != serde_json::to_value(&candidate.named_ranges).unwrap())
        .then(|| (self.named_ranges.clone(), candidate.named_ranges.clone()));
        let commit = GuardedStructureCommit {
            sheet: self.active_sheet_id(),
            steps: Vec::new(),
            cells,
            metadata,
            names,
            before_fingerprint: scoped_fingerprint(self, fingerprint_sheets.as_deref()),
            after_fingerprint: scoped_fingerprint(candidate, fingerprint_sheets.as_deref()),
            fingerprint_sheets,
            append_table,
            before_cycles: self.dep_graph.find_cycle_members(),
            after_cycles: candidate.dep_graph.find_cycle_members(),
            sheets: identity(self),
            renamed_sheets,
            sheet_change,
            views: views(self),
        };
        Ok(commit)
    }
}

impl GuardedStructureCommit {
    pub fn estimated_history_bytes(&self) -> usize {
        self.cells.iter().map(CellPatch::estimated_history_bytes).sum::<usize>()
            .saturating_add(crate::history_size::serialized_bytes(&(&self.metadata, &self.names,
                &self.before_cycles, &self.after_cycles, &self.views, &self.sheets,
                &self.steps, &self.renamed_sheets, &self.fingerprint_sheets)))
            .saturating_add(self.sheet_change.as_ref().map_or(0, |c| crate::history_size::sheet_bytes(&c.sheet)))
            .saturating_add(self.metadata.iter().map(|m| crate::history_size::serialized_bytes(&(&m.before.validations, &m.after.validations))).sum::<usize>())
    }

    pub fn is_empty(&self) -> bool {
        self.cells.is_empty() && self.metadata.is_empty() && self.names.is_none() && self.renamed_sheets.is_empty() && self.sheet_change.is_none()
    }
    pub fn changed_cell_count(&self) -> usize {
        self.cells.len()
    }
    /// Validate a full candidate before publishing. Derived values and view
    /// mappings are recomputed rather than retained in history.
    pub fn candidate(&self, wb: &Workbook, undo: bool) -> Result<Workbook, String> {
        wb.ensure_writable()?;
        let mut expected_sheets = self.sheets.clone();
        let mut expected_views = self.views.clone();
        if undo {
            for (id, _, after) in &self.renamed_sheets {
                expected_sheets.iter_mut().find(|s| s.0 == *id).unwrap().1 = after.clone();
            }
            if let Some(change) = &self.sheet_change {
                let s = &change.sheet;
                if change.added {
                    expected_sheets.insert(change.index, (s.id, s.name.clone(), s.rows, s.cols));
                    expected_views.insert(change.index, (s.id, s.table_view_spec().cloned()));
                } else {
                    expected_sheets.remove(change.index);
                    expected_views.remove(change.index);
                }
            }
        }
        if identity(wb) != expected_sheets || views(wb) != expected_views {
            return Err("Sheets or Table criteria changed since this structural edit.".into());
        }
        if let (Some(ids), Some(table)) = (&self.fingerprint_sheets, self.append_table) {
            let (owner, _) = wb.table(table).ok_or("Table no longer exists")?;
            let patch = self.metadata.iter().find(|p| p.sheet == owner).ok_or("Missing Table history")?;
            let mut readers = wb.table_readers.get(&table).map(|r| (**r).clone()).unwrap_or_default();
            readers.extend(wb.volatile_cells.iter().copied());
            for sheet in wb.sheets() {
                readers.extend(sheet.exceptional_reference_sources().into_iter().map(|(row,col)| crate::cell_id::CellId::new(sheet.id,row,col)));
            }
            for state in [&patch.before, &patch.after] {
                if let Some(t) = state.tables.iter().find(|t| t.id == table) {
                    if let Some(row) = t.totals_row() {
                        for col in t.range.start_col..=t.range.end_col {
                            readers.extend(wb.dep_graph.dependents(crate::cell_id::CellId::new(owner, row, col)));
                        }
                    }
                }
            }
            if readers.iter().any(|cell| !ids.contains(&cell.sheet)) {
                return Err("New footer references require a fresh Table operation. Undo/redo was not applied.".into());
            }
        }
        if scoped_fingerprint(wb, self.fingerprint_sheets.as_deref())
            != if undo {
                self.after_fingerprint
            } else {
                self.before_fingerprint
            }
        {
            return Err(
                "Workbook cells or structural metadata changed. Undo/redo was not applied.".into(),
            );
        }
        for p in &self.cells {
            let s = wb
                .sheet_by_id(p.sheet)
                .ok_or("A dependent sheet no longer exists.")?;
            if !p.matches(s, undo) {
                return Err("A rewritten cell changed. Undo/redo was not applied.".into());
            }
        }
        for p in &self.metadata {
            if Metadata::capture(wb.sheet_by_id(p.sheet).unwrap()).signature()
                != (if undo { &p.after } else { &p.before }).signature()
            {
                return Err("Structural metadata changed. Undo/redo was not applied.".into());
            }
        }
        if let Some((before, after)) = &self.names {
            if serde_json::to_value(wb.named_ranges()).unwrap()
                != serde_json::to_value(if undo { after } else { before }).unwrap()
            {
                return Err("Named ranges changed. Undo/redo was not applied.".into());
            }
        }
        let mut candidate = wb.clone();
        if let Some(change) = &self.sheet_change {
            let active = candidate.active_sheet_id();
            if change.added != undo {
                candidate.next_sheet_id = candidate.next_sheet_id.max(change.sheet.id.0.saturating_add(1));
                candidate.sheets.insert(change.index, *change.sheet.clone());
            } else {
                candidate.sheets.remove(change.index);
            }
            candidate.active_sheet = candidate.sheet_index_by_id(active)
                .unwrap_or(change.index.min(candidate.sheets.len() - 1));
        }
        for (id, before, after) in &self.renamed_sheets {
            candidate.sheet_by_id_mut(*id).unwrap().set_name(if undo { before } else { after });
        }
        for p in &self.metadata {
            (if undo { &p.before } else { &p.after })
                .install(candidate.sheet_by_id_mut(p.sheet).unwrap());
        }
        if let Some((before, after)) = &self.names {
            candidate.named_ranges = if undo { before } else { after }.clone();
        }
        for p in &self.cells {
            p.install(candidate.sheet_by_id_mut(p.sheet).unwrap(), undo);
        }
        candidate.rebuild_dep_graph();
        let report = candidate.recompute_full_ordered();
        let allowed_cycles = if undo { &self.before_cycles } else { &self.after_cycles };
        if (report.had_cycles && !candidate.dep_graph.find_cycle_members().is_subset(allowed_cycles))
            || report.errors.iter().any(|e| e.error.contains("not settled")) {
            return Err("History replay would create a cycle or an unsettled calculation.".into());
        }
        if self.sheet_change.is_some() || !self.renamed_sheets.is_empty() || self.names.is_some() || candidate.tables().any(|(_, table)| table.totals.is_some()) {
            // A pivot can source an unchanged formula whose result depends on
            // the restored footer or name. Authored-cell patches alone miss
            // that sheet when only a name definition changed.
            let changed: Vec<_> = candidate.sheets().iter().filter_map(|sheet| {
                let before = wb.sheet_by_id(sheet.id)?;
                sheet.cells_iter().any(|((row, col), cell)| {
                    matches!(cell.value(), ValueRef::Formula { .. })
                        && sheet.get_computed_value(row, col) != before.get_computed_value(row, col)
                }).then_some(sheet.id)
            }).collect();
            for id in changed { candidate.sheet_by_id_mut(id).unwrap().mark_table_changed(); }
        }
        if let Some(error) = candidate.take_incremental_errors().first() {
            return Err(format!("Structural replay failed: {error:?}"));
        }
        validate_views(&candidate)?;
        candidate.increment_revision();
        Ok(candidate)
    }
    pub fn replay(&self, wb: &mut Workbook, undo: bool) -> Result<(), String> {
        let candidate = self.candidate(wb, undo)?;
        wb.restore_snapshot_monotonic(&candidate);
        Ok(())
    }
}

impl serde::Serialize for GuardedStructureCommit {
    fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        crate::history_size::serialize_count(self.estimated_history_bytes(), serializer)
    }
}
