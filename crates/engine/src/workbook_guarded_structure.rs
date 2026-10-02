//! Atomic structural edits with sparse authored-cell and metadata history.
//! Temporary candidates are discarded; history never retains a Workbook/Sheet.
use super::Workbook;
use crate::{
    cell::{Cell, CellFormat, ValueRef},
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
use std::collections::{BTreeMap, BTreeSet, HashMap};

#[derive(Clone, Copy, Debug)]
pub struct StructureStep {
    pub axis: Axis,
    pub at: usize,
    pub count: usize,
    pub delete: bool,
}

#[derive(Clone, Debug, Serialize)]
struct Metadata {
    tables: Vec<DataTable>,
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
#[derive(Clone, Debug)]
struct CellPatch {
    sheet: SheetId,
    row: usize,
    col: usize,
    before: Option<Cell>,
    after: Option<Cell>,
}
#[derive(Clone, Debug)]
struct MetaPatch {
    sheet: SheetId,
    before: Metadata,
    after: Metadata,
}

#[derive(Clone, Debug)]
pub struct GuardedStructureCommit {
    pub sheet: SheetId,
    pub steps: Vec<StructureStep>,
    cells: Vec<CellPatch>,
    metadata: Vec<MetaPatch>,
    names: Option<(NamedRangeStore, NamedRangeStore)>,
    before_fingerprint: [u8; 32],
    after_fingerprint: [u8; 32],
    sheets: Vec<(SheetId, String, usize, usize)>,
    views: Vec<(SheetId, Option<TableViewSpec>)>,
}

fn image(s: &Sheet, r: usize, c: usize) -> Option<Cell> {
    s.get_cell_opt(r, c).map(|cell| {
        let mut cell = cell.to_cell();
        cell.clear_spill_state();
        cell
    })
}
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
    let cells: BTreeMap<_, _> = s
        .cells_iter()
        .map(|(p, _)| (p, signature(&image(s, p.0, p.1))))
        .collect();
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
fn workbook_fingerprint(wb: &Workbook) -> [u8; 32] {
    let mut h = Sha256::new();
    for s in wb.sheets() {
        h.update(fingerprint(s));
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
                        if step.axis == Axis::Row {
                            m.end.0 >= edge
                        } else {
                            m.end.1 >= edge
                        }
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
        let mut cells = Vec::new();
        let mut metadata = Vec::new();
        for (b, a) in self.sheets.iter().zip(&candidate.sheets) {
            let positions: BTreeSet<_> = b
                .cells_iter()
                .map(|(p, _)| p)
                .chain(a.cells_iter().map(|(p, _)| p))
                .collect();
            for (row, col) in positions {
                let before = image(b, row, col);
                let after = image(a, row, col);
                if signature(&before) != signature(&after) {
                    if cells.len() >= 100_000 {
                        return Err("This structural edit changes more than 100,000 stored cells. Use a smaller selection or clear Table views first.".into());
                    }
                    cells.push(CellPatch {
                        sheet: b.id,
                        row,
                        col,
                        before,
                        after,
                    });
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
            sheet: before.id,
            steps,
            cells,
            metadata,
            names,
            before_fingerprint: workbook_fingerprint(self),
            after_fingerprint: workbook_fingerprint(&candidate),
            sheets: identity(self),
            views: views(self),
        };
        Ok((candidate, commit))
    }
}
impl GuardedStructureCommit {
    pub fn changed_cell_count(&self) -> usize {
        self.cells.len()
    }
    /// Validate a full candidate before publishing. Derived values and view
    /// mappings are recomputed rather than retained in history.
    pub fn candidate(&self, wb: &Workbook, undo: bool) -> Result<Workbook, String> {
        wb.ensure_writable()?;
        if identity(wb) != self.sheets || views(wb) != self.views {
            return Err("Sheets or Table criteria changed since this structural edit.".into());
        }
        if workbook_fingerprint(wb)
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
            if signature(&image(s, p.row, p.col))
                != signature(if undo { &p.after } else { &p.before })
            {
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
        for p in &self.metadata {
            (if undo { &p.before } else { &p.after })
                .install(candidate.sheet_by_id_mut(p.sheet).unwrap());
        }
        if let Some((before, after)) = &self.names {
            candidate.named_ranges = if undo { before } else { after }.clone();
        }
        for p in &self.cells {
            candidate
                .sheet_by_id_mut(p.sheet)
                .unwrap()
                .restore_history_cell(
                    p.row,
                    p.col,
                    if undo {
                        p.before.clone()
                    } else {
                        p.after.clone()
                    },
                );
        }
        candidate.rebuild_dep_graph();
        candidate.recompute_full_ordered();
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
