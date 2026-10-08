//! A workbook that stays alive between calls.
//!
//! `recompute` is stateless: it deserialises the sheets, writes every cell,
//! rebuilds the dependency graph and recomputes the lot, on every call. That is
//! correct for what it was built for — verifying a saved document once — and
//! ruinous as a per-keystroke path. Measured natively at 200,000 formulas, one
//! call is ~470 ms of writes, ~170 ms of graph rebuild and ~630 ms of
//! recompute, none of which the edit needed: the engine can recalculate the
//! cells a single write actually dirties in microseconds, and does so for the
//! desktop app already.
//!
//! What stood in the way was not a missing engine capability but this
//! boundary. A `Session` holds the `Workbook`, so construction is paid once at
//! load and an edit is an edit:
//!
//! ```text
//! const s = new Session(sheets);           // pays the load cost, once
//! const delta = s.set_cell(0, 0, 0, "42"); // pays only for what changed
//! ```
//!
//! `set_cell` returns just the cells that were re-evaluated, which is the other
//! half of the problem. Handing back the whole workbook so the caller could
//! diff it would put the cost right back, in serialisation instead of
//! evaluation.
//!
//! Nothing here replaces `recompute`. Verification still wants a cold rebuild
//! from the document of record — that independence is the point of the check.

use serde::{Deserialize, Serialize};
use visigrid_engine::cell_id::CellId;
use visigrid_engine::sheet::SheetId;
use visigrid_engine::workbook::{Recalculated, Workbook};
use wasm_bindgen::prelude::*;

use crate::{build_workbook, out_result, InSheet, OutResult};

fn to_js<T: Serialize>(value: &T) -> Result<JsValue, JsValue> {
    serde_wasm_bindgen::to_value(value).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[derive(Deserialize)]
pub(crate) struct InEdit {
    #[serde(default)]
    pub(crate) sheet: usize,
    pub(crate) row: usize,
    pub(crate) col: usize,
    /// Raw content as the user typed it: a formula with its leading `=`, or a
    /// literal. `null` clears the cell.
    pub(crate) raw: Option<String>,
}

/// What one edit (or one batch) changed.
#[derive(Serialize, Debug)]
pub(crate) struct Delta {
    /// The workbook revision after the edit, so a caller applying deltas can
    /// tell it has not missed one.
    pub(crate) revision: u64,
    /// What the caller must repaint, with current values: first the cells
    /// written — every written formula (its computed value is news to the
    /// caller) and every written literal whose value, error or display
    /// changed (a clear, a coercion such as "1e3" to 1000) — then the cells
    /// re-evaluated as a consequence, including cells an array spilled into or
    /// out of. Each cell appears once.
    pub(crate) cells: Vec<OutResult>,
    /// Set when a circular reference forced a full recompute, so the extent of
    /// the change is the whole workbook and `cells` is not a delta to trust.
    /// The caller should re-read everything with `all_results`.
    ///
    /// A flag rather than an empty list: silence and "everything moved" must
    /// not look the same to a renderer.
    pub(crate) resync: bool,
}

/// A workbook held across calls, edited in place.
#[wasm_bindgen]
pub struct Session {
    wb: Workbook,
}

#[wasm_bindgen]
impl Session {
    /// Build the workbook. Takes the same shape as `recompute`:
    /// `[{ name?, cells: [{ row, col, raw }] }]`.
    ///
    /// This is where the whole per-call cost of the stateless path now lives —
    /// once, at load, instead of on every keystroke.
    #[wasm_bindgen(constructor)]
    pub fn new(input: JsValue) -> Result<Session, JsValue> {
        console_error_panic_hook::set_once();
        let sheets: Vec<InSheet> =
            serde_wasm_bindgen::from_value(input).map_err(|e| JsValue::from_str(&e.to_string()))?;
        Ok(Session::from_sheets(&sheets))
    }

    /// Construction without the JS boundary. See `apply_one`.
    pub(crate) fn from_sheets(sheets: &[InSheet]) -> Session {
        Session { wb: build_workbook(sheets) }
    }

    /// The workbook revision. Advances on every edit that changes anything.
    #[wasm_bindgen(getter)]
    pub fn revision(&self) -> u64 {
        self.wb.revision()
    }

    /// Number of sheets, so a caller can bounds-check before addressing one.
    #[wasm_bindgen(getter)]
    pub fn sheet_count(&self) -> usize {
        self.wb.sheets().len()
    }

    /// Write one cell and report what it changed.
    ///
    /// `raw` is the content as typed. Pass `null` to clear the cell.
    #[wasm_bindgen]
    pub fn set_cell(
        &mut self,
        sheet: usize,
        row: usize,
        col: usize,
        raw: Option<String>,
    ) -> Result<JsValue, JsValue> {
        let delta = self
            .apply_one(sheet, row, col, raw.as_deref())
            .map_err(|e| JsValue::from_str(&e))?;
        to_js(&delta)
    }

    /// `set_cell` without the JS boundary, so its behaviour can be asserted in
    /// an ordinary test — the same split `recompute_core` uses next door, and
    /// for the same reason: everything interesting here is what comes back,
    /// and `JsValue` cannot be inspected off-target.
    pub(crate) fn apply_one(
        &mut self,
        sheet: usize,
        row: usize,
        col: usize,
        raw: Option<&str>,
    ) -> Result<Delta, String> {
        if sheet >= self.wb.sheets().len() {
            return Err(format!("no sheet at index {sheet}"));
        }
        let before = self.snapshot(&[(sheet, row, col)]);
        let recalculated = match raw {
            Some(value) => self.wb.set_cell_value_tracked(sheet, row, col, value),
            None => self.wb.clear_cell_tracked(sheet, row, col),
        };
        Ok(self.delta(&before, recalculated))
    }

    /// Apply several edits with a single recalculation at the end.
    ///
    /// Shape: `[{ sheet?, row, col, raw }]`. This is the op-burst path — a
    /// paste, an import, or a run of operations arriving together — and it
    /// matters because recalculating once for fifty writes is not fifty times
    /// cheaper than recalculating fifty times, it is very much cheaper: the
    /// dirty sets overlap and are evaluated as one.
    #[wasm_bindgen]
    pub fn set_cells(&mut self, edits: JsValue) -> Result<JsValue, JsValue> {
        let edits: Vec<InEdit> =
            serde_wasm_bindgen::from_value(edits).map_err(|e| JsValue::from_str(&e.to_string()))?;
        let delta = self.apply_many(&edits).map_err(|e| JsValue::from_str(&e))?;
        to_js(&delta)
    }

    /// `set_cells` without the JS boundary. See `apply_one`.
    pub(crate) fn apply_many(&mut self, edits: &[InEdit]) -> Result<Delta, String> {
        let sheet_count = self.wb.sheets().len();
        if let Some(bad) = edits.iter().find(|e| e.sheet >= sheet_count) {
            // Checked before any write, so a bad index in the middle of a
            // batch cannot leave half of it applied.
            return Err(format!("no sheet at index {}", bad.sheet));
        }

        let written: Vec<(usize, usize, usize)> = edits.iter().map(|e| (e.sheet, e.row, e.col)).collect();
        let before = self.snapshot(&written);
        self.wb.begin_batch();
        for edit in edits {
            match edit.raw {
                Some(ref value) => {
                    self.wb.set_cell_value_tracked(edit.sheet, edit.row, edit.col, value)
                }
                None => self.wb.clear_cell_tracked(edit.sheet, edit.row, edit.col),
            };
        }
        let outcome = self.wb.end_batch_outcome();
        Ok(self.delta(&before, outcome.recalculated))
    }

    /// Every formula cell in the workbook with its current value.
    ///
    /// The resync path behind `Delta::resync`, and the way a caller that has
    /// lost track can start again without rebuilding the session.
    #[wasm_bindgen]
    pub fn all_results(&self) -> Result<JsValue, JsValue> {
        to_js(&self.all_results_core())
    }

    /// `all_results` without the JS boundary. See `apply_one`.
    pub(crate) fn all_results_core(&self) -> Delta {
        let mut cells = Vec::new();
        for (idx, sheet) in self.wb.sheets().iter().enumerate() {
            let mut coords: Vec<(usize, usize)> = sheet
                .cells_iter()
                .filter(|(_, cell)| cell.value().formula_ast().is_some())
                .map(|((row, col), _)| (row, col))
                .collect();
            // Sparse storage iterates in hash order; sort so the same workbook
            // always produces the same list.
            coords.sort_unstable();
            cells.extend(coords.into_iter().map(|(row, col)| out_result(idx, sheet, row, col)));
        }
        Delta { revision: self.wb.revision(), cells, resync: false }
    }

    /// The written cells as they stood before an edit, in edit order, once
    /// each (a batch may write one cell twice).
    fn snapshot(&self, written: &[(usize, usize, usize)]) -> Vec<OutResult> {
        let mut seen = std::collections::HashSet::new();
        written
            .iter()
            .filter(|w| seen.insert(**w))
            .filter_map(|&(sheet, row, col)| self.wb.sheets().get(sheet).map(|s| out_result(sheet, s, row, col)))
            .collect()
    }

    fn delta(&self, before: &[OutResult], recalculated: Recalculated) -> Delta {
        let cells = match recalculated {
            Recalculated::Cells(cells) => cells,
            Recalculated::All => {
                return Delta { revision: self.wb.revision(), cells: Vec::new(), resync: true };
            }
        };
        let mut out: Vec<OutResult> = Vec::with_capacity(before.len() + cells.len());
        let mut reported = std::collections::HashSet::new();
        for old in before {
            let sheet = &self.wb.sheets()[old.sheet];
            let now = out_result(old.sheet, sheet, old.row, old.col);
            let is_formula = sheet
                .get_cell_opt(old.row, old.col)
                .is_some_and(|c| c.value().formula_ast().is_some());
            let changed = now.value != old.value || now.error != old.error || now.display != old.display;
            if is_formula || changed {
                reported.insert((now.sheet, now.row, now.col));
                out.push(now);
            }
        }
        for result in cells.iter().filter_map(|id| self.project(*id)) {
            if reported.insert((result.sheet, result.row, result.col)) {
                out.push(result);
            }
        }
        Delta { revision: self.wb.revision(), cells: out, resync: false }
    }

    /// `CellId` carries a `SheetId`; the wire carries an index. Linear over
    /// sheets, which is a handful, not over cells.
    fn project(&self, id: CellId) -> Option<OutResult> {
        let idx = self.sheet_index(id.sheet)?;
        Some(out_result(idx, &self.wb.sheets()[idx], id.row, id.col))
    }

    fn sheet_index(&self, id: SheetId) -> Option<usize> {
        self.wb.sheets().iter().position(|s| s.id == id)
    }
}

/// A row or column shift, as applied to the workbook. Mirrors must apply the
/// same shift to their own grid (cells at or after `at` on `sheet` move by
/// `count`; a delete removes `[at, at + count)` first).
#[derive(Serialize, Debug, Clone, Copy, PartialEq)]
pub(crate) struct Shift {
    pub(crate) sheet: usize,
    /// "rows" or "cols".
    pub(crate) axis: &'static str,
    pub(crate) at: usize,
    pub(crate) count: usize,
    pub(crate) delete: bool,
}

/// A formula whose text the operation rewrote (its references moved, or became
/// `#REF!`), so a mirror can repaint the new text at its new position.
#[derive(Serialize, Debug, Clone, PartialEq)]
pub(crate) struct Rewritten {
    pub(crate) sheet: usize,
    pub(crate) row: usize,
    pub(crate) col: usize,
    pub(crate) formula: String,
}

/// What a structural operation changed.
///
/// `cells` follows `Delta::cells`: every formula cell or spill receiver whose
/// value, error or display differs from before, at its position AFTER the
/// operation, each once. Cells that only moved, with the same value, are not
/// listed — the mirror moves them with `shift`.
#[derive(Serialize, Debug)]
pub(crate) struct StructuralDelta {
    pub(crate) revision: u64,
    pub(crate) shift: Option<Shift>,
    pub(crate) rewritten: Vec<Rewritten>,
    pub(crate) cells: Vec<OutResult>,
    /// The sheet the operation created, for `add_sheet`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) sheet: Option<usize>,
}

/// One formula's text, for `formulas`.
#[derive(Serialize, Debug, Clone, PartialEq)]
pub(crate) struct FormulaText {
    pub(crate) row: usize,
    pub(crate) col: usize,
    pub(crate) formula: String,
}

/// Values keyed by stable sheet id, so a delete that renumbers sheets does
/// not mismatch them.
type Values = std::collections::HashMap<(u64, usize, usize), (Option<serde_json::Value>, Option<String>, String)>;

#[wasm_bindgen]
impl Session {
    /// Insert `count` empty rows before row `at` on `sheet`.
    #[wasm_bindgen]
    pub fn insert_rows(&mut self, sheet: usize, at: usize, count: usize) -> Result<JsValue, JsValue> {
        to_js(&self.structural(sheet, Axis::Row, at, count, false).map_err(|e| JsValue::from_str(&e))?)
    }

    /// Delete rows `[at, at + count)` on `sheet`.
    #[wasm_bindgen]
    pub fn delete_rows(&mut self, sheet: usize, at: usize, count: usize) -> Result<JsValue, JsValue> {
        to_js(&self.structural(sheet, Axis::Row, at, count, true).map_err(|e| JsValue::from_str(&e))?)
    }

    /// Insert `count` empty columns before column `at` on `sheet`.
    #[wasm_bindgen]
    pub fn insert_cols(&mut self, sheet: usize, at: usize, count: usize) -> Result<JsValue, JsValue> {
        to_js(&self.structural(sheet, Axis::Col, at, count, false).map_err(|e| JsValue::from_str(&e))?)
    }

    /// Delete columns `[at, at + count)` on `sheet`.
    #[wasm_bindgen]
    pub fn delete_cols(&mut self, sheet: usize, at: usize, count: usize) -> Result<JsValue, JsValue> {
        to_js(&self.structural(sheet, Axis::Col, at, count, true).map_err(|e| JsValue::from_str(&e))?)
    }

    /// Append a sheet. `name` is optional (the next free "SheetN" otherwise).
    #[wasm_bindgen]
    pub fn add_sheet(&mut self, name: Option<String>) -> Result<JsValue, JsValue> {
        to_js(&self.add_sheet_core(name.as_deref()).map_err(|e| JsValue::from_str(&e))?)
    }

    /// Rename a sheet, exactly as the desktop does (formula text is not
    /// rewritten), then recalculate so nothing stale is reported.
    #[wasm_bindgen]
    pub fn rename_sheet(&mut self, sheet: usize, name: String) -> Result<JsValue, JsValue> {
        to_js(&self.rename_sheet_core(sheet, &name).map_err(|e| JsValue::from_str(&e))?)
    }

    /// Delete a sheet, then recalculate what referenced it.
    #[wasm_bindgen]
    pub fn delete_sheet(&mut self, sheet: usize) -> Result<JsValue, JsValue> {
        to_js(&self.delete_sheet_core(sheet).map_err(|e| JsValue::from_str(&e))?)
    }

    /// Formula text of every formula cell in a rectangle (inclusive), sorted
    /// by row then column: what a mirror repaints after formulas moved.
    #[wasm_bindgen]
    pub fn formulas(&self, sheet: usize, top: usize, left: usize, bottom: usize, right: usize) -> Result<JsValue, JsValue> {
        to_js(&self.formulas_core(sheet, top, left, bottom, right).map_err(|e| JsValue::from_str(&e))?)
    }

    /// Install the clock and random seed volatile functions read, so replicas
    /// recalculate identically. Pass `undefined` for any field to use the
    /// machine's: `now_ms` (Unix ms), `utc_offset_minutes`, `seed` (an
    /// integer below 2^53). Affects recalculations from the next edit on.
    #[wasm_bindgen]
    pub fn set_clock(&mut self, now_ms: Option<f64>, utc_offset_minutes: Option<i32>, seed: Option<f64>) {
        self.set_clock_core(now_ms, utc_offset_minutes, seed);
    }
}

use visigrid_engine::structural::Axis;

impl Session {
    pub(crate) fn set_clock_core(&mut self, now_ms: Option<f64>, utc_offset_minutes: Option<i32>, seed: Option<f64>) {
        let clock = visigrid_engine::RecalcClock {
            now_ms: now_ms.filter(|v| v.is_finite()).map(|v| v as i64),
            utc_offset_seconds: utc_offset_minutes.map(|m| m as i64 * 60),
            seed: seed.filter(|v| v.is_finite() && *v >= 0.0).map(|v| v as u64),
        };
        let any = clock.now_ms.is_some() || clock.utc_offset_seconds.is_some() || clock.seed.is_some();
        self.wb.set_recalc_clock(any.then_some(clock));
    }

    /// Every formula cell and spill receiver, with its current reading.
    fn values(&self) -> Values {
        let mut out = Values::new();
        for (idx, sheet) in self.wb.sheets().iter().enumerate() {
            let id = sheet.id.0;
            let formulas = sheet
                .cells_iter()
                .filter(|(_, cell)| cell.value().formula_ast().is_some())
                .map(|((r, c), _)| (r, c));
            for (r, c) in formulas.chain(sheet.spill_receiver_coords()) {
                let o = out_result(idx, sheet, r, c);
                out.insert((id, r, c), (o.value, o.error, o.display));
            }
        }
        out
    }

    /// Cells whose reading differs between `before` (keys mapped through
    /// `moved`; `None` = the cell was deleted) and now, plus cells that held
    /// something before and now read empty. Sorted by sheet, row, column.
    fn changed_since(&self, before: Values, moved: impl Fn(u64, usize, usize) -> Option<(usize, usize)>) -> Vec<OutResult> {
        let mut prior: Values = Values::with_capacity(before.len());
        for ((id, r, c), reading) in before {
            if let Some((r2, c2)) = moved(id, r, c) {
                prior.insert((id, r2, c2), reading);
            }
        }
        let now = self.values();
        let mut keys: Vec<(u64, usize, usize)> = now.keys().chain(prior.keys()).copied().collect();
        keys.sort_unstable();
        keys.dedup();
        let mut out = Vec::new();
        for key in keys {
            if now.get(&key) == prior.get(&key) {
                continue;
            }
            let Some(idx) = self.sheet_index(SheetId(key.0)) else { continue };
            out.push(out_result(idx, &self.wb.sheets()[idx], key.1, key.2));
        }
        out
    }

    pub(crate) fn structural(&mut self, sheet: usize, axis: Axis, at: usize, count: usize, delete: bool) -> Result<StructuralDelta, String> {
        if sheet >= self.wb.sheets().len() {
            return Err(format!("no sheet at index {sheet}"));
        }
        if count == 0 {
            return Err("count must be at least 1".into());
        }
        let moved_sheet = self.wb.sheets()[sheet].id.0;
        let before = self.values();
        let rewrites = self.wb.structural_edit(sheet, axis, at, count, delete)?;
        let is_row = matches!(axis, Axis::Row);
        let moved = |id: u64, r: usize, c: usize| -> Option<(usize, usize)> {
            if id != moved_sheet {
                return Some((r, c));
            }
            let pos = if is_row { r } else { c };
            let new_pos = if !delete {
                if pos >= at { pos + count } else { pos }
            } else if pos < at {
                pos
            } else if pos < at + count {
                return None;
            } else {
                pos - count
            };
            Some(if is_row { (new_pos, c) } else { (r, new_pos) })
        };
        let cells = self.changed_since(before, moved);
        // The engine reports rewrites at their PRE-edit positions (for undo);
        // a mirror wants them where they are now.
        let mut rewritten: Vec<Rewritten> = rewrites
            .into_iter()
            .filter_map(|(idx, row, col, _old, new)| {
                let id = self.wb.sheets().get(idx)?.id.0;
                let (row, col) = moved(id, row, col)?;
                Some(Rewritten { sheet: idx, row, col, formula: new })
            })
            .collect();
        rewritten.sort_by_key(|r| (r.sheet, r.row, r.col));
        Ok(StructuralDelta {
            revision: self.wb.revision(),
            shift: Some(Shift { sheet, axis: if is_row { "rows" } else { "cols" }, at, count, delete }),
            rewritten,
            cells,
            sheet: None,
        })
    }

    pub(crate) fn add_sheet_core(&mut self, name: Option<&str>) -> Result<StructuralDelta, String> {
        let before = self.values();
        let (candidate, _) = self.wb.prepare_sheet_add(name)?;
        let idx = candidate.sheet_count() - 1;
        self.wb.restore_snapshot_monotonic(&candidate);
        let cells = self.changed_since(before, |_, r, c| Some((r, c)));
        Ok(StructuralDelta { revision: self.wb.revision(), shift: None, rewritten: Vec::new(), cells, sheet: Some(idx) })
    }

    pub(crate) fn rename_sheet_core(&mut self, sheet: usize, name: &str) -> Result<StructuralDelta, String> {
        if sheet >= self.wb.sheets().len() {
            return Err(format!("no sheet at index {sheet}"));
        }
        let before = self.values();
        let target = &self.wb.sheets()[sheet];
        let (candidate, commit) = self.wb.prepare_sheet_rename(target.id, &target.name, name)?;
        let mut rewritten = Vec::new();
        for (index, updated) in candidate.sheets().iter().enumerate() {
            let previous = &self.wb.sheets()[index];
            for ((row, col), cell) in updated.cells_iter() {
                if cell.value().formula_ast().is_some() {
                    let formula = updated.get_raw(row, col);
                    if formula != previous.get_raw(row, col) {
                        rewritten.push(Rewritten { sheet: index, row, col, formula });
                    }
                }
            }
        }
        rewritten.sort_by_key(|r| (r.sheet, r.row, r.col));
        if !commit.is_empty() {
            self.wb.restore_snapshot_monotonic(&candidate);
        }
        let cells = self.changed_since(before, |_, r, c| Some((r, c)));
        Ok(StructuralDelta { revision: self.wb.revision(), shift: None, rewritten, cells, sheet: None })
    }

    pub(crate) fn delete_sheet_core(&mut self, sheet: usize) -> Result<StructuralDelta, String> {
        if sheet >= self.wb.sheets().len() {
            return Err(format!("no sheet at index {sheet}"));
        }
        let gone = self.wb.sheets()[sheet].id.0;
        let before = self.values();
        let (candidate, _) = self.wb.prepare_sheet_delete(SheetId(gone))?;
        let mut rewritten = Vec::new();
        for (index, updated) in candidate.sheets().iter().enumerate() {
            let previous = self.wb.sheet_by_id(updated.id).unwrap();
            for ((row, col), cell) in updated.cells_iter() {
                if cell.value().formula_ast().is_some() {
                    let formula = updated.get_raw(row, col);
                    if formula != previous.get_raw(row, col) {
                        rewritten.push(Rewritten { sheet: index, row, col, formula });
                    }
                }
            }
        }
        rewritten.sort_by_key(|r| (r.sheet, r.row, r.col));
        self.wb.restore_snapshot_monotonic(&candidate);
        let cells = self.changed_since(before, |id, r, c| (id != gone).then_some((r, c)));
        Ok(StructuralDelta { revision: self.wb.revision(), shift: None, rewritten, cells, sheet: None })
    }

    pub(crate) fn formulas_core(&self, sheet: usize, top: usize, left: usize, bottom: usize, right: usize) -> Result<Vec<FormulaText>, String> {
        let s = self.wb.sheets().get(sheet).ok_or_else(|| format!("no sheet at index {sheet}"))?;
        let mut out: Vec<FormulaText> = s
            .cells_iter()
            .filter(|((r, c), cell)| *r >= top && *r <= bottom && *c >= left && *c <= right && cell.value().formula_ast().is_some())
            .map(|((row, col), _)| FormulaText { row, col, formula: s.get_raw(row, col) })
            .collect();
        out.sort_by_key(|f| (f.row, f.col));
        Ok(out)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::InCell;

    fn sheet(cells: &[(usize, usize, &str)]) -> InSheet {
        InSheet {
            name: None,
            cells: cells
                .iter()
                .map(|(row, col, raw)| InCell { row: *row, col: *col, raw: raw.to_string() })
                .collect(),
        }
    }

    /// (row, col, display) for the cells a delta reports.
    fn reported(delta: &Delta) -> Vec<(usize, usize, String)> {
        delta.cells.iter().map(|c| (c.row, c.col, c.display.clone())).collect()
    }

    #[test]
    fn an_edit_returns_only_what_it_changed() {
        // The whole point of the session. A1 has one dependent chain and the
        // sheet has an unrelated formula; a caller repainting from this delta
        // must be given the chain and not the workbook.
        let mut s = Session::from_sheets(&[sheet(&[
            (0, 0, "10"),
            (0, 1, "=A1*2"),
            (0, 2, "=B1+1"),
            (5, 5, "=1+1"), // untouched by the edit
        ])]);

        let delta = s.apply_one(0, 0, 0, Some("20")).unwrap();

        // The written cell first (its value changed), then the chain.
        assert_eq!(
            reported(&delta),
            vec![(0, 0, "20".to_string()), (0, 1, "40".to_string()), (0, 2, "41".to_string())]
        );
        assert!(!delta.resync);
    }

    #[test]
    fn an_edit_nothing_reads_returns_nothing() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "10"), (0, 1, "=A1*2")])]);

        let delta = s.apply_one(0, 4, 4, Some("hello")).unwrap();

        // Only the written cell itself: nothing reads it.
        assert_eq!(reported(&delta), vec![(4, 4, "hello".to_string())]);
        assert!(!delta.resync);
    }

    #[test]
    fn clearing_a_cell_recalculates_what_read_it() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "10"), (0, 1, "=A1*2")])]);

        let delta = s.apply_one(0, 0, 0, None).unwrap();

        // The clear is visible as the cell going empty.
        assert_eq!(reported(&delta), vec![(0, 0, String::new()), (0, 1, "0".to_string())]);
    }

    #[test]
    fn a_cycle_asks_for_a_resync_rather_than_reporting_a_delta() {
        // The engine falls back to a full recompute here, so any cell may have
        // moved. Reporting an empty delta would tell a renderer nothing
        // changed, which is the opposite of what happened.
        let mut s = Session::from_sheets(&[sheet(&[
            (0, 0, "1"),
            (0, 1, "=A1+C1"),
            (0, 2, "=B1"),
        ])]);

        let delta = s.apply_one(0, 0, 0, Some("20")).unwrap();

        assert!(delta.resync, "a full recompute must not be reported as a delta");
        assert!(delta.cells.is_empty());
    }

    #[test]
    fn a_batch_reports_the_union_once() {
        let mut s = Session::from_sheets(&[sheet(&[
            (0, 0, "1"),
            (1, 0, "2"),
            (0, 1, "=A1+A2"),
        ])]);

        let delta = s
            .apply_many(&[
                InEdit { sheet: 0, row: 0, col: 0, raw: Some("10".into()) },
                InEdit { sheet: 0, row: 1, col: 0, raw: Some("20".into()) },
            ])
            .unwrap();

        // The two writes, then B1 once: it is dirtied by both and evaluated once.
        assert_eq!(
            reported(&delta),
            vec![(0, 0, "10".to_string()), (1, 0, "20".to_string()), (0, 1, "30".to_string())]
        );
    }

    #[test]
    fn a_bad_sheet_index_applies_nothing() {
        // Rejected before the batch opens, so a bad index late in a run cannot
        // leave the earlier writes applied and the workbook half-edited.
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "1"), (0, 1, "=A1*2")])]);
        let before = s.revision();

        let err = s
            .apply_many(&[
                InEdit { sheet: 0, row: 0, col: 0, raw: Some("99".into()) },
                InEdit { sheet: 7, row: 0, col: 0, raw: Some("1".into()) },
            ])
            .unwrap_err();

        assert!(err.contains("no sheet at index 7"), "{err}");
        assert_eq!(s.revision(), before, "nothing should have been applied");
        assert_eq!(reported(&s.all_results_core()), vec![(0, 1, "2".to_string())]);
    }

    #[test]
    fn the_revision_advances_with_edits_so_a_caller_can_spot_a_gap() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "1"), (0, 1, "=A1*2")])]);

        let first = s.apply_one(0, 0, 0, Some("2")).unwrap().revision;
        let second = s.apply_one(0, 0, 0, Some("3")).unwrap().revision;

        assert!(second > first, "{first} -> {second}");
    }

    #[test]
    fn all_results_lists_every_formula_cell_in_a_stable_order() {
        // Sparse storage iterates in hash order, so an unsorted list would
        // shuffle between runs and make a resync diff against itself.
        let s = Session::from_sheets(&[sheet(&[
            (9, 0, "=1+1"),
            (0, 0, "5"),
            (0, 1, "=A1*2"),
            (3, 2, "=A1+1"),
        ])]);

        let listed = reported(&s.all_results_core());

        assert_eq!(
            listed,
            vec![
                (0, 1, "10".to_string()),
                (3, 2, "6".to_string()),
                (9, 0, "2".to_string()),
            ]
        );
        assert_eq!(listed, reported(&s.all_results_core()), "and stable across calls");
    }

    #[test]
    fn edits_reach_across_sheets() {
        let mut s = Session::from_sheets(&[
            sheet(&[(0, 0, "5")]),
            InSheet {
                name: Some("Two".into()),
                cells: vec![InCell { row: 0, col: 0, raw: "=Sheet1!A1+10".into() }],
            },
        ]);

        let delta = s.apply_one(0, 0, 0, Some("100")).unwrap();

        assert_eq!(delta.cells.len(), 2);
        assert_eq!((delta.cells[0].sheet, delta.cells[0].display.as_str()), (0, "100"), "the write");
        assert_eq!(delta.cells[1].sheet, 1, "the dependent is on the second sheet");
        assert_eq!(delta.cells[1].display, "110");
    }

    #[test]
    fn a_new_formula_reports_its_own_value() {
        // Entering a formula must not require all_results to learn its value.
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "4")])]);

        let delta = s.apply_one(0, 0, 1, Some("=A1*3")).unwrap();

        assert_eq!(reported(&delta), vec![(0, 1, "12".to_string())]);
        assert_eq!(delta.cells[0].value, Some(serde_json::json!(12.0)));
    }

    #[test]
    fn a_formula_is_reported_even_when_its_value_did_not_change() {
        // A literal 2 replaced by =1+1: same value, but the caller needs to
        // know the cell is now computed by the engine.
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "2")])]);

        let delta = s.apply_one(0, 0, 0, Some("=1+1")).unwrap();

        assert_eq!(reported(&delta), vec![(0, 0, "2".to_string())]);
    }

    #[test]
    fn rewriting_a_literal_with_the_same_value_reports_nothing() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "7")])]);

        let delta = s.apply_one(0, 0, 0, Some("7")).unwrap();

        assert!(delta.cells.is_empty(), "{:?}", reported(&delta));
    }

    #[test]
    fn a_coerced_literal_reports_its_stored_value() {
        let mut s = Session::from_sheets(&[sheet(&[])]);

        let delta = s.apply_one(0, 0, 0, Some("1e3")).unwrap();

        assert_eq!(delta.cells.len(), 1);
        assert_eq!(delta.cells[0].value, Some(serde_json::json!(1000.0)));
    }

    #[test]
    fn a_batch_reports_each_written_formula_once() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "1")])]);

        let delta = s
            .apply_many(&[
                InEdit { sheet: 0, row: 0, col: 1, raw: Some("=A1+1".into()) },
                InEdit { sheet: 0, row: 0, col: 1, raw: Some("=A1+2".into()) },
                InEdit { sheet: 0, row: 0, col: 2, raw: Some("=B1*10".into()) },
            ])
            .unwrap();

        assert_eq!(reported(&delta), vec![(0, 1, "3".to_string()), (0, 2, "30".to_string())]);
    }

    #[test]
    fn spilled_cells_are_in_the_delta() {
        // An array formula's receivers change too; the delta must carry them
        // so a caller need not fall back to all_results for spilling formulas.
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "2")])]);

        let grown = s.apply_one(0, 0, 1, Some("=SEQUENCE(A1)")).unwrap();
        assert_eq!(reported(&grown), vec![(0, 1, "1".to_string()), (1, 1, "2".to_string())]);

        let more = s.apply_one(0, 0, 0, Some("3")).unwrap();
        assert!(reported(&more).contains(&(2, 1, "3".to_string())), "a new receiver: {:?}", reported(&more));

        let fewer = s.apply_one(0, 0, 0, Some("1")).unwrap();
        let cells = reported(&fewer);
        assert!(cells.contains(&(1, 1, String::new())), "a retired receiver reads empty: {cells:?}");
        assert!(cells.contains(&(2, 1, String::new())), "a retired receiver reads empty: {cells:?}");
    }

    #[test]
    fn a_spill_in_a_batch_is_in_the_delta() {
        let mut s = Session::from_sheets(&[sheet(&[])]);

        let delta = s
            .apply_many(&[InEdit { sheet: 0, row: 0, col: 0, raw: Some("={1;2;3}".into()) }])
            .unwrap();

        assert_eq!(
            reported(&delta),
            vec![(0, 0, "1".to_string()), (1, 0, "2".to_string()), (2, 0, "3".to_string())]
        );
    }

    // --- #91: spill anchors rewritten, blocked and unblocked ---------------

    #[test]
    fn rewriting_a_spill_anchor_reports_every_cell_it_vacated() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "=SEQUENCE(5)")])]);
        let delta = s.apply_one(0, 0, 0, Some("=SEQUENCE(2)")).unwrap();
        let cells = reported(&delta);
        for row in 2..5 {
            assert!(cells.contains(&(row, 0, String::new())), "A{} vacated: {cells:?}", row + 1);
        }
        assert!(cells.contains(&(0, 0, "1".to_string())) && cells.contains(&(1, 0, "2".to_string())), "{cells:?}");
    }

    #[test]
    fn clearing_a_spill_anchor_reports_every_cell_it_vacated() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "=SEQUENCE(3)")])]);
        let delta = s.apply_one(0, 0, 0, None).unwrap();
        let cells = reported(&delta);
        for row in 0..3 {
            assert!(cells.contains(&(row, 0, String::new())), "A{} vacated: {cells:?}", row + 1);
        }
    }

    #[test]
    fn rewriting_an_anchor_with_a_literal_reports_every_cell_it_vacated() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "=SEQUENCE(3)")])]);
        let delta = s.apply_one(0, 0, 0, Some("7")).unwrap();
        let cells = reported(&delta);
        assert!(cells.contains(&(0, 0, "7".to_string())), "{cells:?}");
        for row in 1..3 {
            assert!(cells.contains(&(row, 0, String::new())), "A{} vacated: {cells:?}", row + 1);
        }
    }

    #[test]
    fn typing_into_a_spill_blocks_it_and_clearing_the_blocker_respills() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "=SEQUENCE(3)"), (0, 1, "=A3")])]);
        let blocked = s.apply_one(0, 1, 0, Some("x")).unwrap();
        let cells = reported(&blocked);
        assert!(cells.contains(&(1, 0, "x".to_string())), "the typed value: {cells:?}");
        let anchor = blocked.cells.iter().find(|c| (c.row, c.col) == (0, 0)).expect("anchor reported");
        assert_eq!(anchor.error.as_deref(), Some("#SPILL!"), "{cells:?}");
        assert!(cells.contains(&(2, 0, String::new())), "the rest of the spill is vacated: {cells:?}");
        assert!(blocked.cells.iter().any(|c| (c.row, c.col) == (0, 1)), "a reader of a vacated cell: {cells:?}");

        let unblocked = s.apply_one(0, 1, 0, None).unwrap();
        let cells = reported(&unblocked);
        assert_eq!(
            [(0, 0), (1, 0), (2, 0)].map(|(r, c)| cells.iter().find(|x| (x.0, x.1) == (r, c)).map(|x| x.2.clone())),
            [Some("1".to_string()), Some("2".to_string()), Some("3".to_string())],
            "re-spilled: {cells:?}"
        );
        assert!(cells.contains(&(0, 1, "3".to_string())), "the reader sees the spill again: {cells:?}");
    }

    #[test]
    fn a_batch_that_blocks_and_rewrites_reports_the_settled_state() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "=SEQUENCE(4)")])]);
        let delta = s
            .apply_many(&[
                InEdit { sheet: 0, row: 0, col: 0, raw: Some("=SEQUENCE(2)".into()) },
                InEdit { sheet: 0, row: 3, col: 0, raw: Some("z".into()) },
            ])
            .unwrap();
        let cells = reported(&delta);
        assert!(cells.contains(&(2, 0, String::new())), "{cells:?}");
        assert!(cells.contains(&(3, 0, "z".to_string())), "{cells:?}");
        assert!(cells.contains(&(1, 0, "2".to_string())), "{cells:?}");
    }

    // --- structural operations ---------------------------------------------

    #[test]
    fn insert_rows_shifts_rewrites_and_reports_only_changed_values() {
        let mut s = Session::from_sheets(&[sheet(&[
            (0, 0, "1"),
            (1, 0, "2"),
            (2, 0, "=SUM(A1:A2)"),
            (3, 0, "=A3*10"),
        ])]);
        let d = s.structural(0, Axis::Row, 1, 2, false).unwrap();
        assert_eq!(d.shift, Some(Shift { sheet: 0, axis: "rows", at: 1, count: 2, delete: false }));
        // A3 moved to A5 and its range grew to A1:A4; A4 moved to A6.
        assert!(d.rewritten.contains(&Rewritten { sheet: 0, row: 4, col: 0, formula: "=SUM(A1:A4)".into() }), "{:?}", d.rewritten);
        assert!(d.rewritten.contains(&Rewritten { sheet: 0, row: 5, col: 0, formula: "=A5*10".into() }), "{:?}", d.rewritten);
        // Same values, only moved: nothing to repaint beyond the shift.
        assert!(d.cells.is_empty(), "{:?}", reported_structural(&d));
        assert_eq!(s.formulas_core(0, 0, 0, 10, 0).unwrap().len(), 2);
    }

    fn reported_structural(d: &StructuralDelta) -> Vec<(usize, usize, String)> {
        d.cells.iter().map(|c| (c.row, c.col, c.display.clone())).collect()
    }

    #[test]
    fn delete_rows_reports_ref_errors_and_new_values() {
        let mut s = Session::from_sheets(&[sheet(&[
            (0, 0, "1"),
            (1, 0, "2"),
            (2, 0, "3"),
            (3, 0, "=SUM(A1:A3)"),
            (4, 0, "=A2"),
        ])]);
        let d = s.structural(0, Axis::Row, 1, 1, true).unwrap();
        let cells = reported_structural(&d);
        assert!(cells.contains(&(2, 0, "4".into())), "SUM lost a row: {cells:?}");
        let gone = d.cells.iter().find(|c| (c.row, c.col) == (3, 0)).expect("the reference to a deleted row");
        assert_eq!(gone.error.as_deref(), Some("#REF!"), "{cells:?}");
    }

    #[test]
    fn column_ops_move_spills_with_their_anchor() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 1, "=SEQUENCE(1,3)")])]);
        let d = s.structural(0, Axis::Col, 0, 1, false).unwrap();
        assert!(d.cells.is_empty(), "the spill only moved: {:?}", reported_structural(&d));
        let all = s.all_results_core();
        assert!(all.cells.iter().any(|c| (c.row, c.col, c.display.as_str()) == (0, 2, "1")));
    }

    #[test]
    fn sheet_add_rename_delete_recalculate_what_named_them() {
        let mut s = Session::from_sheets(&[sheet(&[(0, 0, "=Data!A1+1")])]);
        let added = s.add_sheet_core(Some("Data")).unwrap();
        assert_eq!(added.sheet, Some(1));
        assert_eq!(s.apply_one(1, 0, 0, Some("41")).unwrap().cells.first().map(|c| c.sheet), Some(1));
        assert_eq!(s.all_results_core().cells[0].display, "42");

        let deleted = s.delete_sheet_core(1).unwrap();
        assert_eq!(deleted.rewritten.len(), 1);
        assert_eq!(deleted.rewritten[0].formula, "=#REF!+1");
        let first = deleted.cells.iter().find(|c| (c.sheet, c.row, c.col) == (0, 0, 0)).expect("A1 recalculated");
        assert!(first.error.is_some(), "a reference to a deleted sheet is an error: {:?}", reported_structural(&deleted));
        assert!(s.delete_sheet_core(0).is_err(), "the last sheet cannot go");
        assert!(s.rename_sheet_core(0, "Renamed").is_ok());
        assert!(s.rename_sheet_core(0, "   ").is_err(), "a blank name is refused");
        s.add_sheet_core(Some("Data")).unwrap();
        s.apply_one(1, 0, 0, Some("99")).unwrap();
        assert_eq!(s.wb.sheet(0).unwrap().get_display(0, 0), "#REF!");
    }

    #[test]
    fn sheet_rename_reports_source_rewrites_even_when_values_do_not_change() {
        let mut s = Session::from_sheets(&[sheet(&[
            (0, 0, "7"), (0, 1, "=( Sheet1!A1 + 1 )*2"),
            (1, 1, "=INDIRECT(\"Sheet1!A1\")"),
        ])]);
        s.add_sheet_core(Some("Summary")).unwrap();
        s.apply_one(1, 0, 0, Some("=Sheet1!A1")).unwrap();
        let before = s.revision();
        let d = s.rename_sheet_core(0, "New Data").unwrap();
        assert!(d.revision > before);
        assert_eq!(d.rewritten.len(), 2);
        assert_eq!((d.rewritten[0].sheet, d.rewritten[0].row, d.rewritten[0].col), (0, 0, 1));
        assert_eq!(d.rewritten[0].formula, "=( 'New Data'!A1 + 1 )*2");
        assert_eq!(d.rewritten[1].formula, "='New Data'!A1");
        assert_eq!(s.wb.sheet(0).unwrap().get_display(0, 1), "16");
        assert_eq!(s.wb.sheet(1).unwrap().get_display(0, 0), "7");
        assert!(d.cells.iter().any(|c| (c.sheet, c.row, c.col) == (0, 1, 1)
            && c.error.as_deref() == Some("#REF!")));
        let noop = s.rename_sheet_core(0, "New Data").unwrap();
        assert_eq!(noop.revision, d.revision);
        assert!(noop.rewritten.is_empty() && noop.cells.is_empty());
        assert!(s.rename_sheet_core(0, "Summary").is_err());
        assert_eq!(s.revision(), d.revision);
        s.apply_one(0, 0, 0, Some("8")).unwrap();
        assert_eq!(s.wb.sheet(0).unwrap().get_display(0, 1), "18");
        assert_eq!(s.wb.sheet(1).unwrap().get_display(0, 0), "8");
    }

    #[test]
    fn a_seeded_clock_makes_two_sessions_agree() {
        let doc = [sheet(&[(0, 0, "=RAND()"), (1, 0, "=NOW()")])];
        let run = || {
            let mut s = Session::from_sheets(&doc);
            s.set_clock_core(Some(1_791_028_800_000.0), Some(0), Some(7.0));
            s.apply_one(0, 5, 5, Some("x")).unwrap();
            s.all_results_core().cells.into_iter().map(|c| c.display).collect::<Vec<_>>()
        };
        assert_eq!(run(), run());
    }

    /// One-dependent insert_rows at 200k formulas vs rebuilding the session.
    /// `cargo test -p visigrid-engine-wasm --release -- --ignored --nocapture structural_bench`
    #[test]
    #[ignore]
    fn structural_bench() {
        let n = 200_000usize;
        let mut cells: Vec<(usize, usize, String)> = Vec::with_capacity(n + 1);
        cells.push((0, 0, "1".into()));
        for r in 1..n {
            cells.push((r, 0, format!("=A{}+1", r)));
        }
        let doc = InSheet {
            name: None,
            cells: cells.iter().map(|(row, col, raw)| InCell { row: *row, col: *col, raw: raw.clone() }).collect(),
        };
        let t = std::time::Instant::now();
        let mut s = Session::from_sheets(std::slice::from_ref(&doc));
        let build = t.elapsed();
        let t = std::time::Instant::now();
        let d = s.structural(0, Axis::Row, n + 10, 1, false).unwrap();
        let insert_far = t.elapsed();
        let t = std::time::Instant::now();
        let d2 = s.structural(0, Axis::Row, 5, 1, false).unwrap();
        let insert_near = t.elapsed();
        println!(
            "formulas={n} rebuild={:?} insert_below_all={:?} (rewrites {}) insert_near_top={:?} (rewrites {}, cells {})",
            build, insert_far, d.rewritten.len(), insert_near, d2.rewritten.len(), d2.cells.len()
        );
        let t = std::time::Instant::now();
        s.wb.structural_edit(0, Axis::Row, 5, 1, false).unwrap();
        let engine_only = t.elapsed();
        let t = std::time::Instant::now();
        let _ = s.values();
        println!("engine structural_edit alone={engine_only:?} one values() snapshot={:?}", t.elapsed());
    }
}
