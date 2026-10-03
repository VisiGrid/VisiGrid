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
}
