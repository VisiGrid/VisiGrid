//! Formula inspection for the web grid's inspector: a cell's precedents
//! (as its formula names them: cells and ranges) and dependents from the
//! engine's dependency graph, the whole chain either way on request, and a
//! what-if (the dependents' new values if an input changed, computed on a
//! copy-on-write copy of the workbook, never applied).

use std::collections::{BTreeSet, HashSet, VecDeque};

use serde_json::{json, Value};
use visigrid_engine::cell_id::CellId;
use visigrid_engine::sheet::SheetId;
use visigrid_engine::workbook::Workbook;

/// Cells listed in a chain at most (a wide model stays readable).
pub(crate) const CHAIN_LIMIT: usize = 500;

fn cell_json(wb: &Workbook, c: CellId) -> Option<Value> {
    let idx = wb.idx_for_sheet_id(c.sheet)?;
    let s = &wb.sheets()[idx];
    Some(json!({
        "sheet": c.sheet.0, "sheet_name": s.name, "row": c.row, "col": c.col,
        "raw": s.get_raw(c.row, c.col), "shown": s.get_formatted_display(c.row, c.col),
    }))
}

/// Single references first, then ranges, each once, in a stable order.
fn direct_precedents(wb: &Workbook, cell: CellId) -> Vec<Value> {
    let g = wb.dep_graph();
    let mut singles: Vec<CellId> = g.precedents(cell).collect();
    singles.sort_by_key(|c| (c.sheet.0, c.row, c.col));
    let mut out: Vec<Value> = singles.into_iter().filter_map(|c| cell_json(wb, c).map(|mut v| { v["kind"] = json!("cell"); v })).collect();
    let mut seen = BTreeSet::new();
    for r in g.precedent_ranges(cell) {
        if !seen.insert((r.sheet.0, r.start_row, r.start_col, r.end_row, r.end_col)) {
            continue;
        }
        let Some(idx) = wb.idx_for_sheet_id(r.sheet) else { continue };
        out.push(json!({
            "kind": "range", "sheet": r.sheet.0, "sheet_name": wb.sheets()[idx].name,
            "r0": r.start_row, "c0": r.start_col, "r1": r.end_row, "c1": r.end_col,
        }));
    }
    out
}

fn direct_dependents(wb: &Workbook, cell: CellId) -> Vec<CellId> {
    let mut v: Vec<CellId> = wb.dep_graph().dependents(cell).collect();
    v.sort_by_key(|c| (c.sheet.0, c.row, c.col));
    v.dedup();
    v
}

/// Every cell `cell` leads to (dependents) or comes from (precedents,
/// through formulas only), breadth first with each one's distance; at most
/// `CHAIN_LIMIT`.
fn chain(wb: &Workbook, cell: CellId, up: bool) -> (Vec<Value>, bool) {
    let mut seen: HashSet<CellId> = HashSet::from([cell]);
    let mut queue: VecDeque<(CellId, usize)> = VecDeque::from([(cell, 0)]);
    let mut out = Vec::new();
    let mut truncated = false;
    while let Some((c, depth)) = queue.pop_front() {
        let next: Vec<CellId> = if up { wb.get_precedents(c.sheet, c.row, c.col) } else { direct_dependents(wb, c) };
        for n in next {
            if !seen.insert(n) {
                continue;
            }
            if out.len() >= CHAIN_LIMIT {
                truncated = true;
                break;
            }
            if let Some(mut v) = cell_json(wb, n) {
                v["depth"] = json!(depth + 1);
                out.push(v);
            }
            queue.push_back((n, depth + 1));
        }
    }
    (out, truncated)
}

/// What the inspector shows for a cell. `full`: the whole chain both ways.
pub(crate) fn inspect(wb: &Workbook, sheet: u64, row: usize, col: usize, full: bool) -> Option<Value> {
    let id = SheetId(sheet);
    wb.idx_for_sheet_id(id)?;
    let cell = CellId::new(id, row, col);
    let mut v = cell_json(wb, cell)?;
    v["precedents"] = json!(direct_precedents(wb, cell));
    v["dependents"] = json!(direct_dependents(wb, cell).into_iter().filter_map(|c| cell_json(wb, c)).collect::<Vec<_>>());
    if full {
        let (up, t1) = chain(wb, cell, true);
        let (down, t2) = chain(wb, cell, false);
        v["chain"] = json!({"precedents": up, "dependents": down, "truncated": t1 || t2});
    }
    Some(v)
}

/// The dependents (whole chain) whose shown value would change if `raw`
/// were typed into the cell, before and after; nothing is applied.
pub(crate) fn what_if(wb: &Workbook, sheet: u64, row: usize, col: usize, raw: &str) -> Option<Value> {
    let id = SheetId(sheet);
    let idx = wb.idx_for_sheet_id(id)?;
    let cell = CellId::new(id, row, col);
    let (down, truncated) = chain(wb, cell, false);
    let mut copy = wb.clone();
    copy.set_cell_value_tracked_at(idx, row, col, raw);
    copy.recompute_full_ordered();
    let mut changes = Vec::new();
    for d in &down {
        let (s, r, c) = (SheetId(d["sheet"].as_u64()?), d["row"].as_u64()? as usize, d["col"].as_u64()? as usize);
        let i = copy.idx_for_sheet_id(s)?;
        let after = copy.sheets()[i].get_formatted_display(r, c);
        if after != d["shown"].as_str().unwrap_or("") {
            changes.push(json!({"sheet": s.0, "sheet_name": d["sheet_name"], "row": r, "col": c, "before": d["shown"], "after": after}));
        }
    }
    let shown = copy.sheets()[idx].get_formatted_display(row, col);
    Some(json!({"shown": shown, "changes": changes, "truncated": truncated}))
}

#[cfg(test)]
mod tests {
    use super::*;
    use visigrid_engine::sheet::{Sheet, NUM_COLS, NUM_ROWS};

    /// Inputs B1:B3 → subtotal B4 = SUM(B1:B3) → total B6 = B4*(1+B5),
    /// and Summary!A1 = Model!B6.
    fn model() -> Workbook {
        let mut m = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
        m.set_name("Model");
        for (r, v) in [(0, "100"), (1, "200"), (2, "300"), (4, "0.1")] {
            m.set_value_deferred(r, 1, v);
        }
        m.set_value_deferred(3, 1, "=SUM(B1:B3)");
        m.set_value_deferred(5, 1, "=B4*(1+B5)");
        let mut s = Sheet::new(SheetId(2), NUM_ROWS, NUM_COLS);
        s.set_name("Summary");
        s.set_value_deferred(0, 0, "=Model!B6");
        let mut wb = Workbook::from_sheets(vec![m, s], 0);
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
        wb
    }

    #[test]
    fn direct_precedents_keep_ranges_and_dependents_cross_sheets() {
        let wb = model();
        let total = inspect(&wb, 1, 5, 1, false).unwrap();
        assert_eq!(total["shown"], "660");
        let p: Vec<(String, u64)> = total["precedents"].as_array().unwrap().iter().map(|x| (x["kind"].as_str().unwrap().to_string(), x["row"].as_u64().unwrap_or(99))).collect();
        assert_eq!(p, [("cell".to_string(), 3), ("cell".to_string(), 4)], "B4 and B5");
        let deps = total["dependents"].as_array().unwrap();
        assert_eq!((deps.len(), deps[0]["sheet_name"].as_str()), (1, Some("Summary")));

        let sub = inspect(&wb, 1, 3, 1, false).unwrap();
        let range = &sub["precedents"][0];
        assert_eq!((range["kind"].as_str(), range["r0"].as_u64(), range["r1"].as_u64()), (Some("range"), Some(0), Some(2)));
        // An input's dependents come through the range.
        let input = inspect(&wb, 1, 1, 1, false).unwrap();
        assert_eq!(input["dependents"][0]["row"], 3);
    }

    #[test]
    fn the_full_chain_goes_both_ways_with_depths() {
        let wb = model();
        let input = inspect(&wb, 1, 0, 1, true).unwrap();
        let down: Vec<(u64, u64, u64)> = input["chain"]["dependents"].as_array().unwrap().iter().map(|x| (x["sheet"].as_u64().unwrap(), x["row"].as_u64().unwrap(), x["depth"].as_u64().unwrap())).collect();
        assert_eq!(down, [(1, 3, 1), (1, 5, 2), (2, 0, 3)]);
        let summary = inspect(&wb, 2, 0, 0, true).unwrap();
        assert_eq!(summary["chain"]["precedents"].as_array().unwrap().len(), 6, "B6, B4, B5, B1..B3");
    }

    #[test]
    fn what_if_previews_the_chain_and_changes_nothing() {
        let wb = model();
        let w = what_if(&wb, 1, 4, 1, "0.5").unwrap();
        let changes: Vec<(u64, &str, &str)> = w["changes"].as_array().unwrap().iter().map(|x| (x["row"].as_u64().unwrap(), x["before"].as_str().unwrap(), x["after"].as_str().unwrap())).collect();
        assert_eq!(changes, [(5, "660", "900"), (0, "660", "900")]);
        assert_eq!(wb.sheets()[0].get_formatted_display(5, 1), "660", "the real workbook is untouched");
    }
}
