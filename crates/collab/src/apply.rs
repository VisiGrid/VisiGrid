//! Applying operations to a real `visigrid_engine::Workbook`.
//!
//! Every replica — every client and the server — applies sequenced ops with
//! exactly this code, so this is where determinism lives. An op whose target
//! no longer exists (its sheet is gone) is a no-op on every replica alike.

use crate::op::{CollabOp, SheetKey};
use visigrid_engine::cell_id::CellId;
use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};
use visigrid_engine::workbook::Workbook;

/// Why an op had no effect. Informational only: the outcome is identical on
/// every replica either way.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Skipped {
    NoSuchSheet(SheetKey),
    Refused(String),
}

fn index_of(wb: &Workbook, sheet: SheetKey) -> Result<usize, Skipped> {
    wb.idx_for_sheet_id(SheetId(sheet))
        .ok_or(Skipped::NoSuchSheet(sheet))
}

/// After a sheet is added, renamed or removed, formulas that name sheets can
/// resolve differently: rebuild dependencies and recalculate everything.
fn after_sheet_change(wb: &mut Workbook) {
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
}

pub fn apply_op(wb: &mut Workbook, op: &CollabOp) -> Result<(), Skipped> {
    match op {
        CollabOp::SetCell {
            sheet,
            row,
            col,
            content,
            ..
        } => {
            let idx = index_of(wb, *sheet)?;
            // Clear is "clear contents" (Excel's Delete key): the value
            // goes, the cell's format stays. The engine's clear_cell removes
            // the whole cell including its format, which would not commute
            // with a concurrent format change.
            wb.set_cell_value_tracked(idx, *row, *col, content.raw());
            Ok(())
        }
        CollabOp::SetBold { sheet, rect, bold } => {
            let idx = index_of(wb, *sheet)?;
            let id = wb.sheets()[idx].id;
            for r in rect.r0..=rect.r1 {
                for c in rect.c0..=rect.c1 {
                    if let Some(s) = wb.sheet_mut(idx) {
                        s.set_bold(r, c, *bold);
                    }
                    wb.note_format_changed(CellId::new(id, r, c));
                }
            }
            Ok(())
        }
        CollabOp::Structural {
            sheet,
            axis,
            at,
            count,
            delete,
            ..
        } => {
            let idx = index_of(wb, *sheet)?;
            wb.structural_edit(idx, (*axis).into(), *at, *count, *delete)
                .map(|_| ())
                .map_err(Skipped::Refused)
        }
        CollabOp::AddSheet { sheet, name, index } => {
            let s = Sheet::new_with_name(SheetId(*sheet), NUM_ROWS, NUM_COLS, name);
            let at = (*index).min(wb.sheets().len());
            if !wb.restore_sheet(at, s) {
                return Err(Skipped::Refused(format!("could not add sheet {name}")));
            }
            after_sheet_change(wb);
            Ok(())
        }
        CollabOp::RenameSheet { sheet, name } => {
            let idx = index_of(wb, *sheet)?;
            if !wb.rename_sheet(idx, name) {
                return Err(Skipped::Refused(format!("could not rename to {name}")));
            }
            after_sheet_change(wb);
            Ok(())
        }
        CollabOp::DeleteSheet { sheet, .. } => {
            let idx = index_of(wb, *sheet)?;
            if wb.take_sheet(idx).is_none() {
                return Err(Skipped::Refused("could not delete sheet".into()));
            }
            after_sheet_change(wb);
            Ok(())
        }
        CollabOp::ReplaceRange {
            sheet,
            row,
            col,
            values,
        } => {
            let idx = index_of(wb, *sheet)?;
            for (dr, line) in values.iter().enumerate() {
                for (dc, content) in line.iter().enumerate() {
                    wb.set_cell_value_tracked(idx, row + dr, col + dc, content.raw());
                }
            }
            Ok(())
        }
    }
}

/// Apply a list in order. Skips are collected, never fatal.
/// Keep only the ops that can apply to `wb`: the target sheet exists, an
/// added sheet's id and name are free, a rename's name is free (names
/// compare with the engine's own `normalize_sheet_name`), and a delete does
/// not remove the last sheet: exactly what the engine itself refuses. Ops in an
/// envelope are sequential, so earlier adds, renames and deletes in the list
/// count. The sequencer runs this on its own replica before sequencing, so
/// an op that cannot apply is dropped there and never reaches any client;
/// kept, it would be a no-op everywhere but still shift positions and
/// rewrite names in the transforms of every concurrent op. Clients run it
/// when they rebuild, for the same reason.
/// Returns the kept ops and the reason for each dropped one.
pub fn filter_unappliable(wb: &Workbook, ops: &[CollabOp]) -> (Vec<CollabOp>, Vec<&'static str>) {
    let key = visigrid_engine::sheet::normalize_sheet_name;
    let mut sheets: Vec<(SheetKey, String)> = wb.sheets().iter().map(|s| (s.id.0, key(&s.name))).collect();
    let mut kept = Vec::with_capacity(ops.len());
    let mut dropped = Vec::new();
    for op in ops {
        let exists = |sheets: &Vec<(SheetKey, String)>, id: SheetKey| sheets.iter().any(|(s, _)| *s == id);
        let taken = |sheets: &Vec<(SheetKey, String)>, name: &str, except: SheetKey| {
            sheets.iter().any(|(s, n)| *s != except && *n == key(name))
        };
        match op {
            CollabOp::AddSheet { sheet, name, .. } => {
                if exists(&sheets, *sheet) || taken(&sheets, name, *sheet) || key(name).is_empty() {
                    dropped.push(NAME_TAKEN);
                    continue;
                }
                sheets.push((*sheet, key(name)));
            }
            _ if !exists(&sheets, op.sheet()) => {
                dropped.push(NO_SUCH_SHEET);
                continue;
            }
            CollabOp::RenameSheet { sheet, name } => {
                if taken(&sheets, name, *sheet) || key(name).is_empty() {
                    dropped.push(NAME_TAKEN);
                    continue;
                }
                if let Some(entry) = sheets.iter_mut().find(|(s, _)| s == sheet) {
                    entry.1 = key(name);
                }
            }
            CollabOp::DeleteSheet { sheet, .. } => {
                if sheets.len() <= 1 {
                    dropped.push(LAST_SHEET);
                    continue;
                }
                sheets.retain(|(s, _)| s != sheet);
            }
            _ => {}
        }
        kept.push(op.clone());
    }
    (kept, dropped)
}

/// Backwards-compatible count form of [`filter_unappliable`].
pub fn filter_missing_sheets(wb: &Workbook, ops: &[CollabOp]) -> (Vec<CollabOp>, usize) {
    let (kept, dropped) = filter_unappliable(wb, ops);
    (kept, dropped.len())
}

/// Drop reason: an added or renamed sheet's name (or id) is in use.
pub const NAME_TAKEN: &str = "name_taken";
/// Drop reason: deleting the only sheet.
pub const LAST_SHEET: &str = "last_sheet";

/// The drop reason for ops whose sheet does not exist.
pub const NO_SUCH_SHEET: &str = "no_such_sheet";

pub fn apply_ops(wb: &mut Workbook, ops: &[CollabOp]) -> Vec<Skipped> {
    ops.iter().filter_map(|op| apply_op(wb, op).err()).collect()
}

/// Everything convergence compares: tab order, ids and names, then every
/// non-empty cell's raw text, cached computed display, and bold flag.
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize)]
pub struct Fingerprint {
    pub sheets: Vec<SheetPrint>,
}

#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize)]
pub struct SheetPrint {
    pub id: u64,
    pub name: String,
    /// (row, col, raw, computed, bold), sorted by position.
    pub cells: Vec<(usize, usize, String, String, bool)>,
}

pub fn fingerprint(wb: &Workbook) -> Fingerprint {
    let sheets = wb
        .sheets()
        .iter()
        .map(|s| {
            let mut cells: Vec<(usize, usize, String, String, bool)> = s
                .cells_iter()
                .map(|((r, c), _)| {
                    (
                        r,
                        c,
                        s.get_raw(r, c),
                        s.get_formatted_display(r, c),
                        s.get_format(r, c).bold,
                    )
                })
                .filter(|(_, _, raw, shown, bold)| !raw.is_empty() || !shown.is_empty() || *bold)
                .collect();
            cells.sort_by_key(|(r, c, ..)| (*r, *c));
            SheetPrint {
                id: s.id.0,
                name: s.name.clone(),
                cells,
            }
        })
        .collect();
    Fingerprint { sheets }
}

/// The convergence checksum the server publishes (protocol v2 `checksum`):
/// lowercase hex SHA-256 of the fingerprint's JSON form. Field order is fixed
/// by the struct definitions, so every build computes the same bytes.
pub fn checksum(wb: &Workbook) -> String {
    checksum_of(&fingerprint(wb))
}

pub fn checksum_of(print: &Fingerprint) -> String {
    use sha2::{Digest, Sha256};
    let bytes = serde_json::to_vec(print).expect("fingerprint serializes");
    format!("{:x}", Sha256::digest(&bytes))
}

/// First difference between two fingerprints, for failure reports.
pub fn first_difference(a: &Fingerprint, b: &Fingerprint) -> Option<String> {
    if a.sheets.len() != b.sheets.len() {
        return Some(format!(
            "sheet count {} vs {}: {:?} vs {:?}",
            a.sheets.len(),
            b.sheets.len(),
            a.sheets.iter().map(|s| &s.name).collect::<Vec<_>>(),
            b.sheets.iter().map(|s| &s.name).collect::<Vec<_>>()
        ));
    }
    for (sa, sb) in a.sheets.iter().zip(&b.sheets) {
        if sa.id != sb.id || sa.name != sb.name {
            return Some(format!(
                "sheet ({}, {}) vs ({}, {})",
                sa.id, sa.name, sb.id, sb.name
            ));
        }
        let mut i = 0;
        let mut j = 0;
        while i < sa.cells.len() || j < sb.cells.len() {
            match (sa.cells.get(i), sb.cells.get(j)) {
                (Some(x), Some(y)) if x == y => {
                    i += 1;
                    j += 1;
                }
                (x, y) => return Some(format!("sheet {}: {:?} vs {:?}", sa.name, x, y)),
            }
        }
    }
    None
}
