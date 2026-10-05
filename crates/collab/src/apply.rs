//! Applying operations to a real `visigrid_engine::Workbook`.
//!
//! Every replica — every client and the server — applies sequenced ops with
//! exactly this code, so this is where determinism lives. An op whose target
//! no longer exists (its sheet is gone) is a no-op on every replica alike.

use crate::op::{parse_hex_color, Axis, BorderLine, BorderSpec, CollabOp, FormatProps, HAlign, Rect, SheetKey, VAlign};
use visigrid_engine::cell::{Alignment, BorderStyle, CellBorder, CellFormat, NumberFormat, TextOverflow, VerticalAlignment};
use visigrid_engine::sheet::MergedRegion;
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

/// What applying ops changed, for a client mirroring the workbook in a UI.
///
/// Recording never affects what is applied: every replica applies the same
/// ops with the same code whether or not it records.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Changes {
    /// Cells written, then the cells their writes re-evaluated (spill
    /// receivers included), by stable sheet key, in order. May repeat.
    pub cells: Vec<(SheetKey, usize, usize)>,
    /// The change is not describable cell by cell (a structural or sheet op,
    /// a full recompute, a rebuild): the mirror must repaint everything.
    pub full: bool,
    /// Sheets were added, renamed or removed.
    pub sheets: bool,
    /// Structural ops applied, in order.
    pub structural: Vec<CollabOp>,
    /// Format ops applied, in order: the rectangle and the properties set or
    /// cleared (for a mirror to restyle incrementally).
    pub formats: Vec<(SheetKey, Rect, FormatProps)>,
    /// Line layout (sizes, hidden, frozen) or merges changed: the mirror
    /// re-reads the layout.
    pub layout: bool,
}

impl Changes {
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty() && !self.full && !self.sheets && self.structural.is_empty() && self.formats.is_empty() && !self.layout
    }

    fn recalculated(&mut self, r: visigrid_engine::workbook::Recalculated) {
        match r {
            visigrid_engine::workbook::Recalculated::Cells(ids) => {
                self.cells.extend(ids.into_iter().map(|id| (id.sheet.0, id.row, id.col)));
            }
            visigrid_engine::workbook::Recalculated::All => self.full = true,
        }
    }
}

/// A formatting rectangle larger than this is reported as `full` rather than
/// cell by cell.
const MAX_TRACKED_CELLS: usize = 10_000;

pub fn apply_op(wb: &mut Workbook, op: &CollabOp) -> Result<(), Skipped> {
    apply_op_tracked(wb, op, None)
}

/// [`apply_op`], recording what changed into `changes` when given.
pub fn apply_op_tracked(wb: &mut Workbook, op: &CollabOp, mut changes: Option<&mut Changes>) -> Result<(), Skipped> {
    match op {
        CollabOp::SetCell {
            sheet,
            row,
            col,
            content,
            ..
        } => {
            let idx = index_of(wb, *sheet)?;
            // Writes land on exactly this cell, even one hidden by a merge:
            // where merges are must not change what a positional op means.
            // Clear is "clear contents" (Excel's Delete key): the value
            // goes, the cell's format stays. The engine's clear_cell removes
            // the whole cell including its format, which would not commute
            // with a concurrent format change.
            let r = wb.set_cell_value_tracked_at(idx, *row, *col, content.raw());
            if let Some(ch) = changes.as_deref_mut() {
                ch.cells.push((*sheet, *row, *col));
                ch.recalculated(r);
            }
            Ok(())
        }
        CollabOp::SetBold { .. } | CollabOp::SetFormat { .. } => {
            let (sheet, rect, props) = op.format().expect("format op");
            let idx = index_of(wb, sheet)?;
            let id = wb.sheets()[idx].id;
            if let Some(ch) = changes.as_deref_mut() {
                ch.formats.push((sheet, rect, props.clone()));
                // Number formats change what a cell displays: report the cells.
                let area = (rect.r1.saturating_sub(rect.r0) + 1).saturating_mul(rect.c1.saturating_sub(rect.c0) + 1);
                if area > MAX_TRACKED_CELLS {
                    ch.full = true;
                } else {
                    for r in rect.r0..=rect.r1 {
                        for c in rect.c0..=rect.c1 {
                            ch.cells.push((sheet, r, c));
                        }
                    }
                }
            }
            for r in rect.r0..=rect.r1 {
                for c in rect.c0..=rect.c1 {
                    if let Some(s) = wb.sheet_mut(idx) {
                        let mut fmt = s.get_format(r, c);
                        apply_props(&mut fmt, &props);
                        s.set_format(r, c, fmt);
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
            let done = wb
                .structural_edit(idx, (*axis).into(), *at, *count, *delete)
                .map(|_| ())
                .map_err(Skipped::Refused);
            if let Some(ch) = changes.as_deref_mut() {
                if done.is_ok() {
                    ch.full = true;
                    ch.structural.push(op.clone());
                }
            }
            done
        }
        CollabOp::AddSheet { sheet, name, index } => {
            let s = Sheet::new_with_name(SheetId(*sheet), NUM_ROWS, NUM_COLS, name);
            let at = (*index).min(wb.sheets().len());
            if !wb.restore_sheet(at, s) {
                return Err(Skipped::Refused(format!("could not add sheet {name}")));
            }
            after_sheet_change(wb);
            sheet_changed(changes.as_deref_mut());
            Ok(())
        }
        CollabOp::RenameSheet { sheet, name } => {
            let idx = index_of(wb, *sheet)?;
            if !wb.rename_sheet(idx, name) {
                return Err(Skipped::Refused(format!("could not rename to {name}")));
            }
            after_sheet_change(wb);
            sheet_changed(changes.as_deref_mut());
            Ok(())
        }
        CollabOp::MoveSheet { sheet, index } => {
            let idx = index_of(wb, *sheet)?;
            let to = (*index).min(wb.sheets().len() - 1);
            if to != idx {
                let s = wb.take_sheet(idx).ok_or_else(|| Skipped::Refused("could not move sheet".into()))?;
                if !wb.restore_sheet(to, s) {
                    return Err(Skipped::Refused("could not move sheet".into()));
                }
                after_sheet_change(wb);
                sheet_changed(changes.as_deref_mut());
            }
            Ok(())
        }
        CollabOp::DeleteSheet { sheet, .. } => {
            let idx = index_of(wb, *sheet)?;
            if wb.take_sheet(idx).is_none() {
                return Err(Skipped::Refused("could not delete sheet".into()));
            }
            after_sheet_change(wb);
            sheet_changed(changes.as_deref_mut());
            Ok(())
        }
        CollabOp::SetLines { sheet, axis, lo, hi, props } => {
            let idx = index_of(wb, *sheet)?;
            let limit = if *axis == Axis::Row { NUM_ROWS } else { NUM_COLS };
            if let Some(s) = wb.sheet_mut(idx) {
                let l = &mut s.layout;
                let (sizes, hidden) = match axis {
                    Axis::Row => (&mut l.row_heights, &mut l.hidden_rows),
                    Axis::Col => (&mut l.col_widths, &mut l.hidden_cols),
                };
                for i in *lo..=(*hi).min(limit - 1) {
                    match props.size {
                        Some(Some(v)) => {
                            sizes.insert(i, v as f32);
                        }
                        Some(None) => {
                            sizes.remove(&i);
                        }
                        None => {}
                    }
                    match props.hidden {
                        Some(true) => {
                            hidden.insert(i);
                        }
                        Some(false) => {
                            hidden.remove(&i);
                        }
                        None => {}
                    }
                }
            }
            layout_changed(changes.as_deref_mut(), false);
            Ok(())
        }
        CollabOp::SetFreeze { sheet, rows, cols } => {
            let idx = index_of(wb, *sheet)?;
            if let Some(s) = wb.sheet_mut(idx) {
                s.layout.frozen_rows = *rows;
                s.layout.frozen_cols = *cols;
            }
            layout_changed(changes.as_deref_mut(), false);
            Ok(())
        }
        CollabOp::Merge { sheet, rect } => {
            let idx = index_of(wb, *sheet)?;
            let region = MergedRegion::new(rect.r0, rect.c0, rect.r1, rect.c1);
            let done = wb.sheet_mut(idx).map_or(Ok(()), |s| s.add_merge(region)).map_err(Skipped::Refused);
            if done.is_ok() {
                layout_changed(changes.as_deref_mut(), true);
            }
            done
        }
        CollabOp::Unmerge { sheet, rect } => {
            let idx = index_of(wb, *sheet)?;
            if let Some(s) = wb.sheet_mut(idx) {
                let starts: Vec<(usize, usize)> = s
                    .merged_regions
                    .iter()
                    .filter(|m| rect.intersects(&Rect::new(m.start.0, m.start.1, m.end.0, m.end.1)))
                    .map(|m| m.start)
                    .collect();
                for st in starts {
                    s.remove_merge(st);
                }
            }
            layout_changed(changes.as_deref_mut(), true);
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
                    let r = wb.set_cell_value_tracked_at(idx, row + dr, col + dc, content.raw());
                    if let Some(ch) = changes.as_deref_mut() {
                        ch.cells.push((*sheet, row + dr, col + dc));
                        ch.recalculated(r);
                    }
                }
            }
            Ok(())
        }
    }
}

/// Set or clear each property `props` names on `fmt`; a clear restores the
/// engine's default for that property.
pub fn apply_props(fmt: &mut CellFormat, props: &FormatProps) {
    let d = CellFormat::default();
    if let Some(v) = props.bold {
        fmt.bold = v.unwrap_or(d.bold);
    }
    if let Some(v) = props.italic {
        fmt.italic = v.unwrap_or(d.italic);
    }
    if let Some(v) = props.underline {
        fmt.underline = v.unwrap_or(d.underline);
    }
    if let Some(v) = props.strikethrough {
        fmt.strikethrough = v.unwrap_or(d.strikethrough);
    }
    if let Some(v) = &props.font_family {
        fmt.font_family = v.clone();
    }
    if let Some(v) = props.font_size {
        fmt.font_size = v.map(|s| s as f32);
    }
    if let Some(v) = &props.color {
        fmt.font_color = v.as_deref().and_then(parse_hex_color);
    }
    if let Some(v) = &props.background {
        fmt.background_color = v.as_deref().and_then(parse_hex_color);
    }
    if let Some(v) = &props.number_format {
        fmt.number_format = match v.as_deref() {
            None => d.number_format.clone(),
            Some(code) if code.is_empty() || code.eq_ignore_ascii_case("general") => NumberFormat::General,
            Some(code) => NumberFormat::Custom(code.to_string()),
        };
    }
    if let Some(v) = props.h_align {
        fmt.alignment = match v {
            None | Some(HAlign::General) => Alignment::General,
            Some(HAlign::Left) => Alignment::Left,
            Some(HAlign::Center) => Alignment::Center,
            Some(HAlign::Right) => Alignment::Right,
        };
    }
    if let Some(v) = props.v_align {
        fmt.vertical_alignment = match v {
            None => d.vertical_alignment,
            Some(VAlign::Top) => VerticalAlignment::Top,
            Some(VAlign::Middle) => VerticalAlignment::Middle,
            Some(VAlign::Bottom) => VerticalAlignment::Bottom,
        };
    }
    if let Some(v) = props.wrap {
        fmt.text_overflow = match v {
            Some(true) => TextOverflow::Wrap,
            Some(false) => TextOverflow::Clip,
            None => d.text_overflow,
        };
    }
    for (prop, edge) in [
        (&props.border_top, &mut fmt.border_top),
        (&props.border_right, &mut fmt.border_right),
        (&props.border_bottom, &mut fmt.border_bottom),
        (&props.border_left, &mut fmt.border_left),
    ] {
        if let Some(v) = prop {
            *edge = v.as_ref().map(cell_border).unwrap_or_default();
        }
    }
}

/// A border edge as the engine stores it.
pub fn cell_border(b: &BorderSpec) -> CellBorder {
    CellBorder {
        style: match b.style {
            BorderLine::Thin => BorderStyle::Thin,
            BorderLine::Medium => BorderStyle::Medium,
            BorderLine::Thick => BorderStyle::Thick,
        },
        color: b.color.as_deref().and_then(parse_hex_color),
    }
}

/// An engine border edge as op properties (None when the edge has no line).
pub fn border_spec(b: &CellBorder) -> Option<BorderSpec> {
    let style = match b.style {
        BorderStyle::None => return None,
        BorderStyle::Thin => BorderLine::Thin,
        BorderStyle::Medium => BorderLine::Medium,
        BorderStyle::Thick => BorderLine::Thick,
    };
    Some(BorderSpec { style, color: b.color.map(|c| format!("#{:02X}{:02X}{:02X}", c[0], c[1], c[2])) })
}

fn layout_changed(changes: Option<&mut Changes>, cells: bool) {
    if let Some(ch) = changes {
        ch.layout = true;
        if cells {
            ch.full = true;
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

/// [`apply_ops`], recording what changed.
pub fn apply_ops_tracked(wb: &mut Workbook, ops: &[CollabOp], changes: &mut Changes) -> Vec<Skipped> {
    ops.iter().filter_map(|op| apply_op_tracked(wb, op, Some(changes)).err()).collect()
}

fn sheet_changed(changes: Option<&mut Changes>) {
    if let Some(ch) = changes {
        ch.full = true;
        ch.sheets = true;
    }
}

/// Everything convergence compares: tab order, ids and names, then every
/// non-empty cell's raw text, cached computed display, and full format (as
/// JSON, empty for the default format).
#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize)]
pub struct Fingerprint {
    pub sheets: Vec<SheetPrint>,
}

#[derive(Clone, Debug, PartialEq, Eq, serde::Serialize)]
pub struct SheetPrint {
    pub id: u64,
    pub name: String,
    /// (row, col, raw, computed, format), sorted by position.
    pub cells: Vec<(usize, usize, String, String, String)>,
    /// Line layout as JSON (empty when default) and merges, sorted.
    pub layout: String,
    pub merges: Vec<(usize, usize, usize, usize)>,
}

pub fn fingerprint(wb: &Workbook) -> Fingerprint {
    let sheets = wb
        .sheets()
        .iter()
        .map(|s| {
            let default = CellFormat::default();
            let mut cells: Vec<(usize, usize, String, String, String)> = s
                .cells_iter()
                .map(|((r, c), _)| {
                    let fmt = s.get_format(r, c);
                    let fmt = if fmt == default {
                        String::new()
                    } else {
                        serde_json::to_string(&fmt).expect("format serializes")
                    };
                    (r, c, s.get_raw(r, c), s.get_formatted_display(r, c), fmt)
                })
                .filter(|(_, _, raw, shown, fmt)| !raw.is_empty() || !shown.is_empty() || !fmt.is_empty())
                .collect();
            cells.sort_by_key(|(r, c, ..)| (*r, *c));
            let layout = if s.layout.is_default() {
                String::new()
            } else {
                serde_json::to_string(&s.layout).expect("layout serializes")
            };
            let mut merges: Vec<(usize, usize, usize, usize)> =
                s.merged_regions.iter().map(|m| (m.start.0, m.start.1, m.end.0, m.end.1)).collect();
            merges.sort_unstable();
            SheetPrint {
                id: s.id.0,
                name: s.name.clone(),
                cells,
                layout,
                merges,
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

/// Workbooks over this many cells are not checksummed in collaboration: the
/// fingerprint costs O(cells) per call (5 s and 900 MB of WASM memory at
/// 300,000 x 20), and the host computes one per operation. The same as the
/// size at which sheets are stored as bands, whose content-addressed keys
/// compare snapshots instead. An incremental checksum is follow-up work.
pub const CHECKSUM_CELL_LIMIT: usize = 200_000;

/// Whether `wb` is over [`CHECKSUM_CELL_LIMIT`] (counts at most limit + 1 cells).
pub fn too_large_to_checksum(wb: &Workbook) -> bool {
    let mut n = 0usize;
    for s in wb.sheets() {
        n += s.cells_iter().take(CHECKSUM_CELL_LIMIT + 1 - n).count();
        if n > CHECKSUM_CELL_LIMIT {
            return true;
        }
    }
    false
}

/// [`checksum`], or empty for a workbook over [`CHECKSUM_CELL_LIMIT`]: what
/// the host reports and clients compare (an empty checksum is not compared).
pub fn collab_checksum(wb: &Workbook) -> String {
    if too_large_to_checksum(wb) {
        String::new()
    } else {
        checksum(wb)
    }
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
        if sa.layout != sb.layout || sa.merges != sb.merges {
            return Some(format!("sheet {}: layout {} vs {}, merges {:?} vs {:?}", sa.name, sa.layout, sb.layout, sa.merges, sb.merges));
        }
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
