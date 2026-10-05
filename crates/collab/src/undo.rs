//! Per-user undo (spec §Undo).
//!
//! Undo never restores a snapshot. When a user makes an edit, the client
//! records its *inverse* against the state the edit applied to, plus what the
//! edit left in each cell it wrote (the "expectation"). Every op applied to
//! the optimistic state afterwards, the user's own or anyone else's,
//! transforms the recorded entries forward like any concurrent op, so an
//! entry always describes the current document. Undo submits the inverse as
//! a new, ordinary operation, so convergence needs nothing new: the server
//! sequences and transforms it like any other edit.
//!
//! Before submitting, cells someone else changed since (the content or
//! format property is no longer what this user wrote) are left alone: undo
//! reverts the user's own change, never a collaborator's later one. An entry
//! whose target is gone (its row, column or sheet was deleted) transforms to
//! nothing, and undo reports that it no longer applies.
//!
//! Known limits (V1): restoring a deleted row, column or sheet brings back
//! its cells' contents and the format properties this vocabulary carries,
//! not borders, comments, merges or other sheets' references that became
//! `#REF!`. Undoing an edit after undoing a later row delete can lose cells
//! in that band (the entry was transformed past the delete first).

use std::collections::HashSet;

use uuid::Uuid;
use visigrid_engine::cell::{Alignment, CellFormat, DateStyle, NumberFormat, TextOverflow, VerticalAlignment};
use visigrid_engine::workbook::Workbook;

use crate::apply::{apply_ops, apply_props};
use crate::apply::border_spec;
use crate::op::{Axis, CellContent, CollabOp, FormatProps, HAlign, LineProps, Rect, SheetKey, VAlign};

/// Cells per format rectangle that undo inspects one by one. Larger
/// rectangles restore their non-default cells sparsely and skip the guard.
const MAX_INSPECTED_CELLS: usize = 10_000;

/// One undoable (or redoable) user action.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct UndoEntry {
    /// The envelope that made the change (a refused envelope's entry goes).
    pub origin: Uuid,
    /// Ops that revert the change, in the current frame.
    pub inverse: Vec<CollabOp>,
    /// What the change left behind, in the current frame: the content of
    /// every cell it wrote and the format properties it set. Only cells and
    /// properties that still match are reverted.
    pub expect: Vec<CollabOp>,
}

impl UndoEntry {
    /// The change left every cell it wrote as it was (an editor writing the
    /// same value twice): there is nothing to undo.
    pub fn is_noop(&self) -> bool {
        !self.inverse.is_empty()
            && self.inverse.iter().all(|op| match op {
                CollabOp::SetCell { sheet, row, col, content, .. } => self.expect.iter().any(|e| {
                    matches!(e, CollabOp::SetCell { sheet: s, row: r, col: c, content: x, .. }
                        if s == sheet && r == row && c == col && x == content)
                }),
                _ => false,
            })
    }
}

/// The result of an undo or redo request.
#[derive(Clone, Debug, Default, PartialEq, Eq, serde::Serialize)]
pub struct UndoOutcome {
    /// Ops were submitted.
    pub applied: bool,
    /// Cells left alone because someone else changed them since.
    pub kept_others: usize,
    /// Nothing to do: the stack was empty (`"empty"`) or the change no
    /// longer applies (`"gone"`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reason: Option<&'static str>,
}

fn content_of(raw: String) -> CellContent {
    if raw.is_empty() {
        CellContent::Clear
    } else if raw.starts_with('=') {
        CellContent::Formula(raw)
    } else {
        CellContent::Value(raw)
    }
}

/// What a cell holds, as the write that puts it back: text stays text, so
/// undoing over "007" restores the text, never the number 7.
fn content_at(s: &visigrid_engine::sheet::Sheet, row: usize, col: usize) -> CellContent {
    let raw = s.get_raw(row, col);
    if !raw.is_empty() && matches!(s.get_cell(row, col).value, visigrid_engine::cell::CellValue::Text(_)) {
        return CellContent::Text(raw);
    }
    content_of(raw)
}

fn index_of(wb: &Workbook, sheet: SheetKey) -> Option<usize> {
    wb.idx_for_sheet_id(visigrid_engine::sheet::SheetId(sheet))
}

fn hex(c: [u8; 4]) -> String {
    format!("#{:02X}{:02X}{:02X}", c[0], c[1], c[2])
}

/// An engine number format as the Excel code `FormatProps` carries.
pub fn number_format_code(nf: &NumberFormat) -> Option<String> {
    let decimals = |d: u8| if d == 0 { String::new() } else { format!(".{}", "0".repeat(d as usize)) };
    Some(match nf {
        NumberFormat::General => return None,
        NumberFormat::Custom(code) => code.clone(),
        NumberFormat::Number { decimals: d, thousands, .. } => {
            format!("{}{}", if *thousands { "#,##0" } else { "0" }, decimals(*d))
        }
        NumberFormat::Currency { decimals: d, thousands, symbol, .. } => format!(
            "{}{}{}",
            symbol.as_deref().unwrap_or("$"),
            if *thousands { "#,##0" } else { "0" },
            decimals(*d)
        ),
        NumberFormat::Percent { decimals: d } => format!("0{}%", decimals(*d)),
        NumberFormat::Date { style } => match style {
            DateStyle::Short => "m/d/yyyy".into(),
            DateStyle::Long => "mmmm d, yyyy".into(),
            DateStyle::Iso => "yyyy-mm-dd".into(),
        },
        NumberFormat::Time => "h:mm:ss".into(),
        NumberFormat::DateTime => "m/d/yyyy h:mm:ss".into(),
    })
}

/// The values `fmt` has for the properties `mask` names (every property when
/// `mask` is `None`), as props that would set them.
pub fn props_of(fmt: &CellFormat, mask: Option<&FormatProps>) -> FormatProps {
    let want = |m: fn(&FormatProps) -> bool| mask.is_none_or(m);
    let mut p = FormatProps::default();
    if want(|m| m.bold.is_some()) {
        p.bold = Some(Some(fmt.bold));
    }
    if want(|m| m.italic.is_some()) {
        p.italic = Some(Some(fmt.italic));
    }
    if want(|m| m.underline.is_some()) {
        p.underline = Some(Some(fmt.underline));
    }
    if want(|m| m.strikethrough.is_some()) {
        p.strikethrough = Some(Some(fmt.strikethrough));
    }
    if want(|m| m.font_family.is_some()) {
        p.font_family = Some(fmt.font_family.clone());
    }
    if want(|m| m.font_size.is_some()) {
        p.font_size = Some(fmt.font_size.map(f64::from));
    }
    if want(|m| m.color.is_some()) {
        p.color = Some(fmt.font_color.map(hex));
    }
    if want(|m| m.background.is_some()) {
        p.background = Some(fmt.background_color.map(hex));
    }
    if want(|m| m.number_format.is_some()) {
        p.number_format = Some(number_format_code(&fmt.number_format));
    }
    if want(|m| m.h_align.is_some()) {
        p.h_align = Some(match fmt.alignment {
            Alignment::General => None,
            Alignment::Left => Some(HAlign::Left),
            // Center across selection has no op form; centre is closest.
            Alignment::Center | Alignment::CenterAcrossSelection => Some(HAlign::Center),
            Alignment::Right => Some(HAlign::Right),
        });
    }
    if want(|m| m.v_align.is_some()) {
        p.v_align = Some(Some(match fmt.vertical_alignment {
            VerticalAlignment::Top => VAlign::Top,
            VerticalAlignment::Middle => VAlign::Middle,
            VerticalAlignment::Bottom => VAlign::Bottom,
        }));
    }
    if want(|m| m.wrap.is_some()) {
        p.wrap = Some(match fmt.text_overflow {
            TextOverflow::Wrap => Some(true),
            TextOverflow::Clip => Some(false),
            TextOverflow::Overflow => None,
        });
    }
    if want(|m| m.border_top.is_some()) {
        p.border_top = Some(border_spec(&fmt.border_top));
    }
    if want(|m| m.border_right.is_some()) {
        p.border_right = Some(border_spec(&fmt.border_right));
    }
    if want(|m| m.border_bottom.is_some()) {
        p.border_bottom = Some(border_spec(&fmt.border_bottom));
    }
    if want(|m| m.border_left.is_some()) {
        p.border_left = Some(border_spec(&fmt.border_left));
    }
    p
}

/// The inverse of `SetLines` against `wb`: clear what it sets over the run,
/// then put back each line that had its own size or was hidden.
fn invert_lines(wb: &Workbook, sheet: SheetKey, axis: Axis, lo: usize, hi: usize, props: &LineProps) -> Vec<CollabOp> {
    let Some(idx) = index_of(wb, sheet) else { return Vec::new() };
    let l = &wb.sheets()[idx].layout;
    let (sizes, hidden) = match axis {
        Axis::Row => (&l.row_heights, &l.hidden_rows),
        Axis::Col => (&l.col_widths, &l.hidden_cols),
    };
    let cleared = LineProps {
        size: props.size.map(|_| None),
        hidden: props.hidden.map(|_| false),
    };
    let mut out = vec![CollabOp::SetLines { sheet, axis, lo, hi, props: cleared }];
    let mut lines: Vec<usize> = Vec::new();
    if props.size.is_some() {
        lines.extend(sizes.range(lo..=hi).map(|(k, _)| *k));
    }
    if props.hidden.is_some() {
        lines.extend(hidden.range(lo..=hi).copied());
    }
    lines.sort_unstable();
    lines.dedup();
    let mut run: Option<(usize, usize, LineProps)> = None;
    for i in lines {
        let p = LineProps {
            size: props.size.map(|_| sizes.get(&i).map(|v| *v as f64)),
            hidden: props.hidden.map(|_| hidden.contains(&i)),
        };
        match &mut run {
            Some((_, h, rp)) if *h + 1 == i && *rp == p => *h = i,
            _ => {
                if let Some((a, b, rp)) = run.take() {
                    out.push(CollabOp::SetLines { sheet, axis, lo: a, hi: b, props: rp });
                }
                run = Some((i, i, p));
            }
        }
    }
    if let Some((a, b, rp)) = run {
        out.push(CollabOp::SetLines { sheet, axis, lo: a, hi: b, props: rp });
    }
    out
}

/// Each property `props` sets, as its own single-property props.
fn single_props(props: &FormatProps) -> Vec<(&'static str, FormatProps)> {
    let mut out = Vec::new();
    macro_rules! one {
        ($f:ident) => {
            if props.$f.is_some() {
                out.push((stringify!($f), FormatProps { $f: props.$f.clone(), ..Default::default() }));
            }
        };
    }
    one!(bold);
    one!(italic);
    one!(underline);
    one!(strikethrough);
    one!(font_family);
    one!(font_size);
    one!(color);
    one!(background);
    one!(number_format);
    one!(h_align);
    one!(v_align);
    one!(wrap);
    one!(border_top);
    one!(border_right);
    one!(border_bottom);
    one!(border_left);
    out
}

/// `props` without the named properties.
fn without_names(props: &FormatProps, names: &HashSet<&'static str>) -> FormatProps {
    let mut drop = FormatProps::default();
    for (name, p) in single_props(props) {
        if names.contains(name) {
            drop = merge(&drop, &p);
        }
    }
    props.without(&drop)
}

fn merge(a: &FormatProps, b: &FormatProps) -> FormatProps {
    let mut out = a.clone();
    macro_rules! take {
        ($f:ident) => {
            if b.$f.is_some() {
                out.$f = b.$f.clone();
            }
        };
    }
    take!(bold);
    take!(italic);
    take!(underline);
    take!(strikethrough);
    take!(font_family);
    take!(font_size);
    take!(color);
    take!(background);
    take!(number_format);
    take!(h_align);
    take!(v_align);
    take!(wrap);
    take!(border_top);
    take!(border_right);
    take!(border_bottom);
    take!(border_left);
    out
}

fn area(r: &Rect) -> usize {
    (r.r1 - r.r0 + 1).saturating_mul(r.c1 - r.c0 + 1)
}

/// Single-cell format ops, merging horizontal runs with equal props.
fn per_cell_formats(sheet: SheetKey, mut cells: Vec<(usize, usize, FormatProps)>) -> Vec<CollabOp> {
    cells.sort_by_key(|(r, c, _)| (*r, *c));
    let mut out: Vec<CollabOp> = Vec::new();
    for (r, c, props) in cells {
        if props.is_empty() {
            continue;
        }
        if let Some(CollabOp::SetFormat { rect, props: last, .. }) = out.last_mut() {
            if rect.r0 == r && rect.r1 == r && rect.c1 + 1 == c && *last == props {
                rect.c1 = c;
                continue;
            }
        }
        out.push(CollabOp::SetFormat { sheet, rect: Rect::new(r, c, r, c), props });
    }
    out
}

/// Ops that put back every cell (content and carried format properties)
/// listed in `cells`, on a sheet named `sheet_name`.
fn restore_cells(s: &visigrid_engine::sheet::Sheet, sheet: SheetKey, sheet_name: &str, cells: Vec<(usize, usize, String, CellFormat)>) -> Vec<CollabOp> {
    let default = CellFormat::default();
    let mut out = Vec::new();
    let mut formats = Vec::new();
    for (r, c, raw, fmt) in cells {
        if !raw.is_empty() {
            out.push(CollabOp::SetCell {
                sheet,
                sheet_name: sheet_name.to_string(),
                row: r,
                col: c,
                content: content_at(s, r, c),
            });
        }
        if fmt != default {
            formats.push((r, c, props_of(&fmt, None)));
        }
    }
    out.extend(per_cell_formats(sheet, formats));
    out
}

/// The inverse of a format op against `wb`: clear the touched properties
/// over the rectangle, then put back each cell whose old value differs.
fn invert_format(wb: &Workbook, sheet: SheetKey, rect: Rect, props: &FormatProps) -> Vec<CollabOp> {
    let Some(idx) = index_of(wb, sheet) else { return Vec::new() };
    let s = &wb.sheets()[idx];
    let cleared = props_of(&CellFormat::default(), Some(props));
    let coords: Vec<(usize, usize)> = if area(&rect) <= MAX_INSPECTED_CELLS {
        (rect.r0..=rect.r1).flat_map(|r| (rect.c0..=rect.c1).map(move |c| (r, c))).collect()
    } else {
        s.cells_in_range(rect.r0, rect.r1, rect.c0, rect.c1)
    };
    let mut cells = Vec::new();
    for (r, c) in coords {
        let old = props_of(&s.get_format(r, c), Some(props));
        if old != cleared {
            cells.push((r, c, old));
        }
    }
    let mut out = vec![CollabOp::SetFormat { sheet, rect, props: cleared }];
    out.extend(per_cell_formats(sheet, cells));
    out
}

/// Record the inverse of `ops`, which are about to apply to `before` (a
/// copy is stepped through; the client uses [`apply_recording`] instead).
pub fn record(origin: Uuid, ops: &[CollabOp], before: &Workbook) -> UndoEntry {
    let mut scratch = before.clone();
    apply_recording(&mut scratch, origin, ops, |wb, op| {
        apply_ops(wb, op);
    })
}

/// Apply `ops` to `wb` one at a time with `apply`, recording each op's
/// inverse against the state just before it and what it left just after.
/// No copy of the workbook is made.
pub fn apply_recording(
    scratch: &mut Workbook,
    origin: Uuid,
    ops: &[CollabOp],
    mut apply: impl FnMut(&mut Workbook, &[CollabOp]),
) -> UndoEntry {
    let mut groups: Vec<Vec<CollabOp>> = Vec::new();
    let mut expect = Vec::new();
    for op in ops {
        let name_of = |wb: &Workbook, sheet: SheetKey| {
            index_of(wb, sheet).map(|i| wb.sheets()[i].name.clone())
        };
        let mut inv = Vec::new();
        match op {
            CollabOp::SetCell { sheet, sheet_name, row, col, .. } => {
                if let Some(idx) = index_of(scratch, *sheet) {
                    inv.push(CollabOp::SetCell {
                        sheet: *sheet,
                        sheet_name: sheet_name.clone(),
                        row: *row,
                        col: *col,
                        content: content_at(&scratch.sheets()[idx], *row, *col),
                    });
                }
            }
            CollabOp::SetBold { .. } | CollabOp::SetFormat { .. } => {
                let (sheet, rect, props) = op.format().expect("format op");
                inv = invert_format(scratch, sheet, rect, &props);
                expect.push(op.normalized());
            }
            CollabOp::ReplaceRange { sheet, row, col, values } => {
                // Restored as cell writes, not a ReplaceRange: the old cells
                // can hold formulas, and only a SetCell carries the sheet name
                // the transform needs to rewrite a formula's references.
                if let Some(idx) = index_of(scratch, *sheet) {
                    let s = &scratch.sheets()[idx];
                    for (dr, line) in values.iter().enumerate() {
                        for dc in 0..line.len() {
                            inv.push(CollabOp::SetCell {
                                sheet: *sheet,
                                sheet_name: s.name.clone(),
                                row: row + dr,
                                col: col + dc,
                                content: content_at(s, row + dr, col + dc),
                            });
                        }
                    }
                }
            }
            CollabOp::Structural { sheet, sheet_name, axis, at, count, delete } => {
                inv.push(CollabOp::Structural {
                    sheet: *sheet,
                    sheet_name: sheet_name.clone(),
                    axis: *axis,
                    at: *at,
                    count: *count,
                    delete: !*delete,
                });
                if *delete {
                    if let Some(idx) = index_of(scratch, *sheet) {
                        let s = &scratch.sheets()[idx];
                        let cells = match axis {
                            crate::op::Axis::Row => s.occupied_cells_in_rows(*at, *count),
                            crate::op::Axis::Col => s.occupied_cells_in_cols(*at, *count),
                        };
                        inv.extend(restore_cells(s, *sheet, &s.name, cells));
                    }
                }
            }
            CollabOp::AddSheet { sheet, index, .. } => {
                inv.push(CollabOp::DeleteSheet { sheet: *sheet, index: *index });
            }
            CollabOp::MoveSheet { sheet, .. } => {
                if let Some(idx) = index_of(scratch, *sheet) {
                    inv.push(CollabOp::MoveSheet { sheet: *sheet, index: idx });
                }
            }
            CollabOp::SetLines { sheet, axis, lo, hi, props } => {
                inv = invert_lines(scratch, *sheet, *axis, *lo, *hi, props);
            }
            CollabOp::SetFreeze { sheet, .. } => {
                if let Some(idx) = index_of(scratch, *sheet) {
                    let l = &scratch.sheets()[idx].layout;
                    inv.push(CollabOp::SetFreeze { sheet: *sheet, rows: l.frozen_rows, cols: l.frozen_cols });
                }
            }
            CollabOp::Merge { sheet, rect } => {
                inv.push(CollabOp::Unmerge { sheet: *sheet, rect: *rect });
            }
            CollabOp::Unmerge { sheet, rect } => {
                if let Some(idx) = index_of(scratch, *sheet) {
                    for m in &scratch.sheets()[idx].merged_regions {
                        let r = Rect::new(m.start.0, m.start.1, m.end.0, m.end.1);
                        if rect.intersects(&r) {
                            inv.push(CollabOp::Merge { sheet: *sheet, rect: r });
                        }
                    }
                }
            }
            CollabOp::RenameSheet { sheet, .. } => {
                if let Some(name) = name_of(scratch, *sheet) {
                    inv.push(CollabOp::RenameSheet { sheet: *sheet, name });
                }
            }
            CollabOp::DeleteSheet { sheet, .. } => {
                if let Some(idx) = index_of(scratch, *sheet) {
                    let s = &scratch.sheets()[idx];
                    inv.push(CollabOp::AddSheet { sheet: *sheet, name: s.name.clone(), index: idx });
                    let cells: Vec<_> = s
                        .cells_iter()
                        .map(|((r, c), _)| (r, c, s.get_raw(r, c), s.get_format(r, c)))
                        .filter(|(_, _, raw, fmt)| !raw.is_empty() || *fmt != CellFormat::default())
                        .collect();
                    let mut cells = cells;
                    cells.sort_by_key(|(r, c, ..)| (*r, *c));
                    inv.extend(restore_cells(s, *sheet, &s.name, cells));
                }
            }
        }
        apply(scratch, std::slice::from_ref(op));
        // What the write left, as the engine stores it (it may normalize).
        match op {
            CollabOp::SetCell { sheet, sheet_name, row, col, .. } => {
                if let Some(idx) = index_of(scratch, *sheet) {
                    expect.push(CollabOp::SetCell {
                        sheet: *sheet,
                        sheet_name: sheet_name.clone(),
                        row: *row,
                        col: *col,
                        content: content_at(&scratch.sheets()[idx], *row, *col),
                    });
                }
            }
            CollabOp::ReplaceRange { sheet, row, col, values } => {
                if let Some(idx) = index_of(scratch, *sheet) {
                    let s = &scratch.sheets()[idx];
                    let name = s.name.clone();
                    for (dr, line) in values.iter().enumerate() {
                        for dc in 0..line.len() {
                            expect.push(CollabOp::SetCell {
                                sheet: *sheet,
                                sheet_name: name.clone(),
                                row: row + dr,
                                col: col + dc,
                                content: content_at(s, row + dr, col + dc),
                            });
                        }
                    }
                }
            }
            _ => {}
        }
        groups.push(inv);
    }
    groups.reverse();
    UndoEntry { origin, inverse: groups.into_iter().flatten().collect(), expect }
}

/// The ops that undo `entry` now: its inverse, minus every cell and format
/// property someone else changed since. Returns the ops and how many cells
/// were left alone.
pub fn resolve(entry: &UndoEntry, wb: &Workbook) -> (Vec<CollabOp>, usize) {
    let mut foreign_cells: HashSet<(SheetKey, usize, usize)> = HashSet::new();
    let mut foreign_props: HashSet<(SheetKey, usize, usize, &'static str)> = HashSet::new();
    for e in &entry.expect {
        match e {
            CollabOp::SetCell { sheet, row, col, content, .. } => {
                if let Some(idx) = index_of(wb, *sheet) {
                    if wb.sheets()[idx].get_raw(*row, *col) != content.raw() {
                        foreign_cells.insert((*sheet, *row, *col));
                    }
                }
            }
            CollabOp::SetFormat { sheet, rect, props } if area(rect) <= MAX_INSPECTED_CELLS => {
                let Some(idx) = index_of(wb, *sheet) else { continue };
                let s = &wb.sheets()[idx];
                let singles = single_props(props);
                for r in rect.r0..=rect.r1 {
                    for c in rect.c0..=rect.c1 {
                        let cur = s.get_format(r, c);
                        for (name, p) in &singles {
                            let mut f = cur.clone();
                            apply_props(&mut f, p);
                            if f != cur {
                                foreign_props.insert((*sheet, r, c, name));
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }
    let mut kept: HashSet<(SheetKey, usize, usize)> = HashSet::new();
    let mut out = Vec::with_capacity(entry.inverse.len());
    for op in &entry.inverse {
        match op {
            CollabOp::SetCell { sheet, row, col, .. } if foreign_cells.contains(&(*sheet, *row, *col)) => {
                kept.insert((*sheet, *row, *col));
            }
            CollabOp::SetFormat { sheet, rect, props }
                if foreign_props.iter().any(|(s, r, c, _)| s == sheet && rect.contains(*r, *c)) =>
            {
                if area(rect) > MAX_INSPECTED_CELLS {
                    out.push(op.clone());
                    continue;
                }
                let mut cells = Vec::new();
                for r in rect.r0..=rect.r1 {
                    for c in rect.c0..=rect.c1 {
                        let names: HashSet<&'static str> = foreign_props
                            .iter()
                            .filter(|(s, fr, fc, _)| s == sheet && *fr == r && *fc == c)
                            .map(|(.., n)| *n)
                            .collect();
                        if !names.is_empty() {
                            kept.insert((*sheet, r, c));
                        }
                        cells.push((r, c, without_names(props, &names)));
                    }
                }
                out.extend(per_cell_formats(*sheet, cells));
            }
            other => out.push(other.clone()),
        }
    }
    (out, kept.len())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::apply::fingerprint;
    use visigrid_engine::sheet::SheetId;

    fn wb() -> (Workbook, SheetKey) {
        let wb = Workbook::new();
        let key = wb.sheets()[0].id.0;
        (wb, key)
    }

    fn set(sheet: SheetKey, row: usize, col: usize, text: &str) -> CollabOp {
        CollabOp::SetCell { sheet, sheet_name: "Sheet1".into(), row, col, content: content_of(text.into()) }
    }

    fn undo_now(entry: &UndoEntry, wb: &mut Workbook) -> usize {
        let (ops, kept) = resolve(entry, wb);
        apply_ops(wb, &ops);
        kept
    }

    #[test]
    fn number_formats_map_to_codes_that_display_alike() {
        let (mut w, k) = wb();
        let idx = w.idx_for_sheet_id(SheetId(k)).unwrap();
        w.set_cell_value_tracked(idx, 0, 0, "1234.5");
        let fmt = CellFormat { number_format: NumberFormat::number(2), ..Default::default() };
        w.sheet_mut(idx).unwrap().set_format(0, 0, fmt.clone());
        let shown = w.sheets()[idx].get_formatted_display(0, 0);
        let mut back = CellFormat::default();
        apply_props(&mut back, &props_of(&fmt, None));
        w.sheet_mut(idx).unwrap().set_format(0, 0, back);
        assert_eq!(w.sheets()[idx].get_formatted_display(0, 0), shown);
    }

    #[test]
    fn every_op_kind_round_trips() {
        let (mut w, k) = wb();
        apply_ops(&mut w, &[set(k, 0, 0, "1"), set(k, 1, 0, "=A1+1"), set(k, 4, 2, "keep")]);
        apply_ops(
            &mut w,
            &[CollabOp::SetFormat { sheet: k, rect: Rect::new(4, 2, 4, 2), props: FormatProps::bold(true) }],
        );
        let start = fingerprint(&w);
        let cases: Vec<Vec<CollabOp>> = vec![
            vec![set(k, 0, 0, "5"), set(k, 9, 9, "new")],
            vec![CollabOp::SetFormat {
                sheet: k,
                rect: Rect::new(0, 0, 5, 3),
                props: FormatProps { italic: Some(Some(true)), bold: Some(None), ..Default::default() },
            }],
            vec![CollabOp::Structural {
                sheet: k,
                sheet_name: "Sheet1".into(),
                axis: crate::op::Axis::Row,
                at: 3,
                count: 2,
                delete: true,
            }],
            vec![CollabOp::Structural {
                sheet: k,
                sheet_name: "Sheet1".into(),
                axis: crate::op::Axis::Col,
                at: 1,
                count: 1,
                delete: false,
            }],
            vec![CollabOp::ReplaceRange {
                sheet: k,
                row: 0,
                col: 0,
                values: vec![vec![CellContent::Value("x".into()), CellContent::Clear]],
            }],
            vec![CollabOp::RenameSheet { sheet: k, name: "Renamed".into() }],
            vec![CollabOp::AddSheet { sheet: 99, name: "Extra".into(), index: 1 }],
        ];
        for ops in cases {
            let mut w2 = w.clone();
            let entry = record(Uuid::nil(), &ops, &w2);
            apply_ops(&mut w2, &ops);
            assert_eq!(undo_now(&entry, &mut w2), 0);
            assert_eq!(fingerprint(&w2), start, "undo of {ops:?}");
        }
    }

    #[test]
    fn deleting_a_sheet_round_trips() {
        let (mut w, k) = wb();
        apply_ops(&mut w, &[CollabOp::AddSheet { sheet: 7, name: "Two".into(), index: 1 }]);
        let two = |r, c, t: &str| CollabOp::SetCell {
            sheet: 7,
            sheet_name: "Two".into(),
            row: r,
            col: c,
            content: content_of(t.into()),
        };
        apply_ops(&mut w, &[two(0, 0, "a"), two(2, 1, "=Sheet1!A1"), set(k, 0, 0, "3")]);
        let start = fingerprint(&w);
        let ops = vec![CollabOp::DeleteSheet { sheet: 7, index: 1 }];
        let entry = record(Uuid::nil(), &ops, &w);
        apply_ops(&mut w, &ops);
        undo_now(&entry, &mut w);
        assert_eq!(fingerprint(&w), start);
    }

    #[test]
    fn leaves_cells_others_changed_since() {
        let (mut w, k) = wb();
        let ops = vec![set(k, 0, 0, "mine"), set(k, 0, 1, "mine too")];
        let entry = record(Uuid::nil(), &ops, &w);
        apply_ops(&mut w, &ops);
        apply_ops(&mut w, &[set(k, 0, 1, "theirs")]);
        assert_eq!(undo_now(&entry, &mut w), 1);
        let s = &w.sheets()[0];
        assert_eq!(s.get_raw(0, 0), "");
        assert_eq!(s.get_raw(0, 1), "theirs");
    }

    #[test]
    fn leaves_format_properties_others_changed_since() {
        let (mut w, k) = wb();
        let ops = vec![CollabOp::SetFormat {
            sheet: k,
            rect: Rect::new(0, 0, 0, 1),
            props: FormatProps { bold: Some(Some(true)), italic: Some(Some(true)), ..Default::default() },
        }];
        let entry = record(Uuid::nil(), &ops, &w);
        apply_ops(&mut w, &ops);
        // Someone else un-bolds B1; its italic is still ours to undo.
        apply_ops(
            &mut w,
            &[CollabOp::SetFormat { sheet: k, rect: Rect::new(0, 1, 0, 1), props: FormatProps::bold(false) }],
        );
        assert_eq!(undo_now(&entry, &mut w), 1);
        let s = &w.sheets()[0];
        assert!(!s.get_format(0, 0).bold && !s.get_format(0, 0).italic);
        assert!(!s.get_format(0, 1).bold && !s.get_format(0, 1).italic);
    }
}
