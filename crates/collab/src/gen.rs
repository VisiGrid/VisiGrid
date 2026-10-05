//! Random operations, generated from one replica's own view and weighted
//! toward the conflicts the transform table has to get right: a small hot
//! region, edits next to concurrent inserts and deletes, overlapping deletes,
//! cross-sheet formulas, renames, and atomic ranges.
//!
//! Formulas avoid NOW/TODAY/RAND/INDIRECT/OFFSET: their determinism is Phase
//! 0 work (VisiGrid#88), not the transform's.

use rand::rngs::StdRng;
use rand::Rng;
use visigrid_engine::formula::parser::{format_parsed_expr, parse};
use visigrid_engine::workbook::Workbook;

use crate::op::{Axis, CellContent, CollabOp, FormatProps, HAlign, Rect, VAlign, BorderLine, BorderSpec, LineProps};

/// Rows and columns most edits land in.
pub const HOT_ROWS: usize = 10;
pub const HOT_COLS: usize = 5;
/// Sheet names are drawn from a small pool so concurrent adds and renames
/// collide.
const NAME_POOL: usize = 6;

fn col_letter(c: usize) -> char {
    (b'A' + c as u8) as char
}

/// The engine's canonical spelling, so a generated formula is stored exactly
/// as written.
fn canonical(f: &str) -> String {
    match parse(f) {
        Ok(p) => format_parsed_expr(&p),
        Err(_) => f.to_string(),
    }
}

/// A formula for a cell in row `row` (0-based, so row >= 1) that references
/// only rows above it (1-based rows 1..=row). Structural edits keep row order,
/// so the workbook stays acyclic: the engine's full recompute destroys the
/// text of formulas in a cycle (see the crate report), which is an engine
/// bug the transform cannot fix.
fn formula(rng: &mut StdRng, wb: &Workbook, row: usize) -> String {
    let above = row; // 1-based rows strictly above this cell: 1..=row
    let r = |rng: &mut StdRng| rng.gen_range(1..=above);
    let c = |rng: &mut StdRng| col_letter(rng.gen_range(0..HOT_COLS));
    let f = match rng.gen_range(0..6) {
        0 => format!("={}{}+{}{}", c(rng), r(rng), c(rng), r(rng)),
        1 => {
            let (a, b) = (r(rng), r(rng));
            format!("=SUM({}{}:{}{})", c(rng), a.min(b), c(rng), a.max(b))
        }
        2 => format!("={}{}*2", c(rng), r(rng)),
        3 => {
            let col = c(rng);
            let (a, b) = (r(rng), r(rng));
            format!("=SUM({col}{}:{col}{})", a.min(b), a.max(b))
        }
        4 => {
            // Cross-sheet: names in the pool are plain identifiers.
            let sheets = wb.sheets();
            let s = &sheets[rng.gen_range(0..sheets.len())];
            format!("={}!{}{}+1", s.name, c(rng), r(rng))
        }
        _ => format!("=IF({}{}>5,{}{},0)", c(rng), r(rng), c(rng), r(rng)),
    };
    canonical(&f)
}

fn value(rng: &mut StdRng) -> String {
    if rng.gen_bool(0.85) {
        rng.gen_range(-20..100).to_string()
    } else {
        ["apple", "pear", "x", "total"][rng.gen_range(0..4)].to_string()
    }
}

fn fresh_name(rng: &mut StdRng, wb: &Workbook) -> Option<String> {
    for _ in 0..8 {
        let n = format!("S{}", rng.gen_range(0..NAME_POOL));
        if !wb.sheet_name_exists(&n) {
            return Some(n);
        }
    }
    None
}

/// One to four properties, each set (or sometimes cleared) from a small
/// pool so concurrent formats of the same cells really collide.
fn format_props(rng: &mut StdRng) -> FormatProps {
    let mut p = FormatProps::default();
    let n = rng.gen_range(1..=4);
    for _ in 0..n {
        let clear = rng.gen_bool(0.15);
        macro_rules! pick {
            ($v:expr) => {
                if clear { Some(None) } else { Some(Some($v)) }
            };
        }
        match rng.gen_range(0..14) {
            0 => p.bold = pick!(rng.gen_bool(0.7)),
            1 => p.italic = pick!(rng.gen_bool(0.7)),
            2 => p.underline = pick!(rng.gen_bool(0.5)),
            3 => p.strikethrough = pick!(rng.gen_bool(0.5)),
            4 => p.font_family = pick!(["Inter", "Georgia", "Menlo"][rng.gen_range(0..3)].to_string()),
            5 => p.font_size = pick!([9.0, 11.0, 14.5, 24.0][rng.gen_range(0..4)]),
            6 => p.color = pick!(["#FF0000", "#00aa00", "#123456"][rng.gen_range(0..3)].to_string()),
            7 => p.background = pick!(["#FFFF00", "#eeeeee"][rng.gen_range(0..2)].to_string()),
            // Number formats change what numbers display.
            8 => p.number_format = pick!(["General", "0.00", "#,##0", "0%"][rng.gen_range(0..4)].to_string()),
            9 => p.h_align = pick!([HAlign::General, HAlign::Left, HAlign::Center, HAlign::Right][rng.gen_range(0..4)]),
            10 => p.v_align = pick!([VAlign::Top, VAlign::Middle, VAlign::Bottom][rng.gen_range(0..3)]),
            11 => p.wrap = pick!(rng.gen_bool(0.5)),
            12 => p.border_bottom = pick!(border(rng)),
            _ => {
                p.border_left = pick!(border(rng));
                p.border_top = pick!(border(rng));
            }
        }
    }
    p
}

fn border(rng: &mut StdRng) -> BorderSpec {
    BorderSpec {
        style: [BorderLine::Thin, BorderLine::Medium, BorderLine::Thick][rng.gen_range(0..3)],
        color: rng.gen_bool(0.3).then(|| "#3355FF".to_string()),
    }
}

/// Line layout, freeze or merge changes.
fn layout_op(rng: &mut StdRng, sheet: u64) -> CollabOp {
    let rect = |rng: &mut StdRng| {
        let (r0, c0) = (rng.gen_range(0..HOT_ROWS), rng.gen_range(0..HOT_COLS));
        Rect::new(r0, c0, (r0 + rng.gen_range(0..3)).min(HOT_ROWS), (c0 + rng.gen_range(0..2)).min(HOT_COLS))
    };
    match rng.gen_range(0..100) {
        0..=49 => {
            let axis = if rng.gen_bool(0.6) { Axis::Row } else { Axis::Col };
            let limit = if axis == Axis::Row { HOT_ROWS } else { HOT_COLS };
            let lo = rng.gen_range(0..limit);
            let hi = (lo + rng.gen_range(0..3)).min(limit);
            let mut props = LineProps::default();
            if rng.gen_bool(0.6) {
                props.size = Some(if rng.gen_bool(0.2) { None } else { Some([18.0, 40.0, 120.0][rng.gen_range(0..3)]) });
            }
            if props.size.is_none() || rng.gen_bool(0.3) {
                props.hidden = Some(rng.gen_bool(0.6));
            }
            CollabOp::SetLines { sheet, axis, lo, hi, props }
        }
        50..=64 => CollabOp::SetFreeze { sheet, rows: rng.gen_range(0..4), cols: rng.gen_range(0..3) },
        65..=84 => CollabOp::Merge { sheet, rect: rect(rng) },
        _ => CollabOp::Unmerge { sheet, rect: rect(rng) },
    }
}

/// One user action against `wb`, as an envelope's op list. `sheet_key`
/// supplies a new stable id for an added sheet.
pub fn random_ops(rng: &mut StdRng, wb: &Workbook, sheet_key: u64) -> Vec<CollabOp> {
    let sheets = wb.sheets();
    let idx = rng.gen_range(0..sheets.len());
    let s = &sheets[idx];
    let (sheet, sheet_name) = (s.id.0, s.name.clone());
    let roll = rng.gen_range(0..100);
    let op = if roll < 45 {
        let row = rng.gen_range(0..HOT_ROWS);
        let content = match rng.gen_range(0..10) {
            0..=4 => CellContent::Value(value(rng)),
            5..=8 if row > 0 => CellContent::Formula(formula(rng, wb, row)),
            5..=8 => CellContent::Value(value(rng)),
            _ => CellContent::Clear,
        };
        CollabOp::SetCell {
            sheet,
            sheet_name,
            row,
            col: rng.gen_range(0..HOT_COLS),
            content,
        }
    } else if roll < 55 {
        let (r0, c0) = (rng.gen_range(0..HOT_ROWS), rng.gen_range(0..HOT_COLS));
        let (r1, c1) = (
            (r0 + rng.gen_range(0..3)).min(HOT_ROWS),
            (c0 + rng.gen_range(0..2)).min(HOT_COLS),
        );
        let rect = Rect::new(r0, c0, r1, c1);
        if rng.gen_bool(0.25) {
            // Legacy logs still carry SetBold.
            CollabOp::SetBold { sheet, rect, bold: rng.gen_bool(0.7) }
        } else {
            CollabOp::SetFormat { sheet, rect, props: format_props(rng) }
        }
    } else if roll < 63 {
        layout_op(rng, sheet)
    } else if roll < 80 {
        let axis = if rng.gen_bool(0.75) {
            Axis::Row
        } else {
            Axis::Col
        };
        let limit = if axis == Axis::Row {
            HOT_ROWS
        } else {
            HOT_COLS
        };
        CollabOp::Structural {
            sheet,
            sheet_name,
            axis,
            at: rng.gen_range(0..limit),
            count: rng.gen_range(1..=3),
            delete: rng.gen_bool(0.45),
        }
    } else if roll < 85 {
        match fresh_name(rng, wb) {
            Some(name) => CollabOp::AddSheet {
                sheet: sheet_key,
                name,
                index: rng.gen_range(0..=sheets.len()),
            },
            None => return Vec::new(),
        }
    } else if roll < 90 {
        match fresh_name(rng, wb) {
            Some(name) => CollabOp::RenameSheet { sheet, name },
            None => return Vec::new(),
        }
    } else if roll < 92 {
        if sheets.len() < 2 {
            return Vec::new();
        }
        CollabOp::DeleteSheet { sheet, index: idx }
    } else if roll < 95 {
        if sheets.len() < 2 {
            return Vec::new();
        }
        CollabOp::MoveSheet { sheet, index: rng.gen_range(0..sheets.len()) }
    } else {
        let (h, w) = (rng.gen_range(1..=3), rng.gen_range(1..=2));
        let values = (0..h)
            .map(|_| (0..w).map(|_| CellContent::Value(value(rng))).collect())
            .collect();
        CollabOp::ReplaceRange {
            sheet,
            row: rng.gen_range(0..HOT_ROWS),
            col: rng.gen_range(0..HOT_COLS),
            values,
        }
    };
    vec![op]
}
