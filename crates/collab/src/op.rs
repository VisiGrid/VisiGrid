//! Operations and the protocol v2 envelope.
//!
//! An operation is an *input* to the workbook — raw cell text, a format, a
//! structural edit — never a computed value. Every replica recalculates.
//!
//! Cells and structural edits name their sheet by `SheetId` (stable) and also
//! carry the sheet's *name* as of the state the op was written against.
//! Formula reference rewriting is name-based (that is how the engine stores
//! references), so the transform needs names without consulting any state;
//! a concurrent rename updates the carried name.

use serde::{Deserialize, Serialize};
use uuid::Uuid;

/// Stable sheet identity (`visigrid_engine::sheet::SheetId.0`).
pub type SheetKey = u64;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum Axis {
    Row,
    Col,
}

impl From<Axis> for visigrid_engine::structural::Axis {
    fn from(a: Axis) -> Self {
        match a {
            Axis::Row => visigrid_engine::structural::Axis::Row,
            Axis::Col => visigrid_engine::structural::Axis::Col,
        }
    }
}

/// What a cell write puts in the cell.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum CellContent {
    /// Literal input (a number or text, as typed).
    Value(String),
    /// Formula text, starting with `=`.
    Formula(String),
    Clear,
}

impl CellContent {
    pub fn raw(&self) -> &str {
        match self {
            CellContent::Value(s) | CellContent::Formula(s) => s,
            CellContent::Clear => "",
        }
    }
}

/// Inclusive rectangle.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct Rect {
    pub r0: usize,
    pub c0: usize,
    pub r1: usize,
    pub c1: usize,
}

impl Rect {
    pub fn new(r0: usize, c0: usize, r1: usize, c1: usize) -> Self {
        Rect { r0, c0, r1, c1 }
    }

    pub fn contains(&self, r: usize, c: usize) -> bool {
        r >= self.r0 && r <= self.r1 && c >= self.c0 && c <= self.c1
    }

    pub fn intersects(&self, o: &Rect) -> bool {
        self.r0 <= o.r1 && o.r0 <= self.r1 && self.c0 <= o.c1 && o.c0 <= self.c1
    }

    /// `self` minus `o`, as up to four disjoint rectangles.
    pub fn subtract(&self, o: &Rect) -> Vec<Rect> {
        if !self.intersects(o) {
            return vec![*self];
        }
        let mut out = Vec::new();
        if o.r0 > self.r0 {
            out.push(Rect::new(self.r0, self.c0, o.r0 - 1, self.c1));
        }
        if o.r1 < self.r1 {
            out.push(Rect::new(o.r1 + 1, self.c0, self.r1, self.c1));
        }
        let (mr0, mr1) = (self.r0.max(o.r0), self.r1.min(o.r1));
        if o.c0 > self.c0 {
            out.push(Rect::new(mr0, self.c0, mr1, o.c0 - 1));
        }
        if o.c1 < self.c1 {
            out.push(Rect::new(mr0, o.c1 + 1, mr1, self.c1));
        }
        out
    }
}

/// One operation. V1 vocabulary (see the spec's pair table).
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum CollabOp {
    /// Write a value or formula, or clear, one cell.
    SetCell {
        sheet: SheetKey,
        sheet_name: String,
        row: usize,
        col: usize,
        content: CellContent,
    },
    /// Format range (V1: the bold property; others follow the same rules).
    SetBold {
        sheet: SheetKey,
        rect: Rect,
        bold: bool,
    },
    /// Insert or delete `count` rows/columns starting at `at`.
    Structural {
        sheet: SheetKey,
        sheet_name: String,
        axis: Axis,
        at: usize,
        count: usize,
        delete: bool,
    },
    /// Add a sheet with a pre-chosen stable id at a tab position.
    AddSheet {
        sheet: SheetKey,
        name: String,
        index: usize,
    },
    RenameSheet {
        sheet: SheetKey,
        name: String,
    },
    /// `index` is the sheet's tab position in the state the op was written
    /// against; transforms keep it current so tab order converges.
    DeleteSheet {
        sheet: SheetKey,
        index: usize,
    },
    /// V1 atomic range op standing in for sort / move / large paste: it
    /// replaces a block of values. Concurrent overlapping ops are refused.
    ReplaceRange {
        sheet: SheetKey,
        row: usize,
        col: usize,
        values: Vec<Vec<CellContent>>,
    },
}

impl CollabOp {
    pub fn sheet(&self) -> SheetKey {
        match self {
            CollabOp::SetCell { sheet, .. }
            | CollabOp::SetBold { sheet, .. }
            | CollabOp::Structural { sheet, .. }
            | CollabOp::AddSheet { sheet, .. }
            | CollabOp::RenameSheet { sheet, .. }
            | CollabOp::DeleteSheet { sheet, .. }
            | CollabOp::ReplaceRange { sheet, .. } => *sheet,
        }
    }

    /// The block a `ReplaceRange` covers.
    pub fn replace_rect(&self) -> Option<Rect> {
        match self {
            CollabOp::ReplaceRange {
                row, col, values, ..
            } => {
                let h = values.len().max(1);
                let w = values.iter().map(|r| r.len()).max().unwrap_or(0).max(1);
                Some(Rect::new(*row, *col, row + h - 1, col + w - 1))
            }
            _ => None,
        }
    }
}

/// Protocol v2 envelope. `ops` is applied atomically, in order. A single
/// user action is usually one op, but a transform can split one (a delete
/// around a concurrent insert), so the wire unit is a list.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct Envelope {
    /// Idempotency key: a resend after reconnect must not apply twice.
    pub client_op_id: Uuid,
    /// The last sequence number the sender had applied.
    pub base_seq: u64,
    /// Set by the server from the authenticated connection.
    pub actor: u64,
    pub ops: Vec<CollabOp>,
}

/// Map a v1 protocol op onto the collaboration vocabulary. Protocol v1 names
/// sheets by index, so the caller resolves the stable key and the name.
/// Ops outside the V1 collaboration vocabulary (number formats, italic,
/// underline, pivots) return `None`; they are later rows of the table.
pub fn from_protocol(
    op: &visigrid_protocol::Op,
    sheet_key: SheetKey,
    sheet_name: &str,
) -> Option<CollabOp> {
    use visigrid_protocol::Op;
    Some(match op {
        Op::SetCellValue {
            row, col, value, ..
        } => CollabOp::SetCell {
            sheet: sheet_key,
            sheet_name: sheet_name.into(),
            row: *row,
            col: *col,
            content: CellContent::Value(value.clone()),
        },
        Op::SetCellFormula {
            row, col, formula, ..
        } => CollabOp::SetCell {
            sheet: sheet_key,
            sheet_name: sheet_name.into(),
            row: *row,
            col: *col,
            content: CellContent::Formula(formula.clone()),
        },
        Op::ClearCell { row, col, .. } => CollabOp::SetCell {
            sheet: sheet_key,
            sheet_name: sheet_name.into(),
            row: *row,
            col: *col,
            content: CellContent::Clear,
        },
        Op::SetStyle {
            start_row,
            start_col,
            end_row,
            end_col,
            bold: Some(b),
            italic: None,
            underline: None,
            ..
        } => CollabOp::SetBold {
            sheet: sheet_key,
            rect: Rect::new(*start_row, *start_col, *end_row, *end_col),
            bold: *b,
        },
        _ => return None,
    })
}
