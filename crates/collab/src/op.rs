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
    /// Literal text, never read as a number, date or formula ("007", an ID
    /// that looks numeric): what an import's text columns write.
    Text(String),
}

impl CellContent {
    pub fn raw(&self) -> &str {
        match self {
            CellContent::Value(s) | CellContent::Formula(s) | CellContent::Text(s) => s,
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

    /// The overlap of two rectangles, if any.
    pub fn intersection(&self, o: &Rect) -> Option<Rect> {
        if !self.intersects(o) {
            return None;
        }
        Some(Rect::new(self.r0.max(o.r0), self.c0.max(o.c0), self.r1.min(o.r1), self.c1.min(o.c1)))
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

/// `Option<Option<T>>` from JSON: an absent field is `None` (unchanged), an
/// explicit `null` is `Some(None)` (clear to the default), a value is
/// `Some(Some(v))` (set).
fn double_option<'de, T, D>(d: D) -> Result<Option<Option<T>>, D::Error>
where
    T: Deserialize<'de>,
    D: serde::Deserializer<'de>,
{
    Option::<T>::deserialize(d).map(Some)
}

/// Horizontal alignment.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum HAlign {
    General,
    Left,
    Center,
    Right,
}

/// Vertical alignment.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum VAlign {
    Top,
    Middle,
    Bottom,
}

/// A border edge's line style.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum BorderLine {
    Thin,
    Medium,
    Thick,
}

/// One border edge: a line style and an optional `#RRGGBB` colour (black when absent).
#[derive(Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct BorderSpec {
    pub style: BorderLine,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

/// The format properties one op sets or clears. Every field is
/// `None` = unchanged, `Some(None)` = clear to the default, `Some(Some(v))` =
/// set. Colors are `#RRGGBB`. `number_format` is an Excel format code
/// (`"General"` or e.g. `"#,##0.00"`).
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct FormatProps {
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub bold: Option<Option<bool>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub italic: Option<Option<bool>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub underline: Option<Option<bool>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<Option<bool>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub font_family: Option<Option<String>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub font_size: Option<Option<f64>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub color: Option<Option<String>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub background: Option<Option<String>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub number_format: Option<Option<String>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub h_align: Option<Option<HAlign>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub v_align: Option<Option<VAlign>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub wrap: Option<Option<bool>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub border_top: Option<Option<BorderSpec>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub border_right: Option<Option<BorderSpec>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub border_bottom: Option<Option<BorderSpec>>,
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub border_left: Option<Option<BorderSpec>>,
}

/// What a `SetLines` op sets on a run of rows or columns. `size`: `None`
/// unchanged, `Some(None)` back to the default, `Some(Some(points))` set.
/// `hidden`: `None` unchanged, else hide or show.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct LineProps {
    #[serde(default, deserialize_with = "double_option", skip_serializing_if = "Option::is_none")]
    pub size: Option<Option<f64>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
}

/// Sizes are validated finite, so equality is reflexive.
impl Eq for LineProps {}

impl LineProps {
    pub fn is_empty(&self) -> bool {
        self.size.is_none() && self.hidden.is_none()
    }

    /// `self` without the properties `other` sets.
    pub fn without(&self, other: &LineProps) -> LineProps {
        LineProps {
            size: if other.size.is_some() { None } else { self.size },
            hidden: if other.hidden.is_some() { None } else { self.hidden },
        }
    }

    pub fn validate(&self) -> Result<(), String> {
        if let Some(Some(s)) = self.size {
            if !s.is_finite() || s < 0.0 || s > 2000.0 {
                return Err(format!("line size {s} out of range"));
            }
        }
        Ok(())
    }
}

/// `font_size` is the only float. Ops are compared exactly (tests, dedupe);
/// `validate` rejects non-finite sizes, so equality is reflexive.
impl Eq for FormatProps {}

macro_rules! each_prop {
    ($m:ident) => {
        $m!(bold);
        $m!(italic);
        $m!(underline);
        $m!(strikethrough);
        $m!(font_family);
        $m!(font_size);
        $m!(color);
        $m!(background);
        $m!(number_format);
        $m!(h_align);
        $m!(v_align);
        $m!(wrap);
        $m!(border_top);
        $m!(border_right);
        $m!(border_bottom);
        $m!(border_left);
    };
}

impl FormatProps {
    pub fn bold(bold: bool) -> Self {
        FormatProps { bold: Some(Some(bold)), ..Default::default() }
    }

    /// No property is set or cleared.
    pub fn is_empty(&self) -> bool {
        let mut empty = true;
        macro_rules! check {
            ($f:ident) => {
                empty &= self.$f.is_none();
            };
        }
        each_prop!(check);
        empty
    }

    /// `self` without the properties `other` sets or clears: what survives
    /// of an earlier op where a later one overlaps it.
    pub fn without(&self, other: &FormatProps) -> FormatProps {
        let mut out = self.clone();
        macro_rules! drop_shared {
            ($f:ident) => {
                if other.$f.is_some() {
                    out.$f = None;
                }
            };
        }
        each_prop!(drop_shared);
        out
    }

    /// Reject values no replica could apply identically.
    pub fn validate(&self) -> Result<(), String> {
        if let Some(Some(size)) = self.font_size {
            if !size.is_finite() || size <= 0.0 || size > 409.0 {
                return Err(format!("font_size {size} out of range"));
            }
        }
        for (name, c) in [("color", &self.color), ("background", &self.background)] {
            if let Some(Some(c)) = c {
                if parse_hex_color(c).is_none() {
                    return Err(format!("{name} must be #RRGGBB"));
                }
            }
        }
        for b in [&self.border_top, &self.border_right, &self.border_bottom, &self.border_left] {
            if let Some(Some(BorderSpec { color: Some(c), .. })) = b {
                if parse_hex_color(c).is_none() {
                    return Err("border colour must be #RRGGBB".into());
                }
            }
        }
        Ok(())
    }
}

/// `#RRGGBB` (case-insensitive) as RGBA with full opacity.
pub fn parse_hex_color(s: &str) -> Option<[u8; 4]> {
    let hex = s.strip_prefix('#')?;
    if hex.len() != 6 || !hex.bytes().all(|b| b.is_ascii_hexdigit()) {
        return None;
    }
    let byte = |i: usize| u8::from_str_radix(&hex[i..i + 2], 16).ok();
    Some([byte(0)?, byte(2)?, byte(4)?, 255])
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
    /// Legacy bold-only format range, kept so existing logs replay. Every
    /// transform and apply treats it as `SetFormat { bold }`.
    SetBold {
        sheet: SheetKey,
        rect: Rect,
        bold: bool,
    },
    /// Format range: set or clear any of `props` on every cell of `rect`.
    SetFormat {
        sheet: SheetKey,
        rect: Rect,
        props: FormatProps,
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
    /// Move a sheet's tab to `index` (its position after the move). V1
    /// serializes it against any concurrent add, delete or move.
    MoveSheet {
        sheet: SheetKey,
        index: usize,
    },
    /// Set the size and/or visibility of rows or columns `lo..=hi`. Merges
    /// per property with concurrent ones, like `SetFormat`.
    SetLines {
        sheet: SheetKey,
        axis: Axis,
        lo: usize,
        hi: usize,
        props: LineProps,
    },
    /// Freeze the first `rows` rows and `cols` columns (0 = none). The later
    /// of two concurrent freezes wins.
    SetFreeze {
        sheet: SheetKey,
        rows: usize,
        cols: usize,
    },
    /// Merge `rect` into one cell. V1 serializes it against concurrent edits
    /// that touch the rectangle and against structural edits on the sheet.
    Merge {
        sheet: SheetKey,
        rect: Rect,
    },
    /// Remove every merge that intersects `rect`. Serialized like `Merge`.
    Unmerge {
        sheet: SheetKey,
        rect: Rect,
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
            | CollabOp::SetFormat { sheet, .. }
            | CollabOp::Structural { sheet, .. }
            | CollabOp::AddSheet { sheet, .. }
            | CollabOp::RenameSheet { sheet, .. }
            | CollabOp::DeleteSheet { sheet, .. }
            | CollabOp::MoveSheet { sheet, .. }
            | CollabOp::SetLines { sheet, .. }
            | CollabOp::SetFreeze { sheet, .. }
            | CollabOp::Merge { sheet, .. }
            | CollabOp::Unmerge { sheet, .. }
            | CollabOp::ReplaceRange { sheet, .. } => *sheet,
        }
    }

    /// `SetBold` as the `SetFormat` it stands for; every other op unchanged.
    pub fn normalized(&self) -> CollabOp {
        match self {
            CollabOp::SetBold { sheet, rect, bold } => CollabOp::SetFormat {
                sheet: *sheet,
                rect: *rect,
                props: FormatProps::bold(*bold),
            },
            other => other.clone(),
        }
    }

    /// The rectangle and properties of a format op (`SetBold` or `SetFormat`).
    pub fn format(&self) -> Option<(SheetKey, Rect, FormatProps)> {
        match self.normalized() {
            CollabOp::SetFormat { sheet, rect, props } => Some((sheet, rect, props)),
            _ => None,
        }
    }

    /// Reject ops no replica could apply identically.
    pub fn validate(&self) -> Result<(), String> {
        match self {
            CollabOp::SetFormat { rect, props, .. } => {
                if rect.r0 > rect.r1 || rect.c0 > rect.c1 {
                    return Err("format rect is inverted".into());
                }
                props.validate()
            }
            CollabOp::SetLines { lo, hi, props, .. } => {
                if lo > hi {
                    return Err("line span is inverted".into());
                }
                props.validate()
            }
            CollabOp::Merge { rect, .. } | CollabOp::Unmerge { rect, .. } => {
                if rect.r0 > rect.r1 || rect.c0 > rect.c1 {
                    return Err("merge rect is inverted".into());
                }
                Ok(())
            }
            _ => Ok(()),
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

#[cfg(feature = "protocol")]
/// Map a v1 protocol op onto the collaboration vocabulary. Protocol v1 names
/// sheets by index, so the caller resolves the stable key and the name.
/// Ops outside the collaboration vocabulary (pivots, borders) return `None`.
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
            bold,
            italic,
            underline,
            ..
        } if bold.is_some() || italic.is_some() || underline.is_some() => CollabOp::SetFormat {
            sheet: sheet_key,
            rect: Rect::new(*start_row, *start_col, *end_row, *end_col),
            props: FormatProps {
                bold: bold.map(Some),
                italic: italic.map(Some),
                underline: underline.map(Some),
                ..Default::default()
            },
        },
        _ => return None,
    })
}

/// The opaque `op` payload of protocol v2 and the engine host protocol: the
/// envelope's atomic op list as a JSON array. A single op object is accepted
/// on input too, so a writer with one op need not wrap it.
pub fn ops_from_json(v: &serde_json::Value) -> Result<Vec<CollabOp>, String> {
    let ops: Vec<CollabOp> = if v.is_array() {
        serde_json::from_value(v.clone()).map_err(|e| format!("invalid op list: {e}"))?
    } else {
        serde_json::from_value::<CollabOp>(v.clone())
            .map(|op| vec![op])
            .map_err(|e| format!("invalid op: {e}"))?
    };
    for op in &ops {
        op.validate()?;
    }
    Ok(ops)
}

pub fn ops_to_json(ops: &[CollabOp]) -> serde_json::Value {
    serde_json::to_value(ops).expect("ops serialize")
}
