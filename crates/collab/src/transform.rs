//! `T(a, b)`: operation `a`, written without knowing about the concurrent
//! operation `b`, rewritten to apply after `b`.
//!
//! Client-server OT needs only this one function plus an explicit order:
//!
//! - On the **server**, an incoming op is transformed against each op already
//!   sequenced after its base: the incoming op is [`Order::Later`].
//! - On a **client**, an incoming sequenced op meets the client's pending
//!   (unsequenced) ops. The pending ops will be sequenced *after* it, so the
//!   pending op is transformed with [`Order::Later`] and the incoming op with
//!   [`Order::Earlier`].
//!
//! Correctness is TP1: for every pair defined on the same state `S`,
//! `apply(S, [b, T(a, b, Later)]) == apply(S, [a, T(b, a, Earlier)])`.
//!
//! Ties are broken by sequence order: the later op wins a same-cell write;
//! the earlier of two inserts at the same index goes first (above / left).
//! Pairs that V1 serializes return a refusal *for the later op only*; the
//! earlier op is left unchanged (the later one will never be sequenced).

use crate::op::{Axis, CellContent, CollabOp, Rect, SheetKey};
use visigrid_engine::structural::{adjust_formula_text, StructuralEdit};

/// Whether `a` is sequenced after `b`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Order {
    Later,
    Earlier,
}

impl Order {
    pub fn flip(self) -> Order {
        match self {
            Order::Later => Order::Earlier,
            Order::Earlier => Order::Later,
        }
    }
}

/// The result of transforming one op.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Transformed {
    /// Usually one op; a delete split around a concurrent insert is two;
    /// a format range cut by a concurrent one can be up to four.
    Ops(Vec<CollabOp>),
    /// The op no longer has a target (its row or sheet was deleted, or a
    /// later write of the same cell won). Not an error: it simply does nothing.
    Dropped(&'static str),
    /// V1 serializes this pair: the later op is refused and its client
    /// refreshes.
    Refused(String),
}

/// Transform `a` past `b`.
pub fn transform(a: &CollabOp, b: &CollabOp, order: Order) -> Transformed {
    if let Some(reason) = conflict(a, b) {
        return match order {
            Order::Later => Transformed::Refused(reason),
            Order::Earlier => Transformed::Ops(vec![a.clone()]),
        };
    }
    let later = order == Order::Later;
    use CollabOp::*;
    match (a, b) {
        // ---- a is a cell write ----
        (
            SetCell {
                sheet, row, col, ..
            },
            SetCell {
                sheet: bs,
                row: br,
                col: bc,
                ..
            },
        ) => {
            if sheet == bs && row == br && col == bc && !later {
                Transformed::Dropped("a later write of the same cell won")
            } else {
                one(a)
            }
        }
        (
            SetCell {
                sheet,
                sheet_name,
                row,
                col,
                content,
            },
            Structural {
                sheet: bs,
                sheet_name: bn,
                axis,
                at,
                count,
                delete,
            },
        ) => {
            let edit = edit(bn, *axis, *at, *count, *delete);
            let content = match content {
                CellContent::Formula(f) => CellContent::Formula(
                    adjust_formula_text(f, &edit, sheet_name).unwrap_or_else(|| f.clone()),
                ),
                other => other.clone(),
            };
            let (mut r, mut c) = (*row, *col);
            if sheet == bs {
                let v = if *axis == Axis::Row { &mut r } else { &mut c };
                match map_index(*v, *at, *count, *delete) {
                    Some(n) => *v = n,
                    None => return Transformed::Dropped("its row or column was deleted"),
                }
            }
            one(&SetCell {
                sheet: *sheet,
                sheet_name: sheet_name.clone(),
                row: r,
                col: c,
                content,
            })
        }
        (
            SetCell {
                sheet,
                row,
                col,
                content,
                ..
            },
            RenameSheet { sheet: bs, name },
        ) if sheet == bs => one(&SetCell {
            sheet: *sheet,
            sheet_name: name.clone(),
            row: *row,
            col: *col,
            content: content.clone(),
        }),
        (SetCell { sheet, .. }, DeleteSheet { sheet: bs, .. }) if sheet == bs => {
            Transformed::Dropped("its sheet was deleted")
        }
        (SetCell { .. }, _) => one(a),

        // ---- a is a format range ----
        (
            SetBold { sheet, rect, bold },
            SetBold {
                sheet: bs,
                rect: br,
                ..
            },
        ) => {
            if sheet == bs && !later && rect.intersects(br) {
                let rest: Vec<CollabOp> = rect
                    .subtract(br)
                    .into_iter()
                    .map(|r| SetBold {
                        sheet: *sheet,
                        rect: r,
                        bold: *bold,
                    })
                    .collect();
                if rest.is_empty() {
                    Transformed::Dropped("a later format of the same cells won")
                } else {
                    Transformed::Ops(rest)
                }
            } else {
                one(a)
            }
        }
        (
            SetBold { sheet, rect, bold },
            Structural {
                sheet: bs,
                axis,
                at,
                count,
                delete,
                ..
            },
        ) if sheet == bs => {
            let pieces = map_rect(rect, *axis, *at, *count, *delete);
            if pieces.is_empty() {
                Transformed::Dropped("its cells were deleted")
            } else {
                Transformed::Ops(
                    pieces
                        .into_iter()
                        .map(|r| SetBold {
                            sheet: *sheet,
                            rect: r,
                            bold: *bold,
                        })
                        .collect(),
                )
            }
        }
        (SetBold { sheet, .. }, DeleteSheet { sheet: bs, .. }) if sheet == bs => {
            Transformed::Dropped("its sheet was deleted")
        }
        (SetBold { .. }, _) => one(a),

        // ---- a is a structural edit ----
        (
            Structural {
                sheet,
                sheet_name,
                axis,
                at,
                count,
                delete,
            },
            Structural {
                sheet: bs,
                axis: bx,
                at: bat,
                count: bcount,
                delete: bdel,
                ..
            },
        ) if sheet == bs && axis == bx => {
            let mk = |at: usize, count: usize| Structural {
                sheet: *sheet,
                sheet_name: sheet_name.clone(),
                axis: *axis,
                at,
                count,
                delete: *delete,
            };
            match (*delete, *bdel) {
                // insert × insert: shift past an earlier insert at or before us.
                (false, false) => {
                    if *at > *bat || (*at == *bat && later) {
                        one(&mk(at + bcount, *count))
                    } else {
                        one(a)
                    }
                }
                // insert × delete
                (false, true) => {
                    if *at <= *bat {
                        one(a)
                    } else if *at >= bat + bcount {
                        one(&mk(at - bcount, *count))
                    } else {
                        one(&mk(*bat, *count))
                    }
                }
                // delete × insert: split around an insert strictly inside.
                (true, false) => {
                    if *bat <= *at {
                        one(&mk(at + bcount, *count))
                    } else if *bat >= at + count {
                        one(a)
                    } else {
                        let first = bat - at;
                        Transformed::Ops(vec![mk(*at, first), mk(at + bcount, count - first)])
                    }
                }
                // delete × delete: delete only what is left, where it now is.
                (true, true) => {
                    let (a0, a1) = (*at, at + count); // [a0, a1)
                    let (b0, b1) = (*bat, bat + bcount);
                    let overlap = a1.min(b1).saturating_sub(a0.max(b0));
                    let remaining = count - overlap;
                    if remaining == 0 {
                        return Transformed::Dropped("its rows or columns were already deleted");
                    }
                    let start = if a0 < b0 {
                        a0
                    } else if a0 >= b1 {
                        a0 - bcount
                    } else {
                        b0
                    };
                    one(&mk(start, remaining))
                }
            }
        }
        (
            Structural {
                sheet,
                axis,
                at,
                count,
                delete,
                ..
            },
            RenameSheet { sheet: bs, name },
        ) if sheet == bs => one(&Structural {
            sheet: *sheet,
            sheet_name: name.clone(),
            axis: *axis,
            at: *at,
            count: *count,
            delete: *delete,
        }),
        (Structural { sheet, .. }, DeleteSheet { sheet: bs, .. }) if sheet == bs => {
            Transformed::Dropped("its sheet was deleted")
        }
        (Structural { .. }, _) => one(a),

        // ---- sheet operations ----
        (AddSheet { sheet, name, index }, AddSheet { index: bi, .. }) => {
            if *bi < *index || (*bi == *index && later) {
                one(&AddSheet {
                    sheet: *sheet,
                    name: name.clone(),
                    index: index + 1,
                })
            } else {
                one(a)
            }
        }
        (AddSheet { sheet, name, index }, DeleteSheet { index: bi, .. }) => {
            if *bi < *index {
                one(&AddSheet {
                    sheet: *sheet,
                    name: name.clone(),
                    index: index - 1,
                })
            } else {
                one(a)
            }
        }
        (AddSheet { .. }, _) => one(a),

        (RenameSheet { sheet, .. }, RenameSheet { sheet: bs, .. }) if sheet == bs => {
            if later {
                one(a)
            } else {
                Transformed::Dropped("a later rename of the same sheet won")
            }
        }
        (RenameSheet { sheet, .. }, DeleteSheet { sheet: bs, .. }) if sheet == bs => {
            Transformed::Dropped("its sheet was deleted")
        }
        (RenameSheet { .. }, _) => one(a),

        (DeleteSheet { sheet, .. }, DeleteSheet { sheet: bs, .. }) if sheet == bs => {
            Transformed::Dropped("the sheet was already deleted")
        }
        (DeleteSheet { sheet, index }, AddSheet { index: bi, .. }) => {
            if *bi <= *index {
                one(&DeleteSheet {
                    sheet: *sheet,
                    index: index + 1,
                })
            } else {
                one(a)
            }
        }
        (DeleteSheet { .. }, _) => one(a),

        // ---- atomic range ----
        (
            ReplaceRange {
                sheet,
                row,
                col,
                values,
            },
            Structural {
                sheet: bs,
                axis,
                at,
                count,
                delete,
                ..
            },
        ) if sheet == bs => {
            // conflict() already refused edits that cut into the block.
            let (mut r, mut c) = (*row, *col);
            let v = if *axis == Axis::Row { &mut r } else { &mut c };
            if *delete {
                if at + count <= *v {
                    *v -= count;
                }
            } else if *at <= *v {
                *v += count;
            }
            one(&ReplaceRange {
                sheet: *sheet,
                row: r,
                col: c,
                values: values.clone(),
            })
        }
        (ReplaceRange { sheet, .. }, DeleteSheet { sheet: bs, .. }) if sheet == bs => {
            Transformed::Dropped("its sheet was deleted")
        }
        (ReplaceRange { .. }, _) => one(a),
    }
}

fn one(op: &CollabOp) -> Transformed {
    Transformed::Ops(vec![op.clone()])
}

fn edit(sheet_name: &str, axis: Axis, at: usize, count: usize, delete: bool) -> StructuralEdit {
    StructuralEdit {
        sheet_name: sheet_name.to_string(),
        axis: axis.into(),
        at,
        count,
        delete,
    }
}

/// Where index `v` lands after a structural edit; `None` if it was deleted.
pub fn map_index(v: usize, at: usize, count: usize, delete: bool) -> Option<usize> {
    if delete {
        if v >= at + count {
            Some(v - count)
        } else if v >= at {
            None
        } else {
            Some(v)
        }
    } else if v >= at {
        Some(v + count)
    } else {
        Some(v)
    }
}

/// A format rectangle after a structural edit. Inserted lines start
/// unformatted, so an insert strictly inside splits the rectangle.
fn map_rect(rect: &Rect, axis: Axis, at: usize, count: usize, delete: bool) -> Vec<Rect> {
    let (lo, hi) = match axis {
        Axis::Row => (rect.r0, rect.r1),
        Axis::Col => (rect.c0, rect.c1),
    };
    let spans: Vec<(usize, usize)> = if delete {
        let survivors: Vec<usize> = (lo..=hi)
            .filter(|v| map_index(*v, at, count, true).is_some())
            .collect();
        match (survivors.first(), survivors.last()) {
            (Some(f), Some(l)) => vec![(
                map_index(*f, at, count, true).unwrap(),
                map_index(*l, at, count, true).unwrap(),
            )],
            _ => vec![],
        }
    } else if at <= lo {
        vec![(lo + count, hi + count)]
    } else if at > hi {
        vec![(lo, hi)]
    } else {
        vec![(lo, at - 1), (at + count, hi + count)]
    };
    spans
        .into_iter()
        .map(|(s, e)| match axis {
            Axis::Row => Rect::new(s, rect.c0, e, rect.c1),
            Axis::Col => Rect::new(rect.r0, s, rect.r1, e),
        })
        .collect()
}

fn name_key(name: &str) -> String {
    name.trim().to_lowercase()
}

/// Pairs V1 serializes. Returns the reason when `a` and `b` conflict; the
/// caller refuses whichever is later.
fn conflict(a: &CollabOp, b: &CollabOp) -> Option<String> {
    use CollabOp::*;
    // Two sheets may not end up with the same name.
    let named = |op: &CollabOp| -> Option<(SheetKey, String)> {
        match op {
            AddSheet { sheet, name, .. } | RenameSheet { sheet, name } => {
                Some((*sheet, name_key(name)))
            }
            _ => None,
        }
    };
    if let (Some((sa, na)), Some((sb, nb))) = (named(a), named(b)) {
        if sa != sb && na == nb {
            return Some(format!("another sheet was just named \"{}\"", na));
        }
    }
    // A rename concurrent with a row/column edit (on any sheet) is
    // serialized. References are stored by sheet *name* and the engine does
    // not rewrite them on rename, so whether a structural edit adjusts a
    // `Name!A1` reference depends on the name map at the moment it applies;
    // concurrent renames would make that differ between replicas (found by
    // the simulator: same ops, different formula text).
    let renames = |x: &CollabOp, y: &CollabOp| {
        matches!(x, RenameSheet { .. }) && matches!(y, Structural { .. })
    };
    if renames(a, b) || renames(b, a) {
        return Some("a sheet was renamed while rows or columns changed".into());
    }
    // Deleting a sheet while another client inserts or deletes rows/columns
    // on it is serialized. A structural edit also rewrites references to
    // its sheet from *other* sheets, so it cannot simply be dropped when the
    // sheet goes away: the replica that saw the edit first would keep the
    // rewritten references and the one that saw the delete first would not
    // (found by the simulator).
    let deletes_under = |x: &CollabOp, y: &CollabOp| matches!((x, y), (DeleteSheet { sheet: s, .. }, Structural { sheet: t, .. }) if s == t);
    if deletes_under(a, b) || deletes_under(b, a) {
        return Some("the sheet was deleted while its rows or columns changed".into());
    }
    // An insert that touches a concurrent delete (inside it, at its first
    // line, or at the line just after it) is serialized. Formula ranges use
    // grid-line semantics (an insert at a range's first line shifts it, one
    // inside expands it), and a delete can move which line is "first": the
    // two orders then rewrite `SUM(C3:E6)` differently (found by the
    // simulator). Away from the deleted block the pair commutes.
    if let (
        Structural {
            sheet: sa,
            axis: xa,
            at: ka,
            count: na,
            delete: da,
            ..
        },
        Structural {
            sheet: sb,
            axis: xb,
            at: kb,
            count: nb,
            delete: db,
            ..
        },
    ) = (a, b)
    {
        if sa == sb && xa == xb && da != db {
            let (k, r, n) = if *da {
                (*kb, *ka, *na)
            } else {
                (*ka, *kb, *nb)
            };
            if k >= r && k <= r + n {
                return Some("rows or columns were inserted where others were deleted".into());
            }
        }
    }
    // Concurrent deletes of different sheets are serialized, so two clients
    // can never together delete the last sheet.
    if let (DeleteSheet { sheet: sa, .. }, DeleteSheet { sheet: sb, .. }) = (a, b) {
        if sa != sb {
            return Some("another sheet was deleted at the same time".into());
        }
    }
    let atomic_vs = |atomic: &CollabOp, other: &CollabOp| -> Option<String> {
        let rect = atomic.replace_rect()?;
        if atomic.sheet() != other.sheet() {
            return None;
        }
        let hit = match other {
            SetCell { row, col, .. } => rect.contains(*row, *col),
            SetBold { rect: r, .. } => rect.intersects(r),
            ReplaceRange { .. } => rect.intersects(&other.replace_rect().unwrap()),
            Structural {
                axis,
                at,
                count,
                delete,
                ..
            } => {
                let (lo, hi) = match axis {
                    Axis::Row => (rect.r0, rect.r1),
                    Axis::Col => (rect.c0, rect.c1),
                };
                if *delete {
                    *at <= hi && at + count > lo
                } else {
                    *at > lo && *at <= hi
                }
            }
            _ => false,
        };
        hit.then(|| "the range was being replaced at the same time".to_string())
    };
    atomic_vs(a, b).or_else(|| atomic_vs(b, a))
}

/// Why a list transform failed: an op on the later side was refused.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Refusal {
    pub reason: String,
}

/// Transform list `a` (every op of it `order` relative to every op of `b`)
/// past list `b`, returning `(a', b')` with
/// `apply(S, b ++ a') == apply(S, a ++ b')`. Ops inside a list are
/// sequential: each is defined on the state after the previous one, so a
/// later op of one list meets the other list's ops already transformed past
/// the earlier ops. A refusal anywhere refuses the whole exchange (an
/// envelope is atomic).
pub fn transform_lists(
    a: &[CollabOp],
    b: &[CollabOp],
    order: Order,
) -> Result<(Vec<CollabOp>, Vec<CollabOp>), Refusal> {
    if a.is_empty() || b.is_empty() {
        return Ok((a.to_vec(), b.to_vec()));
    }
    if a.len() == 1 && b.len() == 1 {
        let a2 = outcome(transform(&a[0], &b[0], order))?;
        let b2 = outcome(transform(&b[0], &a[0], order.flip()))?;
        return Ok((a2, b2));
    }
    if a.len() > 1 {
        let (a0, b1) = transform_lists(&a[..1], b, order)?;
        let (rest, b2) = transform_lists(&a[1..], &b1, order)?;
        let mut out = a0;
        out.extend(rest);
        return Ok((out, b2));
    }
    let (a1, b0) = transform_lists(a, &b[..1], order)?;
    let (a2, brest) = transform_lists(&a1, &b[1..], order)?;
    let mut out = b0;
    out.extend(brest);
    Ok((a2, out))
}

fn outcome(t: Transformed) -> Result<Vec<CollabOp>, Refusal> {
    match t {
        Transformed::Ops(v) => Ok(v),
        Transformed::Dropped(_) => Ok(Vec::new()),
        Transformed::Refused(reason) => Err(Refusal { reason }),
    }
}
