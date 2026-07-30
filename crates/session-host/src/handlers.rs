//! Protocol request handlers over a bare `Workbook` — host-independent.
//!
//! Ported from the GUI's `handle_session_apply_ops` / `handle_session_inspect`
//! (gpui-app/src/app.rs) 2026-07-29. The GUI wraps these with undo-history
//! recording and view notification; headless hosts use them directly.
//!
//! These constants are the GOVERNING grid bounds (moved here with the
//! validators that enforce them — one owner). gpui-app re-exports them.

use std::collections::HashMap;

use visigrid_engine::cell::CellFormat;
use visigrid_engine::cell_id::CellId;
use visigrid_engine::workbook::Workbook;
use visigrid_protocol::{InspectResult, InspectTarget, Op, OpError, CellInfo, WorkbookInfo, StructureOp};

use crate::bridge::{ApplyOpsError, ApplyOpsRequest, ApplyOpsResponse, InspectError, InspectRequest, InspectResponse};
use crate::wire_ext::CellRef;

// Grid bounds (governing).
pub const NUM_ROWS: usize = 65536;
pub const NUM_COLS: usize = 256;

/// Largest cell count a single session format op (SetNumberFormat/SetStyle)
/// may cover. Bounds memory for undo patches; agents get a precise error
/// telling them to split larger ranges.
pub const MAX_SESSION_FORMAT_CELLS: usize = 250_000;

/// Largest cell count a single Inspect range may cover. Keeps the response
/// comfortably under the protocol's 10 MB message cap.
pub const MAX_SESSION_INSPECT_CELLS: usize = 65_536;

/// A value edit, for hosts that record undo history.
#[derive(Debug, Clone)]
pub struct ValueChange {
    pub row: usize,
    pub col: usize,
    pub old_value: String,
    pub new_value: String,
}

/// A format edit (before/after), for hosts that record undo history.
/// Deduped per cell: first `before` kept, last `after` wins.
#[derive(Debug, Clone)]
pub struct FormatPatch {
    pub row: usize,
    pub col: usize,
    pub before: CellFormat,
    pub after: CellFormat,
}

/// Everything a host needs after an apply: the wire response plus the
/// change lists (keyed by sheet index) for undo recording and broadcast.
#[derive(Debug)]
pub struct ApplyOutcome {
    pub response: ApplyOpsResponse,
    pub value_changes: HashMap<usize, Vec<ValueChange>>,
    pub format_patches: HashMap<usize, Vec<FormatPatch>>,
    pub changed_cells: Vec<CellRef>,
}

/// Validate one session-protocol op against the workbook's sheet list and the
/// governing grid bounds. Returns (code, message, suggestion) on failure.
/// Called for every op BEFORE any op is applied — a failure here rejects the
/// whole batch, which is what makes `atomic` semantics honest.
pub fn validate_session_op(
    op: &Op,
    sheet_count: usize,
) -> Option<(&'static str, String, Option<String>)> {
    let check_sheet = |sheet: usize| -> Option<(&'static str, String, Option<String>)> {
        if sheet >= sheet_count {
            Some((
                "sheet_not_found",
                format!("sheet index {} does not exist (workbook has {} sheet{})",
                    sheet, sheet_count, if sheet_count == 1 { "" } else { "s" }),
                Some("Inspect the workbook to list its sheets".to_string()),
            ))
        } else {
            None
        }
    };
    let check_cell = |row: usize, col: usize| -> Option<(&'static str, String, Option<String>)> {
        if row >= NUM_ROWS || col >= NUM_COLS {
            Some((
                "out_of_bounds",
                format!("cell (row {}, col {}) is outside the grid of {} rows × {} columns",
                    row, col, NUM_ROWS, NUM_COLS),
                Some(format!("Rows are 0..={}, columns 0..={}", NUM_ROWS - 1, NUM_COLS - 1)),
            ))
        } else {
            None
        }
    };
    let check_range = |sr: usize, sc: usize, er: usize, ec: usize| -> Option<(&'static str, String, Option<String>)> {
        if sr > er || sc > ec {
            return Some((
                "invalid_op",
                format!("range start (row {}, col {}) is after its end (row {}, col {})", sr, sc, er, ec),
                Some("Ensure start_row <= end_row and start_col <= end_col".to_string()),
            ));
        }
        check_cell(sr, sc).or_else(|| check_cell(er, ec)).or_else(|| {
            let cells = (er - sr + 1) * (ec - sc + 1);
            if cells > MAX_SESSION_FORMAT_CELLS {
                Some((
                    "cells_limit_exceeded",
                    format!("range covers {} cells; format ops are limited to {} cells per op",
                        cells, MAX_SESSION_FORMAT_CELLS),
                    Some("Split the range into smaller ops in the same batch".to_string()),
                ))
            } else {
                None
            }
        })
    };

    match op {
        Op::SetCellValue { sheet, row, col, .. }
        | Op::SetCellFormula { sheet, row, col, .. }
        | Op::ClearCell { sheet, row, col } => {
            check_sheet(*sheet).or_else(|| check_cell(*row, *col))
        }
        Op::SetNumberFormat { sheet, start_row, start_col, end_row, end_col, format } => {
            check_sheet(*sheet)
                .or_else(|| check_range(*start_row, *start_col, *end_row, *end_col))
                .or_else(|| {
                    let t = format.trim();
                    if t.is_empty() {
                        return Some((
                            "invalid_op",
                            "number format string is empty".to_string(),
                            Some("Use a named format (general, number, currency, percent, date, time, datetime — optionally with :decimals) or an Excel format code like \"#,##0.00\"".to_string()),
                        ));
                    }
                    // A known keyword with an unparseable decimals suffix is a
                    // client mistake — reject rather than store it as a Custom code.
                    if let Some((name, dec)) = t.split_once(':') {
                        let known = matches!(name.trim().to_ascii_lowercase().as_str(),
                            "general" | "number" | "currency" | "percent" | "date" | "time" | "datetime");
                        if known && dec.trim().parse::<u8>().map(|d| d > 10).unwrap_or(true) {
                            return Some((
                                "invalid_op",
                                format!("\"{}\" has an invalid decimals suffix (must be an integer 0..=10)", t),
                                Some("Example: \"number:2\" or \"percent:1\"".to_string()),
                            ));
                        }
                    }
                    None
                })
        }
        Op::SetStyle { sheet, start_row, start_col, end_row, end_col, .. } => {
            check_sheet(*sheet).or_else(|| check_range(*start_row, *start_col, *end_row, *end_col))
        }
    }
}

/// Validate an inspect target against the workbook's sheet list and the
/// governing grid bounds. Returns (code, message) on failure, using the same
/// error taxonomy as the write path.
pub fn validate_inspect_target(
    target: &InspectTarget,
    sheet_count: usize,
) -> Option<(&'static str, String)> {
    let check_sheet = |sheet: usize| -> Option<(&'static str, String)> {
        if sheet >= sheet_count {
            Some((
                "sheet_not_found",
                format!("sheet index {} does not exist (workbook has {} sheet{})",
                    sheet, sheet_count, if sheet_count == 1 { "" } else { "s" }),
            ))
        } else {
            None
        }
    };
    let check_cell = |row: usize, col: usize| -> Option<(&'static str, String)> {
        if row >= NUM_ROWS || col >= NUM_COLS {
            Some((
                "out_of_bounds",
                format!("cell (row {}, col {}) is outside the grid of {} rows × {} columns",
                    row, col, NUM_ROWS, NUM_COLS),
            ))
        } else {
            None
        }
    };

    match target {
        InspectTarget::Workbook => None,
        InspectTarget::Cell { sheet, row, col } => {
            check_sheet(*sheet).or_else(|| check_cell(*row, *col))
        }
        InspectTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
            check_sheet(*sheet)
                .or_else(|| {
                    if start_row > end_row || start_col > end_col {
                        Some((
                            "invalid_op",
                            format!("range start (row {}, col {}) is after its end (row {}, col {})",
                                start_row, start_col, end_row, end_col),
                        ))
                    } else {
                        None
                    }
                })
                .or_else(|| check_cell(*start_row, *start_col))
                .or_else(|| check_cell(*end_row, *end_col))
                .or_else(|| {
                    let cells = (end_row - start_row + 1) * (end_col - start_col + 1);
                    if cells > MAX_SESSION_INSPECT_CELLS {
                        Some((
                            "cells_limit_exceeded",
                            format!("range covers {} cells; inspect is limited to {} cells per request",
                                cells, MAX_SESSION_INSPECT_CELLS),
                        ))
                    } else {
                        None
                    }
                })
        }
    }
}

/// Map a session-protocol number-format string to an engine NumberFormat.
/// Named formats: "general", "number[:decimals]", "currency[:decimals]",
/// "percent[:decimals]", "date", "time", "datetime". Anything else is treated
/// as a raw Excel format code (e.g. "#,##0.00"). Assumes the string already
/// passed validate_session_op.
pub fn parse_session_number_format(s: &str) -> visigrid_engine::cell::NumberFormat {
    use visigrid_engine::cell::{DateStyle, NumberFormat};
    let t = s.trim();
    let (name, dec) = match t.split_once(':') {
        Some((n, d)) => (n.trim(), d.trim().parse::<u8>().ok()),
        None => (t, None),
    };
    match name.to_ascii_lowercase().as_str() {
        "general" => NumberFormat::General,
        "number" => NumberFormat::number(dec.unwrap_or(2)),
        "currency" => NumberFormat::currency(dec.unwrap_or(2)),
        "percent" => NumberFormat::Percent { decimals: dec.unwrap_or(0).min(10) },
        "date" => NumberFormat::Date { style: DateStyle::Short },
        "time" => NumberFormat::Time,
        "datetime" => NumberFormat::DateTime,
        _ => NumberFormat::Custom(t.to_string()),
    }
}

/// Largest row/column count a single structure op may add or remove.
/// Deletes capture their cells for undo, so this bounds that snapshot.
pub const MAX_STRUCTURE_COUNT: usize = 1_000;

/// Resolve a structure op's target sheet against the active sheet.
pub fn structure_target_sheet(op: &StructureOp, active: usize) -> usize {
    match op {
        StructureOp::InsertRows { sheet, .. }
        | StructureOp::DeleteRows { sheet, .. }
        | StructureOp::InsertCols { sheet, .. }
        | StructureOp::DeleteCols { sheet, .. }
        | StructureOp::RenameSheet { sheet, .. } => sheet.unwrap_or(active),
        StructureOp::AddSheet { .. } => active,
    }
}

/// Validate a structure op against the workbook. Returns (code, message,
/// suggestion) on failure — same taxonomy as the cell write path.
pub fn validate_structure_op(
    op: &StructureOp,
    wb: &Workbook,
) -> Option<(&'static str, String, Option<String>)> {
    let sheet_count = wb.sheets().len();
    let target = structure_target_sheet(op, wb.active_sheet_index());
    if !matches!(op, StructureOp::AddSheet { .. }) && target >= sheet_count {
        return Some((
            "sheet_not_found",
            format!("sheet index {} does not exist (workbook has {} sheet{})",
                target, sheet_count, if sheet_count == 1 { "" } else { "s" }),
            Some("Omit `sheet` to target the active sheet".to_string()),
        ));
    }

    let check_span = |at: usize, count: usize, limit: usize, unit: &str| {
        if count == 0 {
            return Some((
                "invalid_op",
                format!("count must be at least 1 {}", unit),
                None,
            ));
        }
        if count > MAX_STRUCTURE_COUNT {
            return Some((
                "cells_limit_exceeded",
                format!("{} {}s in one op; the limit is {}", count, unit, MAX_STRUCTURE_COUNT),
                Some("Split into several calls".to_string()),
            ));
        }
        if at >= limit {
            return Some((
                "out_of_bounds",
                format!("{} index {} is outside the grid (0..={})", unit, at, limit - 1),
                None,
            ));
        }
        if at + count > limit {
            return Some((
                "out_of_bounds",
                format!("{} {}s starting at {} would run past the grid edge ({} {}s total)",
                    count, unit, at, limit, unit),
                Some(format!("The last valid start for {} {}s is {}", count, unit, limit - count)),
            ));
        }
        None
    };

    let check_name = |name: &str, wb: &Workbook, exclude: Option<usize>| {
        let trimmed = name.trim();
        if trimmed.is_empty() {
            return Some((
                "invalid_op",
                "sheet name cannot be empty".to_string(),
                None,
            ));
        }
        if trimmed.chars().count() > 64 {
            return Some((
                "invalid_op",
                "sheet name is longer than 64 characters".to_string(),
                None,
            ));
        }
        let clash = wb.sheets().iter().enumerate().any(|(i, s)| {
            Some(i) != exclude && s.name.eq_ignore_ascii_case(trimmed)
        });
        if clash {
            return Some((
                "invalid_op",
                format!("a sheet named \"{}\" already exists", trimmed),
                Some("Sheet names are compared case-insensitively".to_string()),
            ));
        }
        None
    };

    match op {
        StructureOp::InsertRows { at, count, .. } | StructureOp::DeleteRows { at, count, .. } => {
            check_span(*at, *count, NUM_ROWS, "row")
        }
        StructureOp::InsertCols { at, count, .. } | StructureOp::DeleteCols { at, count, .. } => {
            check_span(*at, *count, NUM_COLS, "column")
        }
        StructureOp::AddSheet { name } => match name {
            Some(n) => check_name(n, wb, None),
            None => None,
        },
        StructureOp::RenameSheet { name, .. } => check_name(name, wb, Some(target)),
    }
}

/// Apply a validated structure op directly to the workbook (headless hosts).
/// GUI hosts route through their own methods instead, so view state — row
/// views, row heights, undo entries — stays consistent.
pub fn apply_structure(wb: &mut Workbook, op: &StructureOp) -> Result<String, String> {
    let active = wb.active_sheet_index();
    let target = structure_target_sheet(op, active);
    use visigrid_engine::structural::Axis;

    // Row/column edits go through Workbook::structural_edit so formulas,
    // validations, and named ranges follow the moved cells — and so an
    // insert that would push data off the grid is refused, not silent.
    let span = |axis: Axis, at: usize, count: usize, delete: bool, wb: &mut Workbook| {
        wb.structural_edit(target, axis, at, count, delete)
    };

    Ok(match op {
        StructureOp::InsertRows { at, count, .. } => {
            match span(Axis::Row, *at, *count, false, wb) {
                Ok(_) => format!("Inserted {} row(s) at row {}", count, at + 1),
                Err(e) => return Err(e),
            }
        }
        StructureOp::DeleteRows { at, count, .. } => {
            match span(Axis::Row, *at, *count, true, wb) {
                Ok(_) => format!("Deleted {} row(s) at row {}", count, at + 1),
                Err(e) => return Err(e),
            }
        }
        StructureOp::InsertCols { at, count, .. } => {
            match span(Axis::Col, *at, *count, false, wb) {
                Ok(_) => format!("Inserted {} column(s) at column {}", count, at + 1),
                Err(e) => return Err(e),
            }
        }
        StructureOp::DeleteCols { at, count, .. } => {
            match span(Axis::Col, *at, *count, true, wb) {
                Ok(_) => format!("Deleted {} column(s) at column {}", count, at + 1),
                Err(e) => return Err(e),
            }
        }
        StructureOp::AddSheet { name } => {
            let idx = match name {
                Some(n) => wb.add_sheet_named(n.trim()).unwrap_or_else(|| wb.add_sheet()),
                None => wb.add_sheet(),
            };
            wb.bump_revision_for_structure();
            format!("Added sheet \"{}\"", wb.sheets()[idx].name)
        }
        StructureOp::RenameSheet { name, .. } => {
            let old = wb.sheets().get(target).map(|s| s.name.clone()).unwrap_or_default();
            wb.rename_sheet(target, name.trim());
            wb.bump_revision_for_structure();
            format!("Renamed sheet \"{}\" to \"{}\"", old, name.trim())
        }
    })
}

/// Apply an ops batch to the workbook. The whole batch is validated up front
/// against the real grid bounds and sheet list; any invalid op rejects the
/// entire request (regardless of `atomic`) before anything is applied — by
/// the time we touch the workbook, no op can fail, so a success response
/// never lies. One batch = one recalc = one revision increment.
pub fn apply_ops(wb: &mut Workbook, req: &ApplyOpsRequest) -> ApplyOutcome {
    let current_rev = wb.revision();

    let reject = |error: Option<ApplyOpsError>, total: usize| ApplyOutcome {
        response: ApplyOpsResponse {
            applied: 0,
            total,
            current_revision: current_rev,
            error,
        },
        value_changes: HashMap::new(),
        format_patches: HashMap::new(),
        changed_cells: Vec::new(),
    };

    // Optimistic concurrency check
    if let Some(expected) = req.expected_revision {
        if expected != current_rev {
            return reject(
                Some(ApplyOpsError::RevisionMismatch { expected, actual: current_rev }),
                req.ops.len(),
            );
        }
    }

    if req.ops.is_empty() {
        return reject(None, 0);
    }

    // Up-front validation of the entire batch
    let sheet_count = wb.sheets().len();
    for (i, op) in req.ops.iter().enumerate() {
        if let Some((code, message, suggestion)) = validate_session_op(op, sheet_count) {
            return reject(
                Some(ApplyOpsError::OpFailed(OpError {
                    code: code.to_string(),
                    message,
                    op_index: i,
                    suggestion,
                })),
                req.ops.len(),
            );
        }
    }

    // Apply within a single batch guard: one recalc, one revision increment.
    let mut applied = 0;
    let mut value_changes: HashMap<usize, Vec<ValueChange>> = HashMap::new();
    // Format patches deduped per cell (first `before` kept, last `after`
    // wins) so a multi-op request undoes correctly.
    let mut format_acc: HashMap<usize, (Vec<FormatPatch>, HashMap<(usize, usize), usize>)> =
        HashMap::new();

    {
        let mut guard = wb.batch_guard();

        let mut push_patch = |acc: &mut HashMap<usize, (Vec<FormatPatch>, HashMap<(usize, usize), usize>)>,
                              sheet_idx: usize,
                              patch: FormatPatch| {
            let (patches, index) = acc.entry(sheet_idx).or_default();
            match index.get(&(patch.row, patch.col)) {
                Some(&i) => patches[i].after = patch.after,
                None => {
                    index.insert((patch.row, patch.col), patches.len());
                    patches.push(patch);
                }
            }
        };

        for op in req.ops.iter() {
            match op {
                Op::SetCellValue { sheet, row, col, value } => {
                    let old_value = guard.sheets()[*sheet].get_raw(*row, *col);
                    value_changes.entry(*sheet).or_default().push(ValueChange {
                        row: *row, col: *col, old_value, new_value: value.clone(),
                    });
                    guard.set_cell_value_tracked(*sheet, *row, *col, value);
                    applied += 1;
                }
                Op::SetCellFormula { sheet, row, col, formula } => {
                    let old_value = guard.sheets()[*sheet].get_raw(*row, *col);
                    value_changes.entry(*sheet).or_default().push(ValueChange {
                        row: *row, col: *col, old_value, new_value: formula.clone(),
                    });
                    guard.set_cell_value_tracked(*sheet, *row, *col, formula);
                    applied += 1;
                }
                Op::ClearCell { sheet, row, col } => {
                    let old_value = guard.sheets()[*sheet].get_raw(*row, *col);
                    value_changes.entry(*sheet).or_default().push(ValueChange {
                        row: *row, col: *col, old_value, new_value: String::new(),
                    });
                    guard.clear_cell_tracked(*sheet, *row, *col);
                    applied += 1;
                }
                Op::SetNumberFormat { sheet, start_row, start_col, end_row, end_col, format } => {
                    let nf = parse_session_number_format(format);
                    let sheet_id = guard.sheets()[*sheet].id;
                    for r in *start_row..=*end_row {
                        for c in *start_col..=*end_col {
                            let s = guard.sheet_mut(*sheet).expect("validated sheet index");
                            let before = s.get_format(r, c);
                            s.set_number_format(r, c, nf.clone());
                            let after = s.get_format(r, c);
                            if after != before {
                                push_patch(&mut format_acc, *sheet, FormatPatch { row: r, col: c, before, after });
                                guard.note_format_changed(CellId::new(sheet_id, r, c));
                            }
                        }
                    }
                    applied += 1;
                }
                Op::SetStyle { sheet, start_row, start_col, end_row, end_col, bold, italic, underline } => {
                    let sheet_id = guard.sheets()[*sheet].id;
                    for r in *start_row..=*end_row {
                        for c in *start_col..=*end_col {
                            let s = guard.sheet_mut(*sheet).expect("validated sheet index");
                            let before = s.get_format(r, c);
                            if let Some(b) = bold { s.set_bold(r, c, *b); }
                            if let Some(b) = italic { s.set_italic(r, c, *b); }
                            if let Some(b) = underline { s.set_underline(r, c, *b); }
                            let after = s.get_format(r, c);
                            if after != before {
                                push_patch(&mut format_acc, *sheet, FormatPatch { row: r, col: c, before, after });
                                guard.note_format_changed(CellId::new(sheet_id, r, c));
                            }
                        }
                    }
                    applied += 1;
                }
            }
        }
    } // guard dropped: single recalc + revision increment

    let mut changed_cells: Vec<CellRef> = value_changes
        .iter()
        .flat_map(|(sheet_idx, changes)| {
            changes.iter().map(move |c| CellRef { sheet: *sheet_idx, row: c.row, col: c.col })
        })
        .collect();
    let format_patches: HashMap<usize, Vec<FormatPatch>> =
        format_acc.into_iter().map(|(k, (v, _))| (k, v)).collect();
    changed_cells.extend(format_patches.iter().flat_map(|(sheet_idx, patches)| {
        patches.iter().map(move |p| CellRef { sheet: *sheet_idx, row: p.row, col: p.col })
    }));

    ApplyOutcome {
        response: ApplyOpsResponse {
            applied,
            total: req.ops.len(),
            current_revision: wb.revision(),
            error: None,
        },
        value_changes,
        format_patches,
        changed_cells,
    }
}

/// Handle an inspect request. `title` is the workbook display name (host-owned).
/// Bad sheet indexes and out-of-bounds coordinates are errors, never silent
/// redirects to the active sheet.
pub fn inspect(wb: &Workbook, req: &InspectRequest, title: &str) -> InspectResponse {
    let current_rev = wb.revision();

    if let Some((code, message)) = validate_inspect_target(&req.target, wb.sheets().len()) {
        return InspectResponse {
            current_revision: current_rev,
            result: Err(InspectError { code: code.to_string(), message }),
        };
    }

    let cell_info = |sheet: &visigrid_engine::sheet::Sheet, row: usize, col: usize| {
        let display = sheet.get_display(row, col);
        let raw = sheet.get_raw(row, col);
        let formula = if raw.starts_with('=') { Some(raw.clone()) } else { None };
        CellInfo { raw, display, formula }
    };

    let result = match &req.target {
        InspectTarget::Cell { sheet, row, col } => {
            InspectResult::Cell(cell_info(&wb.sheets()[*sheet], *row, *col))
        }
        InspectTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
            let sheet_data = &wb.sheets()[*sheet];
            let mut cells = Vec::new();
            for r in *start_row..=*end_row {
                for c in *start_col..=*end_col {
                    cells.push(cell_info(sheet_data, r, c));
                }
            }
            InspectResult::Range { cells }
        }
        InspectTarget::Workbook => InspectResult::Workbook(WorkbookInfo {
            sheet_count: wb.sheets().len(),
            active_sheet: wb.active_sheet_index(),
            title: title.to_string(),
        }),
    };

    InspectResponse { current_revision: current_rev, result: Ok(result) }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::bridge::ApplyOpsRequest;

    fn write(sheet: usize, row: usize, col: usize, value: &str) -> Op {
        Op::SetCellValue { sheet, row, col, value: value.to_string() }
    }

    fn req(ops: Vec<Op>) -> ApplyOpsRequest {
        ApplyOpsRequest {
            request_id: "t".into(),
            batch_name: "test".into(),
            atomic: true,
            expected_revision: None,
            ops,
            client: None,
        }
    }

    #[test]
    fn structure_validation_matrix() {
        use visigrid_protocol::StructureOp;
        let mut wb = Workbook::new();

        let ok = |op: &StructureOp, wb: &Workbook| validate_structure_op(op, wb).is_none();
        let code = |op: &StructureOp, wb: &Workbook| validate_structure_op(op, wb).unwrap().0;

        assert!(ok(&StructureOp::InsertRows { sheet: None, at: 0, count: 1 }, &wb));
        assert!(ok(&StructureOp::InsertRows { sheet: None, at: NUM_ROWS - 1, count: 1 }, &wb));
        // Past the grid edge, and the message says where the last valid start is.
        assert_eq!(code(&StructureOp::InsertRows { sheet: None, at: NUM_ROWS - 1, count: 2 }, &wb), "out_of_bounds");
        assert_eq!(code(&StructureOp::InsertRows { sheet: None, at: NUM_ROWS, count: 1 }, &wb), "out_of_bounds");
        assert_eq!(code(&StructureOp::InsertRows { sheet: None, at: 0, count: 0 }, &wb), "invalid_op");
        assert_eq!(
            code(&StructureOp::DeleteRows { sheet: None, at: 0, count: MAX_STRUCTURE_COUNT + 1 }, &wb),
            "cells_limit_exceeded"
        );
        assert_eq!(code(&StructureOp::InsertCols { sheet: None, at: NUM_COLS, count: 1 }, &wb), "out_of_bounds");
        assert_eq!(code(&StructureOp::InsertRows { sheet: Some(7), at: 0, count: 1 }, &wb), "sheet_not_found");

        // Sheet names: unique case-insensitively, non-empty, bounded.
        assert!(ok(&StructureOp::AddSheet { name: Some("Summary".into()) }, &wb));
        assert_eq!(code(&StructureOp::AddSheet { name: Some("  ".into()) }, &wb), "invalid_op");
        assert_eq!(code(&StructureOp::AddSheet { name: Some("sheet1".into()) }, &wb), "invalid_op");
        // Renaming a sheet to its own name is fine (self is excluded).
        assert!(ok(&StructureOp::RenameSheet { sheet: Some(0), name: "Sheet1".into() }, &wb));

        // Apply path: insert shifts a formula's target and it recomputes.
        wb.sheet_mut(0).unwrap().set_value(0, 0, "10");
        wb.sheet_mut(0).unwrap().set_value(1, 0, "=A1*2");
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
        assert_eq!(wb.sheets()[0].get_display(1, 0), "20");

        let desc = apply_structure(&mut wb, &StructureOp::InsertRows { sheet: None, at: 0, count: 2 }).unwrap();
        assert!(desc.contains("Inserted 2 row"));
        assert_eq!(wb.sheets()[0].get_display(2, 0), "10", "value shifted down 2");
        // Reference adjustment (2026-07-30): the formula follows its content.
        assert_eq!(wb.sheets()[0].get_raw(3, 0), "=A3*2", "reference adjusted");
        assert_eq!(wb.sheets()[0].get_display(3, 0), "20");

        let desc = apply_structure(&mut wb, &StructureOp::AddSheet { name: Some("Summary".into()) }).unwrap();
        assert!(desc.contains("Summary"));
        assert_eq!(wb.sheets().len(), 2);
        // The new sheet's name is now taken.
        assert_eq!(code(&StructureOp::AddSheet { name: Some("SUMMARY".into()) }, &wb), "invalid_op");
    }

    #[test]
    fn apply_and_inspect_headless() {
        let mut wb = Workbook::new();
        let out = apply_ops(&mut wb, &req(vec![
            write(0, 0, 0, "10"),
            write(0, 1, 0, "32"),
            Op::SetCellFormula { sheet: 0, row: 2, col: 0, formula: "=A1+A2".into() },
        ]));
        assert_eq!(out.response.applied, 3);
        assert!(out.response.error.is_none());
        assert_eq!(out.response.current_revision, 1);
        assert_eq!(out.changed_cells.len(), 3);
        assert_eq!(out.value_changes[&0].len(), 3);

        let resp = inspect(&wb, &InspectRequest {
            request_id: "t".into(),
            target: InspectTarget::Cell { sheet: 0, row: 2, col: 0 },
        }, "Test");
        match resp.result.unwrap() {
            InspectResult::Cell(info) => assert_eq!(info.display, "42"),
            other => panic!("unexpected: {:?}", other),
        }
    }

    #[test]
    fn ghost_cell_rejected_without_host() {
        let mut wb = Workbook::new();
        let out = apply_ops(&mut wb, &req(vec![write(0, 70_000, 0, "ghost")]));
        let err = out.response.error.unwrap();
        match err {
            ApplyOpsError::OpFailed(e) => assert_eq!(e.code, "out_of_bounds"),
            other => panic!("unexpected: {:?}", other),
        }
        assert_eq!(wb.revision(), 0, "nothing applied");
    }

    #[test]
    fn revision_mismatch_headless() {
        let mut wb = Workbook::new();
        apply_ops(&mut wb, &req(vec![write(0, 0, 0, "x")]));
        let mut r = req(vec![write(0, 0, 1, "y")]);
        r.expected_revision = Some(0);
        let out = apply_ops(&mut wb, &r);
        assert!(matches!(out.response.error, Some(ApplyOpsError::RevisionMismatch { .. })));
    }

    #[test]
    fn format_ops_bump_revision_and_patch() {
        let mut wb = Workbook::new();
        let out = apply_ops(&mut wb, &req(vec![Op::SetStyle {
            sheet: 0, start_row: 0, start_col: 0, end_row: 0, end_col: 2,
            bold: Some(true), italic: None, underline: None,
        }]));
        assert_eq!(out.response.applied, 1);
        assert_eq!(out.response.current_revision, 1);
        assert_eq!(out.format_patches[&0].len(), 3);
        assert!(wb.sheets()[0].get_format(0, 1).bold);
    }
}

#[cfg(test)]
mod validation_tests {
    use super::*;
    use visigrid_protocol::{InspectTarget, Op};
    use visigrid_engine::cell::NumberFormat;

    fn set_value(sheet: usize, row: usize, col: usize) -> Op {
        Op::SetCellValue { sheet, row, col, value: "x".to_string() }
    }

    #[test]
    fn valid_ops_pass() {
        assert!(validate_session_op(&set_value(0, 0, 0), 1).is_none());
        assert!(validate_session_op(&set_value(0, NUM_ROWS - 1, NUM_COLS - 1), 1).is_none());
        assert!(validate_session_op(&set_value(2, 5, 5), 3).is_none());
    }

    #[test]
    fn out_of_bounds_cell_rejected() {
        // The exact ghost-cell case: row 70,000 on the 65,536-row grid
        let (code, msg, _) = validate_session_op(&set_value(0, 70_000, 0), 1).unwrap();
        assert_eq!(code, "out_of_bounds");
        assert!(msg.contains("70000") && msg.contains("65536"));
        let (code, _, _) = validate_session_op(&set_value(0, 0, NUM_COLS), 1).unwrap();
        assert_eq!(code, "out_of_bounds");
    }

    #[test]
    fn invalid_sheet_rejected_not_redirected() {
        let (code, msg, _) = validate_session_op(&set_value(5, 0, 0), 1).unwrap();
        assert_eq!(code, "sheet_not_found");
        assert!(msg.contains("5"));
    }

    #[test]
    fn format_range_checks() {
        let style = |sr: usize, sc: usize, er: usize, ec: usize| Op::SetStyle {
            sheet: 0, start_row: sr, start_col: sc, end_row: er, end_col: ec,
            bold: Some(true), italic: None, underline: None,
        };
        assert!(validate_session_op(&style(0, 0, 9, 9), 1).is_none());
        let (code, _, _) = validate_session_op(&style(9, 0, 0, 9), 1).unwrap();
        assert_eq!(code, "invalid_op");
        let (code, _, _) = validate_session_op(&style(0, 0, NUM_ROWS, 0), 1).unwrap();
        assert_eq!(code, "out_of_bounds");
        // Whole grid exceeds the per-op cap
        let (code, _, _) = validate_session_op(&style(0, 0, NUM_ROWS - 1, NUM_COLS - 1), 1).unwrap();
        assert_eq!(code, "cells_limit_exceeded");
        // One full column (65,536 cells) is comfortably under the cap
        assert!(NUM_ROWS <= MAX_SESSION_FORMAT_CELLS);
        assert!(validate_session_op(&style(0, 0, NUM_ROWS - 1, 0), 1).is_none());
    }

    #[test]
    fn number_format_string_checks() {
        let nf = |format: &str| Op::SetNumberFormat {
            sheet: 0, start_row: 0, start_col: 0, end_row: 0, end_col: 0,
            format: format.to_string(),
        };
        assert!(validate_session_op(&nf("currency"), 1).is_none());
        assert!(validate_session_op(&nf("number:2"), 1).is_none());
        assert!(validate_session_op(&nf("#,##0.00"), 1).is_none());
        let (code, _, _) = validate_session_op(&nf(""), 1).unwrap();
        assert_eq!(code, "invalid_op");
        let (code, _, _) = validate_session_op(&nf("number:abc"), 1).unwrap();
        assert_eq!(code, "invalid_op");
        let (code, _, _) = validate_session_op(&nf("percent:99"), 1).unwrap();
        assert_eq!(code, "invalid_op");
    }

    #[test]
    fn inspect_target_checks() {
        let range = |sheet: usize, sr: usize, sc: usize, er: usize, ec: usize| InspectTarget::Range {
            sheet, start_row: sr, start_col: sc, end_row: er, end_col: ec,
        };
        assert!(validate_inspect_target(&InspectTarget::Workbook, 1).is_none());
        assert!(validate_inspect_target(&InspectTarget::Cell { sheet: 0, row: 0, col: 0 }, 1).is_none());

        // Bad sheet is an error, not a redirect to the active sheet
        let (code, _) = validate_inspect_target(&InspectTarget::Cell { sheet: 3, row: 0, col: 0 }, 1).unwrap();
        assert_eq!(code, "sheet_not_found");
        let (code, _) = validate_inspect_target(&InspectTarget::Cell { sheet: 0, row: 70_000, col: 0 }, 1).unwrap();
        assert_eq!(code, "out_of_bounds");

        assert!(validate_inspect_target(&range(0, 0, 0, 19, 9), 1).is_none());
        let (code, _) = validate_inspect_target(&range(0, 5, 0, 0, 9), 1).unwrap();
        assert_eq!(code, "invalid_op");
        let (code, _) = validate_inspect_target(&range(0, 0, 0, NUM_ROWS - 1, NUM_COLS - 1), 1).unwrap();
        assert_eq!(code, "cells_limit_exceeded");
        // One full column is exactly at the cap
        assert_eq!(NUM_ROWS, MAX_SESSION_INSPECT_CELLS);
        assert!(validate_inspect_target(&range(0, 0, 0, NUM_ROWS - 1, 0), 1).is_none());
    }

    #[test]
    fn number_format_parsing() {
        assert_eq!(parse_session_number_format("general"), NumberFormat::General);
        assert_eq!(parse_session_number_format("number:3"), NumberFormat::number(3));
        assert_eq!(parse_session_number_format("Currency"), NumberFormat::currency(2));
        assert_eq!(parse_session_number_format("percent:1"), NumberFormat::Percent { decimals: 1 });
        assert_eq!(
            parse_session_number_format("#,##0.00"),
            NumberFormat::Custom("#,##0.00".to_string())
        );
    }
}

