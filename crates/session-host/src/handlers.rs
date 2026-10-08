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
use visigrid_engine::workbook::{Recalculated, Workbook};
use visigrid_protocol::{InspectResult, InspectTarget, Op, OpError, CellInfo, WorkbookInfo, StructureOp};

use crate::bridge::{ApplyOpsError, ApplyOpsRequest, ApplyOpsResponse, InspectError, InspectRequest, InspectResponse};
use crate::wire_ext::CellRef;

// Grid bounds, owned by the engine so every crate agrees.
pub use visigrid_engine::sheet::{NUM_COLS, NUM_ROWS};

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
    pub guarded_commit: Option<visigrid_engine::workbook::GuardedStructureCommit>,
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
        | StructureOp::RenameSheet { sheet, .. }
        | StructureOp::CreatePivot { sheet, .. } => sheet.unwrap_or(active),
        StructureOp::AddSheet { .. } | StructureOp::RefreshPivot { .. } | StructureOp::RefreshRecipeTable { .. } => active,
    }
}

/// Parse "A1" into 0-based (row, col).
fn parse_a1_cell(s: &str) -> Option<(usize, usize)> {
    let s = s.trim().replace('$', "");
    let split = s.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = s.split_at(split);
    if letters.is_empty() || !letters.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    let col = letters.to_ascii_uppercase().bytes().try_fold(0usize, |acc, b| {
        acc.checked_mul(26)?.checked_add((b - b'A' + 1) as usize)
    })?;
    let row: usize = digits.parse().ok()?;
    (row >= 1 && col >= 1).then(|| (row - 1, col - 1))
}

/// Resolve a `create_pivot` op against the workbook: the source rectangle
/// and a definition built from header names. Shared by every host so a GUI
/// window and `vgrid serve` refuse exactly the same requests.
pub fn resolve_create_pivot(
    op: &StructureOp,
    wb: &Workbook,
) -> Result<(visigrid_engine::pivot::PivotSource, visigrid_engine::pivot::PivotDefinition), (&'static str, String)> {
    use visigrid_engine::pivot::{Aggregation, PivotDefinition, PivotSource};
    let StructureOp::CreatePivot { source, rows, column, values, .. } = op else {
        return Err(("invalid_op", "not a create_pivot op".into()));
    };
    let named_table = source.as_deref().and_then(|name| wb.table_by_name(name));
    let target = named_table.and_then(|(id, _)| wb.sheet_index_by_id(id)).unwrap_or_else(|| structure_target_sheet(op, wb.active_sheet_index()));
    let sheet = wb.sheets().get(target).ok_or_else(|| {
        ("sheet_not_found", format!("sheet index {} does not exist (workbook has {} sheets)", target, wb.sheets().len()))
    })?;
    let (r0, c0, r1, c1) = match source {
        _ if named_table.is_some() => {
            let r = named_table.unwrap().1.range;
            (r.start_row, r.start_col, r.end_row, r.end_col)
        }
        Some(range) => {
            let (a, b) = range.split_once(':').unwrap_or((range.as_str(), range.as_str()));
            let bad = || ("invalid_op", format!("source \"{}\" is not a Table name or an A1 range like A1:D100", range));
            let (ar, ac) = parse_a1_cell(a).ok_or_else(bad)?;
            let (br, bc) = parse_a1_cell(b).ok_or_else(bad)?;
            (ar.min(br), ac.min(bc), ar.max(br), ac.max(bc))
        }
        None => {
            let (mr, mc) = sheet.data_extent();
            (0, 0, mr, mc)
        }
    };
    if r1 >= NUM_ROWS || c1 >= NUM_COLS {
        return Err(("out_of_bounds", "the source runs past the grid edge".into()));
    }
    if r1 <= r0 {
        return Err(("invalid_op", "a pivot source needs a header row and at least one data row".into()));
    }
    if let Some(p) = sheet.pivots.iter().find(|p| p.intersects(r0, c0, r1, c1)) {
        return Err(("invalid_op", format!("the source overlaps {}'s output; point it at the data instead", p.name)));
    }
    let headers: Vec<String> = (c0..=c1).map(|c| sheet.get_display(r0, c).trim().to_string()).collect();
    let values = values
        .iter()
        .map(|v| match v.aggregation.as_deref().map(str::trim).filter(|a| !a.is_empty()) {
            None => Ok((None, v.field.clone())),
            Some(a) => Aggregation::parse(a).map(|a| (Some(a), v.field.clone())).ok_or_else(|| {
                ("invalid_op", format!("unknown aggregation \"{a}\" (use sum, count, distinct_count, average, min or max)"))
            }),
        })
        .collect::<Result<Vec<_>, _>>()?;
    let source = PivotSource { table_id: named_table.map(|(_, t)| t.id), sheet_id: sheet.id, start_row: r0 as u32, start_col: c0 as u32, end_row: r1 as u32, end_col: c1 as u32 };
    let profile = visigrid_engine::pivot::column_profile(sheet, &source);
    let mut definition = PivotDefinition::from_names(&headers, rows, column.as_deref(), &values, &profile)
        .map_err(|m| ("invalid_op", m))?;
    if let Some((_, table)) = named_table {
        for field in definition.rows.iter_mut().chain(definition.column.iter_mut()).chain(definition.values.iter_mut().map(|v| &mut v.field)) {
            field.column_id = Some(table.columns[field.offset as usize].id);
        }
    }
    visigrid_engine::pivot::validate(&definition, &headers).map_err(|e| ("invalid_op", e.to_string()))?;
    // Asked-for Sum/Average/Min/Max of a column with no numbers would show
    // zeros or errors: say so instead.
    for v in &definition.values {
        let numeric_only = matches!(v.aggregation, Aggregation::Sum | Aggregation::Average | Aggregation::Min | Aggregation::Max);
        if numeric_only && !visigrid_engine::pivot::column_has_numbers(sheet, &source, v.field.offset) {
            return Err(("invalid_op", format!(
                "\"{}\" has no numbers to {}; use count or distinct_count, or omit the aggregation",
                v.field.header,
                v.aggregation.label().to_lowercase()
            )));
        }
    }
    Ok((source, definition))
}

/// The pivots a `refresh_pivot` op names: one by name or id, or all.
pub fn resolve_refresh_pivots(pivot: Option<&str>, wb: &Workbook) -> Result<Vec<u64>, (&'static str, String)> {
    match pivot {
        Some(name) => wb
            .find_pivot_by_name(name)
            .map(|(_, t)| vec![t.id])
            .ok_or_else(|| {
                let names: Vec<String> = wb.pivots().iter().map(|(_, t)| t.name.clone()).collect();
                ("invalid_op", if names.is_empty() {
                    "this workbook has no pivot tables".to_string()
                } else {
                    format!("no pivot named \"{}\" (pivots: {})", name, names.join(", "))
                })
            }),
        None => {
            let ids: Vec<u64> = wb.pivots().iter().map(|(_, t)| t.id).collect();
            if ids.is_empty() {
                Err(("invalid_op", "this workbook has no pivot tables".to_string()))
            } else {
                Ok(ids)
            }
        }
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
    if !matches!(op, StructureOp::AddSheet { .. } | StructureOp::RefreshPivot { .. } | StructureOp::RefreshRecipeTable { .. })
        && target >= sheet_count
    {
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
        let limit = 31;
        if trimmed.chars().count() > limit {
            return Some((
                "invalid_op",
                format!("sheet name is longer than {limit} characters"),
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
        StructureOp::CreatePivot { .. } => resolve_create_pivot(op, wb).err().map(|(c, m)| (c, m, None)),
        StructureOp::RefreshPivot { pivot } => {
            resolve_refresh_pivots(pivot.as_deref(), wb).err().map(|(c, m)| (c, m, None))
        }
        // The desktop host resolves the Table and the recipe itself
        StructureOp::RefreshRecipeTable { .. } => None,
    }
}

/// Apply a validated structure op directly to the workbook (headless hosts).
/// GUI hosts route through their own methods instead, so view state — row
/// views, row heights, undo entries — stays consistent.
pub fn apply_structure(wb: &mut Workbook, op: &StructureOp) -> Result<String, String> {
    wb.ensure_writable()?;
    if wb.has_table_criteria()
        && !matches!(
            op,
            StructureOp::InsertRows { .. }
                | StructureOp::DeleteRows { .. }
                | StructureOp::InsertCols { .. }
                | StructureOp::DeleteCols { .. }
                | StructureOp::RenameSheet { .. }
                | StructureOp::AddSheet { .. }
        )
    {
        return Err("Clear Table criteria before sheet or pivot automation.".into());
    }
    let active = wb.active_sheet_index();
    let target = structure_target_sheet(op, active);
    use visigrid_engine::structural::Axis;

    // Row/column edits go through Workbook::structural_edit so formulas,
    // validations, and named ranges follow the moved cells — and so an
    // insert that would push data off the grid is refused, not silent.
    let span = |axis: Axis, at: usize, count: usize, delete: bool, wb: &mut Workbook| {
        if wb.has_table_criteria() {
            let (candidate, _) = wb.prepare_guarded_structure(
                target,
                vec![visigrid_engine::workbook::StructureStep {
                    axis,
                    at,
                    count,
                    delete,
                }],
            )?;
            wb.restore_snapshot_monotonic(&candidate);
            Ok(Vec::new())
        } else {
            wb.structural_edit(target, axis, at, count, delete)
        }
    };

    Ok(match op {
        StructureOp::InsertRows { at, count, .. } => {
            match span(Axis::Row, *at, *count, false, wb) {
                Ok(_) => format!("Inserted {} row(s) at row {}", count, at + 1),
                Err(e) => return Err(e),
            }
        }
        StructureOp::DeleteRows { at, count, .. } => match span(Axis::Row, *at, *count, true, wb) {
            Ok(_) => format!("Deleted {} row(s) at row {}", count, at + 1),
            Err(e) => return Err(e),
        },
        StructureOp::InsertCols { at, count, .. } => {
            match span(Axis::Col, *at, *count, false, wb) {
                Ok(_) => format!("Inserted {} column(s) at column {}", count, at + 1),
                Err(e) => return Err(e),
            }
        }
        StructureOp::DeleteCols { at, count, .. } => match span(Axis::Col, *at, *count, true, wb) {
            Ok(_) => format!("Deleted {} column(s) at column {}", count, at + 1),
            Err(e) => return Err(e),
        },
        StructureOp::AddSheet { name } => {
            let (candidate, _) = wb.prepare_sheet_add(name.as_deref())?;
            let idx = candidate.sheet_count() - 1;
            wb.restore_snapshot_monotonic(&candidate);
            format!("Added sheet \"{}\"", wb.sheets()[idx].name)
        }
        StructureOp::RenameSheet { name, .. } => {
            let sheet = wb.sheets().get(target).ok_or("Sheet does not exist.")?;
            let old = sheet.name.clone();
            let (candidate, commit) = wb.prepare_sheet_rename(sheet.id, &old, name)?;
            if !commit.is_empty() {
                wb.restore_snapshot_monotonic(&candidate);
            }
            format!("Renamed sheet \"{}\" to \"{}\"", old, name.trim())
        }
        StructureOp::CreatePivot { .. } => {
            let (source, definition) = resolve_create_pivot(op, wb).map_err(|(_, m)| m)?;
            let (id, idx) = wb.create_pivot(source, definition)?;
            let name = wb
                .find_pivot(id)
                .map(|(_, t)| t.name.clone())
                .unwrap_or_default();
            let (rows, cols) = wb
                .find_pivot(id)
                .and_then(|(_, t)| t.extent)
                .unwrap_or((0, 0));
            format!(
                "Created {} on sheet \"{}\" (index {}): {} × {}",
                name,
                wb.sheets()[idx].name,
                idx,
                rows,
                cols
            )
        }
        StructureOp::RefreshRecipeTable { .. } => {
            return Err("refreshing a recipe-linked Table needs the VisiGrid desktop app (recipes are approved there); run the recipe with `vgrid recipe run` instead".into());
        }
        StructureOp::RefreshPivot { pivot } => {
            let ids = resolve_refresh_pivots(pivot.as_deref(), wb).map_err(|(_, m)| m)?;
            let mut done = Vec::new();
            for id in ids {
                let name = wb
                    .find_pivot(id)
                    .map(|(_, t)| t.name.clone())
                    .unwrap_or_default();
                match wb.refresh_pivot(id) {
                    Ok((rows, cols)) => done.push(format!("{} ({} × {})", name, rows, cols)),
                    // Earlier pivots stay refreshed (each is valid on its own);
                    // say which, as the desktop does.
                    Err(m) if done.is_empty() => return Err(format!("{}: {}", name, m)),
                    Err(m) => {
                        return Err(format!(
                            "{}: {} (already refreshed: {})",
                            name,
                            m,
                            done.join(", ")
                        ))
                    }
                }
            }
            format!("Refreshed {}", done.join(", "))
        }
    })
}

/// Apply an ops batch to the workbook. The whole batch is validated up front
/// against the real grid bounds and sheet list; any invalid op rejects the
/// entire request (regardless of `atomic`) before anything is applied — by
/// the time we touch the workbook, no op can fail, so a success response
/// never lies. One batch = one recalc = one revision increment.
/// Table-aware requests stage every operation and postflight before publication.
pub fn apply_ops(wb: &mut Workbook, req: &ApplyOpsRequest) -> ApplyOutcome {
    if !wb.has_table_criteria() {
        return apply_ops_inner(wb, req);
    }
    let reject = |code: &str, message: String, op_index: usize| ApplyOutcome {
        guarded_commit: None,
        response: ApplyOpsResponse {
            applied: 0,
            total: req.ops.len(),
            current_revision: wb.revision(),
            error: Some(ApplyOpsError::OpFailed(OpError {
                code: code.into(),
                message,
                op_index,
                suggestion: None,
            })),
            warnings: vec![],
        },
        value_changes: HashMap::new(),
        format_patches: HashMap::new(),
        changed_cells: vec![],
    };
    // Preserve protocol errors (including revision mismatch) before view checks.
    if req.expected_revision.is_some_and(|r| r != wb.revision())
        || wb.ensure_writable().is_err()
        || req.ops.is_empty()
    {
        return apply_ops_inner(wb, req);
    }
    if let Err(e) = wb.validate_saved_table_views() {
        return reject("table_view_unsafe", e, 0);
    }
    let mut targets = Vec::new();
    let mut count = 0usize;
    for (i, op) in req.ops.iter().enumerate() {
        if let Some((code, message, _)) = validate_session_op(op, wb.sheet_count()) {
            return reject(code, message, i);
        }
        let (sheet, sr, sc, er, ec, values) = match op {
            Op::SetCellValue {
                sheet, row, col, ..
            }
            | Op::SetCellFormula {
                sheet, row, col, ..
            }
            | Op::ClearCell { sheet, row, col } => (*sheet, *row, *col, *row, *col, true),
            Op::SetStyle {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                ..
            }
            | Op::SetNumberFormat {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                ..
            } => (*sheet, *start_row, *start_col, *end_row, *end_col, false),
        };
        let range = visigrid_engine::validation::CellRange {
            start_row: sr,
            start_col: sc,
            end_row: er,
            end_col: ec,
        };
        count = count.saturating_add((er - sr + 1).saturating_mul(ec - sc + 1));
        if count > 100_000 {
            return reject(
                "cells_limit_exceeded",
                "A Table-aware batch may target at most 100,000 cells.".into(),
                i,
            );
        }
        if let Err(e) = wb.validate_automation_range(sheet, range, values) {
            return reject("table_view_unsafe", e, i);
        }
        targets.push((sheet, range, values));
    }
    let mut candidate = wb.clone();
    let mut outcome = apply_ops_inner(&mut candidate, req);
    if outcome.response.error.is_some() {
        return outcome;
    }
    for (i, &(sheet, range, values)) in targets.iter().enumerate() {
        if let Err(e) = candidate.validate_automation_range(sheet, range, values) {
            return reject("table_view_unsafe", e, i);
        }
    }
    let commit = match wb.capture_guarded_batch(&candidate) {
        Ok(commit) => commit,
        Err(e) => return reject("table_view_unsafe", e, req.ops.len() - 1),
    };
    wb.restore_snapshot_monotonic(&candidate);
    outcome.response.current_revision = wb.revision();
    outcome.guarded_commit = Some(commit);
    outcome
}

fn apply_ops_inner(wb: &mut Workbook, req: &ApplyOpsRequest) -> ApplyOutcome {
    let current_rev = wb.revision();

    let reject = |error: Option<ApplyOpsError>, total: usize| ApplyOutcome {
        guarded_commit: None,
        response: ApplyOpsResponse {
            applied: 0,
            total,
            current_revision: current_rev,
            error,
            warnings: Vec::new(),
        },
        value_changes: HashMap::new(),
        format_patches: HashMap::new(),
        changed_cells: Vec::new(),
    };

    if let Err(message) = wb.ensure_writable() {
        return reject(Some(ApplyOpsError::OpFailed(OpError {
            code: "read_only".into(), message, op_index: 0, suggestion: None,
        })), req.ops.len());
    }

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
        let table_error = match op {
            Op::SetCellValue { sheet, row, col, .. }
            | Op::SetCellFormula { sheet, row, col, .. }
            | Op::ClearCell { sheet, row, col } => wb.sheet(*sheet)
                .and_then(|s| s.table_value_write_error(*row, *col))
                .map(|reason| ("table_header", reason, None)),
            _ => None,
        };
        if let Some((code, message, suggestion)) = validate_session_op(op, sheet_count).or(table_error) {
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

    wb.begin_batch();
    {
        let guard: &mut Workbook = wb;

        let push_patch = |acc: &mut HashMap<usize, (Vec<FormatPatch>, HashMap<(usize, usize), usize>)>,
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
    }
    // Closing the batch is the single recalc and revision increment. Its
    // outcome says which OTHER cells changed as a consequence: dependent
    // formulas, custom-function and =LUA cells, spill receivers. A subscriber
    // or a live viewer only learns about cells it is told about, so the
    // delta carries those too, not just what the caller wrote.
    let outcome = wb.end_batch_outcome();

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
    {
        let mut seen: std::collections::HashSet<(usize, usize, usize)> =
            changed_cells.iter().map(|c| (c.sheet, c.row, c.col)).collect();
        let recalculated: Vec<CellId> = match outcome.recalculated {
            Recalculated::Cells(cells) => cells,
            // A cycle forced a full recompute, which can also resize or retire
            // spills, so "everything" is every cell that holds a value: the
            // stored cells and every spill receiver, on every sheet.
            Recalculated::All => wb
                .sheets()
                .iter()
                .flat_map(|sheet| {
                    let id = sheet.id;
                    sheet
                        .cells_iter()
                        .map(move |((row, col), _)| CellId { sheet: id, row, col })
                        .chain(sheet.spill_receiver_coords().map(move |(row, col)| CellId { sheet: id, row, col }))
                })
                .collect(),
        };
        for cell in recalculated {
            if let Some(sheet) = wb.sheet_index_by_id(cell.sheet) {
                if seen.insert((sheet, cell.row, cell.col)) {
                    changed_cells.push(CellRef { sheet, row: cell.row, col: cell.col });
                }
            }
        }
    }

    let warnings: Vec<String> = outcome
        .errors
        .iter()
        .map(|e| {
            let sheet = wb.sheet_index_by_id(e.cell.sheet).unwrap_or(0);
            format!(
                "sheet {} {}{}: {}",
                sheet,
                visigrid_engine::formula::parser::column_letters_pub(e.cell.col),
                e.cell.row + 1,
                e.error
            )
        })
        .collect();

    ApplyOutcome {
        guarded_commit: None,
        response: ApplyOpsResponse {
            applied,
            total: req.ops.len(),
            current_revision: wb.revision(),
            error: None,
            warnings,
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

    #[test]
    fn read_only_recovery_rejects_session_writes() {
        let mut wb = Workbook::new();
        wb.active_sheet_mut().read_only_reason = Some("Upgrade VisiGrid".into());
        let rev = wb.revision();
        let req = ApplyOpsRequest {
            request_id: String::new(), batch_name: String::new(), atomic: true,
            expected_revision: None, client: None,
            ops: vec![Op::SetCellValue { sheet: 0, row: 1, col: 0, value: "99".into() }],
        };
        let outcome = apply_ops(&mut wb, &req);
        assert!(matches!(outcome.response.error, Some(ApplyOpsError::OpFailed(error)) if error.code == "read_only"));
        assert_eq!(wb.revision(), rev);
        assert_eq!(wb.active_sheet().get_raw(1, 0), "");
    }

    #[test]
    fn filtered_session_structure_uses_canonical_rows_and_guards_final_layout() {
        use visigrid_engine::{
            filter::SortDirection,
            sheet::{Sheet, SheetId},
            table::TableRange,
            table_view::{TableSort, TableViewSpec},
        };
        let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(7), 30, 8)], 0);
        wb.set_cell_value_tracked(0, 2, 1, "Amount");
        wb.set_cell_value_tracked(0, 3, 1, "30");
        wb.set_cell_value_tracked(0, 4, 1, "10");
        let id = wb
            .create_table(
                SheetId(7),
                TableRange {
                    start_row: 2,
                    start_col: 1,
                    end_row: 4,
                    end_col: 1,
                },
                "Sales",
            )
            .unwrap()
            .table_id();
        let mut spec = TableViewSpec::new(id);
        spec.sort = Some(TableSort {
            column: wb.table(id).unwrap().1.columns[0].id,
            direction: SortDirection::Ascending,
        });
        wb.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
        apply_structure(
            &mut wb,
            &StructureOp::DeleteRows {
                sheet: None,
                at: 3,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(wb.active_sheet().get_raw(3, 1), "10");
        let rev = wb.revision();
        assert!(apply_structure(
            &mut wb,
            &StructureOp::DeleteCols {
                sheet: None,
                at: 1,
                count: 1
            }
        )
        .is_err());
        assert!(apply_structure(
            &mut wb,
            &StructureOp::DeleteRows {
                sheet: None,
                at: 2,
                count: 1
            }
        )
        .is_err());
        assert_eq!(wb.revision(), rev);
    }
    #[test]
    fn session_sheet_rename_preserves_table_criteria_and_rewrites_references_atomically() {
        use visigrid_engine::{
            filter::SortDirection,
            table::TableRange,
            table_view::{TableSort, TableViewSpec},
        };
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "Amount");
        wb.set_cell_value_tracked(0, 1, 0, "30");
        wb.set_cell_value_tracked(0, 2, 0, "10");
        let sid = wb.active_sheet_id();
        let id = wb.create_table(sid, TableRange {
            start_row: 0, start_col: 0, end_row: 2, end_col: 0,
        }, "Sales").unwrap().table_id();
        let mut spec = TableViewSpec::new(id);
        spec.sort = Some(TableSort {
            column: wb.table(id).unwrap().1.columns[0].id,
            direction: SortDirection::Ascending,
        });
        wb.set_table_view_spec(sid, Some(spec.clone())).unwrap();
        let summary = wb.add_sheet_named("Summary").unwrap();
        wb.set_cell_value_tracked(summary, 0, 0, "=(Sheet1!A2+Sheet1!A3)*2");
        let op = StructureOp::RenameSheet { sheet: Some(0), name: "New Data".into() };
        let before = wb.revision();
        apply_structure(&mut wb, &op).unwrap();
        assert!(wb.revision() > before);
        assert_eq!(wb.sheet(0).unwrap().name, "New Data");
        assert_eq!(wb.sheet(0).unwrap().table_view_spec(), Some(&spec));
        assert_eq!(wb.sheet(summary).unwrap().get_raw(0, 0), "=('New Data'!A2+'New Data'!A3)*2");
        wb.set_cell_value_tracked(summary, 2, 0, "=IFERROR(New!A1+1,99)");
        apply_structure(&mut wb, &StructureOp::AddSheet { name: Some("New".into()) }).unwrap();
        assert_eq!(wb.sheet(summary).unwrap().get_display(2, 0), "1");
        assert_eq!(wb.sheet(0).unwrap().table_view_spec(), Some(&spec));
        let count = wb.sheet_count();
        assert!(apply_structure(&mut wb, &StructureOp::AddSheet { name: Some("New".into()) }).is_err());
        assert_eq!(wb.sheet_count(), count);
        assert_eq!(wb.sheet(summary).unwrap().get_display(0, 0), "80");
        let revision = wb.revision();
        apply_structure(&mut wb, &op).unwrap();
        assert_eq!(wb.revision(), revision);
        wb.set_cell_value_tracked(summary, 1, 0, "='New Data'!A2+@");
        let revision = wb.revision();
        apply_structure(&mut wb, &StructureOp::RenameSheet {
            sheet: Some(0), name: "Next".into(),
        }).unwrap();
        assert!(wb.revision() > revision);
        assert_eq!(wb.sheet(0).unwrap().name, "Next");
        assert_eq!(wb.sheet(summary).unwrap().get_raw(0, 0), "=('Next'!A2+'Next'!A3)*2");
        assert_eq!(wb.sheet(summary).unwrap().get_raw(1, 0), "=Next!A2+@");
    }

    #[test]
    fn table_header_write_rejects_entire_session_batch() {
        let mut wb = Workbook::new();
        let sid = wb.active_sheet().id;
        wb.create_table(sid, visigrid_engine::table::TableRange { start_row: 0, start_col: 0, end_row: 2, end_col: 0 }, "Sales").unwrap();
        let rev = wb.revision();
        let req = ApplyOpsRequest {
            request_id: String::new(), batch_name: String::new(), atomic: true,
            expected_revision: None, client: None,
            ops: vec![
                Op::SetCellValue { sheet: 0, row: 1, col: 0, value: "99".into() },
                Op::ClearCell { sheet: 0, row: 0, col: 0 },
            ],
        };
        let outcome = apply_ops(&mut wb, &req);
        assert!(outcome.response.error.is_some());
        assert_eq!(wb.revision(), rev);
        assert_eq!(wb.active_sheet().get_raw(1, 0), "");
        assert_eq!(wb.active_sheet().get_raw(0, 0), "Column1");
    }

    /// The delta of an op must name the cells the op changed indirectly, or
    /// a subscriber mirroring the sheet keeps showing the old dependent.
    #[test]
    fn apply_ops_delta_includes_dependents_and_spill_receivers() {
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "3");
        wb.set_cell_value_tracked(0, 0, 1, "=A1*2");
        wb.set_cell_value_tracked(0, 0, 2, "=SEQUENCE(A1)");
        wb.rebuild_dep_graph();
        let req = ApplyOpsRequest {
            request_id: String::new(),
            batch_name: String::new(),
            atomic: false,
            expected_revision: None,
            ops: vec![Op::SetCellValue { sheet: 0, row: 0, col: 0, value: "5".into() }],
            client: None,
        };
        let outcome = apply_ops(&mut wb, &req);
        assert!(outcome.response.error.is_none(), "{:?}", outcome.response.error);
        let has = |row: usize, col: usize| outcome.changed_cells.iter().any(|c| c.sheet == 0 && c.row == row && c.col == col);
        assert!(has(0, 0), "the written cell");
        assert!(has(0, 1), "its dependent");
        assert!(has(4, 2), "a new spill receiver");
        assert_eq!(wb.sheets()[0].get_display(0, 1), "10");
        assert_eq!(wb.sheets()[0].get_display(4, 2), "5");

        let shrink = ApplyOpsRequest {
            request_id: String::new(),
            batch_name: String::new(),
            atomic: false,
            expected_revision: None,
            ops: vec![Op::SetCellValue { sheet: 0, row: 0, col: 0, value: "2".into() }],
            client: None,
        };
        let outcome = apply_ops(&mut wb, &shrink);
        let has = |row: usize, col: usize| outcome.changed_cells.iter().any(|c| c.sheet == 0 && c.row == row && c.col == col);
        assert!(has(2, 2), "retired spill receiver C3");
        assert!(has(3, 2), "retired spill receiver C4");
        assert!(has(4, 2), "retired spill receiver C5");
        assert_eq!(wb.sheets()[0].get_display(2, 2), "");
        assert_eq!(wb.sheets()[0].get_display(3, 2), "");
        assert_eq!(wb.sheets()[0].get_display(4, 2), "");
    }

    /// A spill chain deeper than the settlement bound is reported on the
    /// response, so an agent learns its values are stale.
    #[test]
    fn apply_ops_reports_an_unsettled_spill_chain_as_a_warning() {
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "1");
        wb.set_cell_value_tracked(0, 0, 1, "=SEQUENCE(A1+2)");
        for col in 2..20 {
            let prev = visigrid_engine::formula::parser::column_letters_pub(col - 1);
            wb.set_cell_value_tracked(0, 0, col, &format!("=SEQUENCE({}3)", prev));
        }
        wb.rebuild_dep_graph();
        let _ = wb.take_incremental_errors();
        let req = ApplyOpsRequest {
            request_id: String::new(),
            batch_name: String::new(),
            atomic: false,
            expected_revision: None,
            ops: vec![Op::SetCellValue { sheet: 0, row: 0, col: 0, value: "2".into() }],
            client: None,
        };
        let outcome = apply_ops(&mut wb, &req);
        assert!(outcome.response.error.is_none(), "{:?}", outcome.response.error);
        assert!(
            outcome.response.warnings.iter().any(|w| w.contains("not settled")),
            "{:?}",
            outcome.response.warnings
        );
    }

    /// When a cycle forces a full recompute the delta must cover every cell
    /// that can hold a value, receivers included.
    #[test]
    fn apply_ops_delta_after_a_full_recompute_covers_spill_receivers() {
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "1");
        wb.set_cell_value_tracked(0, 0, 2, "=SEQUENCE(3)");
        wb.rebuild_dep_graph();
        let req = ApplyOpsRequest {
            request_id: String::new(),
            batch_name: String::new(),
            atomic: false,
            expected_revision: None,
            // A1 and B1 reference each other: a cycle among the dirty cells
            // makes the incremental path fall back to a full recompute.
            ops: vec![
                Op::SetCellFormula { sheet: 0, row: 0, col: 0, formula: "=B1".into() },
                Op::SetCellFormula { sheet: 0, row: 0, col: 1, formula: "=A1".into() },
            ],
            client: None,
        };
        let outcome = apply_ops(&mut wb, &req);
        assert!(outcome.response.error.is_none(), "{:?}", outcome.response.error);
        let has = |row: usize, col: usize| outcome.changed_cells.iter().any(|c| c.sheet == 0 && c.row == row && c.col == col);
        assert!(has(2, 2), "receiver C3 is in the delta, {:?}", outcome.changed_cells);
    }
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
        assert_eq!(code(&StructureOp::RenameSheet { sheet: Some(0), name: "x".repeat(32) }, &wb), "invalid_op");

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
        let out = apply_ops(&mut wb, &req(vec![write(0, NUM_ROWS + 1, 0, "ghost")]));
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
        // The ghost-cell case: a row past the end of the grid. 70,000 was that
        // row until 0.35.0 and is ordinary now, which is the point of pinning
        // this to NUM_ROWS rather than to a number.
        let (code, msg, _) = validate_session_op(&set_value(0, NUM_ROWS, 0), 1).unwrap();
        assert_eq!(code, "out_of_bounds");
        assert!(msg.contains(&NUM_ROWS.to_string()));
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
        // A full column used to fit under the cap and no longer does: the grid
        // grew to 1,048,576 rows while the cap stayed where undo-patch memory
        // wants it. Agents format a bounded range, and the error says so.
        assert!(NUM_ROWS > MAX_SESSION_FORMAT_CELLS);
        let (code, _, _) = validate_session_op(&style(0, 0, NUM_ROWS - 1, 0), 1).unwrap();
        assert_eq!(code, "cells_limit_exceeded");
        assert!(validate_session_op(&style(0, 0, MAX_SESSION_FORMAT_CELLS - 1, 0), 1).is_none());
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
        let (code, _) = validate_inspect_target(&InspectTarget::Cell { sheet: 0, row: NUM_ROWS, col: 0 }, 1).unwrap();
        assert_eq!(code, "out_of_bounds");

        assert!(validate_inspect_target(&range(0, 0, 0, 19, 9), 1).is_none());
        let (code, _) = validate_inspect_target(&range(0, 5, 0, 0, 9), 1).unwrap();
        assert_eq!(code, "invalid_op");
        let (code, _) = validate_inspect_target(&range(0, 0, 0, NUM_ROWS - 1, NUM_COLS - 1), 1).unwrap();
        assert_eq!(code, "cells_limit_exceeded");
        // A full column was exactly at the inspect cap before the grid grew.
        assert!(NUM_ROWS > MAX_SESSION_INSPECT_CELLS);
        assert!(validate_inspect_target(
            &range(0, 0, 0, MAX_SESSION_INSPECT_CELLS - 1, 0), 1).is_none());
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

    #[test]
    fn create_and_refresh_pivot_ops_headless() {
        use visigrid_protocol::PivotValueSpec;
        let mut wb = Workbook::new();
        {
            let sh = wb.sheet_mut(0).unwrap();
            for (c, h) in ["Region", "Rep", "Amount"].iter().enumerate() {
                sh.set_value(0, c, h);
            }
            for (r, row) in [["West", "Ann", "10"], ["East", "Bo", "5"], ["West", "Cy", "2"]].iter().enumerate() {
                for (c, v) in row.iter().enumerate() {
                    sh.set_value(r + 1, c, v);
                }
            }
        }
        let create = StructureOp::CreatePivot {
            sheet: None,
            source: None,
            rows: vec!["region".into()],
            column: None,
            values: vec![PivotValueSpec { field: "Amount".into(), aggregation: Some("sum".into()) }],
        };
        assert!(validate_structure_op(&create, &wb).is_none());
        let desc = apply_structure(&mut wb, &create).unwrap();
        assert!(desc.starts_with("Created PivotTable1 on sheet \"Pivot\""), "{desc}");
        assert_eq!(wb.sheets()[1].get_display(2, 1), "12");

        wb.sheet_mut(0).unwrap().set_value(3, 2, "20");
        let refresh = StructureOp::RefreshPivot { pivot: Some("PivotTable1".into()) };
        assert!(validate_structure_op(&refresh, &wb).is_none());
        apply_structure(&mut wb, &refresh).unwrap();
        assert_eq!(wb.sheets()[1].get_display(2, 1), "30");

        let code = |op: &StructureOp, wb: &Workbook| validate_structure_op(op, wb).map(|e| e.1).unwrap_or_default();
        let bad_field = StructureOp::CreatePivot {
            sheet: None, source: Some("A1:C4".into()), rows: vec!["Month".into()], column: None, values: vec![],
        };
        assert!(code(&bad_field, &wb).contains("no column headed \"Month\""));
        let bad_agg = StructureOp::CreatePivot {
            sheet: None, source: None, rows: vec![], column: None,
            values: vec![PivotValueSpec { field: "Amount".into(), aggregation: Some("median".into()) }],
        };
        assert!(code(&bad_agg, &wb).contains("unknown aggregation"));
        let bad_range = StructureOp::CreatePivot {
            sheet: None, source: Some("A1".into()), rows: vec!["Region".into()], column: None, values: vec![],
        };
        assert!(code(&bad_range, &wb).contains("at least one data row"));
        assert!(code(&StructureOp::RefreshPivot { pivot: Some("Nope".into()) }, &wb).contains("pivots: PivotTable1"));

        // Pointing a new pivot at existing output is refused.
        let on_output = StructureOp::CreatePivot {
            sheet: Some(1), source: Some("A1:B3".into()), rows: vec!["Region".into()], column: None, values: vec![],
        };
        assert!(code(&on_output, &wb).contains("overlaps PivotTable1"));
    }

    #[test]
    fn create_pivot_uses_the_desktop_defaults() {
        use visigrid_engine::cell::NumberFormat;
        use visigrid_engine::pivot::Aggregation;
        use visigrid_protocol::PivotValueSpec;
        let mut wb = Workbook::new();
        let currency = NumberFormat::Currency { decimals: 2, thousands: true, negative: Default::default(), symbol: None };
        {
            let sh = wb.sheet_mut(0).unwrap();
            for (c, h) in ["Region", "Amount"].iter().enumerate() {
                sh.set_value(0, c, h);
            }
            for (r, (reg, amt)) in [("West", "10"), ("East", "5")].iter().enumerate() {
                sh.set_value(r + 1, 0, reg);
                sh.set_value(r + 1, 1, amt);
                sh.set_number_format(r + 1, 1, currency.clone());
            }
        }
        let op = |values: Vec<PivotValueSpec>| StructureOp::CreatePivot {
            sheet: None, source: None, rows: vec!["Region".into()], column: None, values,
        };
        let spec = |f: &str, a: Option<&str>| PivotValueSpec { field: f.into(), aggregation: a.map(Into::into) };

        // Aggregation omitted: Sum for the numeric column (keeping its
        // currency format), Count for the text one.
        let (_, def) = resolve_create_pivot(&op(vec![spec("Amount", None), spec("Region", None)]), &wb).unwrap();
        assert_eq!(def.values[0].aggregation, Aggregation::Sum);
        assert_eq!(def.values[0].number_format, Some(currency.clone()));
        assert_eq!(def.values[1].aggregation, Aggregation::Count);
        assert!(matches!(def.values[1].number_format, Some(NumberFormat::Number { decimals: 0, .. })));

        // Asked-for Sum of a text column is refused rather than showing zeros.
        let err = resolve_create_pivot(&op(vec![spec("Region", Some("sum"))]), &wb).unwrap_err().1;
        assert!(err.contains("\"Region\" has no numbers to sum"), "{err}");
        assert!(resolve_create_pivot(&op(vec![spec("Region", Some("distinct_count"))]), &wb).is_ok());

        let id = wb.create_table(wb.sheet(0).unwrap().id, visigrid_engine::table::TableRange {
            start_row: 0, start_col: 0, end_row: 2, end_col: 1,
        }, "Sales").unwrap().table_id();
        let mut named = op(vec![spec("Amount", None)]);
        if let StructureOp::CreatePivot { source, .. } = &mut named { *source = Some("sales".into()); }
        let (source, definition) = resolve_create_pivot(&named, &wb).unwrap();
        assert_eq!(source.table_id, Some(id));
        assert_eq!(definition.values[0].field.column_id, Some(wb.table(id).unwrap().1.columns[1].id));

        // Headless output shows the currency format.
        apply_structure(&mut wb, &op(vec![spec("Amount", None)])).unwrap();
        let out = wb.sheet(1).unwrap();
        assert_eq!(out.get_format(3, 1).number_format, currency);
    }

    #[test]
    fn refresh_all_reports_what_was_refreshed_before_a_failure() {
        let mut wb = Workbook::new();
        {
            let sh = wb.sheet_mut(0).unwrap();
            for (c, h) in ["Region", "Amount"].iter().enumerate() {
                sh.set_value(0, c, h);
            }
            sh.set_value(1, 0, "West");
            sh.set_value(1, 1, "10");
        }
        let create = |col: &str| StructureOp::CreatePivot {
            sheet: Some(0), source: Some("A1:B2".into()), rows: vec![col.into()], column: None, values: vec![],
        };
        apply_structure(&mut wb, &create("Amount")).unwrap(); // PivotTable1
        apply_structure(&mut wb, &create("Region")).unwrap(); // PivotTable2
        // Renaming Region breaks only the second pivot.
        wb.sheet_mut(0).unwrap().set_value(0, 0, "Area");
        let err = apply_structure(&mut wb, &StructureOp::RefreshPivot { pivot: None }).unwrap_err();
        assert!(err.starts_with("PivotTable2:"), "{err}");
        assert!(err.ends_with("(already refreshed: PivotTable1 (2 × 1))"), "{err}");
    }

    #[test]
    fn a_headless_host_refuses_to_refresh_a_recipe_table() {
        // Running a recipe reads files the user approved in the app; a headless
        // host has no approval to check, so it says where to do it instead.
        let mut wb = Workbook::new();
        let err = apply_structure(&mut wb, &StructureOp::RefreshRecipeTable { table: None }).unwrap_err();
        assert!(err.contains("VisiGrid desktop app"), "{err}");
    }
}
