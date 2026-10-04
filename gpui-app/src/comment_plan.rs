//! Comment metadata uses canonical cells and never changes values or formulas.
use crate::history::UndoAction;
use visigrid_engine::{sheet::Sheet, workbook::Workbook};

pub(crate) fn validate_edit(
    wb: &Workbook,
    sheet_index: usize,
    row: usize,
    col: usize,
    has_comment: bool,
    revision: Option<u64>,
) -> Result<(), String> {
    wb.ensure_writable()?;
    if revision.is_some_and(|r| r != wb.revision()) {
        return Err("The workbook changed while this comment was open. Copy your draft, then reopen the comment before saving.".into());
    }
    let sheet = wb
        .sheet(sheet_index)
        .ok_or("The comment sheet no longer exists.")?;
    validate_target(sheet, row, col, has_comment, false)
}

/// A comment in the body belongs to that record. A new comment beside the
/// body would become unrelated content hidden/moved by the whole-row view.
/// History may restore a hidden record; direct editing must first reveal it.
pub(crate) fn validate_target(
    sheet: &Sheet,
    row: usize,
    col: usize,
    has_comment: bool,
    history: bool,
) -> Result<(), String> {
    if row >= sheet.rows || col >= sheet.cols {
        return Err("The comment cell is outside the worksheet.".into());
    }
    if !history {
        if sheet
            .get_merge(row, col)
            .is_some_and(|m| m.start != (row, col))
        {
            return Err("Edit the comment at the merged cell's top-left cell.".into());
        }
        let view = sheet.build_saved_table_view(sheet.rows)?;
        if sheet.manual_hidden_rows().contains(&row)
            || view.is_some_and(|v| !v.rows().is_data_row_visible(row))
        {
            return Err(
                "Reveal this row before editing its comment. You can still read hidden comments."
                    .into(),
            );
        }
    }
    if has_comment {
        if let Some(spec) = sheet.table_view_spec().filter(|s| s.has_criteria()) {
            let table = sheet
                .tables()
                .iter()
                .find(|t| t.id == spec.table)
                .ok_or("The comment's Table no longer exists.")?;
            let r = table.range;
            if row > r.start_row && row <= r.end_row && (col < r.start_col || col > r.end_col) {
                return Err("A comment beside this Table would move or hide with its records. Clear Table sorting and filters before adding it.".into());
            }
        }
    }
    Ok(())
}

/// Preflight every comment in a group before replay mutates any child. This
/// needs only patch destinations, without copying or recalculating the workbook.
pub(crate) fn validate_history(
    wb: &Workbook,
    action: &UndoAction,
    forward: bool,
) -> Result<(), String> {
    match action {
        UndoAction::Comments {
            sheet_index,
            patches,
            ..
        } => {
            let sheet = wb
                .sheet(*sheet_index)
                .ok_or("The comment sheet no longer exists.")?;
            for p in patches {
                let destination = if forward { &p.after } else { &p.before };
                validate_target(sheet, p.row, p.col, destination.is_some(), true)?;
            }
        }
        UndoAction::Group { actions, .. } => {
            for action in actions {
                validate_history(wb, action, forward)?;
            }
        }
        _ => (),
    }
    Ok(())
}

#[cfg(test)]
#[path = "comment_plan_tests.rs"]
mod tests;
