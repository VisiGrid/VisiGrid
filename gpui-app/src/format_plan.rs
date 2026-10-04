//! Plan metadata-only formatting against displayed rows, then publish canonical
//! patches together. No workbook copy or formula recalculation is needed.
use crate::{
    formatting::BorderApplyMode,
    history::{CellFormatPatch, UndoAction},
};
use std::collections::{BTreeMap, BTreeSet};
use visigrid_engine::{
    cell::{
        Alignment, CellBorder, CellFormat, CellStyle, NumberFormat, TextOverflow, VerticalAlignment,
    },
    filter::RowView,
    sheet::Sheet,
    workbook::Workbook,
};

pub(crate) type Range = ((usize, usize), (usize, usize));
const LIMIT: usize = 100_000;

#[derive(Clone, Debug)]
pub(crate) enum Operation {
    Bold(bool),
    Italic(bool),
    Underline(bool),
    Strike(bool),
    Font(Option<String>),
    Size(Option<f32>),
    Color(Option<[u8; 4]>),
    Background(Option<[u8; 4]>),
    Align(Alignment),
    Vertical(VerticalAlignment),
    Overflow(TextOverflow),
    Number(NumberFormat),
    Decimals(i8),
    Style(CellStyle),
    Replace(CellFormat),
    Clear,
    Borders(BorderApplyMode, CellBorder),
}

pub(crate) fn visible_rows<'a>(
    rows: &'a RowView,
    hidden: Option<&'a BTreeSet<usize>>,
    start: usize,
    end: usize,
) -> impl Iterator<Item = (usize, usize)> + 'a {
    let all = rows.visible_rows();
    let a = all.partition_point(|r| *r < start);
    let b = all.partition_point(|r| *r <= end).max(a);
    all[a..b].iter().copied().filter_map(move |v| {
        let r = rows.view_to_data(v);
        (!hidden.is_some_and(|h| h.contains(&r))).then_some((v, r))
    })
}

pub(crate) fn visible_count(
    rows: &RowView,
    hidden: Option<&BTreeSet<usize>>,
    start: usize,
    end: usize,
) -> usize {
    let all = rows.visible_rows();
    let a = all.partition_point(|r| *r < start);
    let b = all.partition_point(|r| *r <= end).max(a);
    let manual = hidden.map_or(0, |h| {
        h.iter()
            .filter(|r| {
                rows.data_to_view(**r)
                    .is_some_and(|v| v >= start && v <= end)
            })
            .count()
    });
    (b - a).saturating_sub(manual)
}

pub(crate) fn row_neighbor(
    rows: &RowView,
    hidden: Option<&BTreeSet<usize>>,
    view: usize,
    forward: bool,
) -> Option<usize> {
    let all = rows.visible_rows();
    if forward {
        all[all.partition_point(|r| *r <= view)..]
            .iter()
            .copied()
            .map(|v| rows.view_to_data(v))
            .find(|r| !hidden.is_some_and(|h| h.contains(r)))
    } else {
        all[..all.partition_point(|r| *r < view)]
            .iter()
            .rev()
            .copied()
            .map(|v| rows.view_to_data(v))
            .find(|r| !hidden.is_some_and(|h| h.contains(r)))
    }
}

pub(crate) fn col_neighbor(
    hidden: Option<&BTreeSet<usize>>,
    col: usize,
    limit: usize,
    forward: bool,
) -> Option<usize> {
    if forward {
        ((col.saturating_add(1))..limit).find(|c| !hidden.is_some_and(|h| h.contains(c)))
    } else {
        (0..col)
            .rev()
            .find(|c| !hidden.is_some_and(|h| h.contains(c)))
    }
}

fn transform(op: &Operation, f: &mut CellFormat) {
    match op {
        Operation::Bold(v) => f.bold = *v,
        Operation::Italic(v) => f.italic = *v,
        Operation::Underline(v) => f.underline = *v,
        Operation::Strike(v) => f.strikethrough = *v,
        Operation::Font(v) => f.font_family = v.clone(),
        Operation::Size(v) => f.font_size = *v,
        Operation::Color(v) => f.font_color = *v,
        Operation::Background(v) => f.background_color = *v,
        Operation::Align(v) => f.alignment = *v,
        Operation::Vertical(v) => f.vertical_alignment = *v,
        Operation::Overflow(v) => f.text_overflow = *v,
        Operation::Number(v) => f.number_format = v.clone(),
        Operation::Decimals(delta) => match &mut f.number_format {
            NumberFormat::Number { decimals, .. }
            | NumberFormat::Currency { decimals, .. }
            | NumberFormat::Percent { decimals } => {
                *decimals = (i16::from(*decimals) + i16::from(*delta)).clamp(0, 10) as u8
            }
            _ => (),
        },
        Operation::Style(v) => {
            f.cell_style = *v;
            if !v.is_none() {
                f.background_color = None;
                f.font_color = None;
            }
        }
        Operation::Replace(v) => *f = v.clone(),
        Operation::Clear => *f = CellFormat::default(),
        Operation::Borders(..) => (),
    }
}

/// Protected Table headers/totals can be formatted without writing values.
/// Only new non-default metadata beside the body invalidates its row view.
pub(crate) fn validate_target(
    sheet: &Sheet,
    row: usize,
    col: usize,
    format: &CellFormat,
) -> Result<(), String> {
    if row >= sheet.rows || col >= sheet.cols {
        return Err("The format target is outside the worksheet.".into());
    }
    if !format.is_default() {
        if let Some(spec) = sheet.table_view_spec().filter(|s| s.has_criteria()) {
            let t = sheet
                .tables()
                .iter()
                .find(|t| t.id == spec.table)
                .ok_or("The formatting Table no longer exists.")?;
            if row > t.range.start_row
                && row <= t.range.end_row
                && (col < t.range.start_col || col > t.range.end_col)
            {
                return Err("Formatting beside this Table would move or hide with its records. Select cells inside the Table, or clear its sorting and filters.".into());
            }
        }
    }
    Ok(())
}

pub(crate) fn plan(
    wb: &Workbook,
    index: usize,
    rows: &RowView,
    hidden_rows: Option<&BTreeSet<usize>>,
    hidden_cols: Option<&BTreeSet<usize>>,
    ranges: &[Range],
    op: &Operation,
) -> Result<Vec<CellFormatPatch>, String> {
    wb.ensure_writable()?;
    let sheet = wb
        .sheet(index)
        .ok_or("The formatting sheet no longer exists.")?;
    let mut formats: BTreeMap<(usize, usize), CellFormat> = BTreeMap::new();
    let mut update = |row: usize, col: usize, f: &dyn Fn(&mut CellFormat)| -> Result<(), String> {
        if row >= sheet.rows || col >= sheet.cols {
            return Err("The format target is outside the worksheet.".into());
        }
        if formats.len() >= LIMIT && !formats.contains_key(&(row, col)) {
            return Err("Select at most 100,000 visible cells for one formatting operation. Nothing was changed.".into());
        }
        f(formats
            .entry((row, col))
            .or_insert_with(|| sheet.get_format(row, col)));
        Ok(())
    };
    for &((r0, c0), (r1, c1)) in ranges {
        if r0 > r1 || c0 > c1 || r1 >= rows.row_count() || c1 >= sheet.cols {
            return Err("The format selection is outside the worksheet.".into());
        }
        let selected_rows: Vec<_> = visible_rows(rows, hidden_rows, r0, r1)
            .take(LIMIT + 1)
            .collect();
        let selected_cols: Vec<_> = (c0..=c1)
            .filter(|c| !hidden_cols.is_some_and(|h| h.contains(c)))
            .collect();
        if selected_rows.is_empty() || selected_cols.is_empty() {
            continue;
        }
        if selected_rows.len().saturating_mul(selected_cols.len()) > LIMIT {
            return Err("Select at most 100,000 visible cells for one formatting operation. Nothing was changed.".into());
        }
        let last_row = selected_rows.len() - 1;
        let last_col = selected_cols.len() - 1;
        for (ri, &(_, row)) in selected_rows.iter().enumerate() {
            for (ci, &col) in selected_cols.iter().enumerate() {
                if matches!(op, Operation::Number(NumberFormat::Percent { .. }))
                    && sheet.table_header_at(row, col).is_none()
                    && crate::table_command_scope::percent_format_value(&sheet.get_raw(row, col))
                        .is_some()
                {
                    return Err("Percent formatting would convert text to numbers. Clear Table sorting and filters before this conversion.".into());
                }
                update(row, col, &|f| match op {
                    Operation::Borders(mode, edge) => match mode {
                        BorderApplyMode::All => {
                            f.border_top = *edge;
                            f.border_bottom = *edge;
                            f.border_left = *edge;
                            f.border_right = *edge;
                        }
                        BorderApplyMode::Clear => {
                            f.border_top = CellBorder::default();
                            f.border_bottom = CellBorder::default();
                            f.border_left = CellBorder::default();
                            f.border_right = CellBorder::default();
                        }
                        BorderApplyMode::Outline => {
                            if ri == 0 {
                                f.border_top = *edge;
                            }
                            if ri == last_row {
                                f.border_bottom = *edge;
                            }
                            if ci == 0 {
                                f.border_left = *edge;
                            }
                            if ci == last_col {
                                f.border_right = *edge;
                            }
                        }
                        BorderApplyMode::Inside => {
                            if ri < last_row {
                                f.border_bottom = *edge;
                            }
                            if ci < last_col {
                                f.border_right = *edge;
                            }
                        }
                        BorderApplyMode::Top => {
                            if ri == 0 {
                                f.border_top = *edge;
                            }
                        }
                        BorderApplyMode::Bottom => {
                            if ri == last_row {
                                f.border_bottom = *edge;
                            }
                        }
                        BorderApplyMode::Left => {
                            if ci == 0 {
                                f.border_left = *edge;
                            }
                        }
                        BorderApplyMode::Right => {
                            if ci == last_col {
                                f.border_right = *edge;
                            }
                        }
                    },
                    _ => transform(op, f),
                })?;
            }
        }
        if matches!(op, Operation::Borders(BorderApplyMode::Clear, _)) {
            for &col in &selected_cols {
                if let Some(row) = row_neighbor(rows, hidden_rows, selected_rows[0].0, false)
                    .filter(|r| *r < sheet.rows)
                {
                    update(row, col, &|f| f.border_bottom = CellBorder::default())?;
                }
                if let Some(row) = row_neighbor(rows, hidden_rows, selected_rows[last_row].0, true)
                    .filter(|r| *r < sheet.rows)
                {
                    update(row, col, &|f| f.border_top = CellBorder::default())?;
                }
            }
            for &(_, row) in &selected_rows {
                if let Some(col) = col_neighbor(hidden_cols, selected_cols[0], sheet.cols, false) {
                    update(row, col, &|f| f.border_right = CellBorder::default())?;
                }
                if let Some(col) =
                    col_neighbor(hidden_cols, selected_cols[last_col], sheet.cols, true)
                {
                    update(row, col, &|f| f.border_left = CellBorder::default())?;
                }
            }
        }
    }
    let mut patches = Vec::new();
    for ((row, col), after) in formats {
        let before = sheet.get_format(row, col);
        if before != after {
            validate_target(sheet, row, col, &after)?;
            patches.push(CellFormatPatch {
                row,
                col,
                before,
                after,
                remove_cell_on_undo: sheet.get_cell_opt(row, col).is_none(),
            });
        }
    }
    Ok(patches)
}

pub(crate) fn validate_history(
    wb: &Workbook,
    action: &UndoAction,
    forward: bool,
) -> Result<(), String> {
    match action {
        UndoAction::Format {
            sheet_index,
            patches,
            ..
        } => {
            let sheet = wb
                .sheet(*sheet_index)
                .ok_or("The formatting sheet no longer exists.")?;
            for p in patches {
                validate_target(
                    sheet,
                    p.row,
                    p.col,
                    if forward { &p.after } else { &p.before },
                )?;
            }
        }
        UndoAction::Group { actions, .. } => {
            for a in actions {
                validate_history(wb, a, forward)?;
            }
        }
        _ => (),
    }
    Ok(())
}

pub(crate) fn apply(wb: &mut Workbook, index: usize, patches: &[CellFormatPatch], forward: bool) {
    if patches.is_empty() {
        return;
    }
    if let Some(sheet) = wb.sheet_mut(index) {
        for p in patches {
            sheet.set_format(
                p.row,
                p.col,
                if forward {
                    p.after.clone()
                } else {
                    p.before.clone()
                },
            );
            if !forward && p.remove_cell_on_undo {
                sheet.remove_empty_metadata_cell(p.row, p.col);
            }
        }
        // set_format raises the flag when adding a border. Only rescan when a
        // cell loses its last border; ordinary bold/color edits stay sparse.
        if patches.iter().any(|p| {
            let (source, target) = if forward {
                (&p.before, &p.after)
            } else {
                (&p.after, &p.before)
            };
            source.has_any_border() && !target.has_any_border()
        }) {
            sheet.scan_border_flag();
        }
        wb.bump_revision_for_structure();
    }
}

#[cfg(test)]
#[path = "format_plan_tests.rs"]
mod tests;
