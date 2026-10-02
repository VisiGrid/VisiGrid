//! Fill plans address canonical records once, before any filter keys change.
use crate::{
    app::{FillAxis, Spreadsheet},
    series_fill::{self, DetectedSource, FillIntent, FillPattern},
    table_edit::TableCellWrite,
};
use gpui::Context;
use visigrid_engine::{
    cell::{interchange_number, ValueRef},
    formula::{eval::Value, parser::adjust_formula_refs},
    sheet::Sheet,
    table_view::TableView,
};

type Position = (usize, usize);
type Rect = (Position, Position);
const LIMIT: usize = 100_000;

fn normalized(a: Position, b: Position) -> Rect {
    ((a.0.min(b.0), a.1.min(b.1)), (a.0.max(b.0), a.1.max(b.1)))
}

fn body_rows(view: &TableView, rect: Rect) -> Result<Vec<usize>, String> {
    let ((r0, c0), (r1, c1)) = rect;
    let range = view.range();
    if r0 <= range.start_row || r1 > range.end_row || c0 < range.start_col || c1 > range.end_col {
        return Err("Fill must stay inside the Table body, including its source cells.".into());
    }
    let rows: Vec<_> = view
        .rows()
        .visible_rows()
        .iter()
        .copied()
        .filter(|r| *r >= r0 && *r <= r1)
        .map(|r| view.rows().view_to_data(r))
        .collect();
    if rows.is_empty() {
        return Err("Select at least one visible Table record.".into());
    }
    if rows.len().saturating_mul(c1 - c0 + 1) > LIMIT {
        return Err("Fill at most 100,000 visible Table cells at once.".into());
    }
    Ok(rows)
}

fn copy_cell(sheet: &Sheet, source: Position, target: Position) -> TableCellWrite {
    let cell = sheet.get_cell(source.0, source.1);
    let mut write = TableCellWrite::value(target.0, target.1, sheet.get_raw(source.0, source.1));
    match cell.value() {
        ValueRef::Formula {
            source: formula, ..
        } => {
            write.value = Some(adjust_formula_refs(
                formula,
                target.0 as i32 - source.0 as i32,
                target.1 as i32 - source.1 as i32,
            ))
        }
        ValueRef::Text(_) => write.literal_text = true,
        ValueRef::Number(n) => write.value = Some(interchange_number(n)),
        ValueRef::Empty => {}
    }
    write
}

fn plan_direction(
    sheet: &Sheet,
    view: &TableView,
    rect: Rect,
    down: bool,
) -> Result<Vec<TableCellWrite>, String> {
    let rows = body_rows(view, rect)?;
    let ((_, c0), (_, c1)) = rect;
    let mut writes = Vec::new();
    if down {
        let source = if rows.len() == 1 {
            let slot = view.rows().data_to_view(rows[0]).unwrap();
            view.rows()
                .visible_rows()
                .iter()
                .copied()
                .rev()
                .find(|r| *r < slot && *r > view.range().start_row)
                .map(|r| view.rows().view_to_data(r))
                .ok_or("No visible Table record above to fill from.")?
        } else {
            rows[0]
        };
        for row in rows.into_iter().filter(|r| *r != source) {
            for col in c0..=c1 {
                writes.push(copy_cell(sheet, (source, col), (row, col)));
            }
        }
    } else {
        let source = if c0 == c1 {
            c0.checked_sub(1)
                .filter(|c| *c >= view.range().start_col)
                .ok_or("No Table field to the left to fill from.")?
        } else {
            c0
        };
        for row in rows {
            for col in source + 1..=c1 {
                writes.push(copy_cell(sheet, (row, source), (row, col)));
            }
        }
    }
    Ok(writes)
}

fn plan_broadcast(
    sheet: &Sheet,
    view: &TableView,
    source: Position,
    rects: &[Rect],
    edit: Option<&str>,
) -> Result<Vec<TableCellWrite>, String> {
    body_rows(view, (source, source))?;
    let source = (view.rows().view_to_data(source.0), source.1);
    let mut targets = std::collections::BTreeSet::new();
    for &rect in rects {
        for row in body_rows(view, rect)? {
            for col in rect.0 .1..=rect.1 .1 {
                targets.insert((row, col));
                if targets.len() > LIMIT {
                    return Err("Fill at most 100,000 visible Table cells at once.".into());
                }
            }
        }
    }
    let base = edit.map(|value| {
        let mut value = if let Some(rest) = value.strip_prefix('+') {
            format!("={rest}")
        } else {
            value.to_owned()
        };
        if value.starts_with('=') {
            let count = value
                .matches('(')
                .count()
                .saturating_sub(value.matches(')').count());
            value.extend(std::iter::repeat_n(')', count));
        }
        value
    });
    Ok(targets
        .into_iter()
        .map(|target| {
            if let Some(base) = &base {
                let value = if base.starts_with('=') {
                    adjust_formula_refs(
                        base,
                        target.0 as i32 - source.0 as i32,
                        target.1 as i32 - source.1 as i32,
                    )
                } else {
                    base.clone()
                };
                TableCellWrite::value(target.0, target.1, value)
            } else {
                copy_cell(sheet, source, target)
            }
        })
        .collect())
}

fn fill_line(
    sheet: &Sheet,
    mut sources: Vec<Position>,
    targets: Vec<Position>,
    backwards: bool,
    ctrl: bool,
    writes: &mut Vec<TableCellWrite>,
) {
    if backwards {
        sources.reverse();
    }
    let has_formula = sources
        .iter()
        .any(|&(r, c)| sheet.get_cell(r, c).value().is_formula());
    let values: Vec<_> = sources
        .iter()
        .map(|&(r, c)| sheet.get_computed_value(r, c))
        .collect();
    let detected = DetectedSource {
        text_tokens: values
            .iter()
            .map(|v| {
                if let Value::Text(s) = v {
                    Some(s.clone())
                } else {
                    None
                }
            })
            .collect(),
        values,
    };
    let series = series_fill::detect_pattern(&detected, FillIntent::Series);
    let single_list = matches!(
        series,
        FillPattern::TextList { .. }
            | FillPattern::AlphaNum { .. }
            | FillPattern::QuarterYear { .. }
    );
    let intent = series_fill::fill_intent(ctrl, sources.len(), single_list);
    let mut pattern = if has_formula {
        FillPattern::Copy
    } else {
        series_fill::detect_pattern(&detected, intent)
    };
    // Reversing multiple source values infers the backward step. A single
    // source needs its default step explicitly reversed.
    if backwards && sources.len() == 1 {
        match &mut pattern {
            FillPattern::Linear { step, .. } => *step = -*step,
            FillPattern::TextList { step, .. } | FillPattern::QuarterYear { step, .. } => {
                *step = -*step
            }
            FillPattern::AlphaNum { step, .. } => *step = -*step,
            _ => {}
        }
    }
    for (i, target) in targets.into_iter().enumerate() {
        if matches!(pattern, FillPattern::Copy | FillPattern::Repeat { .. }) {
            writes.push(copy_cell(sheet, sources[i % sources.len()], target));
        } else {
            let value = series_fill::generate(&pattern, i + 1);
            let mut write = TableCellWrite::value(
                target.0,
                target.1,
                match &value {
                    Value::Number(n) => interchange_number(*n),
                    Value::Text(s) => s.clone(),
                    _ => String::new(),
                },
            );
            write.literal_text = matches!(value, Value::Text(_));
            writes.push(write);
        }
    }
}

fn plan_handle(
    sheet: &Sheet,
    view: &TableView,
    source: Rect,
    end: Position,
    vertical: bool,
    ctrl: bool,
) -> Result<Vec<TableCellWrite>, String> {
    let ((r0, c0), (r1, c1)) = source;
    let source_rows = body_rows(view, source)?;
    let combined = if vertical {
        ((r0.min(end.0), c0), (r1.max(end.0), c1))
    } else {
        ((r0, c0.min(end.1)), (r1, c1.max(end.1)))
    };
    body_rows(view, combined)?;
    let mut writes = Vec::new();
    if vertical {
        if (r0..=r1).contains(&end.0) {
            return Ok(writes);
        }
        let backwards = end.0 < r0;
        let mut target_rows: Vec<_> = view
            .rows()
            .visible_rows()
            .iter()
            .copied()
            .filter(|r| *r >= combined.0 .0 && *r <= combined.1 .0 && !(*r >= r0 && *r <= r1))
            .map(|r| view.rows().view_to_data(r))
            .collect();
        if backwards {
            target_rows.reverse();
        }
        for col in c0..=c1 {
            fill_line(
                sheet,
                source_rows.iter().map(|r| (*r, col)).collect(),
                target_rows.iter().map(|r| (*r, col)).collect(),
                backwards,
                ctrl,
                &mut writes,
            );
        }
    } else {
        if (c0..=c1).contains(&end.1) {
            return Ok(writes);
        }
        let backwards = end.1 < c0;
        let mut cols: Vec<_> = (combined.0 .1..=combined.1 .1)
            .filter(|c| !(*c >= c0 && *c <= c1))
            .collect();
        if backwards {
            cols.reverse();
        }
        for row in source_rows {
            fill_line(
                sheet,
                (c0..=c1).map(|c| (row, c)).collect(),
                cols.iter().map(|c| (row, *c)).collect(),
                backwards,
                ctrl,
                &mut writes,
            );
        }
    }
    Ok(writes)
}

impl Spreadsheet {
    fn fill_table_view(&mut self, cx: &mut Context<Self>) -> Result<TableView, String> {
        if self.block_if_previewing_only(cx) {
            return Err(
                "Fill is unavailable while viewing a read-only workbook or preview.".into(),
            );
        }
        self.sync_table_view(cx);
        self.sheet(cx)
            .build_saved_table_view(crate::app::NUM_ROWS.min(self.sheet(cx).rows))?
            .ok_or(
                "Select an active Table body, or clear the Table views before filling other cells."
                    .into(),
            )
    }

    fn finish_table_fill(
        &mut self,
        result: Result<Vec<TableCellWrite>, String>,
        description: &str,
        selection: Option<Rect>,
        cx: &mut Context<Self>,
    ) -> bool {
        let writes = match result {
            Ok(writes) => writes,
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                return false;
            }
        };
        if writes.is_empty() {
            return true;
        }
        let selected = self.view_state.selected;
        let end = self.view_state.selection_end;
        let additional = self.view_state.additional_selections.clone();
        if !self.apply_table_cell_writes(writes, description, cx) {
            return false;
        }
        // Keep the selected display area. If a filter key removed its endpoint,
        // trim to visible records; an empty area keeps the safe fallback focus.
        let (start, finish) = selection.unwrap_or((selected, end.unwrap_or(selected)));
        if let Some((start, finish)) = self.visible_fill_selection(start, finish) {
            self.view_state.selected = start;
            self.view_state.selection_end = if selection.is_some() || end.is_some() {
                Some(finish)
            } else {
                None
            };
            self.view_state.additional_selections = additional
                .into_iter()
                .filter_map(|(a, b)| {
                    self.visible_fill_selection(a, b.unwrap_or(a))
                        .map(|(a, e)| (a, b.map(|_| e)))
                })
                .collect();
        }
        self.status_message = Some(description.into());
        self.ensure_visible(cx);
        cx.notify();
        true
    }

    fn visible_fill_selection(&self, a: Position, b: Position) -> Option<Rect> {
        let ((r0, _), (r1, _)) = normalized(a, b);
        let mut rows = self
            .row_view
            .visible_rows()
            .iter()
            .copied()
            .filter(|r| *r >= r0 && *r <= r1);
        let first = rows.next()?;
        let last = rows.last().unwrap_or(first);
        Some((
            (if a.0 <= b.0 { first } else { last }, a.1),
            (if a.0 <= b.0 { last } else { first }, b.1),
        ))
    }

    pub(crate) fn fill_table_direction(&mut self, down: bool, cx: &mut Context<Self>) {
        if self.mode.is_editing() {
            return;
        }
        let result = self
            .fill_table_view(cx)
            .and_then(|view| plan_direction(self.sheet(cx), &view, self.selection_range(), down));
        self.finish_table_fill(
            result,
            if down {
                "Filled down through visible Table records"
            } else {
                "Filled right through visible Table records"
            },
            None,
            cx,
        );
    }

    pub(crate) fn fill_table_selection(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) {
            return;
        }
        let editing = self.mode.is_editing();
        if editing {
            let Some((sheet, _, _, revision)) = self.table_edit_target else {
                return;
            };
            if self.wb(cx).revision() != revision {
                self.status_message = Some(
                    "The workbook changed while editing. Cancel this edit and try again.".into(),
                );
                cx.notify();
                return;
            }
            self.restore_formula_home_sheet(cx);
            if self.sheet_index(cx) != sheet {
                return;
            }
        }
        let result = self.fill_table_view(cx).and_then(|view| {
            let source = if editing {
                let (_, row, col, _) = self.table_edit_target.unwrap();
                (
                    view.rows()
                        .data_to_view(row)
                        .ok_or("The source record is no longer visible.")?,
                    col,
                )
            } else {
                self.view_state.selected
            };
            plan_broadcast(
                self.sheet(cx),
                &view,
                source,
                &self.all_selection_ranges(),
                editing.then_some(self.edit_value.as_str()),
            )
        });
        if self.finish_table_fill(result, "Filled visible Table selection", None, cx) && editing {
            self.formula_edit_cell = None;
            self.cancel_edit(cx);
            self.maybe_show_cycle_banner(cx);
        }
    }

    pub(crate) fn start_table_fill(&mut self, cx: &mut Context<Self>) -> bool {
        let result = self
            .fill_table_view(cx)
            .and_then(|view| body_rows(&view, self.selection_range()));
        if let Err(error) = result {
            self.status_message = Some(error);
            cx.notify();
            return false;
        }
        self.table_fill_revision = Some((self.sheet(cx).id, self.wb(cx).revision()));
        true
    }

    pub(crate) fn end_table_fill(
        &mut self,
        anchor: Position,
        source_end: Position,
        end: Position,
        axis: Option<FillAxis>,
        ctrl: bool,
        cx: &mut Context<Self>,
    ) {
        let revision = self.table_fill_revision.take();
        if revision != Some((self.sheet(cx).id, self.wb(cx).revision())) {
            self.status_message =
                Some("The workbook changed during the fill drag. Try again.".into());
            cx.notify();
            return;
        }
        let Some(axis) = axis else {
            return;
        };
        let vertical = matches!(axis, FillAxis::Row);
        let source = normalized(anchor, source_end);
        let result = self
            .fill_table_view(cx)
            .and_then(|view| plan_handle(self.sheet(cx), &view, source, end, vertical, ctrl));
        let selection = if vertical {
            (
                (source.0 .0.min(end.0), source.0 .1),
                (source.1 .0.max(end.0), source.1 .1),
            )
        } else {
            (
                (source.0 .0, source.0 .1.min(end.1)),
                (source.1 .0, source.1 .1.max(end.1)),
            )
        };
        self.finish_table_fill(result, "Filled visible Table records", Some(selection), cx);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        table_cell_history::TableCellsCommit,
        table_edit::{prepare_table_writes, tests::fixture},
    };
    fn view(sheet: &Sheet) -> TableView {
        sheet.build_saved_table_view(30).unwrap().unwrap()
    }

    #[test]
    fn down_uses_visible_order_and_canonical_formula_offsets() {
        let wb = fixture(true);
        let sheet = wb.active_sheet();
        let view = view(sheet);
        assert_eq!(view.visible_body_rows(4, 3).unwrap(), vec![5, 3, 6]);
        let writes = plan_direction(sheet, &view, ((4, 3), (6, 3)), true).unwrap();
        assert_eq!(
            writes
                .iter()
                .map(|w| (w.row, w.value.as_deref().unwrap()))
                .collect::<Vec<_>>(),
            vec![(3, "=C4*2"), (6, "=C7*2")]
        );
        let writes = plan_direction(sheet, &view, ((4, 2), (6, 2)), true).unwrap();
        let mut after = prepare_table_writes(&wb, 0, &writes).unwrap();
        assert_eq!(after.active_sheet().get_raw(4, 2), "10");
        assert_eq!(after.active_sheet().get_raw(3, 2), "20");
        assert_eq!(after.active_sheet().get_raw(6, 2), "20");
        let commit = TableCellsCommit::capture(
            sheet,
            after.active_sheet(),
            writes.iter().map(|w| (w.row, w.col)),
        );
        commit.replay(&mut after, true).unwrap();
        assert_eq!(
            after.active_sheet().get_computed_value(0, 1),
            Value::Number(100.0)
        );
        commit.replay(&mut after, false).unwrap();
        assert_eq!(
            after.active_sheet().get_computed_value(0, 1),
            Value::Number(70.0)
        );
    }

    #[test]
    fn single_row_uses_previous_visible_record_and_edges_refuse() {
        let wb = fixture(true);
        let s = wb.active_sheet();
        let v = view(s);
        let writes = plan_direction(s, &v, ((5, 2), (5, 2)), true).unwrap();
        assert_eq!((writes[0].row, writes[0].value.as_deref()), (3, Some("20")));
        assert!(plan_direction(s, &v, ((4, 2), (4, 2)), true).is_err());
        assert!(plan_direction(s, &v, ((4, 1), (6, 1)), false).is_err());
        assert!(plan_direction(s, &v, ((2, 2), (6, 2)), true).is_err());
        assert!(plan_handle(s, &v, ((4, 2), (5, 2)), (7, 2), true, false).is_err());
        assert!(plan_handle(s, &v, ((4, 2), (5, 2)), (6, 4), false, false).is_err());
        assert_eq!(wb.active_sheet().get_raw(6, 2), "40");
    }

    #[test]
    fn right_fill_preserves_literal_text_in_each_visible_record() {
        let mut wb = fixture(true);
        wb.set_cell_text_tracked(0, 5, 2, "00123");
        wb.set_cell_text_tracked(0, 3, 2, "=1+1");
        let s = wb.active_sheet();
        let v = view(s);
        let writes = plan_direction(s, &v, ((3, 2), (6, 3)), false).unwrap();
        assert!(writes
            .iter()
            .filter(|w| w.row == 3 || w.row == 5)
            .all(|w| w.literal_text));
        let after = prepare_table_writes(&wb, 0, &writes).unwrap();
        assert_eq!(
            after.active_sheet().get_computed_value(3, 3),
            Value::Text("=1+1".into())
        );
        assert_eq!(
            after.active_sheet().get_computed_value(5, 3),
            Value::Text("00123".into())
        );
        assert_eq!(after.active_sheet().get_raw(4, 3), "=C5*2");
    }

    #[test]
    fn broadcast_plans_all_records_before_filter_membership_changes() {
        let wb = fixture(true);
        let s = wb.active_sheet();
        let v = view(s);
        let writes = plan_broadcast(
            s,
            &v,
            (4, 1),
            &[((4, 1), (6, 1)), ((5, 1), (5, 1))],
            Some("East"),
        )
        .unwrap();
        assert_eq!(writes.len(), 3);
        let mut after = prepare_table_writes(&wb, 0, &writes).unwrap();
        assert_eq!(view(after.active_sheet()).rows().visible_count(), 26);
        assert_eq!(after.active_sheet().get_raw(4, 1), "East");
        let commit = TableCellsCommit::capture(
            s,
            after.active_sheet(),
            writes.iter().map(|w| (w.row, w.col)),
        );
        commit.replay(&mut after, true).unwrap();
        assert_eq!(
            view(after.active_sheet()).visible_body_rows(4, 3).unwrap(),
            vec![5, 3, 6]
        );
        let writes =
            plan_broadcast(s, &v, (4, 3), &[((4, 3), (6, 3))], Some("+SUM(C6,$C$4")).unwrap();
        assert_eq!(
            writes.iter().find(|w| w.row == 3).unwrap().value.as_deref(),
            Some("=SUM(C4,$C$4)")
        );
        assert!(plan_broadcast(
            s,
            &v,
            (4, 3),
            &[((4, 3), (6, 3)), ((8, 1), (8, 1))],
            Some("99")
        )
        .is_err());
    }

    #[test]
    fn drag_series_counts_visible_records_and_copies_formulas_per_column() {
        let wb = fixture(true);
        let s = wb.active_sheet();
        let v = view(s);
        let writes = plan_handle(s, &v, ((4, 2), (5, 3)), (6, 3), true, false).unwrap();
        assert_eq!(writes.len(), 2);
        assert_eq!(writes[0].value.as_deref(), Some("40")); // 20,30 -> 40, hidden 10 ignored
        assert_eq!(writes[1].value.as_deref(), Some("=C7*2"));
        let writes = plan_handle(s, &v, ((4, 2), (4, 2)), (6, 2), true, true).unwrap();
        assert_eq!(
            writes
                .iter()
                .map(|w| w.value.as_deref().unwrap())
                .collect::<Vec<_>>(),
            vec!["21", "22"]
        );
        let writes = plan_handle(s, &v, ((4, 2), (5, 2)), (6, 2), true, true).unwrap();
        assert_eq!(writes[0].value.as_deref(), Some("20"));
        let writes = plan_handle(s, &v, ((5, 2), (6, 2)), (4, 2), true, false).unwrap();
        assert_eq!(writes[0].value.as_deref(), Some("20")); // reverse 40,30 -> 20
        let writes = plan_handle(s, &v, ((5, 2), (5, 2)), (4, 2), true, true).unwrap();
        assert_eq!(writes[0].value.as_deref(), Some("29"));
        let writes = plan_handle(s, &v, ((4, 2), (6, 2)), (6, 3), false, true).unwrap();
        assert_eq!(writes.len(), 3);
        assert_eq!(
            writes.iter().map(|w| w.row).collect::<Vec<_>>(),
            vec![5, 3, 6]
        );
        assert_eq!(writes[0].value.as_deref(), Some("21"));
    }
}
