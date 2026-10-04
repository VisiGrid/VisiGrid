//! Sheet-local metadata can change without recalculating another sheet's Table.
//! Value/schema mutations still use the workbook-wide guarded paths.
use crate::{app::Spreadsheet, history::UndoAction};
use gpui::Context;
use visigrid_engine::{sheet::SheetId, workbook::Workbook};

/// The existing Percent command also coerces percentage text into numbers.
/// Keep that value mutation out of the metadata-only exemption.
pub(crate) fn percent_format_value(raw: &str) -> Option<String> {
    let pct = raw.strip_suffix('%')?;
    let clean: String = pct
        .chars()
        .filter(|c| !c.is_whitespace() && *c != ',')
        .collect();
    clean.parse::<f64>().ok().map(|n| (n / 100.0).to_string())
}

pub(crate) fn number_format_changes_values(
    sheet: &visigrid_engine::sheet::Sheet,
    format: &visigrid_engine::cell::NumberFormat,
    ranges: &[((usize, usize), (usize, usize))],
) -> bool {
    matches!(format, visigrid_engine::cell::NumberFormat::Percent { .. })
        && ranges.iter().any(|&((r0, c0), (r1, c1))| {
            (r0..=r1).any(|row| {
                (c0..=c1).any(|col| {
                    sheet.table_header_at(row, col).is_none()
                        && percent_format_value(&sheet.get_raw(row, col)).is_some()
                })
            })
        })
}

pub(crate) fn sheet_metadata_allowed(wb: &Workbook, index: usize) -> bool {
    wb.sheet(index).is_some_and(|sheet| {
        !sheet
            .table_view_spec()
            .is_some_and(|v| v.sort.is_some() || !v.filters.is_empty())
    })
}

/// Check the history target, never the currently displayed sheet. Mixed groups
/// must be admitted as a whole before any child action is replayed.
pub(crate) fn metadata_history_allowed(wb: &Workbook, action: &UndoAction) -> bool {
    match action {
        UndoAction::CondFormatAdded { sheet_index, .. }
        | UndoAction::CondFormatsCleared { sheet_index, .. } => {
            sheet_metadata_allowed(wb, *sheet_index)
        }
        UndoAction::Format { sheet_index, .. } | UndoAction::Comments { sheet_index, .. } => wb.sheet(*sheet_index).is_some(),
        UndoAction::FreezePanesChanged { sheet_id, .. } => wb
            .sheet_index_by_id(*sheet_id)
            .is_some_and(|index| sheet_metadata_allowed(wb, index)),
        UndoAction::Group { actions, .. } => {
            !actions.is_empty()
                && actions
                    .iter()
                    .all(|action| metadata_history_allowed(wb, action))
        }
        _ => false,
    }
}

/// Reject missing freeze targets before any grouped history action can mutate.
pub(crate) fn validate_freeze_history(wb: &Workbook, action: &UndoAction) -> Result<(), String> {
    match action {
        UndoAction::FreezePanesChanged { sheet_id, .. }
            if wb.sheet_index_by_id(*sheet_id).is_none() =>
        {
            Err("The freeze-pane sheet no longer exists.".into())
        }
        UndoAction::Group { actions, .. } => {
            for action in actions {
                validate_freeze_history(wb, action)?;
            }
            Ok(())
        }
        _ => Ok(()),
    }
}

pub(crate) fn restore_freeze_panes(
    wb: &mut Workbook,
    sheet_id: SheetId,
    frozen: (usize, usize),
) -> Result<(), String> {
    wb.ensure_writable()?;
    let index = wb
        .sheet_index_by_id(sheet_id)
        .ok_or("The freeze-pane sheet no longer exists.")?;
    wb.sheet_mut(index).unwrap().frozen_panes = frozen;
    Ok(())
}

impl Spreadsheet {
    pub(crate) fn block_number_format_conversion(
        &mut self,
        format: &visigrid_engine::cell::NumberFormat,
        cx: &mut Context<Self>,
    ) -> bool {
        if self.wb(cx).has_table_criteria()
            && number_format_changes_values(self.sheet(cx), format, &self.format_apply_ranges(cx))
        {
            self.status_message = Some("Percent formatting would convert text to numbers. Clear Table sorting and filters before this conversion.".into());
            cx.notify();
            return true;
        }
        false
    }

    pub(crate) fn block_active_sheet_metadata_edit(&mut self, cx: &mut Context<Self>) -> bool {
        self.block_sheet_metadata_edit(self.sheet_index(cx), cx)
    }

    pub(crate) fn block_sheet_metadata_edit(
        &mut self,
        index: usize,
        cx: &mut Context<Self>,
    ) -> bool {
        if self.block_if_previewing_only(cx) {
            return true;
        }
        if !sheet_metadata_allowed(self.wb(cx), index) {
            self.status_message = Some("Clear Table sorting and filters on this sheet before this operation. Other sheets remain editable.".into());
            cx.notify();
            return true;
        }
        false
    }

    pub(crate) fn restore_sheet_freeze_panes(
        &mut self,
        sheet_id: SheetId,
        frozen: (usize, usize),
        cx: &mut Context<Self>,
    ) {
        let result = self
            .workbook
            .update(cx, |wb, _| restore_freeze_panes(wb, sheet_id, frozen));
        if let Err(error) = result {
            self.status_message = Some(error);
            return;
        }
        if self.sheet(cx).id == sheet_id {
            self.view_state.frozen_rows = frozen.0;
            self.view_state.frozen_cols = frozen.1;
            self.clamp_scroll_to_freeze(cx);
        }
    }
}

#[cfg(test)]
#[path = "table_command_scope_tests.rs"]
mod tests;
