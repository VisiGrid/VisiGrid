use crate::{app::Spreadsheet, history::MutationSource};
use gpui::{App, Context};
use visigrid_engine::{
    filter::RowView,
    named_range::NamedRange,
    workbook::{NamedRangeEdit, Workbook},
};

#[derive(Clone, Debug)]
pub(crate) struct NameDraft {
    pub revision: u64,
    pub range: NamedRange,
}

impl NameDraft {
    fn checked_range(&self, wb: &Workbook) -> Result<NamedRange, String> {
        if self.revision != wb.revision() {
            return Err(
                "The workbook changed while this dialog was open. Reopen it and try again.".into(),
            );
        }
        Ok(self.range.clone())
    }
}

/// As with A1 formula references, a rectangle spans its canonical endpoints
/// and includes intervening hidden records. Never treat view slots as data.
pub(crate) fn selection_range(
    wb: &Workbook,
    rows: &RowView,
    start: (usize, usize),
    end: (usize, usize),
) -> Result<NamedRange, String> {
    let sheet = wb.active_sheet();
    if start.0 >= rows.row_count()
        || end.0 >= rows.row_count()
        || !rows.is_view_row_visible(start.0)
        || !rows.is_view_row_visible(end.0)
    {
        return Err("Select visible endpoints for the named range.".into());
    }
    let a = rows.view_to_data(start.0);
    let b = rows.view_to_data(end.0);
    let (r0, r1, c0, c1) = (a.min(b), a.max(b), start.1.min(end.1), start.1.max(end.1));
    if r1 >= sheet.rows || c1 >= sheet.cols {
        return Err("The named range extends beyond the worksheet.".into());
    }
    Ok(if r0 == r1 && c0 == c1 {
        NamedRange::cell("", wb.active_sheet_index(), r0, c0)
    } else {
        NamedRange::range("", wb.active_sheet_index(), r0, c0, r1, c1)
    })
}

pub(crate) fn project_named_range(
    rows: &RowView,
    start: (usize, usize),
    end: (usize, usize),
) -> Vec<((usize, usize), (usize, usize))> {
    let mut ranges: Vec<((usize, usize), (usize, usize))> = Vec::new();
    for &slot in rows.visible_rows() {
        let data = rows.view_to_data(slot);
        if data < start.0 || data > end.0 {
            continue;
        }
        if let Some((_, last)) = ranges.last_mut().filter(|(_, last)| last.0 + 1 == slot) {
            last.0 = slot;
        } else {
            ranges.push(((slot, start.1), (slot, end.1)));
        }
    }
    ranges
}

impl Spreadsheet {
    pub(crate) fn named_range_draft(&self, cx: &App) -> Result<NamedRange, String> {
        let draft = self
            .name_draft
            .as_ref()
            .ok_or("Reopen the named range dialog and try again.")?;
        draft.checked_range(self.wb(cx))
    }

    pub(crate) fn apply_named_range_edit(
        &mut self,
        edit: NamedRangeEdit,
        description: String,
        cx: &mut Context<Self>,
    ) -> bool {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return false;
        }
        let result = self
            .named_range_draft(cx)
            .and_then(|_| self.wb(cx).prepare_named_range_edit(&edit))
            .and_then(|(candidate, commit)| {
                if commit.is_empty() {
                    return Ok(());
                }
                self.publish_table_batch(candidate, commit, description, MutationSource::Human, cx)
            });
        match result {
            Ok(()) => {
                self.name_draft = None;
                self.name_draft_error = None;
                true
            }
            Err(error) => {
                self.name_draft_error = Some(error.clone());
                self.status_message = Some(error);
                cx.notify();
                false
            }
        }
    }
}

#[cfg(test)]
#[path = "plan_tests.rs"]
mod tests;
