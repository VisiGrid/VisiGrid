//! Pivot tables in the desktop app: commands, the field-list drawer state,
//! refresh, refusals for edits that touch pivot output, and undo/redo glue.
//!
//! The engine owns the rules (crates/engine/src/pivot.rs, workbook_pivot.rs):
//! ownership is enforced there, and every pivot action is a scoped
//! `PivotCommit`. This module decides when to run them and what to tell the
//! user.

use gpui::*;
use visigrid_engine::filter::RowView;
use visigrid_engine::sheet::Sheet;
use visigrid_engine::workbook::PivotCommit;

use crate::app::{Spreadsheet, NUM_ROWS};

impl Spreadsheet {
    /// Undo a pivot action: restore the "before" state, and remove the sheet
    /// the action created, if any.
    pub(crate) fn pivot_undo(
        &mut self,
        commit: &PivotCommit,
        created_sheet: &Option<(usize, Box<Sheet>)>,
        cx: &mut Context<Self>,
    ) {
        let active_before = self.wb(cx).active_sheet_index();
        self.workbook.update(cx, |wb, _| {
            let _ = wb.apply_pivot_state(&commit.before);
            if let Some((_, sheet)) = created_sheet {
                if let Some(i) = wb.sheet_index_by_id(sheet.id) {
                    wb.take_sheet(i);
                }
            }
        });
        self.finish_pivot_restore(active_before, cx);
    }

    /// Redo a pivot action: restore the sheet it created, if any, then apply
    /// the "after" state.
    pub(crate) fn pivot_redo(
        &mut self,
        commit: &PivotCommit,
        created_sheet: &Option<(usize, Box<Sheet>)>,
        cx: &mut Context<Self>,
    ) {
        let active_before = self.wb(cx).active_sheet_index();
        self.workbook.update(cx, |wb, _| {
            if let Some((index, sheet)) = created_sheet {
                if wb.sheet_index_by_id(sheet.id).is_none() {
                    wb.restore_sheet(*index, (**sheet).clone());
                }
            }
            let _ = wb.apply_pivot_state(&commit.after);
        });
        self.finish_pivot_restore(active_before, cx);
    }

    fn finish_pivot_restore(&mut self, active_before: usize, cx: &mut Context<Self>) {
        let active_after = self.wb(cx).active_sheet_index();
        let row_view = if active_after == active_before {
            self.row_view.clone()
        } else {
            RowView::new(NUM_ROWS)
        };
        self.finish_workbook_snapshot_restore(row_view, cx);
        self.is_modified = true;
        cx.notify();
    }
}

// ---------------------------------------------------------------------------
// Refusals: an edit that touches pivot output is refused as a whole.
// ---------------------------------------------------------------------------

impl Spreadsheet {
    /// The pivot whose owned output intersects the rectangle given in view
    /// coordinates (rows as displayed, so sort/filter are mapped) on the active
    /// sheet. Returns its name.
    pub(crate) fn pivot_in_view_rect(
        &self,
        view_r0: usize,
        c0: usize,
        view_r1: usize,
        c1: usize,
        cx: &App,
    ) -> Option<String> {
        let sheet = self.sheet(cx);
        if sheet.pivots.is_empty() {
            return None;
        }
        let (r0, r1) = (view_r0.min(view_r1), view_r0.max(view_r1));
        let (c0, c1) = (c0.min(c1), c0.max(c1));
        if !self.row_view.is_sorted() && !self.row_view.is_filtered() {
            return sheet.pivot_in_rect(r0, c0, r1, c1).map(|p| p.name.clone());
        }
        // Sorted or filtered: map each displayed row to its data row.
        sheet.pivots.iter().find_map(|p| {
            let (pr0, pc0, pr1, pc1) = p.region()?;
            if pc0 > c1 || c0 > pc1 {
                return None;
            }
            (r0..=r1)
                .map(|vr| self.row_view.view_to_data(vr))
                .any(|dr| dr >= pr0 && dr <= pr1)
                .then(|| p.name.clone())
        })
    }

    /// Refuse an edit whose target touches pivot output: set a status message
    /// and return true. `what` names the operation ("edit", "paste", …).
    pub(crate) fn block_if_pivot(
        &mut self,
        view_r0: usize,
        c0: usize,
        view_r1: usize,
        c1: usize,
        what: &str,
        cx: &mut Context<Self>,
    ) -> bool {
        let Some(name) = self.pivot_in_view_rect(view_r0, c0, view_r1, c1, cx) else {
            return false;
        };
        self.status_message = Some(format!(
            "Can't {what} here: this is {name}'s output. Change it from the field list, or copy the values to another sheet."
        ));
        cx.notify();
        true
    }

    /// Refuse if any of the current selection ranges touches pivot output.
    pub(crate) fn block_if_selection_in_pivot(&mut self, what: &str, cx: &mut Context<Self>) -> bool {
        let ranges = self.all_selection_ranges();
        for ((r0, c0), (r1, c1)) in ranges {
            if self.block_if_pivot(r0, c0, r1, c1, what, cx) {
                return true;
            }
        }
        false
    }
}
