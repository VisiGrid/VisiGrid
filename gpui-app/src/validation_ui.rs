//! Validation authoring uses captured worksheet targets and sparse history.
use crate::{app::Spreadsheet, history::UndoAction, mode::Mode};
use gpui::{App, Context};
use visigrid_engine::validation::{CellRange, ValidationEdit, ValidationResult};

#[path = "validation_plan.rs"]
pub(crate) mod plan;

impl Spreadsheet {
    pub(crate) fn navigate_validation_failure(&mut self, backward: bool, cx: &mut Context<Self>) {
        self.sync_table_view(cx);
        let target = plan::failure_target(
            &self.row_view,
            self.display_hidden_rows(),
            self.hidden_cols.get(&self.cached_sheet_id()),
            &self.validation_failures,
            self.view_state.selected,
            backward,
        );
        if let Some((index, pos, rank, total)) = target {
            self.validation_failure_index = index;
            self.select_cell(pos.0, pos.1, false, cx);
            self.ensure_visible(cx);
            let reason = self
                .invalid_cells
                .get(&self.validation_failures[index])
                .map(|r| Self::failure_reason_short(*r))
                .unwrap_or_default();
            self.status_message = Some(format!(
                "Invalid {rank} of {total} visible cells: {reason} — F8 next, Shift+F8 previous"
            ));
        } else {
            self.status_message = Some(
                if self.validation_failures.is_empty() {
                    "No validation failures to navigate"
                } else {
                    "Validation failures are hidden by the current view. Reveal them to navigate."
                }
                .into(),
            );
        }
        cx.notify();
    }

    pub(crate) fn capture_validation_draft(
        &mut self,
        cx: &mut Context<Self>,
    ) -> Result<plan::Draft, String> {
        self.sync_table_view(cx);
        self.validate_saved_view_layout(self.wb(cx))?;
        let ranges = self.all_selection_ranges();
        let projected = self.row_view.is_sorted()
            || self.row_view.is_filtered()
            || self.display_hidden_rows().is_some_and(|h| !h.is_empty())
            || self
                .hidden_cols
                .get(&self.cached_sheet_id())
                .is_some_and(|h| !h.is_empty());
        // Whole-column metadata needs no cell expansion on an ordinary sheet.
        let (ranges, anchor) = if !projected && ranges.len() == 1 {
            let ((r, c), (er, ec)) = ranges[0];
            (vec![CellRange::new(r, c, er, ec)], (r, c))
        } else {
            crate::cond_format_ui::plan::metadata_targets(
                self.sheet(cx),
                &self.row_view,
                self.display_hidden_rows(),
                self.hidden_cols.get(&self.cached_sheet_id()),
                &ranges,
                "validation",
            )?
        };
        plan::Draft::new(self.wb(cx), ranges, anchor)
    }

    pub(crate) fn open_validation_editor(&mut self, cx: &mut Context<Self>) {
        let draft = match self.capture_validation_draft(cx) {
            Ok(draft) => draft,
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                return;
            }
        };
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
        self.lua_console.visible = false;
        let rule = draft.anchor_rule();
        let existing = self
            .sheet(cx)
            .validations
            .iter()
            .any(|(r, _)| draft.ranges.iter().any(|t| r.overlaps(t)));
        self.validation_dialog
            .open_draft(draft, rule.as_ref(), existing);
        self.mode = Mode::ValidationDialog;
        cx.notify();
    }

    pub(crate) fn save_validation_editor(&mut self, clear: bool, cx: &mut Context<Self>) {
        let Some(draft) = self.validation_dialog.draft.clone() else {
            self.validation_dialog.error =
                Some("Reopen the validation dialog to select its cells.".into());
            cx.notify();
            return;
        };
        let rule = if clear {
            Ok(None)
        } else {
            self.validation_dialog.build_rule()
        };
        let result = rule.and_then(|rule| {
            let (edit, label) = match rule {
                Some(rule) => (ValidationEdit::Set(rule), "Set validation"),
                None => (ValidationEdit::Clear, "Clear validation"),
            };
            self.publish_validation_edit(&draft, edit, label, cx)
        });
        match result {
            Ok(()) => self.hide_validation_dialog(cx),
            Err(error) => {
                self.validation_dialog.error = Some(error);
                cx.notify();
            }
        }
    }

    pub(crate) fn edit_validation_exclusions(&mut self, clear: bool, cx: &mut Context<Self>) {
        let result = self.capture_validation_draft(cx).and_then(|draft| {
            let (edit, label) = if clear {
                (
                    ValidationEdit::ClearExclusions,
                    "Clear validation exclusions",
                )
            } else {
                (ValidationEdit::Exclude, "Exclude from validation")
            };
            self.publish_validation_edit(&draft, edit, label, cx)
        });
        if let Err(error) = result {
            self.status_message = Some(error);
        }
        cx.notify();
    }

    fn publish_validation_edit(
        &mut self,
        draft: &plan::Draft,
        edit: ValidationEdit,
        label: &str,
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        if self.block_if_previewing_only(cx) {
            return Err("Validation cannot change in the current mode.".into());
        }
        let Some(commit) = draft.prepare(self.wb(cx), edit)? else {
            self.status_message = Some("Validation is already set for these cells.".into());
            return Ok(());
        };
        let sheet_index = self.wb(cx).sheet_index_by_id(commit.sheet_id).unwrap();
        self.workbook.update(cx, |wb, _| commit.apply(wb, true))?;
        self.history.record_action_with_provenance(
            UndoAction::ValidationChanged {
                sheet_index,
                commit: Box::new(commit),
                description: label.into(),
            },
            None,
        );
        self.bump_cells_rev();
        let invalid = self.revalidate_validation_ranges(&draft.ranges, cx);
        self.status_message = Some(if invalid == 0 {
            label.into()
        } else {
            format!("{label}. {invalid} invalid — press F8 to navigate visible cells")
        });
        self.is_modified = true;
        cx.notify();
        Ok(())
    }

    /// Match Circle Invalid Data's populated-cell semantics, avoiding a scan of
    /// a million blank rows when a whole-column rule changes. Markers use stored
    /// coordinates and are removed from newly blank/unvalidated cells as well.
    pub(crate) fn revalidate_validation_ranges(&mut self, ranges: &[CellRange], cx: &App) -> usize {
        self.invalid_cells
            .retain(|&(r, c), _| !ranges.iter().any(|range| range.contains(r, c)));
        let cells: Vec<_> = self
            .sheet(cx)
            .cells_iter()
            .filter_map(|((r, c), _)| {
                ranges
                    .iter()
                    .any(|range| range.contains(r, c))
                    .then_some((r, c))
            })
            .collect();
        let index = self.sheet_index(cx);
        let mut count = 0;
        for (r, c) in cells {
            let display = self.sheet(cx).get_display(r, c);
            if display.is_empty() {
                continue;
            }
            if let ValidationResult::Invalid { reason, .. } =
                self.wb(cx).validate_cell(index, r, c)
            {
                self.invalid_cells.insert(
                    (r, c),
                    visigrid_engine::workbook::Workbook::classify_failure_reason(&reason),
                );
                count += 1;
            }
        }
        self.validation_failures = self.invalid_cells.keys().copied().collect();
        self.validation_failures.sort_unstable();
        self.validation_failure_index = 0;
        count
    }

    pub(crate) fn replay_validation_edit(
        &mut self,
        commit: &plan::Commit,
        forward: bool,
        cx: &mut Context<Self>,
    ) {
        self.workbook
            .update(cx, |wb, _| commit.apply(wb, forward))
            .expect("preflighted validation history");
        if self.wb(cx).active_sheet_id() == commit.sheet_id {
            self.revalidate_validation_ranges(&commit.ranges, cx);
        }
        self.bump_cells_rev();
    }
}
