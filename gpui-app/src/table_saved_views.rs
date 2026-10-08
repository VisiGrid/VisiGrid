//! Desktop named criteria presets. Applying uses normal Table view history;
//! managing presets uses sparse Table metadata history and leaves cells alone.
use crate::{
    app::Spreadsheet,
    table_ui::{TableDialog, TableDialogKind},
};
use gpui::{Context, Keystroke};
use visigrid_engine::{
    table::TableId,
    table_view::TableViewSpec,
    workbook::{TableCommit, Workbook},
};

pub(crate) fn current_spec(wb: &Workbook, id: TableId) -> Result<TableViewSpec, String> {
    let (sheet, _) = wb.table(id).ok_or("Table no longer exists.")?;
    Ok(wb
        .sheet_by_id(sheet)
        .unwrap()
        .table_view_spec()
        .filter(|s| s.table == id)
        .cloned()
        .unwrap_or_else(|| TableViewSpec::new(id)))
}

pub(crate) fn prepare(
    wb: &Workbook,
    draft: &TableDialog,
) -> Result<Option<(Workbook, TableCommit, TableId, &'static str)>, String> {
    let id = match draft.kind {
        TableDialogKind::SaveView(id)
        | TableDialogKind::RenameView(id, _)
        | TableDialogKind::UpdateView(id, _)
        | TableDialogKind::DeleteView(id, _) => id,
        _ => return Err("Select a saved-view command.".into()),
    };
    let (sheet, table) = wb.table(id).ok_or("Table no longer exists.")?;
    if sheet != draft.sheet {
        return Err("The Table moved to another sheet. Reopen Views.".into());
    }
    let before = table.saved_views.clone();
    let mut candidate = wb.clone();
    let (commit, label) = match draft.kind {
        TableDialogKind::SaveView(_) => (
            candidate.save_named_table_view(id, &draft.name, current_spec(wb, id)?)?,
            "Save Table view",
        ),
        TableDialogKind::RenameView(..) => (
            candidate.rename_named_table_view(id, &draft.range, &draft.name)?,
            "Rename Table view",
        ),
        TableDialogKind::UpdateView(..) => (
            candidate.update_named_table_view(id, &draft.range, current_spec(wb, id)?)?,
            "Update saved Table view",
        ),
        TableDialogKind::DeleteView(..) => (
            candidate.delete_named_table_view(id, &draft.range)?,
            "Delete saved Table view",
        ),
        _ => unreachable!(),
    };
    if candidate.table(id).unwrap().1.saved_views == before {
        return Ok(None);
    }
    for sheet in candidate.sheets() {
        sheet.build_saved_table_view(crate::app::NUM_ROWS.min(sheet.rows))?;
    }
    Ok(Some((candidate, commit, id, label)))
}

impl Spreadsheet {
    pub(crate) fn submit_named_table_view(&mut self, cx: &mut Context<Self>) {
        let Some(draft) = self.table_dialog.clone() else {
            return;
        };
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        if let TableDialogKind::Views(id) = draft.kind {
            self.apply_named_table_view(id, draft.field, cx);
            return;
        }
        let saving = matches!(
            draft.kind,
            TableDialogKind::SaveView(_) | TableDialogKind::UpdateView(..)
        );
        let result = if saving
            && !self.table_view_installed
            && (self.row_view.is_sorted() || self.filter_state.is_enabled())
        {
            Err("Clear worksheet sorting and filters first. Saved Table views capture Table controls only.".into())
        } else {
            prepare(self.wb(cx), &draft)
        };
        match result {
            Ok(Some((candidate, commit, id, label))) => {
                let last = candidate
                    .table(id)
                    .unwrap()
                    .1
                    .saved_views
                    .len()
                    .saturating_sub(1);
                let selected = match draft.kind {
                    TableDialogKind::SaveView(_) => last,
                    TableDialogKind::RenameView(_, i)
                    | TableDialogKind::UpdateView(_, i)
                    | TableDialogKind::DeleteView(_, i) => i.min(last),
                    _ => 0,
                };
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.record_table_commit(commit, label.into(), cx);
                self.open_table_dialog(TableDialogKind::Views(id), cx);
                if let Some(d) = &mut self.table_dialog {
                    d.field = selected;
                }
            }
            Ok(None) => self.table_dialog = None,
            Err(error) => {
                if let Some(d) = &mut self.table_dialog {
                    d.error = Some(error);
                }
            }
        }
        cx.notify();
    }

    pub(crate) fn apply_named_table_view(
        &mut self,
        id: TableId,
        index: usize,
        cx: &mut Context<Self>,
    ) {
        let saved = self.wb(cx).table(id).and_then(|(sheet, t)| {
            (sheet == self.sheet(cx).id)
                .then(|| t.saved_views.get(index).cloned())
                .flatten()
        });
        let Some(saved) = saved else {
            return;
        };
        if self.change_table_view(
            Some(saved.view),
            &format!("Apply Table view: {}", saved.name),
            cx,
        ) {
            self.table_dialog = None;
        } else if let Some(d) = &mut self.table_dialog {
            d.error = self.status_message.clone();
        }
        cx.notify();
    }

    /// List navigation stays inside the modal; Enter applies the highlighted
    /// setup. Editing/confirmation dialogs use the existing text interceptor.
    pub(crate) fn named_table_views_key(
        &mut self,
        key: &Keystroke,
        cx: &mut Context<Self>,
    ) -> bool {
        let Some(d) = &self.table_dialog else {
            return false;
        };
        let TableDialogKind::Views(id) = d.kind else {
            return false;
        };
        let count = self
            .wb(cx)
            .table(id)
            .map_or(0, |(_, t)| t.saved_views.len());
        let index = d.field.min(count.saturating_sub(1));
        match key.key.as_str() {
            "escape" => self.table_dialog = None,
            "enter" if count > 0 => self.apply_named_table_view(id, index, cx),
            "up" | "down" | "tab" if count > 0 => {
                let back = key.key == "up" || (key.key == "tab" && key.modifiers.shift);
                let d = self.table_dialog.as_mut().unwrap();
                d.field = if back {
                    (index + count - 1) % count
                } else {
                    (index + 1) % count
                };
            }
            "f2" if count > 0 => self.open_table_dialog(TableDialogKind::RenameView(id, index), cx),
            "delete" if count > 0 => {
                self.open_table_dialog(TableDialogKind::DeleteView(id, index), cx)
            }
            "n" if key.modifiers.control || key.modifiers.platform => {
                self.open_table_dialog(TableDialogKind::SaveView(id), cx)
            }
            _ => {}
        }
        cx.notify();
        true
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        history::{History, UndoAction},
        table_edit::tests::fixture,
    };

    fn draft(wb: &Workbook, kind: TableDialogKind, name: &str, original: &str) -> TableDialog {
        TableDialog {
            kind,
            sheet: wb.active_sheet_id(),
            name: name.into(),
            range: original.into(),
            has_headers: true,
            field: 0,
            select_all: true,
            error: None,
        }
    }

    #[test]
    fn save_update_rename_delete_and_apply_rewind_with_active_criteria() {
        let base = fixture(true);
        let id = base.active_sheet().table_view_spec().unwrap().table;
        let spec = current_spec(&base, id).unwrap();
        let mut wb = base.clone();
        let mut history = History::new();
        let mut commits = Vec::new();
        for d in [
            draft(&base, TableDialogKind::SaveView(id), "West", ""),
            draft(
                &base,
                TableDialogKind::RenameView(id, 0),
                "West orders",
                "West",
            ),
            draft(
                &base,
                TableDialogKind::DeleteView(id, 0),
                "West orders",
                "West orders",
            ),
        ] {
            let (candidate, commit, _, _) = prepare(&wb, &d).unwrap().unwrap();
            assert_eq!(candidate.active_sheet().table_view_spec(), Some(&spec));
            assert!(commit.is_saved_view_change());
            wb = candidate;
            history.record_action_with_provenance(&visigrid_engine::workbook::Workbook::new(),
                UndoAction::TableCommit {
                    sheet_index: 0,
                    commit: Box::new(commit.clone()),
                    header_layout: None,
                    description: "Saved view".into(),
                },
                None,
            );
            let replay = history
                .build_workbook_before(commits.len() + 1, Some(&base), 100, 10_000)
                .unwrap()
                .workbook;
            assert_eq!(replay.table(id).unwrap().1, wb.table(id).unwrap().1);
            assert_eq!(
                replay.active_sheet().get_display(0, 1),
                base.active_sheet().get_display(0, 1)
            );
            commits.push(commit);
        }
        for commit in commits.iter().rev() {
            wb.apply_table_commit(commit, true).unwrap();
        }
        assert_eq!(wb.table(id).unwrap().1, base.table(id).unwrap().1);
        for commit in &commits {
            wb.apply_table_commit(commit, false).unwrap();
        }
        wb.apply_table_commit(commits.last().unwrap(), true)
            .unwrap();
        let clear = wb.set_table_view_spec(wb.active_sheet_id(), None).unwrap();
        let d = draft(
            &wb,
            TableDialogKind::UpdateView(id, 0),
            "West orders",
            "West orders",
        );
        let (candidate, update, _, _) = prepare(&wb, &d).unwrap().unwrap();
        wb = candidate;
        assert_eq!(
            wb.table(id).unwrap().1.saved_views[0].view,
            TableViewSpec::new(id)
        );
        wb.apply_table_commit(&update, true).unwrap();
        wb.apply_table_view_commit(&clear, true).unwrap();
        assert_eq!(wb.table(id).unwrap().1.saved_views[0].view, spec);
        assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
    }

    #[test]
    fn invalid_or_unchanged_presets_do_not_publish_a_candidate() {
        let mut wb = fixture(true);
        let id = wb.active_sheet().table_view_spec().unwrap().table;
        let spec = current_spec(&wb, id).unwrap();
        wb.save_named_table_view(id, "West", spec).unwrap();
        let before = serde_json::to_string(&wb.saved_tables()).unwrap();
        assert!(prepare(&wb, &draft(&wb, TableDialogKind::SaveView(id), "west", "")).is_err());
        assert!(prepare(
            &wb,
            &draft(&wb, TableDialogKind::RenameView(id, 0), "West", "West")
        )
        .unwrap()
        .is_none());
        assert!(prepare(
            &wb,
            &draft(&wb, TableDialogKind::UpdateView(id, 0), "West", "West")
        )
        .unwrap()
        .is_none());
        assert!(prepare(
            &wb,
            &draft(
                &wb,
                TableDialogKind::DeleteView(id, 0),
                "Missing",
                "Missing"
            )
        )
        .is_err());
        assert_eq!(serde_json::to_string(&wb.saved_tables()).unwrap(), before);
    }
}
