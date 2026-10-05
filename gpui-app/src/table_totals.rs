//! Native totals commands and history, including saved Table views.
use crate::app::Spreadsheet;
use gpui::Context;
use visigrid_engine::{
    table::{TableId, TableTotal},
    workbook::{TableCommit, Workbook},
};

pub(crate) const CHOICES: &[(&str, &str)] = &[
    ("sum", "Sum"),
    ("average", "Average"),
    ("count", "Count"),
    ("countNums", "Count numbers"),
    ("min", "Minimum"),
    ("max", "Maximum"),
    ("stdDev", "Std. deviation"),
    ("var", "Variance"),
    ("label", "Label"),
    ("custom", "Formula"),
    ("none", "None"),
];

pub(crate) fn total_setting(choice: &str, text: &str) -> Result<TableTotal, String> {
    Ok(match choice {
        "none" => TableTotal::default(),
        "label" => TableTotal {
            label: Some(text.into()),
            ..Default::default()
        },
        "custom" => TableTotal {
            function: Some("custom".into()),
            formula: Some(text.trim().into()),
            label: None,
        },
        name if CHOICES.iter().any(|(key, _)| *key == name) => TableTotal {
            function: Some(name.into()),
            ..Default::default()
        },
        _ => return Err("Choose a totals function.".into()),
    })
}

fn validate(wb: &Workbook) -> Result<(), String> {
    wb.ensure_writable()?;
    for sheet in wb.sheets() {
        sheet.build_saved_table_view(sheet.rows)?;
    }
    Ok(())
}

pub(crate) fn prepare_replay(
    wb: &Workbook,
    commit: &TableCommit,
    undo: bool,
) -> Result<Workbook, String> {
    if !commit.is_totals_change() {
        return Err("This history entry is not a totals change.".into());
    }
    validate(wb)?;
    let mut candidate = wb.clone();
    candidate.apply_table_commit(commit, undo)?;
    validate(&candidate)?;
    Ok(candidate)
}

impl Spreadsheet {
    pub(crate) fn change_table_totals(
        &mut self,
        id: TableId,
        column: Option<(usize, TableTotal)>,
        cx: &mut Context<Self>,
    ) -> Result<(), String> {
        self.sync_table_view(cx);
        if !self.table_view_installed && (self.row_view.is_sorted() || self.row_view.is_filtered())
        {
            return Err("Clear worksheet sorting and filters before changing totals.".into());
        }
        self.validate_saved_view_layout(self.wb(cx))?;
        validate(self.wb(cx))?;
        let mut candidate = self.wb(cx).clone();
        let (sheet, table) = candidate.table(id).ok_or("Table no longer exists.")?;
        let showing = table.totals_row().is_none();
        let hidden_footer = (column.is_none() && showing
            && self.hidden_rows.get(&sheet).is_some_and(|rows| rows.contains(&(table.range.end_row + 1))))
            .then_some(table.range.end_row + 1);
        let commit = if let Some((col, total)) = column {
            candidate.set_table_total(id, col, total)?
        } else {
            let hidden = self.hidden_rows.get(&sheet).cloned().unwrap_or_default();
            candidate.set_table_totals_visible(id, showing, hidden)?
        };
        validate(&candidate)?;
        self.validate_saved_view_layout(&candidate)?;
        self.workbook
            .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.table_filter_dropdown = None;
        self.sync_table_view(cx);
        self.record_table_commit(commit, "Change Table totals".into(), cx);
        if let Some(row) = hidden_footer {
            self.status_message = Some(format!("Totals added to hidden row {}. Unhide the row to see them.", row + 1));
        }
        Ok(())
    }

    pub(crate) fn toggle_table_totals(&mut self, id: TableId, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) || self.mode.is_editing() || self.mode.is_overlay() {
            return;
        }
        if let Err(error) = self.change_table_totals(id, None, cx) {
            self.status_message = Some(error);
            cx.notify();
        }
    }

    pub(crate) fn replay_table_totals(
        &mut self,
        commit: &TableCommit,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let result = prepare_replay(self.wb(cx), commit, undo).and_then(|wb| {
            self.validate_saved_view_layout(&wb)?;
            Ok(wb)
        });
        match result {
            Ok(candidate) => {
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.table_filter_dropdown = None;
                self.sync_table_view(cx);
                self.bump_cells_rev();
                self.is_modified = true;
                cx.notify();
                true
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                false
            }
        }
    }
}
