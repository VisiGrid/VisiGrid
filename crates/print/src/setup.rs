//! Resolve saved source-coordinate presentation settings against a captured view.
use crate::{
    snapshot::{self, Scope, SheetView, Snapshot},
    PageSettings, Paper, Scale,
};
use visigrid_engine::{
    print_setup::{PrintArea, PrintPaper, PrintScale, PrintSetup},
    sheet::Sheet,
};

pub fn page_settings(snapshot: &Snapshot, setup: &PrintSetup) -> Result<PageSettings, String> {
    setup.validate()?;
    let repeat_rows = if let Some(rows) = setup.repeat_rows {
        let is_title = |index| index >= rows.start && index <= rows.end;
        let count = snapshot
            .layout
            .rows
            .iter()
            .take_while(|r| is_title(r.source_index))
            .count();
        if count == 0
            || snapshot.layout.rows[count..]
                .iter()
                .any(|r| is_title(r.source_index))
        {
            return Err("Repeated headers must be the top visible rows in the print area. Clear the headers, adjust the area, or clear the sort.".into());
        }
        if snapshot.layout.rows[..count]
            .windows(2)
            .any(|r| r[0].source_index >= r[1].source_index)
        {
            return Err(
                "Repeated header rows must be in sheet order. Clear the sort before printing."
                    .into(),
            );
        }
        count
    } else {
        0
    };
    Ok(PageSettings {
        paper: match setup.paper {
            PrintPaper::A4 => Paper::A4,
            PrintPaper::Letter => Paper::Letter,
            PrintPaper::Legal => Paper::Legal,
        },
        landscape: setup.landscape,
        scale: match setup.scale {
            PrintScale::FitColumns => Scale::FitColumns,
            PrintScale::Actual => Scale::Actual,
            PrintScale::FitSheet => Scale::FitSheet,
        },
        footer: setup.page_numbers,
        gridlines: setup.gridlines,
        repeat_rows,
        ..Default::default()
    })
}

/// A saved area is a source rectangle, even when display rows are sorted.
/// Expand intersecting merges before dropping outside rows/columns; otherwise
/// a merge could be silently truncated by the filtered view.
pub fn capture_area(
    sheet: &Sheet,
    view: &SheetView,
    mut area: PrintArea,
    font: &str,
    size: f32,
) -> Result<Snapshot, String> {
    PrintSetup {
        area: Some(area),
        ..Default::default()
    }
    .validate()?;
    loop {
        let old = area;
        for merge in &sheet.merged_regions {
            if merge.start.0 <= area.end_row
                && merge.end.0 >= area.start_row
                && merge.start.1 <= area.end_col
                && merge.end.1 >= area.start_col
            {
                area.start_row = area.start_row.min(merge.start.0);
                area.end_row = area.end_row.max(merge.end.0);
                area.start_col = area.start_col.min(merge.start.1);
                area.end_col = area.end_col.max(merge.end.1);
            }
        }
        if area == old {
            break;
        }
    }
    let view = SheetView {
        rows: view
            .rows
            .iter()
            .copied()
            .filter(|r| r.source_index >= area.start_row && r.source_index <= area.end_row)
            .collect(),
        columns: view
            .columns
            .iter()
            .copied()
            .filter(|c| c.source_index >= area.start_col && c.source_index <= area.end_col)
            .collect(),
    };
    let scope = Scope {
        rows: 0..view.rows.len(),
        columns: 0..view.columns.len(),
    };
    snapshot::capture(sheet, &view, Some(scope), font, size)
}
