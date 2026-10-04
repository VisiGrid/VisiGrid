//! Recipe results as Tables: load one into a new workbook, and refresh a
//! linked Table in place.
//!
//! Refresh replaces the Table's records and nothing else on the sheet. It
//! works on whatever workbook it is given; the desktop passes a candidate
//! copy and publishes it only if this returns Ok, so a refused refresh leaves
//! the open workbook untouched. Callers must not refresh from a run whose
//! report is not ok.

use std::path::Path;

use visigrid_engine::cell::{CellFormat, ValueRef};
use visigrid_engine::sheet::{SheetId, NUM_COLS, NUM_ROWS};
use visigrid_engine::table::{self, RefreshStamp, TableId, TableRange, TableSource};
use visigrid_engine::workbook::Workbook;

use crate::recipe::{RecipeOutput, RunReport};

/// What a refresh changed, for the confirmation the user sees.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableRefresh {
    pub rows_before: usize,
    pub rows_after: usize,
    /// The recipe's columns differ from the Table's (names or count).
    pub columns_changed: bool,
}

/// The stamp for a run that is about to publish.
pub fn stamp(report: &RunReport, source: &Path) -> RefreshStamp {
    RefreshStamp {
        at: chrono::Utc::now().to_rfc3339_opts(chrono::SecondsFormat::Secs, true),
        source: source.display().to_string(),
        snapshot: report.snapshot.clone(),
        rows: report.rows,
    }
}

/// A new one-sheet workbook holding the result as a Table linked to `link`.
pub fn new_workbook(output: &RecipeOutput, table_name: &str, link: TableSource) -> Result<Workbook, String> {
    if output.columns.is_empty() {
        return Err("The recipe produced no columns.".into());
    }
    let mut wb = Workbook::from_sheets(vec![output.to_sheet()], 0);
    let sheet_id = wb.sheet(0).ok_or("The new workbook has no sheet.")?.id;
    let range = TableRange {
        start_row: 0,
        start_col: 0,
        end_row: output.rows.len(),
        end_col: output.columns.len() - 1,
    };
    let id = wb.create_table(sheet_id, range, table_name)?.table_id();
    wb.set_table_source(id, Some(link))?;
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    Ok(wb)
}

/// A Table name from a recipe's file name: `orders.recipe.toml` -> `orders`.
/// Falls back to `Table1` when nothing usable is left.
pub fn table_name_for(recipe_path: &Path) -> String {
    let file = recipe_path.file_name().and_then(|n| n.to_str()).unwrap_or("");
    let stem = file.strip_suffix(".recipe.toml").or_else(|| file.strip_suffix(".toml")).unwrap_or(file);
    let mut name: String = stem
        .chars()
        .map(|c| if c.is_ascii_alphanumeric() { c } else { '_' })
        .collect();
    if name.starts_with(|c: char| c.is_ascii_digit()) {
        name.insert(0, '_');
    }
    if name.trim_matches('_').is_empty() {
        "Table1".into()
    } else {
        name
    }
}

/// Replace the records of Table `id` with `output` and record `link`.
///
/// The header row and first column stay where they are. The Table grows or
/// shrinks to fit; it refuses to grow over cells that are not empty. Formats
/// a column's first record had are kept for every new record. Calculated
/// columns are refused for now: the recipe owns every column of the Table.
pub fn refresh_table(
    wb: &mut Workbook,
    id: TableId,
    output: &RecipeOutput,
    link: TableSource,
) -> Result<TableRefresh, String> {
    wb.ensure_writable()?;
    let (sheet_id, old) = wb.table(id).map(|(s, t)| (s, t.clone())).ok_or("The Table no longer exists.")?;
    if old.columns.iter().any(|c| c.formula.is_some()) {
        return Err(format!(
            "Table {} has calculated columns. Refresh replaces every column with the recipe's result, so it can't keep them yet. Move the calculations next to the Table, or remove them.",
            old.name
        ));
    }
    let width = output.columns.len();
    if width == 0 {
        return Err("The recipe produced no columns.".into());
    }
    let (r0, c0) = (old.range.start_row, old.range.start_col);
    let range = TableRange {
        start_row: r0,
        start_col: c0,
        end_row: r0 + output.rows.len(),
        end_col: c0 + width - 1,
    };
    let names = table::normalize_headers(&output.columns.iter().map(|c| c.name.clone()).collect::<Vec<_>>());
    let old_names: Vec<String> = old.columns.iter().map(|c| c.name.clone()).collect();
    let columns_changed = names != old_names;

    let sheet = sheet_ref(wb, sheet_id)?;
    // Growing must not cover anything: refuse rather than absorb cells
    // into the Table or write over them.
    for ((row, col), cell) in sheet.cells_iter() {
        let inside_new = row >= range.start_row && row <= range.end_row && col >= range.start_col && col <= range.end_col;
        let inside_old = row >= old.range.start_row && row <= old.range.end_row && col >= old.range.start_col && col <= old.range.end_col;
        if inside_new && !inside_old && (!matches!(cell.value(), ValueRef::Empty) || cell.comment().is_some()) {
            return Err(format!(
                "The new result needs {} rows × {} columns, and cell {}{} is in the way. Move what's there, then refresh again.",
                output.rows.len() + 1,
                width,
                visigrid_engine::formula::parser::column_letters_pub(col),
                row + 1
            ));
        }
    }
    // The first record's format, per column, carries over to every record
    let templates: Vec<Option<CellFormat>> = (0..width)
        .map(|c| {
            (c < old.columns.len() && old.range.end_row > r0)
                .then(|| sheet.get_format(r0 + 1, c0 + c))
                .filter(|f| *f != CellFormat::default())
        })
        .collect();

    let sheet = wb.sheet_by_id_mut(sheet_id).ok_or("The Table's sheet no longer exists.")?;
    for row in r0 + 1..=old.range.end_row {
        for col in old.range.start_col..=old.range.end_col {
            sheet.clear_cell(row, col);
        }
    }
    // A sheet made from a recipe is sized to its first result; a larger
    // refresh needs room before the Table can grow into it
    if range.end_row >= NUM_ROWS || range.end_col >= NUM_COLS {
        return Err(format!(
            "The new result needs {} rows × {} columns from {}, more than a sheet holds",
            output.rows.len() + 1,
            width,
            visigrid_engine::formula::parser::column_letters_pub(c0)
        ));
    }
    sheet.rows = sheet.rows.max(range.end_row + 1);
    sheet.cols = sheet.cols.max(range.end_col + 1);
    // resize_table names added columns from the header cells beside the Table
    for (c, name) in names.iter().enumerate().skip(old.columns.len()) {
        sheet.set_text(r0, c0 + c, name);
    }
    wb.resize_table(id, range)?;
    // Columns the result no longer has leave their old header behind
    let sheet = wb.sheet_by_id_mut(sheet_id).ok_or("The Table's sheet no longer exists.")?;
    for col in range.end_col + 1..=old.range.end_col.max(range.end_col) {
        sheet.clear_cell(r0, col);
    }
    if columns_changed {
        wb.rename_table_columns(id, &names)?;
    }
    let sheet = wb.sheet_by_id_mut(sheet_id).ok_or("The Table's sheet no longer exists.")?;
    for (r, values) in output.rows.iter().enumerate() {
        for (c, value) in values.iter().enumerate() {
            output.write_value(sheet, r0 + 1 + r, c0 + c, c, value);
            if let Some(format) = &templates[c] {
                sheet.set_format(r0 + 1 + r, c0 + c, format.clone());
            }
        }
    }
    wb.set_table_source(id, Some(link))?;
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    Ok(TableRefresh {
        rows_before: old.range.end_row - r0,
        rows_after: output.rows.len(),
        columns_changed,
    })
}

fn sheet_ref(wb: &Workbook, id: SheetId) -> Result<&visigrid_engine::sheet::Sheet, String> {
    wb.sheet_by_id(id).ok_or_else(|| "The Table's sheet no longer exists.".into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::recipe::{run, Recipe, Snapshot};

    const RECIPE: &str = r#"
version = 1
[source]
kind = "csv"
path = "orders.csv"
[[step]]
op = "types"
columns = { Amount = "number" }
"#;

    fn result(csv: &str) -> (RecipeOutput, RunReport) {
        let recipe = Recipe::from_toml(RECIPE).unwrap();
        let snap = Snapshot::from_bytes(Path::new("orders.csv"), csv.as_bytes().to_vec());
        let r = run(&recipe, &snap);
        assert!(r.report.ok, "{:?}", r.report.failures);
        (r.output, r.report)
    }

    fn link(report: &RunReport) -> TableSource {
        TableSource { recipe: "/data/orders.recipe.toml".into(), refreshed: Some(stamp(report, Path::new("orders.csv"))) }
    }

    fn table(wb: &Workbook) -> (SheetId, visigrid_engine::table::DataTable) {
        let (s, t) = wb.tables().next().unwrap();
        (s, t.clone())
    }

    #[test]
    fn loads_a_linked_table_and_names_it_after_the_recipe() {
        let (out, report) = result("ID,Amount\n007,10\n008,20\n");
        let wb = new_workbook(&out, &table_name_for(Path::new("/x/orders.recipe.toml")), link(&report)).unwrap();
        let (_, t) = table(&wb);
        assert_eq!(t.name, "orders");
        assert_eq!((t.range.end_row, t.range.end_col), (2, 1));
        assert_eq!(t.source.as_ref().unwrap().refreshed.as_ref().unwrap().rows, 2);
        assert_eq!(wb.active_sheet().get_display(1, 0), "007");
        assert_eq!(table_name_for(Path::new("2026 export.toml")), "_2026_export");
        assert_eq!(table_name_for(Path::new("--.recipe.toml")), "Table1");
    }

    #[test]
    fn refresh_grows_and_shrinks_and_keeps_column_formats() {
        let (out, report) = result("ID,Amount\n1,10\n2,20\n");
        let mut wb = new_workbook(&out, "orders", link(&report)).unwrap();
        let (sheet_id, t) = table(&wb);
        let bold = CellFormat { bold: true, ..Default::default() };
        wb.sheet_by_id_mut(sheet_id).unwrap().set_format(1, 1, bold.clone());

        let (out, report) = result("ID,Amount\n1,10\n2,20\n3,30\n4,40\n");
        let r = refresh_table(&mut wb, t.id, &out, link(&report)).unwrap();
        assert_eq!((r.rows_before, r.rows_after, r.columns_changed), (2, 4, false));
        let sheet = wb.sheet_by_id(sheet_id).unwrap();
        assert_eq!(sheet.get_display(4, 1), "40");
        assert!(sheet.get_format(4, 1).bold);

        let (out, report) = result("ID,Amount\n9,90\n");
        refresh_table(&mut wb, t.id, &out, link(&report)).unwrap();
        let (_, t) = table(&wb);
        assert_eq!(t.range.end_row, 1);
        let sheet = wb.sheet_by_id(sheet_id).unwrap();
        assert_eq!(sheet.get_display(1, 0), "9");
        // Released rows are empty, not stale records below the Table
        assert_eq!(sheet.get_display(2, 0), "");
        assert_eq!(sheet.get_display(4, 1), "");
    }

    #[test]
    fn refresh_follows_changed_columns() {
        let (out, report) = result("ID,Amount\n1,10\n");
        let mut wb = new_workbook(&out, "orders", link(&report)).unwrap();
        let (sheet_id, t) = table(&wb);
        let (out, report) = result("ID,Amount,Region\n1,10,West\n");
        assert!(refresh_table(&mut wb, t.id, &out, link(&report)).unwrap().columns_changed);
        let (_, t) = table(&wb);
        assert_eq!(t.columns.iter().map(|c| c.name.as_str()).collect::<Vec<_>>(), ["ID", "Amount", "Region"]);
        let (out, report) = result("ID,Amount\n1,10\n");
        refresh_table(&mut wb, t.id, &out, link(&report)).unwrap();
        let (_, t) = table(&wb);
        assert_eq!(t.range.end_col, 1);
        assert_eq!(wb.sheet_by_id(sheet_id).unwrap().get_display(0, 2), "");
    }

    #[test]
    fn refresh_refuses_to_cover_cells_and_changes_nothing() {
        let (out, report) = result("ID,Amount\n1,10\n");
        let mut wb = new_workbook(&out, "orders", link(&report)).unwrap();
        let (sheet_id, t) = table(&wb);
        wb.sheet_by_id_mut(sheet_id).unwrap().set_value(3, 0, "note");
        let before = wb.saved_tables();
        let (out, report) = result("ID,Amount\n1,10\n2,20\n3,30\n");
        let err = refresh_table(&mut wb, t.id, &out, link(&report)).unwrap_err();
        assert!(err.contains("A4"), "{err}");
        assert_eq!(wb.saved_tables().sheets[0].tables, before.sheets[0].tables);
        assert_eq!(wb.sheet_by_id(sheet_id).unwrap().get_display(1, 1), "10");
    }

    #[test]
    fn recipe_link_survives_native_save_as_catalog_v4() {
        let (out, report) = result("ID,Amount\n1,10\n");
        let wb = new_workbook(&out, "orders", link(&report)).unwrap();
        assert_eq!(wb.saved_tables().version, 4);
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("linked.sheet");
        crate::native::save_workbook(&wb, &path).unwrap();
        let back = crate::native::load_workbook(&path).unwrap();
        let (_, t) = table(&back);
        assert_eq!(t.source.unwrap().recipe, "/data/orders.recipe.toml");
    }

    #[test]
    fn refresh_grows_past_the_sheet_size_the_first_result_gave_it() {
        let (out, report) = result("ID,Amount\n1,10\n");
        let mut wb = new_workbook(&out, "orders", link(&report)).unwrap();
        let (sheet_id, t) = table(&wb);
        let before = wb.sheet_by_id(sheet_id).unwrap().rows;
        let mut csv = String::from("ID,Amount\n");
        for i in 0..before + 500 {
            csv.push_str(&format!("{i},{i}\n"));
        }
        let (out, report) = result(&csv);
        refresh_table(&mut wb, t.id, &out, link(&report)).unwrap();
        let (_, t) = table(&wb);
        assert_eq!(t.range.end_row, before + 500);
        assert_eq!(wb.sheet_by_id(sheet_id).unwrap().get_display(before + 500, 0), (before + 499).to_string());
    }
}
