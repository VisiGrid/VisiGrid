//! Headless XLSX export must remain usable for sorted 0.44 workbooks.
use std::process::Command;
use visigrid_engine::{
    filter::SortDirection,
    table::TableRange,
    table_view::{TableSort, TableViewSpec},
    workbook::Workbook,
};
use visigrid_io::{native, xlsx};

fn sorted_book(lookup: bool) -> Workbook {
    let mut wb = Workbook::new();
    for (row, values) in [["Key", "Amount"], ["2", "20"], ["1", "10"]]
        .iter()
        .enumerate()
    {
        for (col, value) in values.iter().enumerate() {
            wb.set_cell_value_tracked(0, row, col, value);
        }
    }
    let sid = wb.active_sheet_id();
    let id = wb
        .create_table(
            sid,
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 2,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    let mut view = TableViewSpec::new(id);
    view.sort = Some(TableSort {
        column: wb.table(id).unwrap().1.columns[0].id,
        direction: SortDirection::Ascending,
    });
    wb.set_table_view_spec(sid, Some(view)).unwrap();
    let other = wb.add_sheet_named("Lookup").unwrap();
    if lookup {
        wb.set_cell_value_tracked(other, 0, 0, "=VLOOKUP(1,Sheet1!A2:B3,2,FALSE)");
    }
    assert_eq!(wb.saved_tables().version, 3, "0.44-compatible catalog");
    wb
}

#[test]
fn convert_sorted_native_with_vlookup_falls_back_for_file_and_binary_stdout_even_when_quiet() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("input.sheet");
    let output = dir.path().join("output.xlsx");
    let wb = sorted_book(true);
    native::save_workbook(&wb, &input).unwrap();
    let original = std::fs::read(&input).unwrap();
    assert!(xlsx::export_to_buffer(&wb, None)
        .unwrap_err()
        .contains("VLOOKUP"));
    for file in [true, false] {
        let mut cmd = Command::new(env!("CARGO_BIN_EXE_vgrid"));
        cmd.arg("convert")
            .arg(&input)
            .args(["-t", "xlsx", "--quiet"]);
        if file {
            cmd.arg("-o").arg(&output);
        }
        let result = cmd.output().unwrap();
        let stderr = String::from_utf8_lossy(&result.stderr);
        assert!(result.status.success(), "{stderr}");
        assert!(
            stderr.contains("warning:") && stderr.contains("VLOOKUP"),
            "{stderr}"
        );
        assert!(
            stderr.contains("Exported in stored order instead") && stderr.contains("Reapply"),
            "{stderr}"
        );
        if file {
            assert!(result.stdout.is_empty());
        } else {
            assert!(
                result.stdout.starts_with(b"PK"),
                "stdout is only XLSX bytes"
            );
            std::fs::write(&output, result.stdout).unwrap();
        }
        let (loaded, _) = xlsx::import(&output).unwrap();
        assert_eq!(loaded.sheet(0).unwrap().get_raw(1, 0), "2");
        assert_eq!(loaded.sheet(0).unwrap().get_raw(2, 0), "1");
        assert!(loaded
            .sheet(0)
            .unwrap()
            .table_view_spec()
            .unwrap()
            .sort
            .is_some());
        assert_eq!(
            loaded.sheet(1).unwrap().get_raw(0, 0),
            wb.sheet(1).unwrap().get_raw(0, 0)
        );
        assert_eq!(loaded.sheet(1).unwrap().get_display(0, 0), "10");
        assert_eq!(std::fs::read(&input).unwrap(), original);
    }
}

#[test]
fn convert_still_materializes_supported_sorting() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("input.sheet");
    let output = dir.path().join("output.xlsx");
    native::save_workbook(&sorted_book(false), &input).unwrap();
    let result = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .arg("convert")
        .arg(&input)
        .args(["-t", "xlsx", "-o"])
        .arg(&output)
        .output()
        .unwrap();
    let stderr = String::from_utf8_lossy(&result.stderr);
    assert!(result.status.success(), "{stderr}");
    assert!(!stderr.contains("stored order instead"), "{stderr}");
    let (loaded, _) = xlsx::import(&output).unwrap();
    assert_eq!(loaded.sheet(0).unwrap().get_raw(1, 0), "1");
    assert_eq!(loaded.sheet(0).unwrap().get_raw(2, 0), "2");
}

#[test]
fn convert_does_not_fallback_past_schema_refusal_or_touch_destination() {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("input.sheet");
    let output = dir.path().join("output.xlsx");
    let mut wb = sorted_book(true);
    let id = wb.tables().next().unwrap().1.id;
    wb.rename_table(id, &"T".repeat(256)).unwrap();
    native::save_workbook(&wb, &input).unwrap();
    std::fs::write(&output, b"keep destination").unwrap();
    let result = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .arg("convert")
        .arg(&input)
        .args(["-t", "xlsx", "-o"])
        .arg(&output)
        .output()
        .unwrap();
    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("255-character"));
    assert!(result.stdout.is_empty());
    assert_eq!(std::fs::read(&output).unwrap(), b"keep destination");
}
