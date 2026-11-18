//! Synthetic dump-fee workbook: nine rounded fees, row totals and a column
//! total. These are sample tonnages, not the customer's original statement.
use std::process::Command;
use visigrid_engine::{formula::eval::Value, workbook::Workbook};
use visigrid_io::{native, xlsx};

const FEES: [f64; 9] = [
    110.12, 223.0, 0.89, 735.9, 1226.5, 1525.32, 1784.0, 2230.0, 1172.58,
];

fn fees_book(broken: bool) -> Workbook {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 2, 8, "4785");
    wb.define_name_for_cell("tons_over", 0, 3, 7).unwrap();
    for (i, tons) in [
        "1.2345", "2.5", "0.01", "8.25", "13.75", "17.1", "20", "25", "13.1455",
    ]
    .iter()
    .enumerate()
    {
        let r = i + 3;
        wb.set_cell_value_tracked(0, r, 7, tons);
        let input = if i == 0 {
            "tons_over".into()
        } else {
            format!("H{}", r + 1)
        };
        wb.set_cell_value_tracked(0, r, 8, &format!("=MROUND({input}*89.2,0.01)"));
        wb.set_cell_value_tracked(0, r, 9, &format!("=G{}+I{}", r + 1, r + 1));
    }
    if broken {
        wb.set_cell_value_tracked(0, 3, 8, "=UNSUPPORTED_FEE(tons_over)");
    }
    wb.set_cell_value_tracked(0, 32, 8, "=SUM(I3:I32)");
    let summary = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(summary, 0, 0, "=Sheet1!I33");
    // A stale cached total must never hide formula errors during conversion.
    wb.active_sheet()
        .cache_computed(32, 8, Value::Number(4785.0));
    wb
}

fn converted(broken: bool) -> (Workbook, Vec<csv::StringRecord>) {
    let dir = tempfile::tempdir().unwrap();
    let input = dir.path().join("fees.xlsx");
    let output = dir.path().join("fees.sheet");
    let mut layout = xlsx::ExportLayout::default();
    layout.col_widths.insert(7, 180.0);
    layout.row_heights.insert(3, 35.0);
    layout.hidden_rows.push(9);
    layout.hidden_cols.push(6);
    xlsx::export_with_order(
        &fees_book(broken),
        &input,
        Some(&[layout, xlsx::ExportLayout::default()]),
        xlsx::ExportOrder::Stored,
    )
    .unwrap();
    let result = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .arg("convert")
        .arg(&input)
        .args(["-t", "sheet", "-o"])
        .arg(&output)
        .output()
        .unwrap();
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let loaded = native::load_workbook(&output).unwrap();
    let layout = native::load_layout(&output);
    assert!((layout.col_widths[&0][&7] - 180.0).abs() < 2.0);
    assert!((layout.row_heights[&0][&3] - 35.0).abs() < 0.1);
    assert!(layout.hidden_rows[&0].contains(&9));
    assert!(layout.hidden_cols[&0].contains(&6));
    let result = Command::new(env!("CARGO_BIN_EXE_vgrid"))
        .arg("convert")
        .arg(&input)
        .args(["-t", "csv"])
        .output()
        .unwrap();
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let records = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_reader(result.stdout.as_slice())
        .records()
        .collect::<Result<Vec<_>, _>>()
        .unwrap();
    (loaded, records)
}

#[test]
fn convert_recalculates_nine_mround_fees_and_totals() {
    let (wb, csv) = converted(false);
    assert_eq!(wb.sheet_count(), 2);
    assert!(wb.named_ranges().get("tons_over").is_some());
    assert_eq!(
        wb.sheet(1).unwrap().get_computed_value(0, 0),
        Value::Number(13793.31)
    );
    for (i, fee) in FEES.into_iter().enumerate() {
        for col in [8, 9] {
            let actual = wb.active_sheet().get_computed_value(i + 3, col);
            match actual {
                Value::Number(n) => assert!((n - fee).abs() < 1e-9, "{i}, {col}: {n}"),
                other => panic!("{i}, {col}: {other:?}"),
            }
        }
        assert_eq!(csv[i + 3][8].parse::<f64>().unwrap(), fee);
    }
    assert_eq!(
        wb.active_sheet().get_computed_value(32, 8),
        Value::Number(13793.31)
    );
    assert_eq!(csv[32][8].parse::<f64>().unwrap(), 13793.31);
}

#[test]
fn convert_exposes_unknown_fee_errors_in_totals() {
    let (wb, csv) = converted(true);
    let error = wb.active_sheet().get_computed_value(3, 8);
    assert!(matches!(error, Value::Error(_)), "{error:?}");
    assert_eq!(wb.active_sheet().get_computed_value(3, 9), error);
    assert_eq!(wb.active_sheet().get_computed_value(32, 8), error);
    assert_eq!(&csv[32][8], &csv[3][8]);
    assert!(!csv[32][8].is_empty());
    assert!(csv[32][8].parse::<f64>().is_err());
}
