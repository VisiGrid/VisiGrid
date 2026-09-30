//! XNPV, XIRR, RATE and NPER against the worked examples in Microsoft's
//! documentation for each function.

use visigrid_engine::workbook::Workbook;

fn value(wb: &mut Workbook, f: &str) -> f64 {
    wb.set_cell_value_tracked(0, 20, 0, f);
    let shown = wb.active_sheet().get_display(20, 0);
    match wb.active_sheet().get_computed_value(20, 0) {
        visigrid_engine::formula::eval::Value::Number(n) => n,
        other => panic!("{f}: expected a number, got {other:?} ({shown})"),
    }
}

fn close(got: f64, want: f64, tol: f64, what: &str) {
    assert!((got - want).abs() < tol, "{what}: got {got}, want {want}");
}

/// Microsoft's XNPV/XIRR example: five flows over fifteen months.
fn dated_flows() -> Workbook {
    let mut wb = Workbook::new();
    let rows = [
        ("-10000", "=DATE(2008,1,1)"),
        ("2750", "=DATE(2008,3,1)"),
        ("4250", "=DATE(2008,10,30)"),
        ("3250", "=DATE(2009,2,15)"),
        ("2750", "=DATE(2009,4,1)"),
    ];
    for (r, (v, d)) in rows.iter().enumerate() {
        wb.set_cell_value_tracked(0, r, 0, v);
        wb.set_cell_value_tracked(0, r, 1, d);
    }
    wb
}

#[test]
fn xnpv_and_xirr_match_microsofts_example() {
    let mut wb = dated_flows();
    close(value(&mut wb, "=XNPV(0.09,A1:A5,B1:B5)"), 2086.6476, 0.001, "XNPV");
    close(value(&mut wb, "=XIRR(A1:A5,B1:B5)"), 0.373362535, 1e-8, "XIRR");
    close(value(&mut wb, "=XIRR(A1:A5,B1:B5,0.5)"), 0.373362535, 1e-8, "XIRR with a guess");
    // At the XIRR, the XNPV is zero.
    close(value(&mut wb, "=XNPV(XIRR(A1:A5,B1:B5),A1:A5,B1:B5)"), 0.0, 1e-6, "XNPV at XIRR");
}

#[test]
fn xirr_and_xnpv_errors() {
    let mut wb = dated_flows();
    wb.set_cell_value_tracked(0, 10, 0, "=XIRR(A2:A5,B2:B5)");
    assert_eq!(wb.active_sheet().get_display(10, 0), "#NUM!", "no negative flow");
    wb.set_cell_value_tracked(0, 11, 0, "=XNPV(0.09,A1:A5,B1:B4)");
    assert_eq!(wb.active_sheet().get_display(11, 0), "#NUM!", "different counts");
    // A date before the first is #NUM!.
    wb.set_cell_value_tracked(0, 2, 1, "=DATE(2007,1,1)");
    wb.set_cell_value_tracked(0, 12, 0, "=XNPV(0.09,A1:A5,B1:B5)");
    assert_eq!(wb.active_sheet().get_display(12, 0), "#NUM!", "date before the first");
}

#[test]
fn rate_matches_microsofts_example() {
    let mut wb = Workbook::new();
    // 4-year loan of 8,000 paid 200 a month.
    close(value(&mut wb, "=RATE(48,-200,8000)"), 0.00770147, 1e-8, "RATE monthly");
    close(value(&mut wb, "=RATE(48,-200,8000)*12"), 0.09241767, 1e-7, "RATE annual");
    // RATE inverts PMT.
    close(value(&mut wb, "=RATE(360,PMT(0.05/12,360,200000),200000)*12"), 0.05, 1e-9, "RATE of PMT");
    close(value(&mut wb, "=RATE(10,-100,1000)"), 0.0, 1e-9, "zero rate");
    close(value(&mut wb, "=RATE(5,0,-1000,1500)"), 0.08447177, 1e-7, "lump sum growth");
}

#[test]
fn nper_matches_microsofts_example() {
    let mut wb = Workbook::new();
    close(value(&mut wb, "=NPER(0.12/12,-100,-1000,10000,1)"), 59.6738657, 1e-6, "NPER type 1");
    close(value(&mut wb, "=NPER(0.12/12,-100,-1000,10000)"), 60.0821229, 1e-6, "NPER");
    close(value(&mut wb, "=NPER(0.12/12,-100,-1000)"), -9.57859404, 1e-6, "NPER negative");
    close(value(&mut wb, "=NPER(0,-100,1000)"), 10.0, 1e-12, "NPER zero rate");
    // NPER inverts PMT.
    close(value(&mut wb, "=NPER(0.05/12,PMT(0.05/12,360,200000),200000)"), 360.0, 1e-6, "NPER of PMT");
}
