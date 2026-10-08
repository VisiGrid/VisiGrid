//! A value that rounds to zero at the displayed decimals shows as zero, with
//! no minus sign: a balanced total is often -2.8e-17 in floating point, and
//! "-0.00" under it reads as an error in the books.

use visigrid_engine::cell::{CellFormat, CellValue, NegativeStyle, NumberFormat};
use visigrid_engine::workbook::Workbook;

fn two_decimals(negative: NegativeStyle) -> NumberFormat {
    NumberFormat::Number { decimals: 2, thousands: true, negative }
}

#[test]
fn a_balanced_two_decimal_total_shows_zero() {
    // What adding 0.3, -0.2 and -0.1 in floating point leaves behind
    let remainder = 0.3_f64 - 0.2 - 0.1;
    assert!(remainder < 0.0 && remainder > -1e-15, "{remainder}");
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, &format!("{remainder:e}"));
    assert_eq!(wb.active_sheet().get_cell(0, 0).value.as_number(), remainder);
    for format in [two_decimals(NegativeStyle::Minus), two_decimals(NegativeStyle::Parens), NumberFormat::Custom("#,##0.00".into())] {
        wb.active_sheet_mut().set_format(0, 0, CellFormat { number_format: format.clone(), ..Default::default() });
        assert_eq!(wb.active_sheet().get_formatted_display(0, 0), "0.00", "{format:?}");
    }
}

#[test]
fn minus_0_004_at_two_decimals_has_no_sign() {
    let n = -0.004;
    assert_eq!(CellValue::format_number(n, &two_decimals(NegativeStyle::Minus)), "0.00");
    assert_eq!(CellValue::format_number(n, &two_decimals(NegativeStyle::Parens)), "0.00");
    assert!(!two_decimals(NegativeStyle::RedMinus).should_render_red(n), "zero isn't red");
    let usd = NumberFormat::Currency { decimals: 2, thousands: true, negative: NegativeStyle::Minus, symbol: None };
    assert_eq!(CellValue::format_number(n, &usd), "$0.00");
    assert_eq!(CellValue::format_number(n, &NumberFormat::General), "0.00");
    assert_eq!(CellValue::format_number(n, &NumberFormat::Custom("#,##0.00".into())), "0.00");
    assert_eq!(CellValue::format_number(n, &NumberFormat::Custom("#,##0.00;(#,##0.00)".into())), "0.00");
    assert_eq!(CellValue::format_number(-0.00004, &NumberFormat::Percent { decimals: 2 }), "0.00%");
    assert_eq!(CellValue::format_number(-0.0, &two_decimals(NegativeStyle::Minus)), "0.00", "negative zero");

    // A custom code's zero section applies, as it does to a true 0
    let accounting = NumberFormat::Custom("#,##0.00;(#,##0.00);\"-\"".into());
    assert_eq!(CellValue::format_number(0.0, &accounting), "-");
    assert_eq!(CellValue::format_number(n, &accounting), "-");
    assert_eq!(CellValue::format_number(-2.7755575615628914e-17, &accounting), "-");
    assert_eq!(CellValue::format_number(-12.5, &accounting), "(12.50)");

    // Anything that shows a digit keeps its sign
    assert_eq!(CellValue::format_number(-0.005, &two_decimals(NegativeStyle::Minus)), "-0.01");
    assert_eq!(CellValue::format_number(-0.004, &NumberFormat::Number { decimals: 3, thousands: false, negative: NegativeStyle::Minus }), "-0.004");
    assert!(two_decimals(NegativeStyle::RedMinus).should_render_red(-0.005));
    assert_eq!(CellValue::format_number(-12.5, &NumberFormat::Custom("#,##0.00;(#,##0.00)".into())), "(12.50)");
}
