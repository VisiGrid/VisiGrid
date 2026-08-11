// Date/time functions: TODAY, NOW, DATE, DATEVALUE, YEAR, MONTH, DAY, WEEKDAY, DATEDIF,
// EDATE, EOMONTH, HOUR, MINUTE, SECOND

use super::eval::{evaluate, CellLookup, EvalResult};
use super::eval_helpers::{date_to_serial, serial_to_date, days_in_month, try_parse_date_string};
use super::parser::BoundExpr;

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "TODAY" => {
            if !args.is_empty() {
                return Some(EvalResult::Error("TODAY takes no arguments".to_string()));
            }
            // Excel-style date serial (days since 1899-12-30), for the
            // user's own calendar day rather than UTC's — TODAY() is local in
            // Excel, and a UTC rollover puts anyone west of it a day ahead all
            // evening.
            let now = crate::timing::now_since_epoch();
            let local_secs =
                now.as_secs() as i64 + crate::timing::local_utc_offset_seconds();
            // div_euclid, not /: west of UTC before 1970 the offset can push
            // this negative, and truncating division rounds that the wrong way.
            let days_since_unix = local_secs.div_euclid(86400);
            let excel_date = days_since_unix as f64 + 25569.0;
            EvalResult::Number(excel_date)
        }
        "NOW" => {
            if !args.is_empty() {
                return Some(EvalResult::Error("NOW takes no arguments".to_string()));
            }
            let now = crate::timing::now_since_epoch();
            let secs = now.as_secs() as f64
                + now.subsec_nanos() as f64 / 1_000_000_000.0
                + crate::timing::local_utc_offset_seconds() as f64;
            let days_since_unix = secs / 86400.0;
            let excel_datetime = days_since_unix + 25569.0;
            EvalResult::Number(excel_datetime)
        }
        "TIME" => {
            if args.len() != 3 {
                return Some(EvalResult::Error("TIME requires exactly 3 arguments".to_string()));
            }
            let mut nums = [0i64; 3];
            for (slot, arg) in nums.iter_mut().zip(args.iter()) {
                match evaluate(arg, lookup).to_number() {
                    Ok(n) => *slot = n.trunc() as i64,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            }
            let (hours, minutes, seconds) = (nums[0], nums[1], nums[2]);
            // Excel carries minutes and seconds past their range into hours —
            // TIME(0,90,0) is 01:30 — then keeps only the fractional day, so
            // TIME(25,0,0) is 01:00 rather than an error.
            let total = hours
                .saturating_mul(3600)
                .saturating_add(minutes.saturating_mul(60))
                .saturating_add(seconds);
            if total < 0 {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }
            let day_fraction = (total % 86400) as f64 / 86400.0;
            EvalResult::Number(day_fraction)
        }
        "DATE" => {
            // DATE(year, month, day) - returns Excel date serial
            if args.len() != 3 {
                return Some(EvalResult::Error("DATE requires exactly 3 arguments".to_string()));
            }
            let year = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let month = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let day = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            // Adjust year if 0-99 (Excel convention)
            let year = if year < 100 { year + 1900 } else { year };

            // Simple date to Excel serial conversion
            let serial = date_to_serial(year, month, day);
            EvalResult::Number(serial)
        }
        "DATEVALUE" => {
            // DATEVALUE(date_text) - converts a date string to Excel serial number
            // Supports ISO (2023-11-07) and US (11/07/2023) formats
            if args.len() != 1 {
                return Some(EvalResult::Error("DATEVALUE requires exactly 1 argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            match try_parse_date_string(&text) {
                Some(serial) => EvalResult::Number(serial),
                None => EvalResult::Error(format!("#VALUE! Cannot parse '{}' as date", text)),
            }
        }
        "YEAR" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("YEAR requires exactly one argument".to_string()));
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let (year, _, _) = serial_to_date(serial);
            EvalResult::Number(year as f64)
        }
        "MONTH" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("MONTH requires exactly one argument".to_string()));
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let (_, month, _) = serial_to_date(serial);
            EvalResult::Number(month as f64)
        }
        "DAY" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("DAY requires exactly one argument".to_string()));
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let (_, _, day) = serial_to_date(serial);
            EvalResult::Number(day as f64)
        }
        "DAYS" => {
            // DAYS(end, start) — plain subtraction, but Excel has it and a
            // reader reaching for it should not meet "Unknown function".
            if args.len() != 2 {
                return Some(EvalResult::Error("DAYS requires exactly 2 arguments".to_string()));
            }
            let end = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n.trunc(),
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let start = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n.trunc(),
                Err(e) => return Some(EvalResult::Error(e)),
            };
            // Negative when the end precedes the start, as Excel does.
            EvalResult::Number(end - start)
        }
        "NETWORKDAYS" | "WORKDAY" => {
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error(format!("{name} requires 2 or 3 arguments")));
            }
            let start = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n.trunc() as i64,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let second = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n.trunc() as i64,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            // Holidays are skipped like weekends. Excel accepts a range here;
            // a single date is the common case and both arrive as values.
            let mut holidays: Vec<i64> = Vec::new();
            if args.len() == 3 {
                match evaluate(&args[2], lookup) {
                    EvalResult::Array(array) => {
                        for row in 0..array.rows() {
                            for col in 0..array.cols() {
                                if let Some(value) = array.get(row, col) {
                                    if let Ok(n) = value.to_number() {
                                        holidays.push(n.trunc() as i64);
                                    }
                                }
                            }
                        }
                    }
                    other => {
                        if let Ok(n) = other.to_number() {
                            holidays.push(n.trunc() as i64);
                        }
                    }
                }
            }

            // Same weekday derivation as WEEKDAY above: 0 is Sunday.
            let is_working = |serial: i64| {
                let weekday = (serial + 6).rem_euclid(7);
                weekday != 0 && weekday != 6 && !holidays.contains(&serial)
            };

            if name == "NETWORKDAYS" {
                // Inclusive of both ends, and a reversed range counts negative
                // rather than erroring, which is what Excel does.
                let (lo, hi, sign) = if start <= second {
                    (start, second, 1i64)
                } else {
                    (second, start, -1i64)
                };
                let count = (lo..=hi).filter(|d| is_working(*d)).count() as i64;
                EvalResult::Number((count * sign) as f64)
            } else {
                // WORKDAY steps over non-working days; day zero returns the
                // start unchanged even when it is itself a weekend.
                let step = if second >= 0 { 1i64 } else { -1i64 };
                let mut remaining = second.abs();
                let mut cursor = start;
                while remaining > 0 {
                    cursor += step;
                    if is_working(cursor) {
                        remaining -= 1;
                    }
                }
                EvalResult::Number(cursor as f64)
            }
        }
        "WEEKDAY" => {
            // WEEKDAY(date, [type]) - returns day of week
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("WEEKDAY requires 1 or 2 arguments".to_string()));
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n as i64,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let return_type = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as i32,
                    Err(_) => 1,
                }
            } else {
                1
            };

            let weekday = ((serial + 6) % 7) as i32; // 0 = Sunday, 6 = Saturday

            let result = match return_type {
                1 => weekday + 1,        // 1 (Sunday) to 7 (Saturday)
                2 => if weekday == 0 { 7 } else { weekday }, // 1 (Monday) to 7 (Sunday)
                3 => if weekday == 0 { 6 } else { weekday - 1 }, // 0 (Monday) to 6 (Sunday)
                _ => weekday + 1,
            };
            EvalResult::Number(result as f64)
        }
        "DATEDIF" => {
            // DATEDIF(start_date, end_date, unit)
            if args.len() != 3 {
                return Some(EvalResult::Error("DATEDIF requires exactly 3 arguments".to_string()));
            }
            let start_serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let end_serial = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let unit = evaluate(&args[2], lookup).to_text().to_uppercase();

            if start_serial > end_serial {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }

            let (start_y, start_m, start_d) = serial_to_date(start_serial);
            let (end_y, end_m, end_d) = serial_to_date(end_serial);

            let result = match unit.as_str() {
                "Y" => {
                    // Complete years
                    let mut years = end_y - start_y;
                    if end_m < start_m || (end_m == start_m && end_d < start_d) {
                        years -= 1;
                    }
                    years as f64
                }
                "M" => {
                    // Complete months
                    let mut months = (end_y - start_y) * 12 + (end_m - start_m);
                    if end_d < start_d {
                        months -= 1;
                    }
                    months as f64
                }
                "D" => {
                    // Days
                    (end_serial - start_serial).floor()
                }
                "YM" => {
                    // Months ignoring years
                    let mut months = end_m - start_m;
                    if end_d < start_d {
                        months -= 1;
                    }
                    if months < 0 {
                        months += 12;
                    }
                    months as f64
                }
                "YD" => {
                    // Days ignoring years
                    let end_in_start_year = date_to_serial(start_y, end_m, end_d);
                    let mut days = end_in_start_year - start_serial;
                    if days < 0.0 {
                        let end_in_next_year = date_to_serial(start_y + 1, end_m, end_d);
                        days = end_in_next_year - start_serial;
                    }
                    days.floor()
                }
                "MD" => {
                    // Days ignoring months and years
                    let mut days = end_d - start_d;
                    if days < 0 {
                        // Days in previous month (simplified)
                        days += 30;
                    }
                    days as f64
                }
                _ => return Some(EvalResult::Error("#VALUE!".to_string())),
            };
            EvalResult::Number(result)
        }
        "EDATE" => {
            // EDATE(start_date, months) - add months to a date
            if args.len() != 2 {
                return Some(EvalResult::Error("EDATE requires exactly 2 arguments".to_string()));
            }
            let start_serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let months = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            let (year, month, day) = serial_to_date(start_serial);
            let total_months = year * 12 + month + months;
            let new_year = (total_months - 1) / 12;
            let new_month = ((total_months - 1) % 12) + 1;

            // Clamp day to valid range for new month
            let dim = days_in_month(new_year, new_month);
            let new_day = day.min(dim);

            EvalResult::Number(date_to_serial(new_year, new_month, new_day))
        }
        "EOMONTH" => {
            // EOMONTH(start_date, months) - end of month after adding months
            if args.len() != 2 {
                return Some(EvalResult::Error("EOMONTH requires exactly 2 arguments".to_string()));
            }
            let start_serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let months = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n as i32,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            let (year, month, _) = serial_to_date(start_serial);
            let total_months = year * 12 + month + months;
            let new_year = (total_months - 1) / 12;
            let new_month = ((total_months - 1) % 12) + 1;
            let last_day = days_in_month(new_year, new_month);

            EvalResult::Number(date_to_serial(new_year, new_month, last_day))
        }
        "HOUR" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("HOUR requires exactly one argument".to_string()));
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let time_part = serial.fract();
            let hours = (time_part * 24.0).floor() as i32 % 24;
            EvalResult::Number(hours as f64)
        }
        "MINUTE" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("MINUTE requires exactly one argument".to_string()));
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let time_part = serial.fract();
            let total_minutes = (time_part * 24.0 * 60.0).floor() as i32;
            let minutes = total_minutes % 60;
            EvalResult::Number(minutes as f64)
        }
        "SECOND" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("SECOND requires exactly one argument".to_string()));
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let time_part = serial.fract();
            let total_seconds = (time_part * 24.0 * 60.0 * 60.0).floor() as i32;
            let seconds = total_seconds % 60;
            EvalResult::Number(seconds as f64)
        }
        _ => return None,
    };
    Some(result)
}

#[cfg(test)]
mod time_tests {
    use crate::formula::eval::{evaluate, CellLookup, EvalResult};
    use crate::formula::parser::{bind_expr_same_sheet, parse};

    struct Empty;
    impl CellLookup for Empty {
        fn get_value(&self, _r: usize, _c: usize) -> f64 { 0.0 }
        fn get_text(&self, _r: usize, _c: usize) -> String { String::new() }
    }

    fn number(formula: &str) -> f64 {
        match evaluate(&bind_expr_same_sheet(&parse(formula).unwrap()), &Empty) {
            EvalResult::Number(n) => n,
            other => panic!("{formula} gave {other:?}, expected a number"),
        }
    }

    /// DAYS is subtraction, and exists so that reaching for it works.
    #[test]
    fn days_is_the_difference_between_two_dates() {
        assert_eq!(number("=DAYS(DATE(2026,8,7),DATE(2026,8,3))"), 4.0);
        // Four days apart, five working days inclusive — which is the
        // distinction NETWORKDAYS exists to make.
        assert_eq!(number("=NETWORKDAYS(DATE(2026,8,3),DATE(2026,8,7))"), 5.0);
        assert_eq!(number("=DAYS(DATE(2026,8,3),DATE(2026,8,7))"), -4.0);
    }

    /// Business days: weekends skipped, holidays skipped, both ends counted.
    #[test]
    fn networkdays_and_workday_skip_weekends_and_holidays() {
        // 2026-08-03 is a Monday, 08-07 the Friday, 08-08/09 the weekend.
        assert_eq!(number("=NETWORKDAYS(DATE(2026,8,3),DATE(2026,8,7))"), 5.0);
        assert_eq!(number("=NETWORKDAYS(DATE(2026,8,3),DATE(2026,8,9))"), 5.0);
        assert_eq!(number("=NETWORKDAYS(DATE(2026,8,8),DATE(2026,8,9))"), 0.0);
        // Both ends are counted, so a single working day is 1, not 0.
        assert_eq!(number("=NETWORKDAYS(DATE(2026,8,3),DATE(2026,8,3))"), 1.0);
        // Excel returns a negative count for a reversed range rather than an error.
        assert_eq!(number("=NETWORKDAYS(DATE(2026,8,7),DATE(2026,8,3))"), -5.0);
        assert_eq!(number("=NETWORKDAYS(DATE(2026,8,3),DATE(2026,8,7),DATE(2026,8,5))"), 4.0);

        let aug10 = number("=DATE(2026,8,10)");
        assert_eq!(number("=WORKDAY(DATE(2026,8,3),5)"), aug10);
        assert_eq!(number("=WORKDAY(DATE(2026,8,7),1)"), aug10);
        // Zero days returns the start unchanged.
        assert_eq!(number("=WORKDAY(DATE(2026,8,3),0)"), number("=DATE(2026,8,3)"));
        assert_eq!(number("=WORKDAY(DATE(2026,8,10),-1)"), number("=DATE(2026,8,7)"));
        assert_eq!(number("=WORKDAY(DATE(2026,8,7),1,DATE(2026,8,10))"), number("=DATE(2026,8,11)"));
    }

    /// TIME is a fraction of a day, which is what makes it addable to a date.
    #[test]
    fn time_is_a_day_fraction_and_carries_overflow() {
        assert!((number("=TIME(18,30,0)") - 0.770_833_333_333_333_4).abs() < 1e-12);
        assert_eq!(number("=TIME(0,0,0)"), 0.0);
        // 90 minutes is an hour and a half, not an error.
        assert!((number("=TIME(0,90,0)") - 0.0625).abs() < 1e-12);
        // Past a full day it wraps, so 25:00 is 01:00.
        assert!((number("=TIME(25,0,0)") - 1.0 / 24.0).abs() < 1e-12);
    }
}
