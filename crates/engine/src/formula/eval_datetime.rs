// Date/time functions: TODAY, NOW, DATE, DATEVALUE, YEAR, MONTH, DAY, WEEKDAY, DATEDIF,
// EDATE, EOMONTH, HOUR, MINUTE, SECOND, WEEKNUM, ISOWEEKNUM, YEARFRAC

use super::eval::{evaluate, CellLookup, EvalResult};
use super::eval_helpers::{date_to_serial, serial_to_date, days_in_month, excel_error, try_parse_date_string};
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
        "WEEKNUM" => {
            // WEEKNUM(date, [return_type]): the week of the year. Types 1 and
            // 17 start weeks on Sunday, 2 and 11 on Monday, 12-16 Tuesday to
            // Saturday; week 1 is the week holding January 1. Type 21 is the
            // ISO 8601 week (weeks start Monday; week 1 holds the first
            // Thursday), which can belong to the neighbouring year. Any other
            // type, or a negative date, is #NUM!.
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("WEEKNUM requires 1 or 2 arguments".to_string()));
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < 0.0 => return Some(EvalResult::Error("#NUM!".to_string())),
                Ok(n) => n.floor(),
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let return_type = if args.len() == 2 {
                match evaluate(&args[1], lookup) {
                    EvalResult::Empty => 1,
                    other => match other.to_number() {
                        Ok(n) => n.trunc() as i64,
                        Err(e) => return Some(EvalResult::Error(e)),
                    },
                }
            } else {
                1
            };
            let weekday0 = |serial: f64| (serial as i64 + 6).rem_euclid(7); // 0 = Sunday
            let (year, _, _) = serial_to_date(serial);
            let doy = (serial - date_to_serial(year, 1, 1)) as i64 + 1;
            if return_type == 21 {
                return Some(EvalResult::Number(iso_week(serial, year, doy, weekday0) as f64));
            }
            let start = match return_type {
                1 | 17 => 0,
                2 | 11 => 1,
                12..=16 => return_type - 10,
                _ => return Some(EvalResult::Error("#NUM!".to_string())),
            };
            let jan1_offset = (weekday0(date_to_serial(year, 1, 1)) - start).rem_euclid(7);
            EvalResult::Number(((doy - 1 + jan1_offset) / 7 + 1) as f64)
        }
        "ISOWEEKNUM" => {
            // ISOWEEKNUM(date): the ISO 8601 week, WEEKNUM(date, 21). A
            // negative date is #NUM!.
            if args.len() != 1 {
                return Some(EvalResult::Error("ISOWEEKNUM requires exactly one argument".to_string()));
            }
            let serial = match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < 0.0 => return Some(EvalResult::Error("#NUM!".to_string())),
                Ok(n) => n.floor(),
                Err(e) => return Some(EvalResult::Error(excel_error(e))),
            };
            let weekday0 = |serial: f64| (serial as i64 + 6).rem_euclid(7); // 0 = Sunday
            let (year, _, _) = serial_to_date(serial);
            let doy = (serial - date_to_serial(year, 1, 1)) as i64 + 1;
            EvalResult::Number(iso_week(serial, year, doy, weekday0) as f64)
        }
        "YEARFRAC" => {
            // YEARFRAC(start_date, end_date, [basis]): the years between two
            // dates, as a fraction, counted by a day-count basis: 0 or omitted
            // US (NASD) 30/360, 1 actual/actual, 2 actual/360, 3 actual/365,
            // 4 European 30/360. The order of the dates does not matter, and
            // times of day are ignored. A basis outside 0-4, or a negative
            // date, is #NUM!.
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("YEARFRAC requires 2 or 3 arguments".to_string()));
            }
            let mut dates = [0.0; 2];
            for (i, date) in dates.iter_mut().enumerate() {
                *date = match evaluate(&args[i], lookup).to_number() {
                    Ok(n) if n < 0.0 => return Some(EvalResult::Error("#NUM!".to_string())),
                    Ok(n) => n.trunc(),
                    Err(e) => return Some(EvalResult::Error(excel_error(e))),
                };
            }
            let basis = match args.get(2) {
                None => 0.0,
                Some(a) => match evaluate(a, lookup).to_number() {
                    Ok(n) => n.trunc(),
                    Err(e) => return Some(EvalResult::Error(excel_error(e))),
                },
            };
            if !(0.0..=4.0).contains(&basis) {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }
            let (start, end) = (dates[0].min(dates[1]), dates[0].max(dates[1]));
            EvalResult::Number(year_frac(start, end, basis as u8))
        }
        _ => return None,
    };
    Some(result)
}

fn is_leap(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

/// YEARFRAC from `start` to `end` (serials, start <= end) on a day-count
/// basis. These follow Excel's results rather than the textbook rules, which
/// differ at month ends and over leap days; the method is David A. Wheeler's
/// reconstruction, which matches Excel ("YEARFRAC incompatibilities between
/// spreadsheet programs", 2008).
fn year_frac(start: f64, end: f64, basis: u8) -> f64 {
    let (y1, m1, mut d1) = serial_to_date(start);
    let (y2, m2, mut d2) = serial_to_date(end);
    let days = end - start;
    let thirty_360 = |d1: i32, d2: i32| {
        ((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64 / 360.0
    };
    match basis {
        0 => {
            // US (NASD) 30/360, with Excel's handling of the end of February.
            let last_of_feb = |m: i32, d: i32, y: i32| m == 2 && d == days_in_month(y, 2);
            let start_feb = last_of_feb(m1, d1, y1);
            if start_feb && last_of_feb(m2, d2, y2) {
                d2 = 30;
            }
            if start_feb {
                d1 = 30;
            }
            if d2 == 31 && d1 >= 30 {
                d2 = 30;
            }
            if d1 == 31 {
                d1 = 30;
            }
            thirty_360(d1, d2)
        }
        1 => {
            // Actual/actual. Within a year the divisor is that year's length
            // (366 when the span holds a February 29); across years it is the
            // average length of every year the span touches.
            if days == 0.0 {
                return 0.0;
            }
            let within_a_year = y1 == y2 || (y2 == y1 + 1 && (m1 > m2 || (m1 == m2 && d1 >= d2)));
            if within_a_year {
                let march1 = |y: i32| date_to_serial(y, 3, 1);
                let holds_feb29 = (is_leap(y1) && start < march1(y1) && end >= march1(y1))
                    || (is_leap(y2) && end >= march1(y2) && start < march1(y2));
                let length = if (y1 == y2 && is_leap(y1)) || holds_feb29 || (m2 == 2 && d2 == 29) { 366.0 } else { 365.0 };
                days / length
            } else {
                let years = (y2 - y1 + 1) as f64;
                let span = date_to_serial(y2 + 1, 1, 1) - date_to_serial(y1, 1, 1);
                days / (span / years)
            }
        }
        2 => days / 360.0,
        3 => days / 365.0,
        _ => {
            // European 30/360: the 31st is the 30th, at either end.
            thirty_360(d1.min(30), d2.min(30))
        }
    }
}

/// ISO 8601 week number for a serial whose Gregorian year is `year` and day of
/// year is `doy` (1-based).
fn iso_week(serial: f64, year: i32, doy: i64, weekday0: impl Fn(f64) -> i64) -> i64 {
    let iso_weekday = (weekday0(serial) + 6) % 7 + 1; // Monday 1 .. Sunday 7
    let week = (doy - iso_weekday + 10) / 7;
    let weeks_in = |y: i32| -> i64 {
        // A year has 53 ISO weeks when January 1 is a Thursday, or a
        // Wednesday in a leap year.
        let jan1 = (weekday0(date_to_serial(y, 1, 1)) + 6) % 7 + 1;
        let leap = (y % 4 == 0 && y % 100 != 0) || y % 400 == 0;
        if jan1 == 4 || (leap && jan1 == 3) { 53 } else { 52 }
    };
    if week < 1 {
        weeks_in(year - 1)
    } else if week > weeks_in(year) {
        1
    } else {
        week
    }
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

