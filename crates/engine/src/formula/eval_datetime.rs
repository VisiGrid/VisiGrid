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
            // Return Excel-style date serial number (days since 1899-12-30)
            use std::time::{SystemTime, UNIX_EPOCH};
            let now = SystemTime::now().duration_since(UNIX_EPOCH).unwrap();
            let days_since_unix = now.as_secs() / 86400;
            // Excel epoch is 1899-12-30, Unix epoch is 1970-01-01
            // Difference is 25569 days
            let excel_date = days_since_unix as f64 + 25569.0;
            EvalResult::Number(excel_date)
        }
        "NOW" => {
            if !args.is_empty() {
                return Some(EvalResult::Error("NOW takes no arguments".to_string()));
            }
            use std::time::{SystemTime, UNIX_EPOCH};
            let now = SystemTime::now().duration_since(UNIX_EPOCH).unwrap();
            let secs = now.as_secs() as f64 + now.subsec_nanos() as f64 / 1_000_000_000.0;
            let days_since_unix = secs / 86400.0;
            let excel_datetime = days_since_unix + 25569.0;
            EvalResult::Number(excel_datetime)
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
