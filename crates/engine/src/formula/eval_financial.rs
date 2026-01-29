// Financial functions: PMT, IPMT, PPMT, CUMPRINC, CUMIPMT, FV, PV, NPV, IRR

use super::eval::{evaluate, CellLookup, EvalResult};
use super::parser::{BoundExpr, Expr, col_to_letters};

/// Compute PMT (payment for a loan with constant payments and interest rate)
fn compute_pmt(rate: f64, nper: f64, pv: f64, fv: f64, pmt_type: f64) -> f64 {
    if rate == 0.0 {
        -(pv + fv) / nper
    } else {
        let pow = (1.0 + rate).powf(nper);
        let p = (rate * (pv * pow + fv)) / (pow - 1.0);
        if pmt_type != 0.0 {
            -p / (1.0 + rate)
        } else {
            -p
        }
    }
}

/// Compute IPMT (interest portion of a payment for a specific period)
/// Uses FV-based formula: IPMT = FV(rate, per-1, pmt, pv, type) * rate
fn compute_ipmt(rate: f64, per: f64, nper: f64, pv: f64, fv: f64, pmt_type: f64) -> f64 {
    if rate == 0.0 {
        return 0.0;
    }

    let pmt = compute_pmt(rate, nper, pv, fv, pmt_type);

    if pmt_type != 0.0 {
        // Beginning of period payments
        if per == 1.0 {
            0.0
        } else {
            // FV after (per-2) periods with type=1
            let k = per - 2.0;
            let pow_k = (1.0 + rate).powf(k);
            let fv_k = -pv * pow_k - pmt * (1.0 + rate) * (pow_k - 1.0) / rate;
            fv_k * rate
        }
    } else {
        // End of period payments
        // FV after (per-1) periods with type=0
        let k = per - 1.0;
        let pow_k = (1.0 + rate).powf(k);
        let fv_k = -pv * pow_k - pmt * (pow_k - 1.0) / rate;
        fv_k * rate
    }
}

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "PMT" => {
            // PMT(rate, nper, pv, [fv], [type])
            // Returns the payment for a loan based on constant payments and interest rate
            if args.len() < 3 || args.len() > 5 {
                return Some(EvalResult::Error("PMT requires 3 to 5 arguments".to_string()));
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let nper = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pv = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let fv = if args.len() >= 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };
            let pmt_type = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };

            if nper == 0.0 {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }

            let pmt = if rate == 0.0 {
                -(pv + fv) / nper
            } else {
                let pow = (1.0 + rate).powf(nper);
                let pmt = (rate * (pv * pow + fv)) / (pow - 1.0);
                if pmt_type != 0.0 {
                    -pmt / (1.0 + rate)
                } else {
                    -pmt
                }
            };
            EvalResult::Number(pmt)
        }
        "IPMT" => {
            // IPMT(rate, per, nper, pv, [fv], [type])
            // Returns the interest portion of a payment for a given period
            if args.len() < 4 || args.len() > 6 {
                return Some(EvalResult::Error("IPMT requires 4 to 6 arguments".to_string()));
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let per = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let nper = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pv = match evaluate(&args[3], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let fv = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };
            let pmt_type = if args.len() >= 6 {
                match evaluate(&args[5], lookup).to_number() {
                    Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };

            if nper == 0.0 || per < 1.0 || per > nper {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }

            let ipmt = compute_ipmt(rate, per, nper, pv, fv, pmt_type);
            EvalResult::Number(ipmt)
        }
        "PPMT" => {
            // PPMT(rate, per, nper, pv, [fv], [type])
            // Returns the principal portion of a payment for a given period
            if args.len() < 4 || args.len() > 6 {
                return Some(EvalResult::Error("PPMT requires 4 to 6 arguments".to_string()));
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let per = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let nper = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pv = match evaluate(&args[3], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let fv = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };
            let pmt_type = if args.len() >= 6 {
                match evaluate(&args[5], lookup).to_number() {
                    Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };

            if nper == 0.0 || per < 1.0 || per > nper {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }

            let pmt = compute_pmt(rate, nper, pv, fv, pmt_type);
            let ipmt = compute_ipmt(rate, per, nper, pv, fv, pmt_type);
            EvalResult::Number(pmt - ipmt)
        }
        "CUMPRINC" => {
            // CUMPRINC(rate, nper, pv, start_period, end_period, type)
            // Returns cumulative principal paid on a loan between two periods
            if args.len() != 6 {
                return Some(EvalResult::Error("CUMPRINC requires 6 arguments".to_string()));
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let nper = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pv = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let start_period = match evaluate(&args[3], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let end_period = match evaluate(&args[4], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pmt_type = match evaluate(&args[5], lookup).to_number() {
                Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                Err(e) => return Some(EvalResult::Error(e)),
            };

            if rate <= 0.0 || nper <= 0.0 || pv <= 0.0 {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }
            let sp = start_period as i64;
            let ep = end_period as i64;
            if sp < 1 || ep < sp || ep > nper as i64 {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }

            let pmt = compute_pmt(rate, nper, pv, 0.0, pmt_type);
            let mut cum_princ = 0.0;
            for per in sp..=ep {
                let ipmt = compute_ipmt(rate, per as f64, nper, pv, 0.0, pmt_type);
                cum_princ += pmt - ipmt;
            }
            EvalResult::Number(cum_princ)
        }
        "CUMIPMT" => {
            // CUMIPMT(rate, nper, pv, start_period, end_period, type)
            // Returns cumulative interest paid on a loan between two periods
            if args.len() != 6 {
                return Some(EvalResult::Error("CUMIPMT requires 6 arguments".to_string()));
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let nper = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pv = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let start_period = match evaluate(&args[3], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let end_period = match evaluate(&args[4], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pmt_type = match evaluate(&args[5], lookup).to_number() {
                Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                Err(e) => return Some(EvalResult::Error(e)),
            };

            if rate <= 0.0 || nper <= 0.0 || pv <= 0.0 {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }
            let sp = start_period as i64;
            let ep = end_period as i64;
            if sp < 1 || ep < sp || ep > nper as i64 {
                return Some(EvalResult::Error("#NUM!".to_string()));
            }

            let mut cum_ipmt = 0.0;
            for per in sp..=ep {
                cum_ipmt += compute_ipmt(rate, per as f64, nper, pv, 0.0, pmt_type);
            }
            EvalResult::Number(cum_ipmt)
        }
        "FV" => {
            // FV(rate, nper, pmt, [pv], [type])
            // Returns the future value of an investment
            if args.len() < 3 || args.len() > 5 {
                return Some(EvalResult::Error("FV requires 3 to 5 arguments".to_string()));
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let nper = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pmt = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pv = if args.len() >= 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };
            let pmt_type = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };

            let fv = if rate == 0.0 {
                -pv - pmt * nper
            } else {
                let pow = (1.0 + rate).powf(nper);
                let fv_pmt = if pmt_type != 0.0 {
                    pmt * (1.0 + rate) * (pow - 1.0) / rate
                } else {
                    pmt * (pow - 1.0) / rate
                };
                -pv * pow - fv_pmt
            };
            EvalResult::Number(fv)
        }
        "PV" => {
            // PV(rate, nper, pmt, [fv], [type])
            // Returns the present value of an investment
            if args.len() < 3 || args.len() > 5 {
                return Some(EvalResult::Error("PV requires 3 to 5 arguments".to_string()));
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let nper = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pmt = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let fv = if args.len() >= 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };
            let pmt_type = if args.len() >= 5 {
                match evaluate(&args[4], lookup).to_number() {
                    Ok(n) => if n != 0.0 { 1.0 } else { 0.0 },
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.0
            };

            let pv = if rate == 0.0 {
                -fv - pmt * nper
            } else {
                let pow = (1.0 + rate).powf(nper);
                let pv_pmt = if pmt_type != 0.0 {
                    pmt * (1.0 + rate) * (pow - 1.0) / rate
                } else {
                    pmt * (pow - 1.0) / rate
                };
                (-fv - pv_pmt) / pow
            };
            EvalResult::Number(pv)
        }
        "NPV" => {
            // NPV(rate, value1, [value2], ...)
            // Returns the net present value of an investment based on periodic cash flows
            if args.len() < 2 {
                return Some(EvalResult::Error("NPV requires at least 2 arguments".to_string()));
            }
            let rate = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };

            if rate == -1.0 {
                return Some(EvalResult::Error("#DIV/0!".to_string()));
            }

            let mut npv = 0.0;
            let mut period = 1;

            for arg in &args[1..] {
                // Handle both single values and ranges
                match arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                        let (min_row, min_col) = (*start_row.min(end_row), *start_col.min(end_col));
                        let (max_row, max_col) = (*start_row.max(end_row), *start_col.max(end_col));
                        for r in min_row..=max_row {
                            for c in min_col..=max_col {
                                let text = lookup.get_text(r, c);
                                if text.is_empty() {
                                    continue; // Skip blanks
                                }
                                if text.starts_with('#') {
                                    return Some(EvalResult::Error(text)); // Propagate errors
                                }
                                let val = lookup.get_value(r, c);
                                if val.is_finite() {
                                    npv += val / (1.0 + rate).powi(period);
                                    period += 1;
                                }
                            }
                        }
                    }
                    _ => {
                        match evaluate(arg, lookup).to_number() {
                            Ok(n) => {
                                npv += n / (1.0 + rate).powi(period);
                                period += 1;
                            }
                            Err(e) => return Some(EvalResult::Error(e)),
                        }
                    }
                }
            }
            EvalResult::Number(npv)
        }
        "IRR" => {
            // IRR(values, [guess])
            // Returns the internal rate of return for a series of cash flows
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("IRR requires 1 or 2 arguments".to_string()));
            }

            // Collect cash flows from range
            let values: Vec<f64> = match &args[0] {
                Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                    let (min_row, min_col) = (*start_row.min(end_row), *start_col.min(end_col));
                    let (max_row, max_col) = (*start_row.max(end_row), *start_col.max(end_col));
                    let mut vals = Vec::new();
                    eprintln!("[IRR] context: {}", lookup.debug_context());
                    eprintln!("[IRR] range {}{}:{}{} (0-indexed r{}:r{}, c{}:c{})",
                        col_to_letters(min_col), min_row + 1,
                        col_to_letters(max_col), max_row + 1,
                        min_row, max_row, min_col, max_col);
                    for r in min_row..=max_row {
                        for c in min_col..=max_col {
                            let text = lookup.get_text(r, c);
                            let val = lookup.get_value(r, c);
                            eprintln!("[IRR]   {}{}: get_value={}, get_text=\"{}\"",
                                col_to_letters(c), r + 1, val, text);
                            if text.is_empty() {
                                continue; // Skip blanks (Excel ignores blanks in IRR)
                            }
                            if text.starts_with('#') {
                                eprintln!("[IRR]   → error propagated: {}", text);
                                return Some(EvalResult::Error(text));
                            }
                            if val.is_finite() {
                                vals.push(val);
                            }
                        }
                    }
                    vals
                }
                _ => return Some(EvalResult::Error("IRR requires a range of values".to_string())),
            };

            if values.len() < 2 {
                eprintln!("[IRR] #NUM!: only {} cashflow(s) collected: {:?}", values.len(), values);
                return Some(EvalResult::Error("#NUM!".to_string()));
            }

            // Check that there's at least one positive and one negative value
            let has_positive = values.iter().any(|&v| v > 0.0);
            let has_negative = values.iter().any(|&v| v < 0.0);
            if !has_positive || !has_negative {
                eprintln!("[IRR] #NUM!: need both signs. cashflows: {:?}, has_pos={}, has_neg={}", values, has_positive, has_negative);
                return Some(EvalResult::Error("#NUM!".to_string()));
            }

            let guess = if args.len() >= 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0.1 // Default guess of 10%
            };

            // NPV at a given rate: sum(CF_i / (1+rate)^i)
            let npv_at = |rate: f64| -> f64 {
                values.iter().enumerate().map(|(i, &cf)| {
                    cf / (1.0 + rate).powf(i as f64)
                }).sum()
            };

            // Phase 1: Newton-Raphson (fast quadratic convergence when it works)
            let mut rate = guess;
            let mut newton_converged = false;

            for _ in 0..100 {
                let mut npv = 0.0;
                let mut dnpv = 0.0;

                for (i, &cf) in values.iter().enumerate() {
                    let t = i as f64;
                    let base = 1.0 + rate;
                    let divisor = base.powf(t);
                    if divisor.abs() < 1e-30 { break; }
                    npv += cf / divisor;
                    if t > 0.0 {
                        dnpv -= t * cf / base.powf(t + 1.0);
                    }
                }

                if dnpv.abs() < 1e-30 { break; }

                let new_rate = rate - npv / dnpv;

                if (new_rate - rate).abs() < 1e-10 {
                    if new_rate > -1.0 && new_rate.is_finite() {
                        rate = new_rate;
                        newton_converged = true;
                    }
                    break;
                }

                rate = new_rate;
                if rate <= -1.0 || rate > 10.0 || !rate.is_finite() {
                    break;
                }
            }

            if newton_converged {
                return Some(EvalResult::Number(rate));
            }

            // Phase 2: Bisection fallback — guaranteed convergence when bracket exists
            let search_rates: &[f64] = &[
                -0.99, -0.95, -0.9, -0.8, -0.5, -0.3, -0.1,
                0.0, 0.1, 0.2, 0.3, 0.5, 0.8, 1.0, 2.0, 5.0, 10.0,
            ];

            let mut lo = f64::NAN;
            let mut hi = f64::NAN;
            let mut npv_lo_val = 0.0_f64;

            let mut prev_rate = f64::NAN;
            let mut prev_npv = f64::NAN;

            for &r in search_rates {
                let npv = npv_at(r);
                if !npv.is_finite() {
                    prev_rate = f64::NAN;
                    prev_npv = f64::NAN;
                    continue;
                }

                if prev_npv.is_finite() && npv.signum() != prev_npv.signum() && prev_npv != 0.0 {
                    lo = prev_rate;
                    hi = r;
                    npv_lo_val = prev_npv;
                    break;
                }

                prev_rate = r;
                prev_npv = npv;
            }

            if lo.is_nan() {
                eprintln!("[IRR] #NUM!: no bracket found. cashflows: {:?}", values);
                eprintln!("[IRR]   Newton guess={}, rate after Newton={}", guess, rate);
                for &r in search_rates {
                    let npv = npv_at(r);
                    eprintln!("[IRR]   rate={:.4} → NPV={:.4}", r, npv);
                }
                return Some(EvalResult::Error("#NUM!".to_string()));
            }

            for _ in 0..200 {
                let mid = (lo + hi) / 2.0;
                let npv_mid = npv_at(mid);

                if !npv_mid.is_finite() {
                    return Some(EvalResult::Error("#NUM!".to_string()));
                }

                if npv_mid.abs() < 1e-10 || (hi - lo) < 1e-12 {
                    return Some(EvalResult::Number(mid));
                }

                if npv_mid.signum() == npv_lo_val.signum() {
                    lo = mid;
                } else {
                    hi = mid;
                }
            }

            // After 200 bisection iterations, return best midpoint
            EvalResult::Number((lo + hi) / 2.0)
        }
        _ => return None,
    };
    Some(result)
}
