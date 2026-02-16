//! `vgrid verify` — financial verification commands.
//!
//! Currently supports one subcommand:
//! - `vgrid verify totals` — compare truth vs warehouse daily totals

use std::io::Write;
use std::path::PathBuf;

use clap::Subcommand;
use serde::Serialize;

use visigrid_io::truth::{
    hash_raw_row, parse_amount, read_daily_totals_csv, DailyTotals,
};

use crate::exit_codes;
use crate::CliError;

#[derive(Subcommand)]
pub enum VerifyCommands {
    /// Compare truth vs warehouse daily totals (exit 0 = match, exit 1 = mismatch)
    #[command(after_help = "\
Both files must use the truth_daily_totals CSV format:
  date,currency,source_account,total_gross,total_fee,total_net,transaction_count

Tolerance is in currency units (not micro-units). Default: 0 (exact match).

Exit codes:
  0   All rows match within tolerance
  1   Mismatches found

Examples:
  vgrid verify totals truth_daily_totals.csv warehouse_daily_totals.csv
  vgrid verify totals truth.csv warehouse.csv --tolerance 0.01
  vgrid verify totals truth.csv warehouse.csv --output verify.json
  vgrid verify totals truth.csv warehouse.csv --diff diffs.csv")]
    Totals {
        /// Truth daily totals CSV (the external source of truth)
        truth: PathBuf,

        /// Warehouse daily totals CSV (what to verify against truth)
        warehouse: PathBuf,

        /// Tolerance in currency units (e.g. 0.01 for one cent)
        #[arg(long, default_value = "0")]
        tolerance: f64,

        /// Allow count mismatches without failing (default: fail on count mismatch)
        #[arg(long)]
        no_fail_on_count: bool,

        /// Output verification result JSON to file
        #[arg(long)]
        output: Option<PathBuf>,

        /// Output diff CSV to file (rows that mismatch)
        #[arg(long)]
        diff: Option<PathBuf>,

        /// Quiet mode: only exit code, no stderr output
        #[arg(long, short = 'q')]
        quiet: bool,
    },
}

pub fn cmd_verify(cmd: VerifyCommands) -> Result<(), CliError> {
    match cmd {
        VerifyCommands::Totals {
            truth,
            warehouse,
            tolerance,
            no_fail_on_count,
            output,
            diff,
            quiet,
        } => cmd_verify_totals(
            truth,
            warehouse,
            tolerance,
            no_fail_on_count,
            output,
            diff,
            quiet,
        ),
    }
}

// ── Verification result types ───────────────────────────────────────

#[derive(Debug, Serialize)]
struct VerifyResult {
    status: &'static str, // "pass" or "fail"
    truth_file: String,
    warehouse_file: String,
    truth_hash: String,
    warehouse_hash: String,
    tolerance_micro: i64,
    fail_on_count_mismatch: bool,
    summary: VerifySummary,
    mismatches: Vec<Mismatch>,
}

#[derive(Debug, Serialize)]
struct VerifySummary {
    truth_rows: usize,
    warehouse_rows: usize,
    matched: usize,
    mismatched: usize,
    only_in_truth: usize,
    only_in_warehouse: usize,
}

#[derive(Debug, Serialize)]
struct Mismatch {
    date: String,
    currency: String,
    source_account: String,
    kind: MismatchKind,
    truth_value: Option<String>,
    warehouse_value: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "snake_case")]
enum MismatchKind {
    NetDifference,
    GrossDifference,
    FeeDifference,
    CountDifference,
    OnlyInTruth,
    OnlyInWarehouse,
}

impl std::fmt::Display for MismatchKind {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NetDifference => write!(f, "net_difference"),
            Self::GrossDifference => write!(f, "gross_difference"),
            Self::FeeDifference => write!(f, "fee_difference"),
            Self::CountDifference => write!(f, "count_difference"),
            Self::OnlyInTruth => write!(f, "only_in_truth"),
            Self::OnlyInWarehouse => write!(f, "only_in_warehouse"),
        }
    }
}

// ── Core verification logic ─────────────────────────────────────────

fn cmd_verify_totals(
    truth_path: PathBuf,
    warehouse_path: PathBuf,
    tolerance: f64,
    no_fail_on_count: bool,
    output_path: Option<PathBuf>,
    diff_path: Option<PathBuf>,
    quiet: bool,
) -> Result<(), CliError> {
    // Convert tolerance from currency units to micro-units
    let tolerance_micro = (tolerance * 1_000_000.0).round() as i64;

    // Load both files
    let truth_bytes = std::fs::read(&truth_path).map_err(|e| {
        CliError::io(format!("cannot read {}: {e}", truth_path.display()))
    })?;
    let warehouse_bytes = std::fs::read(&warehouse_path).map_err(|e| {
        CliError::io(format!("cannot read {}: {e}", warehouse_path.display()))
    })?;

    // Hash raw file bytes for proof chain
    let truth_hash = hash_raw_row(&truth_bytes);
    let warehouse_hash = hash_raw_row(&warehouse_bytes);

    // Parse CSVs
    let truth = read_daily_totals_csv(truth_bytes.as_slice()).map_err(|e| {
        CliError::parse(format!("truth file: {e}"))
    })?;
    let warehouse = read_daily_totals_csv(warehouse_bytes.as_slice()).map_err(|e| {
        CliError::parse(format!("warehouse file: {e}"))
    })?;

    // Build lookup by (date, currency, source_account)
    type Key = (String, String, String);
    fn make_key(row: &DailyTotals) -> Key {
        (
            row.date.format("%Y-%m-%d").to_string(),
            row.currency.clone(),
            row.source_account.clone(),
        )
    }

    let truth_map: std::collections::HashMap<Key, &DailyTotals> =
        truth.iter().map(|r| (make_key(r), r)).collect();
    let warehouse_map: std::collections::HashMap<Key, &DailyTotals> =
        warehouse.iter().map(|r| (make_key(r), r)).collect();

    let mut mismatches = Vec::new();
    let mut matched = 0usize;

    // Check each truth row against warehouse
    for (key, t) in &truth_map {
        match warehouse_map.get(key) {
            None => {
                mismatches.push(Mismatch {
                    date: key.0.clone(),
                    currency: key.1.clone(),
                    source_account: key.2.clone(),
                    kind: MismatchKind::OnlyInTruth,
                    truth_value: Some(format_micro(t.total_net)),
                    warehouse_value: None,
                });
            }
            Some(w) => {
                let mut row_ok = true;

                // Net comparison
                if (t.total_net - w.total_net).abs() > tolerance_micro {
                    mismatches.push(Mismatch {
                        date: key.0.clone(),
                        currency: key.1.clone(),
                        source_account: key.2.clone(),
                        kind: MismatchKind::NetDifference,
                        truth_value: Some(format_micro(t.total_net)),
                        warehouse_value: Some(format_micro(w.total_net)),
                    });
                    row_ok = false;
                }

                // Gross comparison
                if (t.total_gross - w.total_gross).abs() > tolerance_micro {
                    mismatches.push(Mismatch {
                        date: key.0.clone(),
                        currency: key.1.clone(),
                        source_account: key.2.clone(),
                        kind: MismatchKind::GrossDifference,
                        truth_value: Some(format_micro(t.total_gross)),
                        warehouse_value: Some(format_micro(w.total_gross)),
                    });
                    row_ok = false;
                }

                // Fee comparison
                if (t.total_fee - w.total_fee).abs() > tolerance_micro {
                    mismatches.push(Mismatch {
                        date: key.0.clone(),
                        currency: key.1.clone(),
                        source_account: key.2.clone(),
                        kind: MismatchKind::FeeDifference,
                        truth_value: Some(format_micro(t.total_fee)),
                        warehouse_value: Some(format_micro(w.total_fee)),
                    });
                    row_ok = false;
                }

                // Count comparison
                if t.transaction_count != w.transaction_count {
                    mismatches.push(Mismatch {
                        date: key.0.clone(),
                        currency: key.1.clone(),
                        source_account: key.2.clone(),
                        kind: MismatchKind::CountDifference,
                        truth_value: Some(t.transaction_count.to_string()),
                        warehouse_value: Some(w.transaction_count.to_string()),
                    });
                    row_ok = false;
                }

                if row_ok {
                    matched += 1;
                }
            }
        }
    }

    // Check for rows only in warehouse
    for key in warehouse_map.keys() {
        if !truth_map.contains_key(key) {
            let w = warehouse_map[key];
            mismatches.push(Mismatch {
                date: key.0.clone(),
                currency: key.1.clone(),
                source_account: key.2.clone(),
                kind: MismatchKind::OnlyInWarehouse,
                truth_value: None,
                warehouse_value: Some(format_micro(w.total_net)),
            });
        }
    }

    // Count categories
    let only_in_truth = mismatches
        .iter()
        .filter(|m| matches!(m.kind, MismatchKind::OnlyInTruth))
        .count();
    let only_in_warehouse = mismatches
        .iter()
        .filter(|m| matches!(m.kind, MismatchKind::OnlyInWarehouse))
        .count();

    // Determine pass/fail
    let has_material_mismatch = mismatches.iter().any(|m| match m.kind {
        MismatchKind::CountDifference => !no_fail_on_count,
        _ => true,
    });

    let status = if has_material_mismatch { "fail" } else { "pass" };

    let summary = VerifySummary {
        truth_rows: truth.len(),
        warehouse_rows: warehouse.len(),
        matched,
        mismatched: mismatches.len(),
        only_in_truth,
        only_in_warehouse,
    };

    let result = VerifyResult {
        status,
        truth_file: truth_path.display().to_string(),
        warehouse_file: warehouse_path.display().to_string(),
        truth_hash,
        warehouse_hash,
        tolerance_micro,
        fail_on_count_mismatch: !no_fail_on_count,
        summary,
        mismatches,
    };

    // Output to stderr
    if !quiet {
        eprintln!("verify: {} ({} rows truth, {} rows warehouse)",
            status.to_uppercase(),
            result.summary.truth_rows,
            result.summary.warehouse_rows,
        );
        eprintln!("  matched:            {}", result.summary.matched);
        if result.summary.mismatched > 0 {
            eprintln!("  mismatched:         {}", result.summary.mismatched);
        }
        if result.summary.only_in_truth > 0 {
            eprintln!("  only in truth:      {}", result.summary.only_in_truth);
        }
        if result.summary.only_in_warehouse > 0 {
            eprintln!("  only in warehouse:  {}", result.summary.only_in_warehouse);
        }
        eprintln!("  truth hash:         {}", &result.truth_hash[..16]);
        eprintln!("  warehouse hash:     {}", &result.warehouse_hash[..16]);
    }

    // Write JSON output
    if let Some(path) = &output_path {
        let json = serde_json::to_string_pretty(&result).map_err(|e| {
            CliError::io(format!("JSON serialization error: {e}"))
        })?;
        std::fs::write(path, json).map_err(|e| {
            CliError::io(format!("cannot write {}: {e}", path.display()))
        })?;
        if !quiet {
            eprintln!("  result written to:  {}", path.display());
        }
    }

    // Write diff CSV
    if let Some(path) = &diff_path {
        write_diff_csv(&result.mismatches, path)?;
        if !quiet {
            eprintln!("  diff written to:    {}", path.display());
        }
    }

    if has_material_mismatch {
        Err(CliError {
            code: exit_codes::EXIT_ERROR,
            message: String::new(), // already printed above
            hint: None,
        })
    } else {
        Ok(())
    }
}

fn format_micro(micro_units: i64) -> String {
    let is_negative = micro_units < 0;
    let abs = micro_units.unsigned_abs();
    let whole = abs / 1_000_000;
    let frac = abs % 1_000_000;
    if is_negative {
        format!("-{whole}.{frac:06}")
    } else {
        format!("{whole}.{frac:06}")
    }
}

fn write_diff_csv(mismatches: &[Mismatch], path: &PathBuf) -> Result<(), CliError> {
    let file = std::fs::File::create(path).map_err(|e| {
        CliError::io(format!("cannot create {}: {e}", path.display()))
    })?;
    let mut writer = std::io::BufWriter::new(file);

    writeln!(writer, "date,currency,source_account,kind,truth_value,warehouse_value")
        .map_err(|e| CliError::io(format!("write error: {e}")))?;

    for m in mismatches {
        writeln!(
            writer,
            "{},{},{},{},{},{}",
            m.date,
            m.currency,
            m.source_account,
            m.kind,
            m.truth_value.as_deref().unwrap_or(""),
            m.warehouse_value.as_deref().unwrap_or(""),
        )
        .map_err(|e| CliError::io(format!("write error: {e}")))?;
    }

    writer.flush().map_err(|e| CliError::io(format!("flush error: {e}")))?;
    Ok(())
}
