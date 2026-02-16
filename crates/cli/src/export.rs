//! `vgrid export` — export canonical truth data in various formats.
//!
//! Currently supports:
//! - `vgrid export truth --format dbt-seed` — write dbt-compatible seed CSVs + manifest

use std::path::PathBuf;

use clap::Subcommand;
use serde::Serialize;

use visigrid_io::truth::{
    compute_daily_totals, hash_daily_totals, hash_raw_row, read_daily_totals_csv,
    write_daily_totals_csv, write_transactions_csv, DailyTotals, TruthTransaction,
};

use crate::exit_codes;
use crate::CliError;

#[derive(Subcommand)]
pub enum ExportCommands {
    /// Export canonical truth data as dbt-seed CSVs + manifest
    #[command(after_help = "\
Reads a truth_transactions.csv and produces:
  seeds/truth_transactions.csv     (deterministic, sorted)
  seeds/truth_daily_totals.csv     (aggregated by date/currency/account)
  seeds/truth_manifest.json        (hashes, metadata, schema version)

Or, if --daily-totals is provided directly, skips transaction processing
and exports daily totals + manifest only.

Examples:
  vgrid export truth --transactions data/stripe_truth.csv
  vgrid export truth --transactions data/stripe_truth.csv --out seeds/
  vgrid export truth --daily-totals data/daily_totals.csv --out seeds/")]
    Truth {
        /// Input truth_transactions.csv
        #[arg(long)]
        transactions: Option<PathBuf>,

        /// Input truth_daily_totals.csv (skip transaction aggregation)
        #[arg(long, conflicts_with = "transactions")]
        daily_totals: Option<PathBuf>,

        /// Output directory (default: seeds/)
        #[arg(long, default_value = "seeds")]
        out: PathBuf,

        /// Quiet mode
        #[arg(long, short = 'q')]
        quiet: bool,
    },
}

pub fn cmd_export(cmd: ExportCommands) -> Result<(), CliError> {
    match cmd {
        ExportCommands::Truth {
            transactions,
            daily_totals,
            out,
            quiet,
        } => cmd_export_truth(transactions, daily_totals, out, quiet),
    }
}

// ── Manifest ────────────────────────────────────────────────────────

#[derive(Debug, Serialize)]
struct TruthManifest {
    schema_version: &'static str,
    transactions_hash: Option<String>,
    daily_totals_hash: String,
    source_account: Option<String>,
    date_range: Option<DateRange>,
    transaction_count: Option<usize>,
    daily_totals_rows: usize,
    mapping_profile_hash: Option<String>,
}

#[derive(Debug, Serialize)]
struct DateRange {
    min: String,
    max: String,
}

// ── Implementation ──────────────────────────────────────────────────

fn cmd_export_truth(
    transactions_path: Option<PathBuf>,
    daily_totals_path: Option<PathBuf>,
    out_dir: PathBuf,
    quiet: bool,
) -> Result<(), CliError> {
    if transactions_path.is_none() && daily_totals_path.is_none() {
        return Err(CliError::args(
            "provide either --transactions or --daily-totals",
        ));
    }

    // Create output directory
    std::fs::create_dir_all(&out_dir).map_err(|e| {
        CliError::io(format!("cannot create {}: {e}", out_dir.display()))
    })?;

    if let Some(dt_path) = daily_totals_path {
        // Direct daily totals pass-through
        return export_daily_totals_only(dt_path, out_dir, quiet);
    }

    let tx_path = transactions_path.unwrap();

    // Read raw transaction bytes for hashing
    let tx_bytes = std::fs::read(&tx_path).map_err(|e| {
        CliError::io(format!("cannot read {}: {e}", tx_path.display()))
    })?;

    // Parse transactions CSV
    let transactions = read_transactions_csv(&tx_bytes)?;

    if transactions.is_empty() {
        return Err(CliError::parse("no transactions found in input file"));
    }

    // Compute daily totals
    let totals = compute_daily_totals(&transactions).map_err(|e| {
        CliError {
            code: exit_codes::EXIT_ERROR,
            message: e,
            hint: None,
        }
    })?;

    // Write transactions CSV
    let tx_out = out_dir.join("truth_transactions.csv");
    let tx_file = std::fs::File::create(&tx_out).map_err(|e| {
        CliError::io(format!("cannot create {}: {e}", tx_out.display()))
    })?;
    write_transactions_csv(&transactions, std::io::BufWriter::new(tx_file)).map_err(|e| {
        CliError::io(format!("write error: {e}"))
    })?;

    // Write daily totals CSV
    let dt_out = out_dir.join("truth_daily_totals.csv");
    let dt_file = std::fs::File::create(&dt_out).map_err(|e| {
        CliError::io(format!("cannot create {}: {e}", dt_out.display()))
    })?;
    write_daily_totals_csv(&totals, std::io::BufWriter::new(dt_file)).map_err(|e| {
        CliError::io(format!("write error: {e}"))
    })?;

    // Compute hashes
    let tx_hash = hash_raw_row(&tx_bytes);
    let dt_hash = hash_daily_totals(&totals).map_err(|e| CliError::io(e))?;

    // Extract metadata
    let source_account = transactions.first().map(|t| t.source_account.clone());
    let date_range = if !totals.is_empty() {
        Some(DateRange {
            min: totals.first().unwrap().date.format("%Y-%m-%d").to_string(),
            max: totals.last().unwrap().date.format("%Y-%m-%d").to_string(),
        })
    } else {
        None
    };

    // Write manifest
    let manifest = TruthManifest {
        schema_version: "1.0",
        transactions_hash: Some(tx_hash),
        daily_totals_hash: dt_hash,
        source_account,
        date_range,
        transaction_count: Some(transactions.len()),
        daily_totals_rows: totals.len(),
        mapping_profile_hash: None, // placeholder for future
    };

    let manifest_out = out_dir.join("truth_manifest.json");
    let manifest_json = serde_json::to_string_pretty(&manifest).map_err(|e| {
        CliError::io(format!("JSON error: {e}"))
    })?;
    std::fs::write(&manifest_out, &manifest_json).map_err(|e| {
        CliError::io(format!("cannot write {}: {e}", manifest_out.display()))
    })?;

    if !quiet {
        eprintln!("export: wrote {} transactions, {} daily totals", transactions.len(), totals.len());
        eprintln!("  {}", tx_out.display());
        eprintln!("  {}", dt_out.display());
        eprintln!("  {}", manifest_out.display());
    }

    Ok(())
}

fn export_daily_totals_only(
    dt_path: PathBuf,
    out_dir: PathBuf,
    quiet: bool,
) -> Result<(), CliError> {
    let dt_bytes = std::fs::read(&dt_path).map_err(|e| {
        CliError::io(format!("cannot read {}: {e}", dt_path.display()))
    })?;

    let totals = read_daily_totals_csv(dt_bytes.as_slice()).map_err(|e| {
        CliError::parse(format!("daily totals file: {e}"))
    })?;

    // Re-write deterministically
    let dt_out = out_dir.join("truth_daily_totals.csv");
    let dt_file = std::fs::File::create(&dt_out).map_err(|e| {
        CliError::io(format!("cannot create {}: {e}", dt_out.display()))
    })?;
    write_daily_totals_csv(&totals, std::io::BufWriter::new(dt_file)).map_err(|e| {
        CliError::io(format!("write error: {e}"))
    })?;

    let dt_hash = hash_daily_totals(&totals).map_err(|e| CliError::io(e))?;

    let source_account = totals.first().map(|t| t.source_account.clone());
    let date_range = if !totals.is_empty() {
        Some(DateRange {
            min: totals.first().unwrap().date.format("%Y-%m-%d").to_string(),
            max: totals.last().unwrap().date.format("%Y-%m-%d").to_string(),
        })
    } else {
        None
    };

    let manifest = TruthManifest {
        schema_version: "1.0",
        transactions_hash: None,
        daily_totals_hash: dt_hash,
        source_account,
        date_range,
        transaction_count: None,
        daily_totals_rows: totals.len(),
        mapping_profile_hash: None,
    };

    let manifest_out = out_dir.join("truth_manifest.json");
    let manifest_json = serde_json::to_string_pretty(&manifest).map_err(|e| {
        CliError::io(format!("JSON error: {e}"))
    })?;
    std::fs::write(&manifest_out, &manifest_json).map_err(|e| {
        CliError::io(format!("cannot write {}: {e}", manifest_out.display()))
    })?;

    if !quiet {
        eprintln!("export: wrote {} daily totals", totals.len());
        eprintln!("  {}", dt_out.display());
        eprintln!("  {}", manifest_out.display());
    }

    Ok(())
}

// ── Transaction CSV reader ──────────────────────────────────────────

/// Read truth_transactions.csv into TruthTransaction structs.
///
/// This is a minimal reader that expects the canonical header format.
fn read_transactions_csv(data: &[u8]) -> Result<Vec<TruthTransaction>, CliError> {
    use visigrid_io::truth::{parse_amount, Direction};
    use chrono::NaiveDate;

    let mut csv = csv::ReaderBuilder::new()
        .has_headers(true)
        .from_reader(data);

    let mut transactions = Vec::new();

    for (i, result) in csv.records().enumerate() {
        let record = result.map_err(|e| {
            CliError::parse(format!("CSV parse error at row {}: {e}", i + 1))
        })?;

        if record.len() < 13 {
            return Err(CliError::parse(format!(
                "row {} has {} columns, expected 13",
                i + 1,
                record.len()
            )));
        }

        let occurred_at = NaiveDate::parse_from_str(&record[3], "%Y-%m-%d").map_err(|e| {
            CliError::parse(format!("row {}: invalid occurred_at '{}': {e}", i + 1, &record[3]))
        })?;

        let posted_at = if record[4].is_empty() {
            None
        } else {
            Some(NaiveDate::parse_from_str(&record[4], "%Y-%m-%d").map_err(|e| {
                CliError::parse(format!("row {}: invalid posted_at '{}': {e}", i + 1, &record[4]))
            })?)
        };

        let direction = match record[6].as_ref() {
            "credit" => Direction::Credit,
            "debit" => Direction::Debit,
            other => return Err(CliError::parse(format!(
                "row {}: invalid direction '{other}'", i + 1
            ))),
        };

        let amount_gross = if record[7].is_empty() {
            None
        } else {
            Some(parse_amount(&record[7]).ok_or_else(|| {
                CliError::parse(format!("row {}: invalid amount_gross '{}'", i + 1, &record[7]))
            })?)
        };

        let fee_amount = if record[8].is_empty() {
            None
        } else {
            Some(parse_amount(&record[8]).ok_or_else(|| {
                CliError::parse(format!("row {}: invalid fee_amount '{}'", i + 1, &record[8]))
            })?)
        };

        let amount_net = parse_amount(&record[9]).ok_or_else(|| {
            CliError::parse(format!("row {}: invalid amount_net '{}'", i + 1, &record[9]))
        })?;

        transactions.push(TruthTransaction {
            source: record[0].to_string(),
            source_account: record[1].to_string(),
            source_id: record[2].to_string(),
            occurred_at,
            posted_at,
            currency: record[5].to_string(),
            direction,
            amount_gross,
            fee_amount,
            amount_net,
            counterparty: if record[10].is_empty() { None } else { Some(record[10].to_string()) },
            description: if record[11].is_empty() { None } else { Some(record[11].to_string()) },
            raw_hash: record[12].to_string(),
        });
    }

    Ok(transactions)
}
