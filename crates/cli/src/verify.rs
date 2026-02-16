//! `vgrid verify` — financial verification commands.
//!
//! Currently supports one subcommand:
//! - `vgrid verify totals` — compare truth vs warehouse daily totals

use std::io::Write;
use std::path::PathBuf;

use clap::Subcommand;
use serde::{Deserialize, Serialize};

use visigrid_io::truth::{hash_raw_row, read_daily_totals_csv, DailyTotals};

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
  vgrid verify totals truth.csv warehouse.csv --diff diffs.csv
  vgrid verify totals truth.csv warehouse.csv --sign --proof proof.json")]
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

        /// Sign the verification result with Ed25519
        #[arg(long)]
        sign: bool,

        /// Output signed proof JSON to file (implies --sign)
        #[arg(long)]
        proof: Option<PathBuf>,

        /// Path to signing key (default: ~/.config/vgrid/proof_key.json, or VGRID_SIGNING_KEY_PATH env)
        #[arg(long, env = "VGRID_SIGNING_KEY_PATH")]
        signing_key: Option<PathBuf>,
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
            sign,
            proof,
            signing_key,
        } => cmd_verify_totals(
            truth,
            warehouse,
            tolerance,
            no_fail_on_count,
            output,
            diff,
            quiet,
            sign || proof.is_some(),
            proof,
            signing_key,
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
    sign: bool,
    proof_path: Option<PathBuf>,
    signing_key_path: Option<PathBuf>,
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
    let truth = read_daily_totals_csv(truth_bytes.as_slice())
        .map_err(|e| CliError::parse(format!("truth file: {e}")))?;
    let warehouse = read_daily_totals_csv(warehouse_bytes.as_slice())
        .map_err(|e| CliError::parse(format!("warehouse file: {e}")))?;

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
        eprintln!(
            "verify: {} ({} rows truth, {} rows warehouse)",
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
            eprintln!(
                "  only in warehouse:  {}",
                result.summary.only_in_warehouse
            );
        }
        eprintln!("  truth hash:         {}", &result.truth_hash[..16]);
        eprintln!("  warehouse hash:     {}", &result.warehouse_hash[..16]);
    }

    // Write JSON output
    if let Some(path) = &output_path {
        let json = serde_json::to_string_pretty(&result)
            .map_err(|e| CliError::io(format!("JSON serialization error: {e}")))?;
        std::fs::write(path, json)
            .map_err(|e| CliError::io(format!("cannot write {}: {e}", path.display())))?;
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

    // Sign and write proof
    if sign {
        let proof_out = proof_path.unwrap_or_else(|| PathBuf::from("proof.json"));
        let sig_out = proof_out.with_extension("sig");

        let signed = sign_proof(&result, &signing_key_path)?;

        let proof_json = serde_json::to_string_pretty(&signed)
            .map_err(|e| CliError::io(format!("proof JSON error: {e}")))?;
        std::fs::write(&proof_out, &proof_json)
            .map_err(|e| CliError::io(format!("cannot write {}: {e}", proof_out.display())))?;

        std::fs::write(&sig_out, &signed.signature)
            .map_err(|e| CliError::io(format!("cannot write {}: {e}", sig_out.display())))?;

        if !quiet {
            eprintln!("  proof written to:   {}", proof_out.display());
            eprintln!("  signature:          {}", sig_out.display());
            eprintln!("  public key:         {}", &signed.public_key[..16]);
        }
    }

    if has_material_mismatch {
        Err(CliError {
            code: exit_codes::EXIT_ERROR,
            message: String::new(),
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
    let file = std::fs::File::create(path)
        .map_err(|e| CliError::io(format!("cannot create {}: {e}", path.display())))?;
    let mut writer = std::io::BufWriter::new(file);

    writeln!(
        writer,
        "date,currency,source_account,kind,truth_value,warehouse_value"
    )
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

    writer
        .flush()
        .map_err(|e| CliError::io(format!("flush error: {e}")))?;
    Ok(())
}

// ── Proof signing ───────────────────────────────────────────────────

#[derive(Debug, Serialize)]
struct SignedProof {
    schema_version: u32,
    payload: ProofPayload,
    signature: String,
    public_key: String,
}

#[derive(Debug, Serialize)]
struct ProofPayload {
    schema_version: u32,
    verifier: ProofVerifier,
    ran_at: String,
    truth: ProofFileRef,
    warehouse: ProofFileRef,
    params: ProofParams,
    result: ProofOutcome,
}

#[derive(Debug, Serialize)]
struct ProofVerifier {
    name: String,
    version: String,
}

#[derive(Debug, Serialize)]
struct ProofFileRef {
    path: String,
    blake3: String,
    rows: usize,
}

#[derive(Debug, Serialize)]
struct ProofParams {
    tolerance_micro: i64,
    fail_on_count_mismatch: bool,
}

#[derive(Debug, Serialize)]
struct ProofOutcome {
    status: String,
    matched: usize,
    mismatched: usize,
    missing_in_warehouse: usize,
    missing_in_truth: usize,
}

#[derive(Debug, Serialize, Deserialize)]
struct StoredKeypair {
    public_key: String,
    secret_key: String,
}

fn sign_proof(
    result: &VerifyResult,
    signing_key_path: &Option<PathBuf>,
) -> Result<SignedProof, CliError> {
    use base64::Engine;
    use ed25519_dalek::Signer;

    let (signing_key, verifying_key) = load_or_generate_key(signing_key_path)?;

    let payload = ProofPayload {
        schema_version: 1,
        verifier: ProofVerifier {
            name: "vgrid".to_string(),
            version: env!("CARGO_PKG_VERSION").to_string(),
        },
        ran_at: chrono::Utc::now().to_rfc3339(),
        truth: ProofFileRef {
            path: result.truth_file.clone(),
            blake3: result.truth_hash.clone(),
            rows: result.summary.truth_rows,
        },
        warehouse: ProofFileRef {
            path: result.warehouse_file.clone(),
            blake3: result.warehouse_hash.clone(),
            rows: result.summary.warehouse_rows,
        },
        params: ProofParams {
            tolerance_micro: result.tolerance_micro,
            fail_on_count_mismatch: result.fail_on_count_mismatch,
        },
        result: ProofOutcome {
            status: result.status.to_string(),
            matched: result.summary.matched,
            mismatched: result.summary.mismatched,
            missing_in_warehouse: result.summary.only_in_truth,
            missing_in_truth: result.summary.only_in_warehouse,
        },
    };

    // Compact JSON for deterministic signing
    let payload_bytes = serde_json::to_vec(&payload)
        .map_err(|e| CliError::io(format!("proof serialization error: {e}")))?;

    let signature = signing_key.sign(&payload_bytes);
    let b64 = base64::engine::general_purpose::STANDARD;

    Ok(SignedProof {
        schema_version: 1,
        payload,
        signature: b64.encode(signature.to_bytes()),
        public_key: b64.encode(verifying_key.to_bytes()),
    })
}

fn load_or_generate_key(
    key_path: &Option<PathBuf>,
) -> Result<(ed25519_dalek::SigningKey, ed25519_dalek::VerifyingKey), CliError> {
    use base64::Engine;
    let b64 = base64::engine::general_purpose::STANDARD;

    let path = match key_path {
        Some(p) => p.clone(),
        None => {
            let config_dir = dirs::config_dir()
                .ok_or_else(|| {
                    CliError::io("cannot determine config directory".to_string())
                })?
                .join("vgrid");
            config_dir.join("proof_key.json")
        }
    };

    if path.exists() {
        let data = std::fs::read_to_string(&path)
            .map_err(|e| CliError::io(format!("cannot read {}: {e}", path.display())))?;
        let stored: StoredKeypair = serde_json::from_str(&data)
            .map_err(|e| CliError::parse(format!("invalid key file {}: {e}", path.display())))?;
        let secret_bytes = b64
            .decode(&stored.secret_key)
            .map_err(|e| CliError::parse(format!("invalid secret key base64: {e}")))?;
        let secret_array: [u8; 32] = secret_bytes
            .try_into()
            .map_err(|_| CliError::parse("secret key must be 32 bytes".to_string()))?;
        let signing_key = ed25519_dalek::SigningKey::from_bytes(&secret_array);
        let verifying_key = signing_key.verifying_key();
        Ok((signing_key, verifying_key))
    } else {
        let mut rng = rand::thread_rng();
        let signing_key = ed25519_dalek::SigningKey::generate(&mut rng);
        let verifying_key = signing_key.verifying_key();

        let stored = StoredKeypair {
            public_key: b64.encode(verifying_key.to_bytes()),
            secret_key: b64.encode(signing_key.to_bytes()),
        };

        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)
                .map_err(|e| CliError::io(format!("cannot create {}: {e}", parent.display())))?;
        }

        let json = serde_json::to_string_pretty(&stored)
            .map_err(|e| CliError::io(format!("key serialization error: {e}")))?;
        std::fs::write(&path, json)
            .map_err(|e| CliError::io(format!("cannot write {}: {e}", path.display())))?;

        eprintln!("  generated new signing key: {}", path.display());
        Ok((signing_key, verifying_key))
    }
}
