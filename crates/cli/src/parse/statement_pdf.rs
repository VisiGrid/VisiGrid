//! `vgrid parse statement-pdf` — extract canonical CSV from processor statement PDFs.

use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

use crate::fetch::common::{write_csv, CanonicalRow};
use crate::CliError;

use super::fiserv_cardpointe_v1;

const AVAILABLE_TEMPLATES: &[&str] = &["fiserv_cardpointe_v1"];

pub(super) fn cmd_parse_statement_pdf(
    template: &str,
    file: &Path,
    out: &Option<PathBuf>,
    save_raw: &Option<PathBuf>,
    quiet: bool,
) -> Result<(), CliError> {
    // Validate template
    if !AVAILABLE_TEMPLATES.contains(&template) {
        return Err(CliError::args(format!(
            "Unknown template: {} (available: {})",
            template,
            AVAILABLE_TEMPLATES.join(", "),
        )));
    }

    // Run pdftotext
    let text = run_pdftotext(file)?;

    if !quiet {
        eprintln!("Extracted {} bytes of text from {}", text.len(), file.display());
    }

    // Dispatch to template parser
    let parsed = match template {
        "fiserv_cardpointe_v1" => fiserv_cardpointe_v1::parse(&text)?,
        _ => unreachable!(),
    };

    if !quiet {
        eprintln!(
            "Parsed {} daily rows for merchant {} ({} to {})",
            parsed.rows.len(),
            parsed.merchant_id,
            parsed.period_start,
            parsed.period_end,
        );
    }

    // Cross-check: sum of rows + month_end_charge should equal Total line
    let sum: i64 = parsed.rows.iter().map(|r| r.amount_minor).sum();
    let month_end = parsed.month_end_charge_minor.unwrap_or(0);
    let expected_total = sum + month_end;
    if let Some(total) = parsed.total_amount_minor {
        if expected_total != total {
            eprintln!(
                "warning: row sum ({}) + month_end_charge ({}) = {} != Total line ({}) — delta: {} cents",
                sum, month_end, expected_total, total, expected_total - total,
            );
        }
    }

    // Build canonical rows with collision-resistant source_id
    let group_id = format!(
        "stmt:{}:{}_{}", parsed.merchant_id, parsed.period_start, parsed.period_end,
    );

    // Track sequence numbers for rows with identical date+amount
    let mut seen: std::collections::HashMap<String, u32> = std::collections::HashMap::new();
    let mut rows: Vec<CanonicalRow> = Vec::with_capacity(parsed.rows.len());

    for day in &parsed.rows {
        let key = format!("{}:{}", day.date, day.amount_minor);
        let seq = seen.entry(key).or_insert(0);
        *seq += 1;
        let source_id = if *seq == 1 {
            format!("stmt:{}:{}:{}", parsed.merchant_id, day.date, day.amount_minor)
        } else {
            format!(
                "stmt:{}:{}:{}:{}",
                parsed.merchant_id, day.date, day.amount_minor, seq,
            )
        };

        rows.push(CanonicalRow {
            effective_date: day.date.clone(),
            posted_date: day.date.clone(),
            amount_minor: day.amount_minor,
            currency: "USD".to_string(),
            r#type: "statement_daily".to_string(),
            source: "fiserv_statement".to_string(),
            source_id,
            group_id: group_id.clone(),
            description: String::new(),
        });
    }

    // Write CSV
    let out_label = write_csv(&rows, out)?;
    if !quiet {
        eprintln!("Wrote {} rows to {}", rows.len(), out_label);
    }

    // Save raw artifacts
    if let Some(raw_dir) = save_raw {
        fs::create_dir_all(raw_dir).map_err(|e| {
            CliError::io(format!("cannot create {}: {}", raw_dir.display(), e))
        })?;

        // pdftotext output
        let txt_path = raw_dir.join("pdftotext_output.txt");
        fs::write(&txt_path, &text).map_err(|e| {
            CliError::io(format!("cannot write {}: {}", txt_path.display(), e))
        })?;

        // engine_meta.json
        let meta = serde_json::json!({
            "template": template,
            "merchant_id": parsed.merchant_id,
            "period_start": parsed.period_start,
            "period_end": parsed.period_end,
            "row_count": parsed.rows.len(),
            "sum_rows_amount_minor": sum,
            "month_end_charge_minor": parsed.month_end_charge_minor,
            "total_line_amount_minor": parsed.total_amount_minor,
            "has_total_crosscheck": parsed.total_amount_minor.is_some(),
            "delta_minor": parsed.total_amount_minor.map(|t| expected_total - t),
            "vgrid_version": env!("CARGO_PKG_VERSION"),
        });
        let meta_path = raw_dir.join("engine_meta.json");
        let meta_str = serde_json::to_string_pretty(&meta).unwrap();
        fs::write(&meta_path, &meta_str).map_err(|e| {
            CliError::io(format!("cannot write {}: {}", meta_path.display(), e))
        })?;

        if !quiet {
            eprintln!("Saved raw artifacts to {}", raw_dir.display());
        }
    }

    Ok(())
}

/// Run `pdftotext -layout <file> -` and capture stdout.
fn run_pdftotext(file: &Path) -> Result<String, CliError> {
    // Check that pdftotext exists
    which::which("pdftotext").map_err(|_| {
        CliError {
            code: crate::EXIT_IO_ERROR,
            message: "pdftotext not installed (poppler-utils)".to_string(),
            hint: Some("Install with: apt install poppler-utils / brew install poppler".to_string()),
        }
    })?;

    let file_str = file.to_str().ok_or_else(|| {
        CliError::args(format!("invalid file path: {}", file.display()))
    })?;

    let output = Command::new("pdftotext")
        .args(["-layout", file_str, "-"])
        .output()
        .map_err(|e| {
            CliError::io(format!("failed to run pdftotext: {}", e))
        })?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(CliError::io(format!(
            "pdftotext failed (exit {}): {}",
            output.status.code().unwrap_or(-1),
            stderr.trim(),
        )));
    }

    let text = String::from_utf8_lossy(&output.stdout).to_string();

    if text.trim().is_empty() {
        return Err(CliError::parse(
            "PDF appears scanned/image-only — text extraction failed",
        ));
    }

    Ok(text)
}
