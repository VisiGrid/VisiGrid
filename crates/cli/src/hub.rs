//! VisiHub CLI commands: login and publish.
//!
//! `vgrid login`   — store API token
//! `vgrid publish`  — upload file, wait for check, print results

use std::collections::HashMap;
use std::io::{self, Write};
use std::path::PathBuf;
use std::time::Duration;

use visigrid_hub_client::{
    AuthCredentials, save_auth,
    HubClient, HubError, CreateRevisionOptions, RunResult,
    AssertionInput, EngineMetadata,
    hash_file,
};

use crate::{CliError, OutputFormat};
use crate::exit_codes::*;
use crate::sheet_ops::parse_cell_ref;

// ── Login ───────────────────────────────────────────────────────────

pub fn cmd_login(token: Option<String>, api_base: String) -> Result<(), CliError> {
    // Resolve token: --token flag > VISIHUB_API_KEY env > interactive prompt
    let token = if let Some(t) = token {
        t
    } else if let Ok(t) = std::env::var("VISIHUB_API_KEY") {
        t
    } else if atty::is(atty::Stream::Stdin) {
        eprint!("VisiHub API token: ");
        io::stderr().flush().ok();
        let mut buf = String::new();
        io::stdin().read_line(&mut buf)
            .map_err(|e| CliError { code: EXIT_ERROR, message: e.to_string(), hint: None })?;
        let trimmed = buf.trim().to_string();
        if trimmed.is_empty() {
            return Err(CliError {
                code: EXIT_USAGE,
                message: "No token provided".into(),
                hint: Some("pass --token or set VISIHUB_API_KEY".into()),
            });
        }
        trimmed
    } else {
        return Err(CliError {
            code: EXIT_USAGE,
            message: "No token provided and stdin is not a TTY".into(),
            hint: Some("pass --token or set VISIHUB_API_KEY".into()),
        });
    };

    // Verify the token works
    let creds = AuthCredentials::new(token.clone(), api_base.clone());
    let client = HubClient::new(creds.clone());

    let user = client.verify_token().map_err(|e| match e {
        HubError::Http(401, _) | HubError::Http(403, _) => CliError {
            code: EXIT_HUB_NOT_AUTH,
            message: "Invalid API token".into(),
            hint: Some("generate a new token at app.visihub.app/settings/tokens".into()),
        },
        HubError::Network(msg) => CliError {
            code: EXIT_HUB_NETWORK,
            message: format!("Cannot reach VisiHub: {}", msg),
            hint: None,
        },
        other => hub_error(other),
    })?;

    // Save with user info
    let creds = AuthCredentials {
        token,
        api_base,
        user_slug: Some(user.slug.clone()),
        email: Some(user.email.clone()),
    };

    save_auth(&creds).map_err(|e| CliError {
        code: EXIT_ERROR,
        message: e,
        hint: None,
    })?;

    eprintln!("Authenticated as {} ({})", user.slug, user.email);
    Ok(())
}

// ── Publish ─────────────────────────────────────────────────────────

pub fn cmd_publish(
    file: PathBuf,
    repo: String,
    dataset: Option<String>,
    source_type: Option<String>,
    source_identity: Option<String>,
    query_hash: Option<String>,
    wait: bool,
    fail_on_check_failure: bool,
    output_fmt: Option<OutputFormat>,
    assert_sum: Vec<String>,
    assert_cell: Vec<String>,
    reset_baseline: bool,
    row_count_policy: Option<String>,
    columns_added_policy: Option<String>,
    columns_removed_policy: Option<String>,
    strict: bool,
) -> Result<(), CliError> {
    // Validate inputs
    if !file.exists() {
        return Err(CliError {
            code: EXIT_USAGE,
            message: format!("File not found: {}", file.display()),
            hint: None,
        });
    }

    let parts: Vec<&str> = repo.splitn(2, '/').collect();
    if parts.len() != 2 || parts[0].is_empty() || parts[1].is_empty() {
        return Err(CliError {
            code: EXIT_USAGE,
            message: format!("Invalid repo format: '{}' (expected owner/slug)", repo),
            hint: Some("example: --repo acme/payments".into()),
        });
    }
    let owner = parts[0];
    let slug = parts[1];

    let dataset_name = dataset.unwrap_or_else(|| {
        file.file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("data.csv")
            .to_string()
    });

    // Detect file format from extension
    let file_format = match file.extension().and_then(|e| e.to_str()) {
        Some("sheet") => Some("sheet"),
        Some("tsv") => Some("tsv"),
        Some("xlsx") => Some("xlsx"),
        _ => None, // server defaults to csv
    };

    // Determine output mode
    let json_output = match output_fmt {
        Some(OutputFormat::Json) => true,
        Some(OutputFormat::Text) => false,
        None => !atty::is(atty::Stream::Stdout),
    };

    let client = HubClient::from_saved_auth().map_err(|e| match e {
        HubError::NotAuthenticated => CliError {
            code: EXIT_HUB_NOT_AUTH,
            message: "Not authenticated".into(),
            hint: Some("run `vgrid login` first".into()),
        },
        other => hub_error(other),
    })?;

    // Step 1: Hash file
    if !json_output { eprint!("Hashing... "); }
    let byte_size = std::fs::metadata(&file)
        .map_err(|e| CliError { code: EXIT_ERROR, message: e.to_string(), hint: None })?
        .len();
    let content_hash = hash_file(&file).map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("{} ({} bytes)", &content_hash[..15], byte_size); }

    // Step 2: Find or create dataset
    if !json_output { eprint!("Finding dataset '{}'... ", dataset_name); }
    let datasets = client.list_datasets(owner, slug).map_err(|e| hub_error(e))?;
    let dataset_id = if let Some(d) = datasets.iter().find(|d| d.name == dataset_name) {
        if !json_output { eprintln!("found"); }
        d.id.clone()
    } else {
        if !json_output { eprint!("creating... "); }
        let id = client.create_dataset(owner, slug, &dataset_name, file_format).map_err(|e| hub_error(e))?;
        if !json_output { eprintln!("created #{}", id); }
        id
    };

    // Parse --assert-sum flags into AssertionInput
    let mut assertions: Vec<AssertionInput> = assert_sum.iter().map(|s| {
        let parts: Vec<&str> = s.splitn(3, ':').collect();
        AssertionInput {
            kind: "sum".into(),
            column: parts.first().unwrap_or(&"").to_string(),
            expected: parts.get(1).map(|s| s.to_string()),
            tolerance: parts.get(2).map(|s| s.to_string()),
            actual: None,
            origin: None,
            engine: None,
        }
    }).collect();

    // Evaluate --assert-cell flags (requires .sheet file)
    if !assert_cell.is_empty() {
        let cell_assertions = evaluate_cell_assertions(&file, &assert_cell)?;
        assertions.extend(cell_assertions);
    }

    // Build check_policy from flags
    let check_policy = {
        let mut policy = HashMap::new();
        if strict {
            policy.insert("row_count".into(), "fail".into());
            policy.insert("columns_added".into(), "fail".into());
            policy.insert("columns_removed".into(), "fail".into());
        } else {
            if let Some(v) = row_count_policy { policy.insert("row_count".into(), v); }
            if let Some(v) = columns_added_policy { policy.insert("columns_added".into(), v); }
            if let Some(v) = columns_removed_policy { policy.insert("columns_removed".into(), v); }
        }
        if policy.is_empty() { None } else { Some(policy) }
    };

    // Step 3: Create revision
    if !json_output { eprint!("Creating revision... "); }
    let opts = CreateRevisionOptions {
        source_type,
        source_identity,
        query_hash,
        assertions,
        reset_baseline,
        check_policy,
        format: file_format.map(String::from),
    };
    let (revision_id, upload_url, upload_headers) = client
        .create_revision(&dataset_id, &content_hash, byte_size, &opts)
        .map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("#{}", revision_id); }

    // Step 4: Upload
    if !json_output { eprint!("Uploading {} bytes... ", byte_size); }
    let data = std::fs::read(&file)
        .map_err(|e| CliError { code: EXIT_ERROR, message: e.to_string(), hint: None })?;
    client.upload_bytes(&upload_url, data, &upload_headers).map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("done"); }

    // Step 5: Complete
    if !json_output { eprint!("Finalizing... "); }
    client.complete_revision(&revision_id, &content_hash).map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("done"); }

    // Step 6: Poll (optional)
    if !wait {
        let proof_url = client.proof_url(owner, slug, &revision_id);
        if json_output {
            let out = serde_json::json!({
                "run_id": revision_id,
                "status": "processing",
                "proof_url": proof_url,
            });
            println!("{}", serde_json::to_string(&out).unwrap());
        } else {
            eprintln!("Revision #{} submitted (not waiting for results)", revision_id);
            eprintln!("Proof: {}", proof_url);
        }
        return Ok(());
    }

    if !json_output { eprint!("Waiting for import... "); }
    let result = client
        .poll_run(owner, slug, &revision_id, Duration::from_secs(120))
        .map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("{}", result.status); }

    // Step 7: Output results
    if json_output {
        println!("{}", serde_json::to_string(&result).unwrap());
    } else {
        print_human_result(&result);
    }

    // Step 8: Exit code
    if fail_on_check_failure && result.check_status.as_deref() == Some("fail") {
        return Err(CliError {
            code: EXIT_HUB_CHECK_FAILED,
            message: format!(
                "Integrity check failed for {} v{}",
                dataset_name, result.version
            ),
            hint: Some(format!("Proof: {}", result.proof_url)),
        });
    }

    Ok(())
}

// ── Helpers ─────────────────────────────────────────────────────────

fn print_human_result(r: &RunResult) {
    eprintln!();
    eprintln!("  Run:     #{}", r.run_id);
    eprintln!("  Version: v{}", r.version);
    eprintln!("  Status:  {}", r.status);
    if let Some(ref cs) = r.check_status {
        let marker = match cs.as_str() {
            "pass" => "PASS",
            "warn" => "WARN",
            "baseline_created" => "BASELINE CREATED",
            _ => "FAIL",
        };
        eprintln!("  Check:   {}", marker);
    }
    if let Some(ref diff) = r.diff_summary {
        if let Some(rc) = diff.get("row_count_change").and_then(|v| v.as_i64()) {
            if rc != 0 {
                let sign = if rc > 0 { "+" } else { "" };
                eprintln!("  Rows:    {}{}", sign, rc);
            }
        }
        if let Some(cc) = diff.get("col_count_change").and_then(|v| v.as_i64()) {
            if cc != 0 {
                let sign = if cc > 0 { "+" } else { "" };
                eprintln!("  Cols:    {}{}", sign, cc);
            }
        }
    }
    if let Some(ref assertions) = r.assertions {
        for a in assertions {
            let origin_tag = if a.origin.as_deref() == Some("client") { " [client]" } else { "" };
            let label = format!("{}({})", a.kind, a.column);
            match a.status.as_str() {
                "pass" => eprintln!("  {}: PASS (actual={}){}", label, a.actual.as_deref().unwrap_or("?"), origin_tag),
                "fail" => {
                    if let Some(ref msg) = a.message {
                        eprintln!("  {}: FAIL ({}){}", label, msg, origin_tag);
                    } else {
                        eprintln!("  {}: FAIL (expected={} actual={} delta={}){}", label,
                            a.expected.as_deref().unwrap_or("?"),
                            a.actual.as_deref().unwrap_or("?"),
                            a.delta.as_deref().unwrap_or("?"),
                            origin_tag);
                    }
                }
                "baseline_created" => eprintln!("  {}: BASELINE (actual={}){}", label, a.actual.as_deref().unwrap_or("?"), origin_tag),
                _ => eprintln!("  {}: {}{}", label, a.status, origin_tag),
            }
        }
    }
    if let Some(ref hash) = r.content_hash {
        eprintln!("  Hash:    {}", hash);
    }
    eprintln!("  Proof:   {}", r.proof_url);
    eprintln!();
}

// ── Cell Assertions ──────────────────────────────────────────────────

/// Parse and evaluate `--assert-cell` flags against a .sheet file.
///
/// Each spec has the format: `sheet!cell:expected[:tolerance]`
/// The workbook is loaded once, recalculated, and all assertions evaluated.
fn evaluate_cell_assertions(
    file: &std::path::Path,
    specs: &[String],
) -> Result<Vec<AssertionInput>, CliError> {
    use visigrid_engine::formula::eval::Value;
    use visigrid_io::native;

    // Require .sheet extension
    let ext = file.extension().and_then(|e| e.to_str()).unwrap_or("");
    if ext != "sheet" {
        return Err(CliError {
            code: EXIT_USAGE,
            message: "--assert-cell requires a .sheet file".into(),
            hint: Some("cell assertions evaluate formulas in the spreadsheet engine".into()),
        });
    }

    // Load workbook once
    let mut workbook = native::load_workbook(file)
        .map_err(|e| CliError::io(format!("failed to load .sheet for cell assertions: {}", e)))?;
    workbook.rebuild_dep_graph();
    workbook.recompute_full_ordered();

    let fingerprint = native::compute_semantic_fingerprint(&workbook);
    let engine = EngineMetadata {
        name: "visigrid-engine".into(),
        version: env!("CARGO_PKG_VERSION").into(),
        fingerprint: Some(fingerprint),
    };

    let mut assertions = Vec::with_capacity(specs.len());

    for spec in specs {
        // Parse sheet!cell:expected[:tolerance]
        // First split on ':' but we need to handle the sheet!cell part
        // which itself contains no colons.
        let parts: Vec<&str> = spec.splitn(3, ':').collect();
        if parts.is_empty() {
            return Err(CliError::args(format!(
                "invalid --assert-cell format: {:?}",
                spec
            )));
        }

        let cell_ref_str = parts[0]; // e.g., "summary!B7"
        let expected = parts.get(1).map(|s| s.to_string());
        let tolerance = parts.get(2).map(|s| s.to_string());

        // Parse sheet!cell reference
        let (sheet_name, cell_part) = if let Some(bang) = cell_ref_str.find('!') {
            let sheet = &cell_ref_str[..bang];
            let cell = &cell_ref_str[bang + 1..];
            if sheet.is_empty() {
                return Err(CliError::args(format!(
                    "empty sheet name in --assert-cell {:?}",
                    spec
                )));
            }
            (Some(sheet), cell)
        } else {
            (None, cell_ref_str)
        };

        let (row, col) = match parse_cell_ref(cell_part) {
            Some(rc) => rc,
            None => {
                return Err(CliError::args(format!(
                    "invalid cell reference in --assert-cell {:?}",
                    spec
                )));
            }
        };

        // Resolve sheet
        let sheet_idx = if let Some(name) = sheet_name {
            let sheet_id = workbook.sheet_id_by_name(name).ok_or_else(|| {
                CliError::args(format!("sheet {:?} not found in --assert-cell {:?}", name, spec))
            })?;
            workbook.idx_for_sheet_id(sheet_id).ok_or_else(|| {
                CliError::args(format!("sheet {:?} not found", name))
            })?
        } else {
            0
        };

        let sheet = match workbook.sheet(sheet_idx) {
            Some(s) => s,
            None => {
                assertions.push(AssertionInput {
                    kind: "cell".into(),
                    column: cell_ref_str.to_string(),
                    expected,
                    tolerance,
                    actual: None,
                    origin: Some("client".into()),
                    engine: Some(engine.clone()),
                });
                continue;
            }
        };

        // Read computed value
        let value = sheet.get_computed_value(row, col);

        let assertion = match value {
            Value::Number(n) => {
                // Format number without trailing zeros for clean comparison
                let actual_str = format_number(n);
                AssertionInput {
                    kind: "cell".into(),
                    column: cell_ref_str.to_string(),
                    expected,
                    tolerance,
                    actual: Some(actual_str),
                    origin: Some("client".into()),
                    engine: Some(engine.clone()),
                }
            }
            Value::Empty => {
                AssertionInput {
                    kind: "cell".into(),
                    column: cell_ref_str.to_string(),
                    expected: None,
                    tolerance: None,
                    actual: None,
                    origin: Some("client".into()),
                    engine: Some(engine.clone()),
                }
            }
            Value::Error(_) | Value::Text(_) | Value::Boolean(_) => {
                AssertionInput {
                    kind: "cell".into(),
                    column: cell_ref_str.to_string(),
                    expected,
                    tolerance,
                    actual: None,
                    origin: Some("client".into()),
                    engine: Some(engine.clone()),
                }
            }
        };

        assertions.push(assertion);
    }

    Ok(assertions)
}

/// Format a number for assertion comparison: integers as integers,
/// decimals with minimum necessary precision.
fn format_number(n: f64) -> String {
    if n == n.trunc() && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        // Use enough precision to round-trip
        let s = format!("{}", n);
        s
    }
}

fn hub_error(e: HubError) -> CliError {
    match e {
        HubError::NotAuthenticated => CliError {
            code: EXIT_HUB_NOT_AUTH,
            message: "Not authenticated".into(),
            hint: Some("run `vgrid login` first".into()),
        },
        HubError::Network(msg) => CliError {
            code: EXIT_HUB_NETWORK,
            message: msg,
            hint: None,
        },
        HubError::Http(code, msg) => CliError {
            code: EXIT_HUB_NETWORK,
            message: format!("HTTP {}: {}", code, msg),
            hint: None,
        },
        HubError::Validation(msg) => CliError {
            code: EXIT_HUB_VALIDATION,
            message: msg,
            hint: None,
        },
        HubError::Parse(msg) => CliError {
            code: EXIT_HUB_NETWORK,
            message: format!("Unexpected response: {}", msg),
            hint: None,
        },
        HubError::Io(msg) => CliError {
            code: EXIT_ERROR,
            message: msg,
            hint: None,
        },
        HubError::Timeout(msg) => CliError {
            code: EXIT_HUB_TIMEOUT,
            message: msg,
            hint: None,
        },
    }
}
