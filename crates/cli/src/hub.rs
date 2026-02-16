//! VisiHub CLI commands: login, publish, hub publish, and pipeline publish.
//!
//! `vgrid login`              — store API token
//! `vgrid publish`            — upload file, wait for check, print results
//! `vgrid hub publish`        — publish verified .sheet with trust metadata
//! `vgrid pipeline publish`   — import → verify → publish in one step

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

use crate::{CliError, FormulaPolicy, OutputFormat};
use crate::exit_codes::*;
use crate::sheet_ops::{self, parse_cell_ref, CalcOutput, CalcResult};
use crate::util;

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
    let datasets = client.list_datasets(owner, slug).map_err(|e| match &e {
        HubError::Http(404, _) => CliError {
            code: EXIT_HUB_NETWORK,
            message: format!("Repository '{}/{}' not found", owner, slug),
            hint: Some("check the --repo value or create the repository on VisiHub first".into()),
        },
        HubError::Http(403, _) => CliError {
            code: EXIT_HUB_NOT_AUTH,
            message: format!("No permission to access '{}/{}'", owner, slug),
            hint: Some("check your API token permissions or ask the repo owner for access".into()),
        },
        _ => hub_error(e),
    })?;
    let dataset_id = if let Some(d) = datasets.iter().find(|d| d.name == dataset_name) {
        if !json_output { eprintln!("found"); }
        d.id.clone()
    } else {
        if !json_output { eprint!("creating... "); }
        let id = client.create_dataset(owner, slug, &dataset_name, file_format).map_err(|e| match &e {
            HubError::Http(403, _) | HubError::Http(422, _) => CliError {
                code: EXIT_HUB_NOT_AUTH,
                message: format!("Cannot create dataset in '{}/{}'", owner, slug),
                hint: Some("repo not found or you don't have permission to create datasets — check your API token".into()),
            },
            _ => hub_error(e),
        })?;
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
    // Attach CI runner identity when running in a recognized CI environment
    let source_metadata = crate::ci::get_runner_context().map(|runner| {
        serde_json::json!({ "runner": runner })
    });

    let opts = CreateRevisionOptions {
        source_type,
        source_identity,
        query_hash,
        assertions,
        reset_baseline,
        check_policy,
        format: file_format.map(String::from),
        source_metadata,
        message: None,
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

// ── Hub Publish (trust pipeline) ────────────────────────────────────

/// Maximum checks JSON payload size (256 KB).
const MAX_CHECKS_BYTES: usize = 256 * 1024;

pub fn cmd_hub_publish(
    file: PathBuf,
    repo: String,
    message: Option<String>,
    notes: Option<PathBuf>,
    checks: Option<PathBuf>,
    lock: bool,
    json_output: bool,
    dry_run: bool,
    no_wait: bool,
    timeout: u64,
) -> Result<(), CliError> {
    // ── Phase A: Local validation (no auth, no network) ────────────

    // 1. Validate file exists + .sheet extension
    if !file.exists() {
        return Err(CliError {
            code: EXIT_USAGE,
            message: format!("File not found: {}", file.display()),
            hint: None,
        });
    }
    let ext = file.extension().and_then(|e| e.to_str()).unwrap_or("");
    if ext != "sheet" {
        return Err(CliError {
            code: EXIT_USAGE,
            message: format!("Expected a .sheet file, got .{}", ext),
            hint: Some("use `vgrid publish` for CSV/TSV/XLSX".into()),
        });
    }

    // 2. Validate repo format
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

    let filename = file.file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("data.sheet")
        .to_string();

    // 3. Load .sheet + extract trust metadata
    let workbook = visigrid_io::native::load_workbook(&file)
        .map_err(|e| CliError {
            code: EXIT_ERROR,
            message: format!("Failed to load .sheet: {}", e),
            hint: None,
        })?;
    let fingerprint = visigrid_io::native::compute_semantic_fingerprint(&workbook);
    let verification = visigrid_io::native::load_semantic_verification(&file).ok();

    // 4. Load optional checks JSON (with size guard)
    let checks_json: Option<serde_json::Value> = if let Some(ref checks_path) = checks {
        let data = std::fs::read_to_string(checks_path)
            .map_err(|e| CliError {
                code: EXIT_USAGE,
                message: format!("Cannot read checks file: {}", e),
                hint: None,
            })?;
        if data.len() > MAX_CHECKS_BYTES {
            return Err(CliError {
                code: EXIT_USAGE,
                message: format!(
                    "Checks file too large ({} KB, max {} KB)",
                    data.len() / 1024,
                    MAX_CHECKS_BYTES / 1024,
                ),
                hint: Some("reduce checks output or upload as a separate artifact".into()),
            });
        }
        let val: serde_json::Value = serde_json::from_str(&data)
            .map_err(|e| CliError {
                code: EXIT_USAGE,
                message: format!("Invalid checks JSON: {}", e),
                hint: None,
            })?;
        Some(val)
    } else {
        None
    };

    // 5. Load optional notes
    let notes_text: Option<String> = if let Some(ref notes_path) = notes {
        let text = std::fs::read_to_string(notes_path)
            .map_err(|e| CliError {
                code: EXIT_USAGE,
                message: format!("Cannot read notes file: {}", e),
                hint: None,
            })?;
        Some(text)
    } else {
        None
    };

    // 6. Build message
    let message = message.unwrap_or_else(|| format!("Publish {}", filename));

    // 7. Compute stamp semantics
    //    stamped = verification record contains a fingerprint
    //    stamp_matches = that fingerprint equals the computed fingerprint
    let has_stamp = verification.as_ref()
        .map(|v| v.fingerprint.is_some())
        .unwrap_or(false);
    let stamp_matches = verification.as_ref()
        .and_then(|v| v.fingerprint.as_ref())
        .map(|fp| fp == &fingerprint)
        .unwrap_or(false);

    // 8. Build source_metadata JSON
    let timestamp = {
        let now = std::time::SystemTime::now();
        let dur = now.duration_since(std::time::UNIX_EPOCH).unwrap_or_default();
        let secs = dur.as_secs();
        let days = secs / 86400;
        let t = secs % 86400;
        let (y, m, d) = days_to_ymd(days);
        format!("{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z", y, m, d, t / 3600, (t % 3600) / 60, t % 60)
    };

    let mut trust_pipeline = serde_json::Map::new();

    if let Some(ref v) = verification {
        if let Some(ref fp) = v.fingerprint {
            let mut stamp = serde_json::Map::new();
            stamp.insert("expected_fingerprint".into(), serde_json::json!(fp));
            if let Some(ref label) = v.label {
                stamp.insert("label".into(), serde_json::json!(label));
            }
            if let Some(ref ts) = v.timestamp {
                stamp.insert("timestamp".into(), serde_json::json!(ts));
            }
            trust_pipeline.insert("stamp".into(), serde_json::Value::Object(stamp));
        }
    }

    if let Some(ref cj) = checks_json {
        trust_pipeline.insert("checks".into(), cj.clone());
    }

    if let Some(ref nt) = notes_text {
        trust_pipeline.insert("notes".into(), serde_json::json!(nt));
    }

    let mut source_metadata = serde_json::json!({
        "type": "trust_pipeline",
        "fingerprint": fingerprint,
        "timestamp": timestamp,
    });
    if !trust_pipeline.is_empty() {
        source_metadata["trust_pipeline"] = serde_json::Value::Object(trust_pipeline);
    }
    if let Some(runner) = crate::ci::get_runner_context() {
        source_metadata["runner"] = runner;
    }

    // 9. Dry-run exit
    if dry_run {
        if json_output {
            let out = serde_json::json!({
                "schema_version": 1,
                "dry_run": true,
                "repo": repo,
                "file": file.display().to_string(),
                "fingerprint": fingerprint,
                "stamped": has_stamp,
                "stamp_matches": stamp_matches,
                "checks_attached": checks_json.is_some(),
                "notes_attached": notes_text.is_some(),
                "message": message,
                "source_metadata": source_metadata,
            });
            println!("{}", serde_json::to_string_pretty(&out).unwrap());
        } else {
            eprintln!("Dry run — no upload");
            eprintln!("  File:        {}", file.display());
            eprintln!("  Repo:        {}", repo);
            eprintln!("  Fingerprint: {}", fingerprint);
            eprintln!("  Stamped:     {}", if has_stamp { "yes" } else { "no" });
            if has_stamp {
                eprintln!("  Stamp match: {}", if stamp_matches { "yes" } else { "NO — fingerprint drifted" });
            }
            eprintln!("  Checks:      {}", if checks_json.is_some() { "attached" } else { "none" });
            eprintln!("  Notes:       {}", if notes_text.is_some() { "attached" } else { "none" });
            eprintln!("  Message:     {}", message);
        }
        return Ok(());
    }

    // ── Phase B: Network operations (requires auth) ────────────────

    // 10. Authenticate
    let client = HubClient::from_saved_auth().map_err(|e| match e {
        HubError::NotAuthenticated => CliError {
            code: EXIT_HUB_NOT_AUTH,
            message: "Not authenticated".into(),
            hint: Some("run `vgrid login` first".into()),
        },
        other => hub_error(other),
    })?;

    // 11. Lock warning
    if lock {
        eprintln!("warning: snapshot locking is not yet available");
    }

    // 12. Hash file
    if !json_output { eprint!("Hashing... "); }
    let byte_size = std::fs::metadata(&file)
        .map_err(|e| CliError { code: EXIT_ERROR, message: e.to_string(), hint: None })?
        .len();
    let content_hash = hash_file(&file).map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("{} ({} bytes)", &content_hash[..15], byte_size); }

    // 13. Resolve dataset by repo
    if !json_output { eprint!("Finding dataset '{}'... ", slug); }
    let datasets = client.list_datasets(owner, slug).map_err(|e| match &e {
        HubError::Http(404, _) => CliError {
            code: EXIT_HUB_NETWORK,
            message: format!("Repository '{}/{}' not found", owner, slug),
            hint: Some("check the --repo value or create the repository on VisiHub first".into()),
        },
        HubError::Http(403, _) => CliError {
            code: EXIT_HUB_NOT_AUTH,
            message: format!("No permission to access '{}/{}'", owner, slug),
            hint: Some("check your API token permissions or ask the repo owner for access".into()),
        },
        _ => hub_error(e),
    })?;
    let dataset_id = if let Some(d) = datasets.iter().find(|d| d.name == slug) {
        if !json_output { eprintln!("found"); }
        d.id.clone()
    } else {
        if !json_output { eprint!("creating... "); }
        let id = client.create_dataset(owner, slug, slug, Some("sheet")).map_err(|e| match &e {
            HubError::Http(403, _) | HubError::Http(422, _) => CliError {
                code: EXIT_HUB_NOT_AUTH,
                message: format!("Cannot create dataset in '{}/{}'", owner, slug),
                hint: Some("repo not found or you don't have permission to create datasets — check your API token".into()),
            },
            _ => hub_error(e),
        })?;
        if !json_output { eprintln!("created #{}", id); }
        id
    };

    // 14. Idempotency check — only match when latest is trust_pipeline WITH matching fingerprint.
    //     Missing source_metadata, different type, or different fingerprint all proceed to publish.
    let app_base = "https://app.visihub.app";
    let dataset_url = format!("{}/{}/{}", app_base, owner, slug);

    let status = client.get_dataset_status(&dataset_id).map_err(|e| hub_error(e))?;
    if let Some(ref sm) = status.source_metadata {
        if sm["type"].as_str() == Some("trust_pipeline")
            && sm["fingerprint"].as_str() == Some(&fingerprint)
        {
            let revision_url = status.current_revision_id.as_ref()
                .map(|id| format!("{}/{}/{}/revisions/{}", app_base, owner, slug, id));
            if json_output {
                let out = serde_json::json!({
                    "schema_version": 1,
                    "ok": true,
                    "already_published": true,
                    "repo": repo,
                    "fingerprint": fingerprint,
                    "stamped": has_stamp,
                    "stamp_matches": stamp_matches,
                    "revision_id": status.current_revision_id,
                    "dataset_url": dataset_url,
                    "revision_url": revision_url,
                });
                println!("{}", serde_json::to_string(&out).unwrap());
            } else {
                eprintln!("Already published — fingerprint matches current revision");
                eprintln!("  Repo:        {}", repo);
                eprintln!("  Fingerprint: {}", fingerprint);
                if let Some(ref id) = status.current_revision_id {
                    eprintln!("  Revision:    #{}", id);
                }
                eprintln!("  Dataset:     {}", dataset_url);
            }
            return Ok(());
        }
    }

    // 15. Create revision
    if !json_output { eprint!("Creating revision... "); }
    let opts = CreateRevisionOptions {
        format: Some("sheet".into()),
        source_metadata: Some(source_metadata),
        message: Some(message.clone()),
        ..Default::default()
    };
    let (revision_id, upload_url, upload_headers) = client
        .create_revision(&dataset_id, &content_hash, byte_size, &opts)
        .map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("#{}", revision_id); }

    // 16. Upload
    if !json_output { eprint!("Uploading {} bytes... ", byte_size); }
    let data = std::fs::read(&file)
        .map_err(|e| CliError { code: EXIT_ERROR, message: e.to_string(), hint: None })?;
    client.upload_bytes(&upload_url, data, &upload_headers).map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("done"); }

    // 17. Complete
    if !json_output { eprint!("Finalizing... "); }
    client.complete_revision(&revision_id, &content_hash).map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("done"); }

    let revision_url = format!("{}/{}/{}/revisions/{}", app_base, owner, slug, revision_id);

    // 18. No-wait exit — return after complete, skip polling
    if no_wait {
        if json_output {
            let out = serde_json::json!({
                "schema_version": 1,
                "ok": true,
                "status": "submitted",
                "repo": repo,
                "revision_id": revision_id,
                "fingerprint": fingerprint,
                "stamped": has_stamp,
                "stamp_matches": stamp_matches,
                "locked": false,
                "dataset_url": dataset_url,
                "revision_url": revision_url,
            });
            println!("{}", serde_json::to_string(&out).unwrap());
        } else {
            eprintln!("Submitted (not waiting for processing)");
            eprintln!("  Repo:        {}", repo);
            eprintln!("  Revision:    #{}", revision_id);
            eprintln!("  Fingerprint: {}", fingerprint);
            eprintln!("  Dataset:     {}", dataset_url);
            eprintln!("  Revision:    {}", revision_url);
        }
        return Ok(());
    }

    // 19. Poll
    if !json_output { eprint!("Waiting for processing... "); }
    let result = client
        .poll_run(owner, slug, &revision_id, Duration::from_secs(timeout))
        .map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("{}", result.status); }

    // 20. Output
    if json_output {
        let out = serde_json::json!({
            "schema_version": 1,
            "ok": true,
            "repo": repo,
            "revision_id": revision_id,
            "version": result.version,
            "fingerprint": fingerprint,
            "stamped": has_stamp,
            "stamp_matches": stamp_matches,
            "locked": false,
            "dataset_url": dataset_url,
            "revision_url": revision_url,
        });
        println!("{}", serde_json::to_string(&out).unwrap());
    } else {
        eprintln!();
        eprintln!("  Repo:        {}", repo);
        eprintln!("  Revision:    #{}", revision_id);
        eprintln!("  Version:     v{}", result.version);
        eprintln!("  Fingerprint: {}", fingerprint);
        eprintln!("  Stamped:     {}", if has_stamp { "yes" } else { "no" });
        if has_stamp {
            eprintln!("  Stamp match: {}", if stamp_matches { "yes" } else { "NO — fingerprint drifted" });
        }
        eprintln!("  Locked:      no");
        eprintln!("  Dataset:     {}", dataset_url);
        eprintln!("  Revision:    {}", revision_url);
        eprintln!();
    }

    Ok(())
}

// ── Pipeline Publish (import → verify → hub publish) ─────────────

pub fn cmd_pipeline_publish(
    source: PathBuf,
    repo: String,
    headers: bool,
    formulas: FormulaPolicy,
    stamp: Option<String>,
    checks_calc: Vec<String>,
    checks_file: Option<PathBuf>,
    delimiter: Option<String>,
    sheet_arg: Option<String>,
    message: Option<String>,
    notes: Option<PathBuf>,
    out: Option<PathBuf>,
    json_output: bool,
    dry_run: bool,
    no_wait: bool,
    timeout: u64,
) -> Result<(), CliError> {
    use std::collections::BTreeMap;
    use visigrid_io::native::{
        compute_semantic_fingerprint, save_workbook, save_workbook_with_metadata,
        save_semantic_verification, CellMetadata, SemanticVerification,
    };

    // ── 1. Validate inputs ────────────────────────────────────────────

    if !source.exists() {
        return Err(CliError {
            code: EXIT_USAGE,
            message: format!("File not found: {}", source.display()),
            hint: None,
        });
    }

    let ext = source.extension().and_then(|e| e.to_str()).unwrap_or("");
    if ext == "sheet" {
        return Err(CliError {
            code: EXIT_USAGE,
            message: "Source is already .sheet format".into(),
            hint: Some("use `vgrid hub publish` for .sheet files".into()),
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

    let is_csv = ext == "csv" || ext == "";
    let is_tsv = ext == "tsv";
    let is_xlsx = ext == "xlsx";

    if !is_csv && !is_tsv && !is_xlsx {
        return Err(CliError {
            code: EXIT_USAGE,
            message: format!("Unsupported source format: .{}", ext),
            hint: Some("supported: .csv, .tsv, .xlsx".into()),
        });
    }

    if sheet_arg.is_some() && !is_xlsx {
        return Err(CliError::args("--sheet is only valid for XLSX sources"));
    }
    if !matches!(formulas, FormulaPolicy::Values) && !is_xlsx {
        return Err(CliError::args("--formulas keep/recalc only valid for XLSX sources"));
    }
    if delimiter.is_some() && !is_csv {
        return Err(CliError::args("--delimiter is only valid for CSV sources"));
    }

    // Load optional checks file
    let checks_file_json: Option<serde_json::Value> = if let Some(ref path) = checks_file {
        let data = std::fs::read_to_string(path)
            .map_err(|e| CliError {
                code: EXIT_USAGE,
                message: format!("Cannot read checks file: {}", e),
                hint: None,
            })?;
        if data.len() > MAX_CHECKS_BYTES {
            return Err(CliError {
                code: EXIT_USAGE,
                message: format!(
                    "Checks file too large ({} KB, max {} KB)",
                    data.len() / 1024, MAX_CHECKS_BYTES / 1024,
                ),
                hint: Some("reduce checks output or upload as a separate artifact".into()),
            });
        }
        Some(serde_json::from_str(&data).map_err(|e| CliError {
            code: EXIT_USAGE,
            message: format!("Invalid checks JSON: {}", e),
            hint: None,
        })?)
    } else {
        None
    };

    // Load optional notes
    let notes_text: Option<String> = if let Some(ref path) = notes {
        Some(std::fs::read_to_string(path).map_err(|e| CliError {
            code: EXIT_USAGE,
            message: format!("Cannot read notes file: {}", e),
            hint: None,
        })?)
    } else {
        None
    };

    if !json_output { eprintln!("Pipeline: import → verify → publish"); }

    // ── 2. Load source ────────────────────────────────────────────────

    if !json_output { eprint!("  Loading {}... ", source.display()); }

    let format_str: &str;
    let (mut workbook, import_result) = match ext {
        "xlsx" => {
            format_str = "xlsx";
            let values_only = !matches!(formulas, FormulaPolicy::Recalc);
            let opts = visigrid_io::xlsx::ImportOptions { values_only, ..Default::default() };
            visigrid_io::xlsx::import_with_options(&source, &opts)
                .map_err(|e| CliError::io(format!("failed to load {}: {}", source.display(), e)))?
        }
        "tsv" => {
            format_str = "tsv";
            let sheet = visigrid_io::csv::import_tsv(&source)
                .map_err(|e| CliError::parse(e))?;
            let wb = visigrid_engine::workbook::Workbook::from_sheets(vec![sheet], 0);
            (wb, visigrid_io::xlsx::ImportResult::default())
        }
        _ => {
            format_str = "csv";
            let sheet = if let Some(ref d) = delimiter {
                let delim = util::parse_delimiter(d)?;
                visigrid_io::csv::import_with_delimiter(&source, delim)
                    .map_err(|e| CliError::parse(e))?
            } else {
                visigrid_io::csv::import(&source)
                    .map_err(|e| CliError::parse(e))?
            };
            let wb = visigrid_engine::workbook::Workbook::from_sheets(vec![sheet], 0);
            (wb, visigrid_io::xlsx::ImportResult::default())
        }
    };

    // Select sheet (xlsx only)
    let selected_sheet_idx: usize;
    if let Some(ref arg) = sheet_arg {
        let idx = sheet_ops::resolve_sheet_by_arg(&workbook, arg)?;
        selected_sheet_idx = idx;
        let extracted = workbook.sheet(idx)
            .ok_or_else(|| CliError::io("sheet not found"))?
            .clone();
        workbook = visigrid_engine::workbook::Workbook::from_sheets(vec![extracted], 0);
    } else {
        selected_sheet_idx = 0;
    }

    if !json_output { eprintln!("done ({})", format_str); }

    // ── 3. Run checks-calc (optional) ─────────────────────────────────

    let computed_checks: Option<serde_json::Value> = if !checks_calc.is_empty() {
        if !json_output { eprint!("  Evaluating {} check(s)... ", checks_calc.len()); }

        let sheet = workbook.sheet(0)
            .ok_or_else(|| CliError::io("no sheets in workbook"))?;
        let sheet_id = workbook.sheet_id_at_idx(0)
            .ok_or_else(|| CliError::io("cannot resolve sheet ID"))?;
        let (max_row, max_col) = sheet_ops::get_data_bounds(sheet);

        let start_row1 = if headers { 2 } else { 1 };
        let end_row1 = if max_row < start_row1 { start_row1 } else { max_row };

        // Build header map for column-name resolution
        let header_map: HashMap<String, String> = if headers {
            let mut map = HashMap::new();
            for col_idx in 0..max_col {
                let val = sheet.get_display(0, col_idx);
                if !val.is_empty() {
                    let key = val.trim().to_ascii_lowercase();
                    let col_letter = util::col_to_letter(col_idx);
                    let col_ref = format!("{}:{}", col_letter, col_letter);
                    map.insert(key, col_ref);
                }
            }
            map
        } else {
            HashMap::new()
        };

        let lookup = visigrid_engine::workbook::WorkbookLookup::new(&workbook, sheet_id);
        let mut results: Vec<CalcResult> = Vec::new();
        let mut any_error = false;

        for expr_str in &checks_calc {
            let with_eq = if expr_str.starts_with('=') {
                expr_str.clone()
            } else {
                format!("={}", expr_str)
            };
            let resolved = sheet_ops::resolve_header_refs(&with_eq, &header_map);
            let formula_str = sheet_ops::translate_column_refs(&resolved, start_row1, end_row1);

            let result = match visigrid_engine::formula::parser::parse(&formula_str) {
                Ok(parsed) => {
                    let bound = visigrid_engine::formula::parser::bind_expr_same_sheet(&parsed);
                    let eval = visigrid_engine::formula::eval::evaluate(&bound, &lookup);
                    let display = eval.to_text();
                    let is_error = matches!(eval, visigrid_engine::formula::eval::EvalResult::Error(_));
                    if is_error { any_error = true; }
                    let value_type = match &eval {
                        visigrid_engine::formula::eval::EvalResult::Number(_) => "number",
                        visigrid_engine::formula::eval::EvalResult::Text(_) => "text",
                        visigrid_engine::formula::eval::EvalResult::Boolean(_) => "boolean",
                        visigrid_engine::formula::eval::EvalResult::Error(_) => "error",
                        visigrid_engine::formula::eval::EvalResult::Empty => "empty",
                        visigrid_engine::formula::eval::EvalResult::Array(_) => "array",
                    };
                    CalcResult {
                        expr: expr_str.clone(),
                        value: display.clone(),
                        value_type: value_type.to_string(),
                        error: if is_error { Some(display) } else { None },
                    }
                }
                Err(e) => {
                    any_error = true;
                    CalcResult {
                        expr: expr_str.clone(),
                        value: format!("#PARSE: {}", e),
                        value_type: "error".to_string(),
                        error: Some(e.to_string()),
                    }
                }
            };
            results.push(result);
        }

        if any_error {
            return Err(CliError {
                code: EXIT_ERROR,
                message: "check formula evaluation failed".into(),
                hint: Some("fix the --checks-calc expressions and retry".into()),
            });
        }

        if !json_output { eprintln!("pass"); }

        let calc_output = CalcOutput {
            format: format_str.to_string(),
            sheet: workbook.sheet(0).map(|s| s.name.clone()).unwrap_or_default(),
            results,
        };
        Some(serde_json::to_value(&calc_output).unwrap())
    } else {
        None
    };

    // Merge checks: --checks-file takes precedence, --checks-calc is fallback
    let final_checks = checks_file_json.or(computed_checks);

    // ── 4. Build cell metadata (--formulas keep) ──────────────────────

    let metadata: CellMetadata = if matches!(formulas, FormulaPolicy::Keep) {
        import_result.formula_strings.iter()
            .filter(|((si, _, _), _)| *si == selected_sheet_idx)
            .map(|((_, r, c), f)| {
                let ref_str = sheet_ops::format_cell_ref(*r, *c);
                let mut map = BTreeMap::new();
                map.insert("formula".to_string(), f.clone());
                (ref_str, map)
            })
            .collect()
    } else {
        BTreeMap::new()
    };

    // ── 5. Compute fingerprint ────────────────────────────────────────

    let fingerprint = compute_semantic_fingerprint(&workbook);

    // ── 6. Write .sheet file ──────────────────────────────────────────

    let sheet_path = out.clone().unwrap_or_else(|| {
        std::env::temp_dir().join(format!("vgrid_pipeline_{}.sheet", std::process::id()))
    });
    let is_temp = out.is_none();

    if !json_output { eprint!("  Writing {}... ", sheet_path.display()); }

    let temp_path = sheet_path.with_extension("sheet.tmp");

    if metadata.is_empty() {
        save_workbook(&workbook, &temp_path)
            .map_err(|e| CliError::io(format!("failed to write: {}", e)))?;
    } else {
        save_workbook_with_metadata(&workbook, &metadata, &temp_path)
            .map_err(|e| CliError::io(format!("failed to write: {}", e)))?;
    }

    // Stamp (always stamp when --stamp is provided, default to repo slug if no explicit label)
    let has_stamp;
    let stamp_matches;
    if let Some(ref label) = stamp {
        let verification = SemanticVerification {
            fingerprint: Some(fingerprint.clone()),
            label: if label.is_empty() { None } else { Some(label.clone()) },
            timestamp: Some(chrono::Utc::now().to_rfc3339()),
        };
        save_semantic_verification(&temp_path, &verification)
            .map_err(|e| CliError::io(format!("failed to write stamp: {}", e)))?;
        has_stamp = true;
        stamp_matches = true; // Freshly stamped always matches
    } else {
        has_stamp = false;
        stamp_matches = false;
    }

    std::fs::rename(&temp_path, &sheet_path)
        .map_err(|e| CliError::io(format!("failed to rename: {}", e)))?;

    if !json_output { eprintln!("done"); }

    // ── 7. Verify (always) ────────────────────────────────────────────

    if !json_output { eprint!("  Verifying fingerprint... "); }

    // Re-load and verify fingerprint matches
    let verify_wb = visigrid_io::native::load_workbook(&sheet_path)
        .map_err(|e| CliError::io(format!("failed to reload .sheet: {}", e)))?;
    let verify_fp = compute_semantic_fingerprint(&verify_wb);

    if verify_fp != fingerprint {
        // Clean up temp file
        if is_temp { std::fs::remove_file(&sheet_path).ok(); }
        return Err(CliError {
            code: EXIT_ERROR,
            message: format!("Verification failed: fingerprint drifted ({} != {})", verify_fp, fingerprint),
            hint: None,
        });
    }
    if !json_output { eprintln!("{}", fingerprint); }

    // ── 8. Compute stats for output ───────────────────────────────────

    let sheet_data = workbook.sheet(0)
        .ok_or_else(|| CliError::io("no sheets"))?;
    let (rows, cols) = sheet_ops::get_data_bounds(sheet_data);
    let source_filename = source.file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("data")
        .to_string();

    let msg = message.unwrap_or_else(|| format!("Publish {}", source_filename));

    // ── 9. Build source_metadata ──────────────────────────────────────

    let timestamp = chrono::Utc::now().to_rfc3339();

    let mut trust_pipeline = serde_json::Map::new();

    if has_stamp {
        let mut stamp_obj = serde_json::Map::new();
        stamp_obj.insert("expected_fingerprint".into(), serde_json::json!(&fingerprint));
        if let Some(ref label) = stamp {
            if !label.is_empty() {
                stamp_obj.insert("label".into(), serde_json::json!(label));
            }
        }
        stamp_obj.insert("timestamp".into(), serde_json::json!(&timestamp));
        trust_pipeline.insert("stamp".into(), serde_json::Value::Object(stamp_obj));
    }

    if let Some(ref cj) = final_checks {
        trust_pipeline.insert("checks".into(), cj.clone());
    }

    if let Some(ref nt) = notes_text {
        trust_pipeline.insert("notes".into(), serde_json::json!(nt));
    }

    let mut source_metadata = serde_json::json!({
        "type": "trust_pipeline",
        "fingerprint": fingerprint,
        "timestamp": timestamp,
    });
    if !trust_pipeline.is_empty() {
        source_metadata["trust_pipeline"] = serde_json::Value::Object(trust_pipeline);
    }
    if let Some(runner) = crate::ci::get_runner_context() {
        source_metadata["runner"] = runner;
    }

    // ── 10. Dry-run exit ──────────────────────────────────────────────

    if dry_run {
        if json_output {
            let out = serde_json::json!({
                "schema_version": 1,
                "dry_run": true,
                "repo": repo,
                "source": source.display().to_string(),
                "format": format_str,
                "rows": rows,
                "cols": cols,
                "fingerprint": fingerprint,
                "stamped": has_stamp,
                "stamp_matches": stamp_matches,
                "checks_attached": final_checks.is_some(),
                "notes_attached": notes_text.is_some(),
                "message": msg,
                "source_metadata": source_metadata,
                "sheet_path": if is_temp { serde_json::Value::Null } else { serde_json::json!(sheet_path.display().to_string()) },
            });
            println!("{}", serde_json::to_string_pretty(&out).unwrap());
        } else {
            eprintln!("Dry run — no upload");
            eprintln!("  Source:      {}", source.display());
            eprintln!("  Format:      {}", format_str);
            eprintln!("  Rows:        {}", rows);
            eprintln!("  Cols:        {}", cols);
            eprintln!("  Fingerprint: {}", fingerprint);
            eprintln!("  Stamped:     {}", if has_stamp { "yes" } else { "no" });
            eprintln!("  Checks:      {}", if final_checks.is_some() { "attached" } else { "none" });
            eprintln!("  Message:     {}", msg);
            if !is_temp {
                eprintln!("  Sheet:       {}", sheet_path.display());
            }
        }
        if is_temp { std::fs::remove_file(&sheet_path).ok(); }
        return Ok(());
    }

    // ── 11. Hub publish (requires auth) ───────────────────────────────

    let client = HubClient::from_saved_auth().map_err(|e| match e {
        HubError::NotAuthenticated => CliError {
            code: EXIT_HUB_NOT_AUTH,
            message: "Not authenticated".into(),
            hint: Some("run `vgrid login` first".into()),
        },
        other => hub_error(other),
    })?;

    // Hash the .sheet file
    if !json_output { eprint!("  Hashing... "); }
    let byte_size = std::fs::metadata(&sheet_path)
        .map_err(|e| CliError { code: EXIT_ERROR, message: e.to_string(), hint: None })?
        .len();
    let content_hash = hash_file(&sheet_path).map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("{} ({} bytes)", &content_hash[..15], byte_size); }

    // Resolve dataset
    if !json_output { eprint!("  Finding dataset '{}'... ", slug); }
    let datasets = client.list_datasets(owner, slug).map_err(|e| match &e {
        HubError::Http(404, _) => CliError {
            code: EXIT_HUB_NETWORK,
            message: format!("Repository '{}/{}' not found", owner, slug),
            hint: Some("check the --repo value or create the repository on VisiHub first".into()),
        },
        HubError::Http(403, _) => CliError {
            code: EXIT_HUB_NOT_AUTH,
            message: format!("No permission to access '{}/{}'", owner, slug),
            hint: Some("check your API token permissions or ask the repo owner for access".into()),
        },
        _ => hub_error(e),
    })?;

    let dataset_id = if let Some(d) = datasets.iter().find(|d| d.name == slug) {
        if !json_output { eprintln!("found"); }
        d.id.clone()
    } else {
        if !json_output { eprint!("creating... "); }
        let id = client.create_dataset(owner, slug, slug, Some("sheet")).map_err(|e| match &e {
            HubError::Http(403, _) | HubError::Http(422, _) => CliError {
                code: EXIT_HUB_NOT_AUTH,
                message: format!("Cannot create dataset in '{}/{}'", owner, slug),
                hint: Some("repo not found or you don't have permission to create datasets — check your API token".into()),
            },
            _ => hub_error(e),
        })?;
        if !json_output { eprintln!("created #{}", id); }
        id
    };

    // Idempotency check
    let app_base = "https://app.visihub.app";
    let dataset_url = format!("{}/{}/{}", app_base, owner, slug);

    let status = client.get_dataset_status(&dataset_id).map_err(|e| hub_error(e))?;
    if let Some(ref sm) = status.source_metadata {
        if sm["type"].as_str() == Some("trust_pipeline")
            && sm["fingerprint"].as_str() == Some(&fingerprint)
        {
            if is_temp { std::fs::remove_file(&sheet_path).ok(); }
            let revision_url = status.current_revision_id.as_ref()
                .map(|id| format!("{}/{}/{}/revisions/{}", app_base, owner, slug, id));
            if json_output {
                let out = serde_json::json!({
                    "schema_version": 1,
                    "ok": true,
                    "already_published": true,
                    "repo": repo,
                    "fingerprint": fingerprint,
                    "stamped": has_stamp,
                    "stamp_matches": stamp_matches,
                    "revision_id": status.current_revision_id,
                    "dataset_url": dataset_url,
                    "revision_url": revision_url,
                });
                println!("{}", serde_json::to_string(&out).unwrap());
            } else {
                eprintln!("  Already published — fingerprint matches current revision");
                eprintln!("    Fingerprint: {}", fingerprint);
                eprintln!("    Dataset:     {}", dataset_url);
            }
            return Ok(());
        }
    }

    // Create revision
    if !json_output { eprint!("  Creating revision... "); }
    let opts = CreateRevisionOptions {
        format: Some("sheet".into()),
        source_metadata: Some(source_metadata),
        message: Some(msg.clone()),
        ..Default::default()
    };
    let (revision_id, upload_url, upload_headers) = client
        .create_revision(&dataset_id, &content_hash, byte_size, &opts)
        .map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("#{}", revision_id); }

    // Upload
    if !json_output { eprint!("  Uploading {} bytes... ", byte_size); }
    let data = std::fs::read(&sheet_path)
        .map_err(|e| CliError { code: EXIT_ERROR, message: e.to_string(), hint: None })?;
    client.upload_bytes(&upload_url, data, &upload_headers).map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("done"); }

    // Clean up temp file after upload
    if is_temp { std::fs::remove_file(&sheet_path).ok(); }

    // Complete
    if !json_output { eprint!("  Finalizing... "); }
    client.complete_revision(&revision_id, &content_hash).map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("done"); }

    let revision_url = format!("{}/{}/{}/revisions/{}", app_base, owner, slug, revision_id);

    // No-wait exit
    if no_wait {
        if json_output {
            let out = serde_json::json!({
                "schema_version": 1,
                "ok": true,
                "status": "submitted",
                "repo": repo,
                "revision_id": revision_id,
                "fingerprint": fingerprint,
                "stamped": has_stamp,
                "stamp_matches": stamp_matches,
                "locked": false,
                "dataset_url": dataset_url,
                "revision_url": revision_url,
            });
            println!("{}", serde_json::to_string(&out).unwrap());
        } else {
            eprintln!("  Submitted (not waiting for processing)");
            eprintln!("    Revision:    #{}", revision_id);
            eprintln!("    Dataset:     {}", dataset_url);
        }
        return Ok(());
    }

    // Poll
    if !json_output { eprint!("  Waiting for processing... "); }
    let result = client
        .poll_run(owner, slug, &revision_id, Duration::from_secs(timeout))
        .map_err(|e| hub_error(e))?;
    if !json_output { eprintln!("{}", result.status); }

    // Output
    if json_output {
        let out = serde_json::json!({
            "schema_version": 1,
            "ok": true,
            "repo": repo,
            "source": source.display().to_string(),
            "format": format_str,
            "rows": rows,
            "cols": cols,
            "revision_id": revision_id,
            "version": result.version,
            "fingerprint": fingerprint,
            "stamped": has_stamp,
            "stamp_matches": stamp_matches,
            "locked": false,
            "dataset_url": dataset_url,
            "revision_url": revision_url,
        });
        println!("{}", serde_json::to_string(&out).unwrap());
    } else {
        eprintln!();
        eprintln!("  Repo:        {}", repo);
        eprintln!("  Source:      {}", source.display());
        eprintln!("  Revision:    #{}", revision_id);
        eprintln!("  Version:     v{}", result.version);
        eprintln!("  Fingerprint: {}", fingerprint);
        eprintln!("  Stamped:     {}", if has_stamp { "yes" } else { "no" });
        eprintln!("  Dataset:     {}", dataset_url);
        eprintln!("  Revision:    {}", revision_url);
        eprintln!();
    }

    Ok(())
}

/// Convert days since epoch to (year, month, day).
fn days_to_ymd(mut days: u64) -> (u64, u64, u64) {
    let mut year = 1970u64;
    loop {
        let dy = if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 { 366 } else { 365 };
        if days < dy { break; }
        days -= dy;
        year += 1;
    }
    let leap = (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
    let md: &[u64] = if leap {
        &[31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        &[31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };
    let mut month = 1u64;
    for &m in md {
        if days < m { break; }
        days -= m;
        month += 1;
    }
    (year, month, days + 1)
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
