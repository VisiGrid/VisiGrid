//! VisiHub CLI commands: login and publish.
//!
//! `vgrid login`   — store API token
//! `vgrid publish`  — upload file, wait for check, print results

use std::io::{self, Write};
use std::path::PathBuf;
use std::time::Duration;

use visigrid_hub_client::{
    AuthCredentials, save_auth,
    HubClient, HubError, CreateRevisionOptions, RunResult,
    hash_file,
};

use crate::{CliError, OutputFormat};
use crate::exit_codes::*;

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
        let id = client.create_dataset(owner, slug, &dataset_name).map_err(|e| hub_error(e))?;
        if !json_output { eprintln!("created #{}", id); }
        id
    };

    // Step 3: Create revision
    if !json_output { eprint!("Creating revision... "); }
    let opts = CreateRevisionOptions {
        source_type,
        source_identity,
        query_hash,
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
    if let Some(ref hash) = r.content_hash {
        eprintln!("  Hash:    {}", hash);
    }
    eprintln!("  Proof:   {}", r.proof_url);
    eprintln!();
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
