//! `vgrid recon` — config-driven multi-source reconciliation.

use std::path::PathBuf;

use clap::Subcommand;

use crate::exit_codes::{EXIT_RECON_INVALID_CONFIG, EXIT_RECON_MISMATCH, EXIT_RECON_RUNTIME, EXIT_RECON_STALE};
use crate::CliError;

#[derive(Subcommand)]
pub enum ReconCommands {
    /// Run reconciliation from a TOML config file
    #[command(after_help = "\
Examples:
  vgrid recon run recon.toml
  vgrid recon run recon.toml --json
  vgrid recon run recon.toml --output result.json
  vgrid recon run daily-close.composite.toml")]
    Run {
        /// Path to the .recon.toml or .composite.toml config file
        config: PathBuf,

        /// Output JSON to stdout instead of human summary
        #[arg(long)]
        json: bool,

        /// Write JSON output to file
        #[arg(long)]
        output: Option<PathBuf>,

        /// Stop on first step failure (composite configs only)
        #[arg(long)]
        fail_fast: bool,
    },

    /// Validate a recon config without running
    #[command(after_help = "\
Examples:
  vgrid recon validate recon.toml")]
    Validate {
        /// Path to the .recon.toml config file
        config: PathBuf,
    },
}

pub fn cmd_recon(cmd: ReconCommands) -> Result<(), CliError> {
    match cmd {
        ReconCommands::Run { config, json, output, fail_fast } => {
            cmd_recon_run(config, json, output, fail_fast)
        }
        ReconCommands::Validate { config } => cmd_recon_validate(config),
    }
}

fn recon_err(code: u8, msg: impl Into<String>) -> CliError {
    CliError { code, message: msg.into(), hint: None }
}

/// Extract the `kind` field from a TOML string, defaulting to "recon".
fn extract_kind(config_str: &str) -> String {
    #[derive(serde::Deserialize)]
    struct KindProbe {
        #[serde(default = "default_kind")]
        kind: String,
    }
    fn default_kind() -> String {
        "recon".into()
    }

    toml::from_str::<KindProbe>(config_str)
        .map(|p| p.kind)
        .unwrap_or_else(|_| "recon".into())
}

fn cmd_recon_run(
    config_path: PathBuf,
    json_output: bool,
    output_file: Option<PathBuf>,
    fail_fast: bool,
) -> Result<(), CliError> {
    // Read config string
    let config_str = std::fs::read_to_string(&config_path)
        .map_err(|e| recon_err(EXIT_RECON_RUNTIME, format!("cannot read config: {e}")))?;

    let kind = extract_kind(&config_str);

    match kind.as_str() {
        "composite" => cmd_recon_run_composite(config_path, &config_str, json_output, output_file, fail_fast),
        "recon" => cmd_recon_run_single(config_path, &config_str, json_output, output_file),
        other => Err(recon_err(
            EXIT_RECON_INVALID_CONFIG,
            format!("unknown config kind: \"{other}\" (expected \"recon\" or \"composite\")"),
        )),
    }
}

fn cmd_recon_run_single(
    config_path: PathBuf,
    config_str: &str,
    json_output: bool,
    output_file: Option<PathBuf>,
) -> Result<(), CliError> {
    use std::collections::HashMap;
    use visigrid_recon::engine::load_csv_rows;

    let config = visigrid_recon::ReconConfig::from_toml(config_str)
        .map_err(|e| recon_err(EXIT_RECON_INVALID_CONFIG, e.to_string()))?;

    // Resolve file paths relative to config file's directory
    let base_dir = config_path.parent().unwrap_or_else(|| std::path::Path::new("."));

    // Load CSV data for each role
    let mut records: HashMap<String, Vec<visigrid_recon::ReconRow>> = HashMap::new();
    for (role_name, role_config) in &config.roles {
        let csv_path = base_dir.join(&role_config.file);
        let csv_data = std::fs::read_to_string(&csv_path).map_err(|e| {
            recon_err(EXIT_RECON_RUNTIME, format!("cannot read {}: {e}", csv_path.display()))
        })?;
        let rows = load_csv_rows(role_name, &csv_data, role_config)
            .map_err(|e| recon_err(EXIT_RECON_RUNTIME, e.to_string()))?;
        records.insert(role_name.clone(), rows);
    }

    let input = visigrid_recon::ReconInput { records };

    // Run engine
    let result = visigrid_recon::run(&config, &input)
        .map_err(|e| recon_err(EXIT_RECON_RUNTIME, e.to_string()))?;

    // Output
    let json_str = serde_json::to_string_pretty(&result)
        .map_err(|e| recon_err(EXIT_RECON_RUNTIME, format!("JSON serialization error: {e}")))?;

    if let Some(ref path) = output_file {
        std::fs::write(path, &json_str)
            .map_err(|e| recon_err(EXIT_RECON_RUNTIME, format!("cannot write output: {e}")))?;
        eprintln!("wrote {}", path.display());
    }

    if json_output {
        println!("{json_str}");
    }

    // Human summary to stderr
    let s = &result.summary;
    eprintln!(
        "{}-way recon: {} groups — {} matched, {} amount mismatches, {} timing mismatches, {} unmatched",
        result.meta.way,
        s.total_groups,
        s.matched,
        s.amount_mismatches,
        s.timing_mismatches,
        s.left_only + s.right_only,
    );

    if let Some(ref settlement) = s.settlement {
        eprintln!(
            "settlement: {} matched, {} pending, {} stale, {} errors",
            settlement.matched, settlement.pending, settlement.stale, settlement.errors,
        );
    }

    // Settlement-aware exit codes when settlement config is present
    if let Some(ref settlement) = s.settlement {
        if settlement.errors > 0 {
            return Err(recon_err(EXIT_RECON_MISMATCH, "settlement errors found"));
        }
        if s.ambiguous > 0 && config.fail_on_ambiguous {
            return Err(recon_err(EXIT_RECON_MISMATCH, "ambiguous matches found (fail_on_ambiguous)"));
        }
        if settlement.stale > 0 {
            return Err(recon_err(EXIT_RECON_STALE, "stale items found"));
        }
        if s.ambiguous > 0 {
            return Err(recon_err(EXIT_RECON_STALE, "ambiguous matches found"));
        }
        // pending only → pass
        return Ok(());
    }

    // Fallback: original logic when no settlement config
    if s.amount_mismatches > 0 || s.timing_mismatches > 0 || s.left_only > 0 || s.right_only > 0 {
        return Err(recon_err(EXIT_RECON_MISMATCH, "mismatches found"));
    }

    if s.ambiguous > 0 {
        if config.fail_on_ambiguous {
            return Err(recon_err(EXIT_RECON_MISMATCH, "ambiguous matches found (fail_on_ambiguous)"));
        }
        return Err(recon_err(EXIT_RECON_STALE, "ambiguous matches found"));
    }

    Ok(())
}

fn cmd_recon_run_composite(
    config_path: PathBuf,
    config_str: &str,
    json_output: bool,
    output_file: Option<PathBuf>,
    fail_fast: bool,
) -> Result<(), CliError> {
    use std::collections::HashMap;
    use std::time::Instant;
    use visigrid_recon::engine::load_csv_rows;
    use visigrid_recon::{
        CompositeResult, CompositeVerdict, StepResult, StepStatus,
    };

    let composite = visigrid_recon::CompositeConfig::from_toml(config_str)
        .map_err(|e| recon_err(EXIT_RECON_INVALID_CONFIG, e.to_string()))?;

    let base_dir = config_path.parent().unwrap_or_else(|| std::path::Path::new("."));
    let mut steps: Vec<StepResult> = Vec::with_capacity(composite.steps.len());

    for step in &composite.steps {
        let step_config_path = base_dir.join(&step.config);
        let start = Instant::now();

        // Attempt to load and run the child recon
        // Returns (result, fail_on_ambiguous) so exit code respects per-config flag
        let step_result = (|| -> Result<(visigrid_recon::ReconResult, bool), String> {
            let child_str = std::fs::read_to_string(&step_config_path)
                .map_err(|e| format!("cannot read {}: {e}", step_config_path.display()))?;
            let child_config = visigrid_recon::ReconConfig::from_toml(&child_str)
                .map_err(|e| e.to_string())?;

            let fail_on_ambiguous = child_config.fail_on_ambiguous;

            let child_base = step_config_path
                .parent()
                .unwrap_or_else(|| std::path::Path::new("."));

            let mut records: HashMap<String, Vec<visigrid_recon::ReconRow>> = HashMap::new();
            for (role_name, role_config) in &child_config.roles {
                let csv_path = child_base.join(&role_config.file);
                let csv_data = std::fs::read_to_string(&csv_path)
                    .map_err(|e| format!("cannot read {}: {e}", csv_path.display()))?;
                let rows = load_csv_rows(role_name, &csv_data, role_config)
                    .map_err(|e| e.to_string())?;
                records.insert(role_name.clone(), rows);
            }

            let input = visigrid_recon::ReconInput { records };
            let result = visigrid_recon::run(&child_config, &input).map_err(|e| e.to_string())?;
            Ok((result, fail_on_ambiguous))
        })();

        let duration_ms = start.elapsed().as_millis() as u64;

        match step_result {
            Ok((result, fail_on_ambiguous)) => {
                let status = StepStatus::from_recon_result_with_options(&result, fail_on_ambiguous);

                eprintln!(
                    "  step '{}': {} ({}ms)",
                    step.name,
                    match status {
                        StepStatus::Pass => "pass",
                        StepStatus::Warn => "warn",
                        StepStatus::Fail => "fail",
                        StepStatus::Error => "error",
                    },
                    duration_ms,
                );

                let should_stop = fail_fast && matches!(status, StepStatus::Fail | StepStatus::Error);

                steps.push(StepResult {
                    name: step.name.clone(),
                    status,
                    duration_ms,
                    config_path: step.config.clone(),
                    result,
                });

                if should_stop {
                    eprintln!("  --fail-fast: stopping after '{}' failure", step.name);
                    break;
                }
            }
            Err(err_msg) => {
                eprintln!("  step '{}': error — {}", step.name, err_msg);

                // Build a minimal error result so we can still include it
                let error_result = visigrid_recon::ReconResult {
                    meta: visigrid_recon::model::ReconMeta {
                        config_name: step.name.clone(),
                        way: 0,
                        engine_version: env!("CARGO_PKG_VERSION").to_string(),
                        run_at: chrono::Utc::now().to_rfc3339(),
                        settlement_clock: None,
                    },
                    summary: visigrid_recon::model::ReconSummary {
                        total_groups: 0,
                        matched: 0,
                        amount_mismatches: 0,
                        timing_mismatches: 0,
                        ambiguous: 0,
                        left_only: 0,
                        right_only: 0,
                        bucket_counts: HashMap::new(),
                        settlement: None,
                    },
                    groups: vec![],
                    derived: visigrid_recon::DerivedOutputs::default(),
                };

                steps.push(StepResult {
                    name: step.name.clone(),
                    status: StepStatus::Error,
                    duration_ms,
                    config_path: step.config.clone(),
                    result: error_result,
                });

                if fail_fast {
                    eprintln!("  --fail-fast: stopping after '{}' error", step.name);
                    break;
                }
            }
        }
    }

    let verdict = CompositeVerdict::from_steps(&steps);
    let exit_code = verdict.exit_code();

    let composite_result = CompositeResult {
        name: composite.name.clone(),
        engine_version: env!("CARGO_PKG_VERSION").to_string(),
        run_at: chrono::Utc::now().to_rfc3339(),
        verdict,
        exit_code,
        steps,
    };

    // Output
    let json_str = serde_json::to_string_pretty(&composite_result)
        .map_err(|e| recon_err(EXIT_RECON_RUNTIME, format!("JSON serialization error: {e}")))?;

    if let Some(ref path) = output_file {
        std::fs::write(path, &json_str)
            .map_err(|e| recon_err(EXIT_RECON_RUNTIME, format!("cannot write output: {e}")))?;
        eprintln!("wrote {}", path.display());
    }

    if json_output {
        println!("{json_str}");
    }

    // Human summary to stderr
    eprintln!(
        "composite '{}': {} step(s), verdict: {}",
        composite.name,
        composite_result.steps.len(),
        match composite_result.verdict {
            CompositeVerdict::Pass => "pass",
            CompositeVerdict::Warn => "warn",
            CompositeVerdict::Fail => "fail",
        },
    );

    match exit_code {
        0 => Ok(()),
        code => Err(recon_err(code, format!("composite verdict: {}", match composite_result.verdict {
            CompositeVerdict::Fail => "fail",
            CompositeVerdict::Warn => "warn (stale items)",
            CompositeVerdict::Pass => unreachable!(),
        }))),
    }
}

fn cmd_recon_validate(config_path: PathBuf) -> Result<(), CliError> {
    let config_str = std::fs::read_to_string(&config_path)
        .map_err(|e| recon_err(EXIT_RECON_RUNTIME, format!("cannot read config: {e}")))?;

    let kind = extract_kind(&config_str);

    match kind.as_str() {
        "composite" => {
            match visigrid_recon::CompositeConfig::from_toml(&config_str) {
                Ok(config) => {
                    eprintln!(
                        "valid: composite '{}' with {} step(s)",
                        config.name,
                        config.steps.len(),
                    );
                    Ok(())
                }
                Err(e) => Err(recon_err(EXIT_RECON_INVALID_CONFIG, e.to_string())),
            }
        }
        _ => {
            match visigrid_recon::ReconConfig::from_toml(&config_str) {
                Ok(config) => {
                    eprintln!(
                        "valid: {}-way recon '{}' with {} role(s), {} pair(s)",
                        config.way,
                        config.name,
                        config.roles.len(),
                        config.pairs.len(),
                    );
                    Ok(())
                }
                Err(e) => Err(recon_err(EXIT_RECON_INVALID_CONFIG, e.to_string())),
            }
        }
    }
}
