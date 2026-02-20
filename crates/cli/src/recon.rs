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
  vgrid recon run recon.toml --output result.json")]
    Run {
        /// Path to the .recon.toml config file
        config: PathBuf,

        /// Output JSON to stdout instead of human summary
        #[arg(long)]
        json: bool,

        /// Write JSON output to file
        #[arg(long)]
        output: Option<PathBuf>,
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
        ReconCommands::Run { config, json, output } => cmd_recon_run(config, json, output),
        ReconCommands::Validate { config } => cmd_recon_validate(config),
    }
}

fn recon_err(code: u8, msg: impl Into<String>) -> CliError {
    CliError { code, message: msg.into(), hint: None }
}

fn cmd_recon_run(
    config_path: PathBuf,
    json_output: bool,
    output_file: Option<PathBuf>,
) -> Result<(), CliError> {
    use std::collections::HashMap;
    use visigrid_recon::engine::load_csv_rows;

    // Read and parse config
    let config_str = std::fs::read_to_string(&config_path)
        .map_err(|e| recon_err(EXIT_RECON_RUNTIME, format!("cannot read config: {e}")))?;
    let config = visigrid_recon::ReconConfig::from_toml(&config_str)
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
        if settlement.stale > 0 {
            return Err(recon_err(EXIT_RECON_STALE, "stale items found"));
        }
        // pending only → pass
        return Ok(());
    }

    // Fallback: original logic when no settlement config
    if s.amount_mismatches > 0 || s.timing_mismatches > 0 || s.left_only > 0 || s.right_only > 0 {
        return Err(recon_err(EXIT_RECON_MISMATCH, "mismatches found"));
    }

    Ok(())
}

fn cmd_recon_validate(config_path: PathBuf) -> Result<(), CliError> {
    let config_str = std::fs::read_to_string(&config_path)
        .map_err(|e| recon_err(EXIT_RECON_RUNTIME, format!("cannot read config: {e}")))?;

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
