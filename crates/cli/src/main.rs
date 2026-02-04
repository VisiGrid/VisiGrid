// VisiGrid CLI - headless spreadsheet operations
// See docs/cli-v1.md for specification

mod exit_codes;
mod replay;
mod session;
mod sheet_ops;

use visigrid_cli::diff;

use std::collections::HashMap;
use std::io::{self, Read, Write};
use std::path::PathBuf;
use std::process::ExitCode;

use clap::{Parser, Subcommand, ValueEnum};

// Re-export exit codes from registry (single source of truth)
use exit_codes::{
    EXIT_SUCCESS, EXIT_ERROR, EXIT_USAGE,
    EXIT_AI_DISABLED, EXIT_AI_MISSING_KEY,
    EXIT_DIFF_DUPLICATE, EXIT_DIFF_AMBIGUOUS, EXIT_DIFF_PARSE,
    session_exit_code,
};

// Legacy aliases for backward compatibility (will be removed)
pub const EXIT_EVAL_ERROR: u8 = EXIT_ERROR;
pub const EXIT_ARGS_ERROR: u8 = EXIT_USAGE;
pub const EXIT_IO_ERROR: u8 = 3;     // TODO: migrate to specific codes
pub const EXIT_PARSE_ERROR: u8 = 4;  // TODO: migrate to specific codes
pub const EXIT_FORMAT_ERROR: u8 = 5; // TODO: migrate to specific codes

#[derive(Parser)]
#[command(name = "visigrid-cli")]
#[command(about = "Fast, native spreadsheet (CLI mode, headless)")]
#[command(long_version = long_version())]
#[command(version)]
#[command(subcommand_required = false)]
struct Cli {
    #[command(subcommand)]
    command: Option<Commands>,
}

#[derive(Subcommand)]
enum Commands {
    /// Evaluate a spreadsheet formula against data read from stdin
    #[command(after_help = "\
Examples:
  cat sales.csv | visigrid calc '=SUM(B:B)' -f csv
  cat data.csv | visigrid calc '=AVERAGE(A:A)' -f csv --headers
  echo '1,2,3' | visigrid calc '=SUM(A1:C1)' -f csv
  cat matrix.csv | visigrid calc '=MMULT(A:B,D:E)' -f csv --spill csv")]
    Calc {
        /// Formula to evaluate (must start with =)
        formula: String,

        /// Input format (required)
        #[arg(long, short = 'f')]
        from: Format,

        /// Load data starting at cell
        #[arg(long, default_value = "A1")]
        into: String,

        /// CSV delimiter
        #[arg(long, default_value = ",")]
        delimiter: char,

        /// First row is headers (excluded from formulas)
        #[arg(long)]
        headers: bool,

        /// Output format for array results (csv or json)
        #[arg(long)]
        spill: Option<SpillFormat>,
    },

    /// Convert between file formats
    #[command(after_help = "\
Examples:
  visigrid convert data.xlsx -t csv
  visigrid convert data.xlsx -t json -o data.json
  cat data.csv | visigrid convert -f csv -t json
  visigrid convert report.xlsx -t csv -o - | head -5
  visigrid convert data.csv -t csv --headers --where 'Status=Pending'
  visigrid convert data.csv -t csv --headers --where 'Amount<0'
  visigrid convert data.csv -t csv --headers --select 'Invoice,Total,Status'
  visigrid convert data.csv -t csv --headers --select Invoice --select Total")]
    Convert {
        /// Input file (omit to read from stdin)
        input: Option<PathBuf>,

        /// Input format (required when reading from stdin)
        #[arg(long, short = 'f')]
        from: Option<Format>,

        /// Output format
        #[arg(long, short = 't')]
        to: Format,

        /// Output file (omit for stdout)
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,

        /// Sheet name for multi-sheet files
        #[arg(long)]
        sheet: Option<String>,

        /// CSV/TSV delimiter
        #[arg(long, default_value = ",")]
        delimiter: char,

        /// First row is headers (affects JSON object keys)
        #[arg(long)]
        headers: bool,

        /// Filter rows (requires --headers). Repeatable.
        /// Examples: 'Status=Pending', 'Amount<0', 'Vendor~cloud'
        #[arg(long, value_name = "EXPR")]
        r#where: Vec<String>,

        /// Select columns to output (requires --headers). Repeatable; comma-separated accepted.
        /// Examples: 'Invoice,Total', or --select Invoice --select Total
        #[arg(long, value_name = "COLS")]
        select: Vec<String>,

        /// Suppress stderr notes (e.g. skipped-row counts)
        #[arg(long, short = 'q')]
        quiet: bool,
    },

    /// List all supported functions
    ListFunctions,

    /// Open file in GUI
    Open {
        /// File to open
        file: Option<PathBuf>,
    },

    /// Replay a provenance script
    #[command(after_help = "\
Examples:
  visigrid replay script.lua
  visigrid replay script.lua --verify
  visigrid replay script.lua -o result.csv
  visigrid replay script.lua -o - -f json | jq .
  visigrid replay script.lua --fingerprint")]
    Replay {
        /// Path to the Lua provenance script
        script: PathBuf,

        /// Verify fingerprint against script header (fail if mismatch)
        #[arg(long)]
        verify: bool,

        /// Output file for resulting spreadsheet (csv, tsv, or json)
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,

        /// Output format (inferred from extension if not specified)
        #[arg(long, short = 'f')]
        format: Option<String>,

        /// Print fingerprint and exit
        #[arg(long)]
        fingerprint: bool,

        /// Quiet mode - only print errors
        #[arg(long, short = 'q')]
        quiet: bool,
    },

    /// AI configuration and diagnostics
    Ai {
        #[command(subcommand)]
        command: AiCommands,
    },

    /// Reconcile two datasets by key (exit 0 = reconciled, exit 1 = material diffs)
    #[command(after_help = "\
Exit code 1 indicates material differences: missing rows or value diffs outside \
tolerance. Within-tolerance diffs are reported but do not cause a non-zero exit.

Examples:
  visigrid diff old.csv new.csv --key id
  visigrid diff old.csv new.csv --key name --tolerance 0.01
  visigrid diff old.csv new.csv --key sku --out csv --output diffs.csv
  visigrid diff old.csv new.csv --key id --compare price,quantity
  visigrid diff old.csv new.csv --key name --match contains
  cat export.csv | visigrid diff - baseline.csv --key id
  docker exec db dump | visigrid diff expected.csv - --key sku")]
    Diff {
        /// Left dataset (file path, or - for stdin)
        left: String,

        /// Right dataset (file path, or - for stdin)
        right: String,

        /// Key column (name, letter, or 1-indexed number)
        #[arg(long)]
        key: String,

        /// Matching mode (exact: keys must match exactly; contains: left key must be substring of right key)
        #[arg(long, default_value = "exact")]
        r#match: DiffMatchMode,

        /// Key transform
        #[arg(long, default_value = "trim")]
        key_transform: DiffKeyTransform,

        /// Columns to compare (comma-separated; omit for all non-key)
        #[arg(long)]
        compare: Option<String>,

        /// Numeric tolerance (absolute)
        #[arg(long, default_value = "0")]
        tolerance: f64,

        /// Policy for duplicate keys
        #[arg(long, default_value = "error")]
        on_duplicate: DiffDuplicatePolicy,

        /// Policy for ambiguous matches (contains mode)
        #[arg(long, default_value = "error")]
        on_ambiguous: DiffAmbiguousPolicy,

        /// Output format
        #[arg(long, alias = "format", default_value = "json")]
        out: DiffOutputFormat,

        /// Output file (default: stdout)
        #[arg(long)]
        output: Option<PathBuf>,

        /// Summary output destination
        #[arg(long, default_value = "stderr")]
        summary: DiffSummaryMode,

        /// Treat first row as data (generate A, B, C headers)
        #[arg(long)]
        no_headers: bool,

        /// Header row number (1-indexed)
        #[arg(long)]
        header_row: Option<usize>,

        /// CSV delimiter
        #[arg(long, default_value = ",")]
        delimiter: char,

        /// Format for stdin when using - (inferred from other file if omitted)
        #[arg(long, value_name = "FORMAT")]
        stdin_format: Option<Format>,

        /// Exit 1 on any diff, even within tolerance (Unix-diff semantics)
        #[arg(long)]
        strict_exit: bool,

        /// Quiet mode - suppress stderr summary and warnings
        #[arg(long, short = 'q')]
        quiet: bool,

        /// Export ambiguous matches to CSV file (written before exit, even on --on-ambiguous error)
        #[arg(long)]
        save_ambiguous: Option<PathBuf>,
    },

    /// List running VisiGrid sessions
    #[command(after_help = "\
Examples:
  visigrid sessions
  visigrid sessions --json")]
    Sessions {
        /// Output as JSON
        #[arg(long)]
        json: bool,
    },

    /// Connect to a running session and show status
    #[command(after_help = "\
Examples:
  visigrid attach
  visigrid attach --session abc123
  VISIGRID_SESSION_TOKEN=xxx visigrid attach --session abc123")]
    Attach {
        /// Session ID (prefix match supported; auto-selects if only one session)
        #[arg(long)]
        session: Option<String>,
    },

    /// Apply operations to a running session
    #[command(after_help = "\
Examples:
  visigrid apply ops.jsonl
  visigrid apply --session abc123 ops.jsonl
  cat ops.jsonl | visigrid apply -
  visigrid apply --atomic --expected-revision 42 ops.jsonl
  visigrid apply --wait --wait-timeout 30 ops.jsonl")]
    Apply {
        /// Operations file (JSONL format, or - for stdin)
        ops: String,

        /// Session ID (prefix match supported; auto-selects if only one session)
        #[arg(long)]
        session: Option<String>,

        /// Apply all-or-nothing (rollback on error)
        #[arg(long)]
        atomic: bool,

        /// Expected revision for optimistic concurrency
        #[arg(long)]
        expected_revision: Option<u64>,

        /// Wait and retry on writer conflict (instead of failing immediately)
        #[arg(long)]
        wait: bool,

        /// Maximum time to wait for writer lease (seconds, default 30)
        #[arg(long, default_value = "30")]
        wait_timeout: u64,
    },

    /// Query cell state from a running session
    #[command(after_help = "\
Examples:
  visigrid inspect A1
  visigrid inspect A1:B10 --json
  visigrid inspect --session abc123 --sheet 1 A1:C5")]
    Inspect {
        /// Cell or range to inspect (e.g., A1, A1:B10, or 'workbook')
        range: String,

        /// Session ID (prefix match supported; auto-selects if only one session)
        #[arg(long)]
        session: Option<String>,

        /// Sheet index (0-based, default: 0)
        #[arg(long, default_value = "0")]
        sheet: usize,

        /// Output as JSON (default: human-readable table)
        #[arg(long)]
        json: bool,
    },

    /// Show session server statistics (health check)
    #[command(after_help = "\
Examples:
  visigrid stats
  visigrid stats --session abc123
  visigrid stats --json")]
    Stats {
        /// Session ID (prefix match supported; auto-selects if only one session)
        #[arg(long)]
        session: Option<String>,

        /// Output as JSON
        #[arg(long)]
        json: bool,
    },

    /// View a live session (read-only grid snapshot)
    #[command(after_help = "\
Examples:
  visigrid view
  visigrid view --range A1:K20
  visigrid view --session abc123 --sheet 1
  visigrid view --follow")]
    View {
        /// Session ID (prefix match supported; auto-selects if only one session)
        #[arg(long)]
        session: Option<String>,

        /// Range to display (default: A1:J20)
        #[arg(long, default_value = "A1:J20")]
        range: String,

        /// Sheet index (0-based, default: 0)
        #[arg(long, default_value = "0")]
        sheet: usize,

        /// Follow mode: refresh on changes (poll every 500ms)
        #[arg(long)]
        follow: bool,

        /// Column width for display (default: 12)
        #[arg(long, default_value = "12")]
        width: usize,
    },

    /// Sheet file operations (headless build/inspect/verify)
    #[command(subcommand)]
    Sheet(SheetCommands),
}

/// Sheet subcommands for agent-ready headless workflows.
#[derive(Subcommand)]
enum SheetCommands {
    /// Build a .sheet file from a Lua script (replacement semantics)
    #[command(after_help = "\
Examples:
  visigrid sheet apply model.sheet --lua build.lua
  visigrid sheet apply model.sheet --lua build.lua --verify v1:42:abc123...
  visigrid sheet apply model.sheet --lua build.lua --dry-run
  visigrid sheet apply model.sheet --lua build.lua --json

The Lua script builds the sheet from scratch using:
  set(cell, value)     -- set cell value or formula
  clear(cell)          -- clear cell
  meta(target, table)  -- semantic metadata (affects fingerprint)
  style(target, table) -- presentation style (excluded from fingerprint)

Example Lua script:
  set(\"A1\", \"Revenue Model\")
  meta(\"A1\", { role = \"title\" })
  style(\"A1\", { bold = true })
  set(\"B2\", 10000)
  set(\"B3\", \"=B2*1.05\")")]
    Apply {
        /// Output .sheet file path
        output: PathBuf,

        /// Path to Lua build script
        #[arg(long)]
        lua: PathBuf,

        /// Verify fingerprint after build (exit 1 if mismatch)
        #[arg(long)]
        verify: Option<String>,

        /// Compute fingerprint but don't write file
        #[arg(long)]
        dry_run: bool,

        /// Output as JSON
        #[arg(long)]
        json: bool,
    },

    /// Inspect cells/ranges in a .sheet file
    #[command(after_help = "\
Examples:
  visigrid sheet inspect model.sheet A1
  visigrid sheet inspect model.sheet A1:D10
  visigrid sheet inspect model.sheet --workbook
  visigrid sheet inspect model.sheet A1 --json
  visigrid sheet inspect model.sheet A1 --include-style")]
    Inspect {
        /// Path to .sheet file
        file: PathBuf,

        /// Target to inspect (cell like A1, range like A1:D10, or omit for workbook)
        target: Option<String>,

        /// Show workbook metadata (fingerprint, sheet count)
        #[arg(long)]
        workbook: bool,

        /// Include style information
        #[arg(long)]
        include_style: bool,

        /// Output as JSON
        #[arg(long)]
        json: bool,
    },

    /// Verify a .sheet file's fingerprint
    #[command(after_help = "\
Examples:
  visigrid sheet verify model.sheet --fingerprint v1:42:abc123...

Exit codes:
  0  Fingerprint matches
  1  Fingerprint mismatch
  2  Usage error")]
    Verify {
        /// Path to .sheet file
        file: PathBuf,

        /// Expected fingerprint
        #[arg(long)]
        fingerprint: String,
    },

    /// Compute and print a .sheet file's fingerprint
    #[command(after_help = "\
Examples:
  visigrid sheet fingerprint model.sheet
  visigrid sheet fingerprint model.sheet --json")]
    Fingerprint {
        /// Path to .sheet file
        file: PathBuf,

        /// Output as JSON
        #[arg(long)]
        json: bool,
    },
}

#[derive(Subcommand)]
enum AiCommands {
    /// Check AI configuration and connectivity
    Doctor {
        /// Output as JSON for machine parsing
        #[arg(long)]
        json: bool,

        /// Test provider connectivity (requires network)
        #[arg(long)]
        test: bool,
    },
}

#[derive(Clone, Copy, ValueEnum)]
enum Format {
    Csv,
    Tsv,
    Json,
    Lines,
    Xlsx,
    Sheet,
}

#[derive(Clone, Copy, ValueEnum)]
enum SpillFormat {
    Csv,
    Json,
}

#[derive(Clone, Copy, ValueEnum)]
enum DiffMatchMode {
    Exact,
    Contains,
}

#[derive(Clone, Copy, ValueEnum)]
enum DiffKeyTransform {
    None,
    Trim,
    Digits,
}

#[derive(Clone, Copy, ValueEnum)]
enum DiffDuplicatePolicy {
    Error,
}

#[derive(Clone, Copy, ValueEnum)]
enum DiffAmbiguousPolicy {
    Error,
    Report,
}

#[derive(Clone, Copy, ValueEnum)]
enum DiffOutputFormat {
    Json,
    Csv,
}

#[derive(Clone, Copy, ValueEnum)]
enum DiffSummaryMode {
    None,
    Stderr,
    Json,
}

// ============================================================================
// --where filtering types and helpers
// ============================================================================

#[derive(Clone, Copy)]
enum WhereOp {
    Eq,
    NotEq,
    Lt,
    Gt,
    Contains,
}

struct WhereClause {
    column: String, // lowercased
    op: WhereOp,
    value: String, // quote-stripped, trimmed
}

struct ResolvedWhere {
    col: usize,
    op: WhereOp,
    value: String,
    /// RHS parsed as f64 (after lenient strip). None if not numeric.
    numeric_value: Option<f64>,
}

/// Strip `$`, `,`, whitespace, then parse as f64.
fn lenient_parse_f64(s: &str) -> Option<f64> {
    let stripped: String = s.chars().filter(|c| *c != '$' && *c != ',').collect();
    stripped.trim().parse::<f64>().ok()
}

fn parse_where(expr: &str) -> Result<WhereClause, CliError> {
    // Reject >= and <= with hint
    if expr.contains(">=") {
        return Err(CliError::args(format!("unsupported operator >= in {:?}", expr))
            .with_hint("use two clauses: --where 'col>value' --where 'col=value'"));
    }
    if expr.contains("<=") {
        return Err(CliError::args(format!("unsupported operator <= in {:?}", expr))
            .with_hint("use two clauses: --where 'col<value' --where 'col=value'"));
    }

    // Try operators in order: != ~ = < >
    let (col, op, raw_value) = if let Some(pos) = expr.find("!=") {
        (&expr[..pos], WhereOp::NotEq, &expr[pos + 2..])
    } else if let Some(pos) = expr.find('~') {
        (&expr[..pos], WhereOp::Contains, &expr[pos + 1..])
    } else if let Some(pos) = expr.find('=') {
        (&expr[..pos], WhereOp::Eq, &expr[pos + 1..])
    } else if let Some(pos) = expr.find('<') {
        (&expr[..pos], WhereOp::Lt, &expr[pos + 1..])
    } else if let Some(pos) = expr.find('>') {
        (&expr[..pos], WhereOp::Gt, &expr[pos + 1..])
    } else {
        return Err(CliError::args(format!("no operator found in --where {:?}", expr))
            .with_hint("syntax: 'Column=value', 'Column<number', 'Column~substring'"));
    };

    let col = col.trim();
    if col.is_empty() {
        return Err(CliError::args(format!("empty column name in --where {:?}", expr)));
    }

    // Strip one layer of surrounding quotes from value
    let value = raw_value.trim();
    let value = if (value.starts_with('"') && value.ends_with('"'))
        || (value.starts_with('\'') && value.ends_with('\''))
    {
        &value[1..value.len() - 1]
    } else {
        value
    };

    Ok(WhereClause {
        column: col.trim().to_lowercase(),
        op,
        value: value.to_string(),
    })
}

fn resolve_where_columns(
    clauses: &[WhereClause],
    canonical_headers: &[String],
) -> Result<Vec<ResolvedWhere>, CliError> {
    let headers: Vec<String> = canonical_headers
        .iter()
        .map(|h| h.trim().to_lowercase())
        .collect();

    let mut resolved = Vec::with_capacity(clauses.len());
    for clause in clauses {
        let col_idx = headers.iter().position(|h| h == &clause.column);
        match col_idx {
            Some(idx) => {
                resolved.push(ResolvedWhere {
                    col: idx,
                    op: clause.op,
                    value: clause.value.clone(),
                    numeric_value: lenient_parse_f64(&clause.value),
                });
            }
            None => {
                let available: Vec<String> = canonical_headers
                    .iter()
                    .map(|h| h.trim().to_string())
                    .filter(|h| !h.is_empty())
                    .collect();
                return Err(CliError::args(format!("unknown column {:?}", clause.column))
                    .with_hint(format!("available columns: {}", available.join(", "))));
            }
        }
    }
    Ok(resolved)
}

fn row_matches(
    sheet: &visigrid_engine::sheet::Sheet,
    row: usize,
    conditions: &[ResolvedWhere],
    skip_counts: &mut [usize],
) -> bool {
    for (i, cond) in conditions.iter().enumerate() {
        let cell = sheet.get_display(row, cond.col);
        let matches = match cond.op {
            WhereOp::Contains => cell.to_lowercase().contains(&cond.value.to_lowercase()),
            WhereOp::Lt => {
                if let Some(rhs) = cond.numeric_value {
                    match lenient_parse_f64(&cell) {
                        Some(lhs) => lhs < rhs,
                        None => {
                            skip_counts[i] += 1;
                            false
                        }
                    }
                } else {
                    false
                }
            }
            WhereOp::Gt => {
                if let Some(rhs) = cond.numeric_value {
                    match lenient_parse_f64(&cell) {
                        Some(lhs) => lhs > rhs,
                        None => {
                            skip_counts[i] += 1;
                            false
                        }
                    }
                } else {
                    false
                }
            }
            WhereOp::Eq => {
                if let Some(rhs) = cond.numeric_value {
                    // Numeric equality
                    match lenient_parse_f64(&cell) {
                        Some(lhs) => lhs == rhs,
                        None => {
                            skip_counts[i] += 1;
                            false
                        }
                    }
                } else {
                    // String equality (case-insensitive)
                    cell.eq_ignore_ascii_case(&cond.value)
                }
            }
            WhereOp::NotEq => {
                if let Some(rhs) = cond.numeric_value {
                    // Numeric not-equals
                    match lenient_parse_f64(&cell) {
                        Some(lhs) => lhs != rhs,
                        None => {
                            skip_counts[i] += 1;
                            false
                        }
                    }
                } else {
                    // String not-equals (case-insensitive)
                    !cell.eq_ignore_ascii_case(&cond.value)
                }
            }
        };
        if !matches {
            return false;
        }
    }
    true
}

fn filter_row_indices(
    sheet: &visigrid_engine::sheet::Sheet,
    conditions: &[ResolvedWhere],
    header_row: usize,
) -> (Vec<usize>, Vec<usize>) {
    let (rows, _) = get_data_bounds(sheet);
    let mut matched = Vec::new();
    let mut skip_counts = vec![0usize; conditions.len()];
    for row in (header_row + 1)..rows {
        if row_matches(sheet, row, conditions, &mut skip_counts) {
            matched.push(row);
        }
    }
    (matched, skip_counts)
}

// ============================================================================
// --select helpers
// ============================================================================

/// Find the first non-empty row in the sheet. Returns 0 if all rows are empty.
fn find_header_row(sheet: &visigrid_engine::sheet::Sheet, rows: usize, cols: usize) -> usize {
    for row in 0..rows {
        for col in 0..cols {
            if !sheet.get_display(row, col).trim().is_empty() {
                return row;
            }
        }
    }
    0
}

fn check_ambiguous_headers(canonical_headers: &[String]) -> Result<(), CliError> {
    let mut seen: HashMap<String, Vec<String>> = HashMap::new();
    for h in canonical_headers {
        let canon = h.trim();
        if canon.is_empty() { continue; }
        seen.entry(canon.to_lowercase())
            .or_default()
            .push(canon.to_string());
    }

    for (key, names) in &seen {
        if names.len() > 1 {
            return Err(CliError::args(format!(
                "ambiguous column name \"{}\" (matches: {})",
                key,
                names.join(", ")
            )));
        }
    }
    Ok(())
}

fn parse_select_args(select_args: &[String]) -> Vec<String> {
    select_args
        .iter()
        .flat_map(|arg| arg.split(','))
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .collect()
}

fn resolve_select_columns(
    select_names: &[String],
    canonical_headers: &[String],
) -> Result<Vec<(usize, String)>, CliError> {
    // Build O(1) lookup: lowercased → (index, canonical name)
    let mut map: HashMap<String, (usize, String)> = HashMap::new();
    for (i, h) in canonical_headers.iter().enumerate() {
        let key = h.trim().to_lowercase();
        if key.is_empty() { continue; }
        map.insert(key, (i, h.clone()));
    }

    let mut result = Vec::with_capacity(select_names.len());
    let mut seen_indices = std::collections::HashSet::new();

    for name in select_names {
        let needle = name.trim().to_lowercase();
        match map.get(&needle) {
            Some((idx, canonical)) => {
                if !seen_indices.insert(*idx) {
                    return Err(CliError::args(
                        format!("duplicate column in --select: \"{}\"", name)
                    ));
                }
                result.push((*idx, canonical.clone()));
            }
            None => {
                let non_empty_count = canonical_headers.iter().filter(|h| !h.trim().is_empty()).count();
                let available: Vec<&str> = canonical_headers
                    .iter()
                    .map(|h| h.as_str())
                    .filter(|h| !h.trim().is_empty())
                    .take(25)
                    .collect();
                let suffix = if non_empty_count > 25 {
                    format!(" (+{} more)", non_empty_count - 25)
                } else {
                    String::new()
                };
                return Err(CliError::args(
                    format!("unknown column in --select: \"{}\"", name)
                ).with_hint(format!("available columns: {}{}", available.join(", "), suffix)));
            }
        }
    }

    Ok(result)
}

fn long_version() -> &'static str {
    if cfg!(debug_assertions) {
        concat!(
            env!("CARGO_PKG_VERSION"),
            " (", env!("GIT_COMMIT_HASH"), ")",
            "\nengine:  visigrid-engine ", env!("CARGO_PKG_VERSION"),
            "\nbuild:   debug",
            "\ntarget:  ", env!("TARGET"),
            "\ncontract_version(diff): 1",
        )
    } else {
        concat!(
            env!("CARGO_PKG_VERSION"),
            " (", env!("GIT_COMMIT_HASH"), ")",
            "\nengine:  visigrid-engine ", env!("CARGO_PKG_VERSION"),
            "\nbuild:   release",
            "\ntarget:  ", env!("TARGET"),
            "\ncontract_version(diff): 1",
        )
    }
}

fn main() -> ExitCode {
    let cli = Cli::parse();

    let result = match cli.command {
        None => {
            // No subcommand = show help
            eprintln!("Usage: visigrid <command> [options]");
            eprintln!("       visigrid --help for more information");
            Ok(())
        }
        Some(Commands::ListFunctions) => cmd_list_functions(),
        Some(Commands::Convert {
            input,
            from,
            to,
            output,
            sheet,
            delimiter,
            headers,
            r#where: where_clauses,
            select: select_args,
            quiet,
        }) => cmd_convert(input, from, to, output, sheet, delimiter, headers, where_clauses, select_args, quiet),
        Some(Commands::Calc {
            formula,
            from,
            into,
            delimiter,
            headers,
            spill,
        }) => cmd_calc(formula, from, into, delimiter, headers, spill),
        Some(Commands::Open { file }) => cmd_open(file),
        Some(Commands::Replay {
            script,
            verify,
            output,
            format,
            fingerprint,
            quiet,
        }) => cmd_replay(script, verify, output, format, fingerprint, quiet),
        Some(Commands::Ai { command }) => match command {
            AiCommands::Doctor { json, test } => cmd_ai_doctor(json, test),
        },
        Some(Commands::Diff {
            left,
            right,
            key,
            r#match,
            key_transform,
            compare,
            tolerance,
            on_duplicate: _,
            on_ambiguous,
            out,
            output,
            summary,
            no_headers,
            header_row,
            delimiter,
            stdin_format,
            strict_exit,
            quiet,
            save_ambiguous,
        }) => cmd_diff(
            left, right, key, r#match, key_transform, compare, tolerance,
            on_ambiguous, out, output, summary, no_headers, header_row, delimiter,
            stdin_format, strict_exit, quiet, save_ambiguous,
        ),
        Some(Commands::Sessions { json }) => cmd_sessions(json),
        Some(Commands::Attach { session }) => cmd_attach(session),
        Some(Commands::Apply { ops, session, atomic, expected_revision, wait, wait_timeout }) => {
            cmd_apply(ops, session, atomic, expected_revision, wait, wait_timeout)
        }
        Some(Commands::Inspect { range, session, sheet, json }) => cmd_inspect(range, session, sheet, json),
        Some(Commands::Stats { session, json }) => cmd_stats(session, json),
        Some(Commands::View { session, range, sheet, follow, width }) => {
            cmd_view(session, range, sheet, follow, width)
        }
        Some(Commands::Sheet(sheet_cmd)) => match sheet_cmd {
            SheetCommands::Apply { output, lua, verify, dry_run, json } => {
                cmd_sheet_apply(output, lua, verify, dry_run, json)
            }
            SheetCommands::Inspect { file, target, workbook, include_style, json } => {
                cmd_sheet_inspect(file, target, workbook, include_style, json)
            }
            SheetCommands::Verify { file, fingerprint } => {
                cmd_sheet_verify(file, fingerprint)
            }
            SheetCommands::Fingerprint { file, json } => {
                cmd_sheet_fingerprint(file, json)
            }
        }
    };

    match result {
        Ok(()) => ExitCode::from(EXIT_SUCCESS),
        Err(CliError { code, message, hint }) => {
            if !message.is_empty() {
                eprintln!("error: {}", message);
            }
            if let Some(hint) = hint {
                eprintln!("hint:  {}", hint);
            }
            ExitCode::from(code)
        }
    }
}

#[derive(Debug)]
pub struct CliError {
    pub code: u8,
    pub message: String,
    pub hint: Option<String>,
}

impl CliError {
    pub fn args(msg: impl Into<String>) -> Self {
        Self { code: EXIT_ARGS_ERROR, message: msg.into(), hint: None }
    }

    pub fn io(msg: impl Into<String>) -> Self {
        Self { code: EXIT_IO_ERROR, message: msg.into(), hint: None }
    }

    pub fn parse(msg: impl Into<String>) -> Self {
        Self { code: EXIT_PARSE_ERROR, message: msg.into(), hint: None }
    }

    pub fn format(msg: impl Into<String>) -> Self {
        Self { code: EXIT_FORMAT_ERROR, message: msg.into(), hint: None }
    }

    pub fn eval(msg: impl Into<String>) -> Self {
        Self { code: EXIT_EVAL_ERROR, message: msg.into(), hint: None }
    }

    /// Create error from session error with proper exit code.
    pub fn session(err: session::SessionError) -> Self {
        let code = session_exit_code(&err);
        let hint = match &err {
            session::SessionError::ConnectionFailed(_) => {
                Some("is VisiGrid GUI running with session server enabled?".to_string())
            }
            session::SessionError::AuthFailed(_) => {
                Some("check VISIGRID_SESSION_TOKEN environment variable".to_string())
            }
            session::SessionError::ServerError { code: err_code, .. }
                if err_code == "writer_conflict" =>
            {
                Some("another client holds the write lease; retry later".to_string())
            }
            session::SessionError::ServerError { code: err_code, .. }
                if err_code == "revision_mismatch" =>
            {
                Some("workbook was modified; re-fetch and retry".to_string())
            }
            _ => None,
        };
        Self { code, message: err.to_string(), hint }
    }

    /// Add a hint to an existing error.
    pub fn with_hint(mut self, hint: impl Into<String>) -> Self {
        self.hint = Some(hint.into());
        self
    }
}

// ============================================================================
// list-functions
// ============================================================================

fn cmd_list_functions() -> Result<(), CliError> {
    let functions = visigrid_engine::formula::functions::list_functions();
    let stdout = io::stdout();
    let mut handle = stdout.lock();

    for name in functions {
        writeln!(handle, "{}", name).map_err(|e| CliError::io(e.to_string()))?;
    }

    Ok(())
}

// ============================================================================
// convert
// ============================================================================

fn cmd_convert(
    input: Option<PathBuf>,
    from: Option<Format>,
    to: Format,
    output: Option<PathBuf>,
    _sheet: Option<String>,
    delimiter: char,
    headers: bool,
    where_clauses: Vec<String>,
    select_args: Vec<String>,
    quiet: bool,
) -> Result<(), CliError> {

    // Validate --where requires --headers
    if !where_clauses.is_empty() && !headers {
        return Err(CliError::args("--where requires --headers")
            .with_hint("add --headers so column names can be resolved"));
    }

    // Validate --select requires --headers
    if !select_args.is_empty() && !headers {
        return Err(CliError::args("--select requires --headers")
            .with_hint("add --headers so column names can be resolved"));
    }

    // Determine input format
    let input_format = match (&input, from) {
        (None, None) => return Err(CliError::args("stdin requires --from to specify the input format")
            .with_hint("visigrid-cli convert --from csv -t json")),
        (None, Some(f)) => f,
        (Some(path), None) => infer_format(path)?,
        (Some(_), Some(f)) => f, // --from overrides extension
    };

    // Read input into sheet (convert always starts at A1)
    let sheet = match &input {
        Some(path) => read_file(path, input_format, delimiter)?,
        None => read_stdin(input_format, delimiter, 0, 0)?,
    };

    let (bounds_rows, bounds_cols) = get_data_bounds(&sheet);

    // Find the actual header row (first non-empty row)
    let header_row = if headers && bounds_rows > 0 && bounds_cols > 0 {
        find_header_row(&sheet, bounds_rows, bounds_cols)
    } else {
        0
    };

    // Build canonical headers list once
    let canonical_headers: Vec<String> = if headers && bounds_cols > 0 {
        (0..bounds_cols).map(|c| sheet.get_display(header_row, c).trim().to_string()).collect()
    } else {
        vec![]
    };

    // Ambiguous header check (once, before --where or --select resolution)
    if (!where_clauses.is_empty() || !select_args.is_empty()) && headers {
        check_ambiguous_headers(&canonical_headers)?;
    }

    // Resolve and apply --where filters
    let row_filter = if !where_clauses.is_empty() {
        let parsed: Vec<WhereClause> = where_clauses
            .iter()
            .map(|e| parse_where(e))
            .collect::<Result<Vec<_>, _>>()?;
        let resolved = resolve_where_columns(&parsed, &canonical_headers)?;
        let (indices, skip_counts) = filter_row_indices(&sheet, &resolved, header_row);

        // Report unparseable cells to stderr (suppressed by --quiet)
        if !quiet {
            for (i, &count) in skip_counts.iter().enumerate() {
                if count > 0 {
                    eprintln!("note: {} rows skipped ({} not numeric)", count, parsed[i].column);
                }
            }
        }

        Some(indices)
    } else {
        None
    };

    // Resolve column selection (after --where, before write)
    let col_filter = if !select_args.is_empty() {
        let select_names = parse_select_args(&select_args);
        if select_names.is_empty() {
            return Err(CliError::args("empty --select list"));
        }
        let resolved = resolve_select_columns(&select_names, &canonical_headers)?;
        Some(resolved)
    } else {
        None
    };

    // Write output
    let output_bytes = write_format(
        &sheet, to, delimiter, headers, header_row,
        row_filter.as_deref(),
        col_filter.as_deref(),
    )?;

    match output {
        Some(path) => {
            std::fs::write(&path, &output_bytes)
                .map_err(|e| CliError::io(format!("{}: {}", path.display(), e)))?;
        }
        None => {
            io::stdout()
                .write_all(&output_bytes)
                .map_err(|e| CliError::io(e.to_string()))?;
        }
    }

    Ok(())
}

fn infer_format(path: &PathBuf) -> Result<Format, CliError> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase());

    match ext.as_deref() {
        Some("csv") => Ok(Format::Csv),
        Some("tsv") => Ok(Format::Tsv),
        Some("json") => Ok(Format::Json),
        Some("xlsx") | Some("xls") | Some("xlsb") | Some("ods") => Ok(Format::Xlsx),
        Some("sheet") => Ok(Format::Sheet),
        _ => Err(CliError::args(format!(
            "cannot infer format from extension {:?}",
            ext.as_deref().unwrap_or("(none)")
        )).with_hint("use --from with one of: csv, tsv, json, xlsx, sheet")),
    }
}

fn read_file(path: &PathBuf, format: Format, _delimiter: char) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    // TODO: Use custom delimiter when io crate supports it
    match format {
        Format::Csv => {
            visigrid_io::csv::import(path)
                .map_err(|e| CliError::parse(e))
        }
        Format::Tsv => {
            visigrid_io::csv::import_tsv(path)
                .map_err(|e| CliError::parse(e))
        }
        Format::Xlsx => {
            let (workbook, _stats) = visigrid_io::xlsx::import(path)
                .map_err(|e| CliError::parse(e))?;
            // Return the first sheet (clone it since we can't move out of workbook)
            workbook.sheet(0)
                .cloned()
                .ok_or_else(|| CliError::parse("xlsx file has no sheets"))
        }
        Format::Sheet => {
            visigrid_io::native::load(path)
                .map_err(|e| CliError::io(e))
        }
        Format::Json => {
            let content = std::fs::read_to_string(path)
                .map_err(|e| CliError::io(e.to_string()))?;
            parse_json(&content, 0, 0)
        }
        Format::Lines => {
            let content = std::fs::read_to_string(path)
                .map_err(|e| CliError::io(e.to_string()))?;
            parse_lines(&content, 0, 0)
        }
    }
}

fn read_stdin(format: Format, delimiter: char, into_row: usize, into_col: usize) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    let mut input = String::new();
    io::stdin()
        .read_to_string(&mut input)
        .map_err(|e| CliError::io(e.to_string()))?;

    if input.is_empty() {
        return Err(CliError::parse("no input received on stdin")
            .with_hint("cat file.csv | visigrid-cli calc '=SUM(A:A)' --from csv"));
    }

    match format {
        Format::Csv => parse_csv(&input, delimiter as u8, into_row, into_col),
        Format::Tsv => parse_csv(&input, b'\t', into_row, into_col),
        Format::Json => parse_json(&input, into_row, into_col),
        Format::Lines => parse_lines(&input, into_row, into_col),
        Format::Xlsx | Format::Sheet => {
            Err(CliError::args("xlsx and sheet formats require file input"))
        }
    }
}

fn parse_csv(content: &str, delimiter: u8, into_row: usize, into_col: usize) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    use visigrid_engine::sheet::{Sheet, SheetId};

    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .from_reader(content.as_bytes());

    let mut sheet = Sheet::new(SheetId(1), 1000, 26);

    for (row_idx, result) in reader.records().enumerate() {
        let record = result.map_err(|e| CliError::parse(format!("line {}: {}", row_idx + 1, e)))?;
        for (col_idx, field) in record.iter().enumerate() {
            if !field.is_empty() {
                sheet.set_value(into_row + row_idx, into_col + col_idx, field);
            }
        }
    }

    Ok(sheet)
}

fn parse_json(content: &str, into_row: usize, into_col: usize) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    use visigrid_engine::sheet::{Sheet, SheetId};

    let value: serde_json::Value = serde_json::from_str(content)
        .map_err(|e| CliError::parse(format!("JSON parse error: {}", e)))?;

    let mut sheet = Sheet::new(SheetId(1), 1000, 26);

    match value {
        serde_json::Value::Array(rows) => {
            if rows.is_empty() {
                return Err(CliError::parse("empty input"));
            }

            // Check if array of arrays or array of objects
            if let Some(serde_json::Value::Object(_)) = rows.first() {
                // Array of objects - collect all keys lexicographically
                let mut all_keys: std::collections::BTreeSet<String> = std::collections::BTreeSet::new();
                for row in &rows {
                    if let serde_json::Value::Object(obj) = row {
                        for key in obj.keys() {
                            all_keys.insert(key.clone());
                        }
                    }
                }
                let keys: Vec<String> = all_keys.into_iter().collect();

                // Write header row
                for (col, key) in keys.iter().enumerate() {
                    sheet.set_value(into_row, into_col + col, key);
                }

                // Write data rows
                for (row_idx, row) in rows.iter().enumerate() {
                    if let serde_json::Value::Object(obj) = row {
                        for (col, key) in keys.iter().enumerate() {
                            if let Some(val) = obj.get(key) {
                                let cell_value = json_value_to_string(val, row_idx + 1, key)?;
                                if !cell_value.is_empty() {
                                    sheet.set_value(into_row + row_idx + 1, into_col + col, &cell_value);
                                }
                            }
                        }
                    }
                }
            } else {
                // Array of arrays
                for (row_idx, row) in rows.iter().enumerate() {
                    if let serde_json::Value::Array(cols) = row {
                        for (col_idx, val) in cols.iter().enumerate() {
                            let cell_value = json_value_to_string(val, row_idx, &col_idx.to_string())?;
                            if !cell_value.is_empty() {
                                sheet.set_value(into_row + row_idx, into_col + col_idx, &cell_value);
                            }
                        }
                    } else {
                        return Err(CliError::parse(format!("row {}: expected array", row_idx)));
                    }
                }
            }
        }
        _ => return Err(CliError::parse("JSON must be array of arrays or array of objects")),
    }

    Ok(sheet)
}

fn json_value_to_string(val: &serde_json::Value, row: usize, key: &str) -> Result<String, CliError> {
    match val {
        serde_json::Value::Null => Ok(String::new()),
        serde_json::Value::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
        serde_json::Value::Number(n) => Ok(n.to_string()),
        serde_json::Value::String(s) => Ok(s.clone()),
        serde_json::Value::Array(_) | serde_json::Value::Object(_) => {
            Err(CliError::parse(format!("non-scalar value at row {}, key \"{}\"", row, key)))
        }
    }
}

fn parse_lines(content: &str, into_row: usize, into_col: usize) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    use visigrid_engine::sheet::{Sheet, SheetId};

    let lines: Vec<&str> = content.lines().collect();
    if lines.is_empty() {
        return Err(CliError::parse("empty input"));
    }

    let mut sheet = Sheet::new(SheetId(1), 1000, 26);
    for (row, line) in lines.iter().enumerate() {
        if !line.is_empty() {
            sheet.set_value(into_row + row, into_col, line);
        }
    }

    Ok(sheet)
}

fn write_format(
    sheet: &visigrid_engine::sheet::Sheet,
    format: Format,
    delimiter: char,
    headers: bool,
    header_row: usize,
    row_filter: Option<&[usize]>,
    col_filter: Option<&[(usize, String)]>,
) -> Result<Vec<u8>, CliError> {
    match format {
        Format::Csv => write_csv(sheet, delimiter as u8, headers, header_row, row_filter, col_filter),
        Format::Tsv => write_csv(sheet, b'\t', headers, header_row, row_filter, col_filter),
        Format::Json => write_json(sheet, headers, header_row, row_filter, col_filter),
        Format::Lines => write_lines(sheet, header_row, row_filter, col_filter),
        Format::Xlsx => Err(CliError::format("xlsx export not yet implemented")
            .with_hint("use -t csv or -t json instead")),
        Format::Sheet => Err(CliError::format("sheet format cannot be written to stdout")
            .with_hint("use -o output.sheet to write to a file")),
    }
}

fn write_csv(
    sheet: &visigrid_engine::sheet::Sheet,
    delimiter: u8,
    headers: bool,
    header_row: usize,
    row_filter: Option<&[usize]>,
    col_filter: Option<&[(usize, String)]>,
) -> Result<Vec<u8>, CliError> {
    let mut writer = csv::WriterBuilder::new()
        .delimiter(delimiter)
        .from_writer(Vec::new());

    let (rows, cols) = get_data_bounds(sheet);

    // Helper: push columns for a given row into the record
    let push_row = |record: &mut Vec<String>, row: usize| {
        match col_filter {
            Some(selected) => {
                for (idx, _) in selected {
                    record.push(sheet.get_display(row, *idx));
                }
            }
            None => {
                for col in 0..cols {
                    record.push(sheet.get_display(row, col));
                }
            }
        }
    };

    match row_filter {
        Some(indices) => {
            // Write header row + filtered data rows
            if rows > 0 {
                let mut record: Vec<String> = Vec::new();
                push_row(&mut record, header_row);
                writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
            }
            for &row in indices {
                let mut record: Vec<String> = Vec::new();
                push_row(&mut record, row);
                writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
            }
        }
        None => {
            if headers {
                // Write header row, then data rows starting after header
                if rows > 0 {
                    let mut record: Vec<String> = Vec::new();
                    push_row(&mut record, header_row);
                    writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
                }
                for row in (header_row + 1)..rows {
                    let mut record: Vec<String> = Vec::new();
                    push_row(&mut record, row);
                    writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
                }
            } else {
                for row in 0..rows {
                    let mut record: Vec<String> = Vec::new();
                    push_row(&mut record, row);
                    writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
                }
            }
        }
    }

    writer.into_inner().map_err(|e| CliError::io(e.to_string()))
}

fn write_json(
    sheet: &visigrid_engine::sheet::Sheet,
    headers: bool,
    header_row: usize,
    row_filter: Option<&[usize]>,
    col_filter: Option<&[(usize, String)]>,
) -> Result<Vec<u8>, CliError> {
    let (rows, cols) = get_data_bounds(sheet);

    if headers && rows > 0 {
        let data_rows: Vec<usize> = match row_filter {
            Some(indices) => indices.to_vec(),
            None => ((header_row + 1)..rows).collect(),
        };

        if let Some(selected) = col_filter {
            // --select path: build JSON manually to preserve key order
            let json_keys: Vec<(usize, String)> = selected.iter().map(|(idx, name)| {
                let sanitized: String = name
                    .to_lowercase()
                    .chars()
                    .map(|c| if c.is_alphanumeric() { c } else { '_' })
                    .collect();
                let key = if sanitized.is_empty() {
                    format!("col{}", idx)
                } else {
                    sanitized
                };
                (*idx, key)
            }).collect();

            let mut rows_json: Vec<Vec<(String, serde_json::Value)>> = Vec::new();
            for row in data_rows {
                let mut pairs = Vec::new();
                for (col_idx, key) in &json_keys {
                    let value = sheet.get_display(row, *col_idx);
                    pairs.push((key.clone(), string_to_json_value(&value)));
                }
                rows_json.push(pairs);
            }

            // Format manually to preserve key order
            let mut bytes = Vec::new();
            bytes.extend_from_slice(b"[\n");
            for (i, pairs) in rows_json.iter().enumerate() {
                bytes.extend_from_slice(b"  {\n");
                for (j, (key, value)) in pairs.iter().enumerate() {
                    let val_str = serde_json::to_string(value).map_err(|e| CliError::io(e.to_string()))?;
                    bytes.extend_from_slice(b"    ");
                    bytes.extend_from_slice(serde_json::to_string(key).map_err(|e| CliError::io(e.to_string()))?.as_bytes());
                    bytes.extend_from_slice(b": ");
                    bytes.extend_from_slice(val_str.as_bytes());
                    if j < pairs.len() - 1 {
                        bytes.push(b',');
                    }
                    bytes.push(b'\n');
                }
                bytes.extend_from_slice(b"  }");
                if i < rows_json.len() - 1 {
                    bytes.push(b',');
                }
                bytes.push(b'\n');
            }
            bytes.extend_from_slice(b"]\n");
            Ok(bytes)
        } else {
            // Standard path: array of objects with BTreeMap key ordering
            let mut header_names: Vec<String> = Vec::new();
            for col in 0..cols {
                let name = sheet.get_display(header_row, col);
                let sanitized: String = name
                    .to_lowercase()
                    .chars()
                    .map(|c| if c.is_alphanumeric() { c } else { '_' })
                    .collect();
                header_names.push(if sanitized.is_empty() {
                    format!("col{}", col)
                } else {
                    sanitized
                });
            }

            let mut objects: Vec<serde_json::Map<String, serde_json::Value>> = Vec::new();
            for row in data_rows {
                let mut obj = serde_json::Map::new();
                for (col, key) in header_names.iter().enumerate() {
                    let value = sheet.get_display(row, col);
                    obj.insert(key.clone(), string_to_json_value(&value));
                }
                objects.push(obj);
            }

            let mut bytes = serde_json::to_vec_pretty(&objects).map_err(|e| CliError::io(e.to_string()))?;
            bytes.push(b'\n');
            Ok(bytes)
        }
    } else {
        // Array of arrays (no col_filter since --select requires --headers)
        let mut rows_vec: Vec<Vec<serde_json::Value>> = Vec::new();
        let all_rows: Vec<usize> = match row_filter {
            Some(indices) => {
                let mut v = vec![header_row];
                v.extend_from_slice(indices);
                v
            }
            None => (0..rows).collect(),
        };
        for row in all_rows {
            let mut row_vec: Vec<serde_json::Value> = Vec::new();
            for col in 0..cols {
                let value = sheet.get_display(row, col);
                row_vec.push(string_to_json_value(&value));
            }
            rows_vec.push(row_vec);
        }

        let mut bytes = serde_json::to_vec_pretty(&rows_vec).map_err(|e| CliError::io(e.to_string()))?;
        bytes.push(b'\n');
        Ok(bytes)
    }
}

/// Convert a display string to a typed JSON value
/// Numbers become JSON numbers, booleans become JSON booleans, rest are strings
fn string_to_json_value(s: &str) -> serde_json::Value {
    if s.is_empty() {
        return serde_json::Value::String(String::new());
    }

    // Try to parse as number first
    if let Ok(n) = s.parse::<f64>() {
        // Check if it's an integer
        if n.fract() == 0.0 && n.abs() < i64::MAX as f64 {
            serde_json::json!(n as i64)
        } else {
            serde_json::json!(n)
        }
    } else if s == "TRUE" || s == "true" {
        serde_json::json!(true)
    } else if s == "FALSE" || s == "false" {
        serde_json::json!(false)
    } else {
        serde_json::json!(s)
    }
}

fn write_lines(
    sheet: &visigrid_engine::sheet::Sheet,
    header_row: usize,
    row_filter: Option<&[usize]>,
    col_filter: Option<&[(usize, String)]>,
) -> Result<Vec<u8>, CliError> {
    let mut output = Vec::new();
    let (rows, _) = get_data_bounds(sheet);

    // With --select: output the first selected column; without: column 0
    let output_col = match col_filter {
        Some(selected) => selected[0].0,
        None => 0,
    };

    let all_rows: Vec<usize> = match row_filter {
        Some(indices) => {
            let mut v = vec![header_row];
            v.extend_from_slice(indices);
            v
        }
        None => (0..rows).collect(),
    };

    for row in all_rows {
        let value = sheet.get_display(row, output_col);
        output.extend_from_slice(value.as_bytes());
        output.push(b'\n');
    }

    Ok(output)
}

fn get_data_bounds(sheet: &visigrid_engine::sheet::Sheet) -> (usize, usize) {
    let mut max_row = 0;
    let mut max_col = 0;

    for row in 0..sheet.rows {
        for col in 0..sheet.cols {
            if !sheet.get_display(row, col).is_empty() {
                max_row = max_row.max(row + 1);
                max_col = max_col.max(col + 1);
            }
        }
    }

    (max_row, max_col)
}

// ============================================================================
// calc
// ============================================================================

fn cmd_calc(
    formula: String,
    from: Format,
    into: String,
    delimiter: char,
    headers: bool,
    spill: Option<SpillFormat>,
) -> Result<(), CliError> {
    // Parse --into cell reference
    let (into_row, into_col) = parse_cell_ref(&into)
        .ok_or_else(|| CliError::args(format!("invalid cell reference: {}", into)))?;

    // Read stdin with offset
    let mut sheet = read_stdin(from, delimiter, into_row, into_col)?;

    // Get data bounds (relative to where we loaded)
    let (data_rows, data_cols) = get_data_bounds(&sheet);

    // If headers, the actual data starts one row after into_row
    // Column refs like A:A should expand to A<start>:A<end> excluding header
    let data_start_row = if headers { into_row + 2 } else { into_row + 1 }; // 1-indexed for formula

    // Translate column references like A:A to explicit ranges
    let formula_str = if formula.starts_with('=') {
        translate_column_refs(&formula, data_start_row, data_rows)
    } else {
        translate_column_refs(&format!("={}", formula), data_start_row, data_rows)
    };

    // Put the formula in a cell outside the data area
    let formula_row = data_rows;
    let formula_col = data_cols;
    sheet.set_value(formula_row, formula_col, &formula_str);

    // Get the result
    let result = sheet.get_display(formula_row, formula_col);

    // Check for error tokens
    if result.starts_with('#') {
        // Formula error - print to stdout, diagnostic to stderr
        println!("{}", result);
        let hint = match result.as_str() {
            "#REF!" => "a cell reference is out of range; check your formula references",
            "#NAME?" => "unrecognized function name; run visigrid-cli list-functions to see all available",
            "#VALUE!" => "wrong argument type; check that referenced cells contain the expected data",
            "#DIV/0!" => "division by zero in your formula",
            "#N/A" => "lookup function did not find a match",
            _ => "check your formula syntax and cell references",
        };
        return Err(CliError::eval(format!("formula returned {}", result))
            .with_hint(hint));
    }

    // Check if result is a spill (array) by checking adjacent cells
    // The engine stores spill results in adjacent cells
    let spill_bounds = detect_spill(&sheet, formula_row, formula_col);

    if let Some((spill_rows, spill_cols)) = spill_bounds {
        if spill_rows * spill_cols > 1 {
            // Result is an array
            match spill {
                None => {
                    return Err(CliError::eval(format!(
                        "result is {}x{} array, use --spill csv or --spill json",
                        spill_rows, spill_cols
                    )));
                }
                Some(SpillFormat::Csv) => {
                    let csv_output = format_spill_csv(&sheet, formula_row, formula_col, spill_rows, spill_cols);
                    print!("{}", csv_output);
                }
                Some(SpillFormat::Json) => {
                    let json_output = format_spill_json(&sheet, formula_row, formula_col, spill_rows, spill_cols);
                    println!("{}", json_output);
                }
            }
            return Ok(());
        }
    }

    // Scalar result (or 1x1 array, which is treated as scalar)
    println!("{}", format_output_value(&result));

    Ok(())
}

fn parse_cell_ref(s: &str) -> Option<(usize, usize)> {
    let s = s.to_uppercase();
    let mut col_str = String::new();
    let mut row_str = String::new();

    for c in s.chars() {
        if c.is_ascii_alphabetic() {
            col_str.push(c);
        } else if c.is_ascii_digit() {
            row_str.push(c);
        } else {
            return None;
        }
    }

    if col_str.is_empty() || row_str.is_empty() {
        return None;
    }

    // Convert column letters to index (A=0, B=1, ..., Z=25, AA=26, ...)
    let mut col: usize = 0;
    for c in col_str.chars() {
        col = col * 26 + (c as usize - 'A' as usize + 1);
    }
    col -= 1; // 0-indexed

    // Convert row to index (1-indexed in input, 0-indexed internally)
    let row: usize = row_str.parse().ok()?;
    if row == 0 {
        return None;
    }

    Some((row - 1, col))
}

fn format_output_value(value: &str) -> String {
    // Try to parse as number and format according to spec:
    // - Integers without decimal point
    // - Floats with minimal representation
    if let Ok(n) = value.parse::<f64>() {
        if n.fract() == 0.0 && n.abs() < i64::MAX as f64 {
            // Integer
            format!("{}", n as i64)
        } else {
            // Float - use default formatting which gives minimal representation
            format!("{}", n)
        }
    } else {
        value.to_string()
    }
}

// ============================================================================
// Column reference translation (A:A → A1:A<max_row>)
// ============================================================================

fn translate_column_refs(formula: &str, start_row: usize, end_row: usize) -> String {
    use std::collections::HashSet;

    // Patterns to translate: A:A, $A:$A, $A:A, A:$A, A:B, etc.
    // Translates to A<start_row>:A<end_row> (1-indexed)
    let mut result = formula.to_string();
    let mut seen: HashSet<String> = HashSet::new();

    let chars: Vec<char> = formula.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        // Check for optional $ before column
        let dollar1 = if i < chars.len() && chars[i] == '$' {
            i += 1;
            true
        } else {
            false
        };

        // Look for letter sequence
        if i < chars.len() && chars[i].is_ascii_alphabetic() {
            let mut col1 = String::new();

            // Collect first column letters
            while i < chars.len() && chars[i].is_ascii_alphabetic() {
                col1.push(chars[i].to_ascii_uppercase());
                i += 1;
            }

            // Check for colon (not followed by digit = column ref, not cell ref like A1:B2)
            if i < chars.len() && chars[i] == ':' {
                i += 1;

                // Check for optional $ before second column
                let dollar2 = if i < chars.len() && chars[i] == '$' {
                    i += 1;
                    true
                } else {
                    false
                };

                let mut col2 = String::new();

                // Collect second column letters
                while i < chars.len() && chars[i].is_ascii_alphabetic() {
                    col2.push(chars[i].to_ascii_uppercase());
                    i += 1;
                }

                // It's a column reference if:
                // - col2 is not empty
                // - Next char is NOT a digit (otherwise it's like A:A1 which is invalid/different)
                if !col2.is_empty() && (i >= chars.len() || !chars[i].is_ascii_digit()) {
                    // Build the original pattern with $ signs
                    let pattern = format!(
                        "{}{}:{}{}",
                        if dollar1 { "$" } else { "" },
                        col1,
                        if dollar2 { "$" } else { "" },
                        col2
                    );

                    if !seen.contains(&pattern) {
                        seen.insert(pattern.clone());
                        // Preserve $ in output, use start_row and end_row
                        let replacement = format!(
                            "{}{}{}:{}{}{}",
                            if dollar1 { "$" } else { "" },
                            col1,
                            start_row,
                            if dollar2 { "$" } else { "" },
                            col2,
                            end_row
                        );
                        result = result.replace(&pattern, &replacement);
                    }
                }
            }
        } else {
            i += 1;
        }
    }

    result
}

// ============================================================================
// Spill detection and formatting
// ============================================================================

fn detect_spill(sheet: &visigrid_engine::sheet::Sheet, start_row: usize, start_col: usize) -> Option<(usize, usize)> {
    // Check if there are adjacent non-empty cells that form a rectangular spill
    // This is a heuristic - the engine doesn't explicitly mark spill boundaries

    // First check if the formula cell itself has a value
    let first_val = sheet.get_display(start_row, start_col);
    if first_val.is_empty() {
        return None;
    }

    // Scan right to find width
    let mut width = 1;
    for col in (start_col + 1)..sheet.cols {
        let val = sheet.get_display(start_row, col);
        if val.is_empty() {
            break;
        }
        width += 1;
    }

    // Scan down to find height
    let mut height = 1;
    for row in (start_row + 1)..sheet.rows {
        // Check if this row has values in all columns of the spill
        let mut row_has_values = false;
        for col in start_col..(start_col + width) {
            if !sheet.get_display(row, col).is_empty() {
                row_has_values = true;
                break;
            }
        }
        if !row_has_values {
            break;
        }
        height += 1;
    }

    Some((height, width))
}

fn format_spill_csv(
    sheet: &visigrid_engine::sheet::Sheet,
    start_row: usize,
    start_col: usize,
    rows: usize,
    cols: usize,
) -> String {
    let mut output = String::new();

    for r in 0..rows {
        for c in 0..cols {
            let val = sheet.get_display(start_row + r, start_col + c);
            // RFC 4180 quoting
            let needs_quote = val.contains(',') || val.contains('"') || val.contains('\n');
            if needs_quote {
                output.push('"');
                output.push_str(&val.replace('"', "\"\""));
                output.push('"');
            } else {
                output.push_str(&val);
            }
            if c < cols - 1 {
                output.push(',');
            }
        }
        output.push('\n');
    }

    output
}

fn format_spill_json(
    sheet: &visigrid_engine::sheet::Sheet,
    start_row: usize,
    start_col: usize,
    rows: usize,
    cols: usize,
) -> String {
    let mut result: Vec<Vec<serde_json::Value>> = Vec::new();

    for r in 0..rows {
        let mut row_vec: Vec<serde_json::Value> = Vec::new();
        for c in 0..cols {
            let val = sheet.get_display(start_row + r, start_col + c);
            // Try to parse as number, otherwise string
            if let Ok(n) = val.parse::<f64>() {
                row_vec.push(serde_json::json!(n));
            } else if val == "TRUE" {
                row_vec.push(serde_json::json!(true));
            } else if val == "FALSE" {
                row_vec.push(serde_json::json!(false));
            } else {
                row_vec.push(serde_json::json!(val));
            }
        }
        result.push(row_vec);
    }

    serde_json::to_string_pretty(&result).unwrap_or_else(|_| "[]".to_string())
}

// ============================================================================
// open
// ============================================================================

fn cmd_open(file: Option<PathBuf>) -> Result<(), CliError> {
    // Find GUI binary
    let gui_binary = if cfg!(target_os = "macos") {
        // Try to find VisiGrid.app
        let app_paths = [
            "/Applications/VisiGrid.app/Contents/MacOS/VisiGrid",
            "~/Applications/VisiGrid.app/Contents/MacOS/VisiGrid",
        ];
        app_paths.iter()
            .map(|p| shellexpand::tilde(p).to_string())
            .find(|p| std::path::Path::new(p).exists())
            .or_else(|| which::which("visigrid-gui").ok().map(|p| p.to_string_lossy().to_string()))
    } else {
        // Linux/Windows - look for visigrid-gui in PATH
        which::which("visigrid-gui").ok().map(|p| p.to_string_lossy().to_string())
    };

    match gui_binary {
        Some(binary) => {
            let mut cmd = std::process::Command::new(&binary);
            if let Some(path) = file {
                cmd.arg(path);
            }
            cmd.spawn().map_err(|e| CliError::io(format!("failed to launch GUI: {}", e)))?;
            Ok(())
        }
        None => {
            Err(CliError::io("GUI binary not found. Install VisiGrid GUI or add visigrid-gui to PATH."))
        }
    }
}

// ============================================================================
// replay (Phase 9B)
// ============================================================================

// ============================================================================
// diff
// ============================================================================

// Diff exit codes imported from exit_codes.rs registry

#[allow(clippy::too_many_arguments)]
fn cmd_diff(
    left_arg: String,
    right_arg: String,
    key: String,
    match_mode: DiffMatchMode,
    key_transform: DiffKeyTransform,
    compare: Option<String>,
    tolerance: f64,
    on_ambiguous: DiffAmbiguousPolicy,
    out: DiffOutputFormat,
    output: Option<PathBuf>,
    summary_mode: DiffSummaryMode,
    no_headers: bool,
    header_row: Option<usize>,
    delimiter: char,
    stdin_format: Option<Format>,
    strict_exit: bool,
    quiet: bool,
    save_ambiguous: Option<PathBuf>,
) -> Result<(), CliError> {
    let left_is_stdin = left_arg == "-";
    let right_is_stdin = right_arg == "-";

    if left_is_stdin && right_is_stdin {
        return Err(CliError::args("cannot read both sides from stdin")
            .with_hint("provide at least one file path: visigrid diff - file.csv --key id"));
    }

    // Resolve formats
    let left_path = if left_is_stdin { None } else { Some(PathBuf::from(&left_arg)) };
    let right_path = if right_is_stdin { None } else { Some(PathBuf::from(&right_arg)) };

    let resolve_stdin_format = |other_path: &Option<PathBuf>| -> Result<Format, CliError> {
        if let Some(fmt) = stdin_format {
            return Ok(fmt);
        }
        if let Some(ref p) = other_path {
            return infer_format(p);
        }
        Err(CliError::args("cannot infer stdin format")
            .with_hint("use --stdin-format to specify the format for stdin input"))
    };

    // Load both sides
    let (left_sheet, left_label) = if left_is_stdin {
        let fmt = resolve_stdin_format(&right_path)?;
        (read_stdin(fmt, delimiter, 0, 0)?, "stdin".to_string())
    } else {
        let p = left_path.as_ref().unwrap();
        let fmt = infer_format(p)?;
        let label = p.display().to_string();
        (read_file(p, fmt, delimiter)?, label)
    };

    let (right_sheet, right_label) = if right_is_stdin {
        let fmt = resolve_stdin_format(&left_path)?;
        (read_stdin(fmt, delimiter, 0, 0)?, "stdin".to_string())
    } else {
        let p = right_path.as_ref().unwrap();
        let fmt = infer_format(p)?;
        let label = p.display().to_string();
        (read_file(p, fmt, delimiter)?, label)
    };

    let (left_bounds_rows, left_bounds_cols) = get_data_bounds(&left_sheet);
    let (right_bounds_rows, right_bounds_cols) = get_data_bounds(&right_sheet);

    if left_bounds_rows == 0 {
        return Err(CliError { code: EXIT_DIFF_PARSE, message: format!("{}: empty or has no data rows", left_label), hint: None });
    }
    if right_bounds_rows == 0 {
        return Err(CliError { code: EXIT_DIFF_PARSE, message: format!("{}: empty or has no data rows", right_label), hint: None });
    }

    // Determine header row (0-indexed internally)
    let hdr_row = if no_headers {
        None
    } else {
        Some(header_row.map(|h| h.saturating_sub(1)).unwrap_or(0))
    };

    // Extract headers
    let max_cols = left_bounds_cols.max(right_bounds_cols);
    let headers: Vec<String> = if let Some(hr) = hdr_row {
        (0..max_cols)
            .map(|c| {
                let lh = left_sheet.get_display(hr, c);
                if !lh.is_empty() {
                    lh
                } else {
                    right_sheet.get_display(hr, c)
                }
            })
            .collect()
    } else {
        // Generate A, B, C, ... headers
        (0..max_cols).map(|c| col_letter(c)).collect()
    };

    // Resolve key column
    let key_col = resolve_column(&key, &headers)?;

    // Resolve compare columns
    let compare_cols = match &compare {
        Some(spec) => {
            let mut cols = Vec::new();
            for part in spec.split(',') {
                let part = part.trim();
                cols.push(resolve_column(part, &headers)?);
            }
            Some(cols)
        }
        None => None,
    };

    // Convert match mode
    let mode = match match_mode {
        DiffMatchMode::Exact => diff::MatchMode::Exact,
        DiffMatchMode::Contains => diff::MatchMode::Contains,
    };

    let kt = match key_transform {
        DiffKeyTransform::None => diff::KeyTransform::None,
        DiffKeyTransform::Trim => diff::KeyTransform::Trim,
        DiffKeyTransform::Digits => diff::KeyTransform::Digits,
    };

    let amb = match on_ambiguous {
        DiffAmbiguousPolicy::Error => diff::AmbiguityPolicy::Error,
        DiffAmbiguousPolicy::Report => diff::AmbiguityPolicy::Report,
    };

    let options = diff::DiffOptions {
        key_col,
        compare_cols,
        match_mode: mode,
        key_transform: kt,
        on_ambiguous: amb,
        tolerance,
    };

    // Extract data rows
    let data_start = hdr_row.map(|h| h + 1).unwrap_or(0);
    let left_rows = extract_data_rows(&left_sheet, data_start, left_bounds_rows, left_bounds_cols, &headers, &options);
    let right_rows = extract_data_rows(&right_sheet, data_start, right_bounds_rows, right_bounds_cols, &headers, &options);

    // Warn when using substring matching
    if !quiet && mode == diff::MatchMode::Contains {
        eprintln!("warning: using substring matching (--match contains); ensure keys are normalized");
    }

    // Run reconciliation
    let result = match diff::reconcile(&left_rows, &right_rows, &headers, &options) {
        Ok(r) => r,
        Err(diff::DiffError::DuplicateKeys(dups)) => {
            let mut msg = String::from("duplicate keys found:\n");
            for dup in &dups {
                msg.push_str(&format!("  {} key {:?} appears {} times\n", dup.side.as_str(), dup.key, dup.count));
            }
            return Err(CliError {
                code: EXIT_DIFF_DUPLICATE,
                message: msg.trim_end().to_string(),
                hint: Some("each key must be unique within its file; deduplicate or choose a different --key column".to_string()),
            });
        }
    };

    // Save ambiguous matches to CSV (before error exit, so the file is always written)
    if let Some(ref amb_path) = save_ambiguous {
        if !result.ambiguous_keys.is_empty() {
            write_ambiguous_csv(amb_path, &result.ambiguous_keys)?;
            if !quiet {
                eprintln!("ambiguous matches exported to: {}", amb_path.display());
            }
        }
    }

    // Check ambiguous error condition
    if !result.ambiguous_keys.is_empty() && amb == diff::AmbiguityPolicy::Error {
        let mut msg = String::from("ambiguous matches found:\n");
        for ak in &result.ambiguous_keys {
            msg.push_str(&format!("  key {:?} matches {} right rows:", ak.key, ak.candidates.len()));
            for c in &ak.candidates {
                msg.push_str(&format!(" {:?}(row {})", c.right_key_raw, c.right_row_index));
            }
            msg.push('\n');
        }
        return Err(CliError {
            code: EXIT_DIFF_AMBIGUOUS,
            message: msg.trim_end().to_string(),
            hint: Some("use --on_ambiguous report to include ambiguous matches in output instead of failing".to_string()),
        });
    }

    // Format output
    let output_bytes = match out {
        DiffOutputFormat::Json => format_diff_json(&result, &options, &headers, &summary_mode)?,
        DiffOutputFormat::Csv => format_diff_csv(&result, &options)?,
    };

    // Write output
    match output {
        Some(path) => {
            std::fs::write(&path, &output_bytes)
                .map_err(|e| CliError::io(format!("{}: {}", path.display(), e)))?;
        }
        None => {
            io::stdout()
                .write_all(&output_bytes)
                .map_err(|e| CliError::io(e.to_string()))?;
        }
    }

    // Write summary to stderr if requested (--quiet suppresses)
    if !quiet && matches!(summary_mode, DiffSummaryMode::Stderr) {
        let s = &result.summary;
        eprintln!("left:  {} rows ({})", s.left_rows, left_label);
        eprintln!("right: {} rows ({})", s.right_rows, right_label);
        eprintln!("matched: {}", s.matched);
        eprintln!("only_left: {}", s.only_left);
        eprintln!("only_right: {}", s.only_right);
        eprintln!("value_diff: {}", s.diff);
        if s.diff > 0 && s.diff != s.diff_outside_tolerance {
            eprintln!("value_diff_outside_tolerance: {}", s.diff_outside_tolerance);
        }
        if s.ambiguous > 0 {
            eprintln!("ambiguous: {}", s.ambiguous);
        }
    }

    // Exit 1 for material differences: missing rows or diffs outside tolerance.
    // Within-tolerance diffs are reported but do not cause a non-zero exit code.
    // --strict-exit: any diff (even within tolerance) causes exit 1.
    let s = &result.summary;
    let diff_count = if strict_exit { s.diff } else { s.diff_outside_tolerance };
    if s.only_left > 0 || s.only_right > 0 || diff_count > 0 {
        return Err(CliError { code: EXIT_EVAL_ERROR, message: String::new(), hint: None });
    }

    Ok(())
}

fn resolve_column(spec: &str, headers: &[String]) -> Result<usize, CliError> {
    // Try by name first (case-insensitive)
    let spec_lower = spec.to_lowercase();
    for (i, h) in headers.iter().enumerate() {
        if h.to_lowercase() == spec_lower {
            return Ok(i);
        }
    }

    // Try as column letter (A=0, B=1, ...)
    if spec.chars().all(|c| c.is_ascii_alphabetic()) {
        let upper = spec.to_uppercase();
        let mut col: usize = 0;
        for c in upper.chars() {
            col = col * 26 + (c as usize - 'A' as usize + 1);
        }
        let idx = col - 1;
        if idx < headers.len() {
            return Ok(idx);
        }
    }

    // Try as 1-indexed number
    if let Ok(n) = spec.parse::<usize>() {
        if n >= 1 && n <= headers.len() {
            return Ok(n - 1);
        }
    }

    let available: Vec<&str> = headers.iter().map(|h| h.as_str()).collect();
    Err(CliError::args(format!("unknown column: {:?}", spec))
        .with_hint(format!("available columns: {}", available.join(", "))))
}

fn col_letter(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;
    loop {
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    result
}

fn extract_data_rows(
    sheet: &visigrid_engine::sheet::Sheet,
    data_start: usize,
    bounds_rows: usize,
    bounds_cols: usize,
    headers: &[String],
    options: &diff::DiffOptions,
) -> Vec<diff::DataRow> {
    let mut rows = Vec::new();
    for r in data_start..bounds_rows {
        // Skip blank rows
        let mut all_blank = true;
        for c in 0..bounds_cols {
            if !sheet.get_display(r, c).is_empty() {
                all_blank = false;
                break;
            }
        }
        if all_blank {
            continue;
        }

        let key_raw = sheet.get_display(r, options.key_col);
        let key_norm = diff::apply_key_transform(&key_raw, options.key_transform);

        let mut values = HashMap::new();
        for (c, header) in headers.iter().enumerate() {
            if c < bounds_cols {
                values.insert(header.clone(), sheet.get_display(r, c));
            }
        }

        rows.push(diff::DataRow {
            key_raw,
            key_norm,
            values,
        });
    }
    rows
}

const DIFF_CONTRACT_VERSION: u32 = 1;

fn format_diff_json(
    result: &diff::DiffResult,
    options: &diff::DiffOptions,
    headers: &[String],
    _summary_mode: &DiffSummaryMode,
) -> Result<Vec<u8>, CliError> {
    let key_name = headers.get(options.key_col).cloned().unwrap_or_default();
    let match_str = match options.match_mode {
        diff::MatchMode::Exact => "exact",
        diff::MatchMode::Contains => "contains",
    };
    let kt_str = match options.key_transform {
        diff::KeyTransform::None => "none",
        diff::KeyTransform::Trim => "trim",
        diff::KeyTransform::Digits => "digits",
    };

    // Build results array
    let results_json: Vec<serde_json::Value> = result.results.iter().map(|row| {
        let diffs_json = if row.diffs.is_empty() {
            serde_json::Value::Null
        } else {
            serde_json::json!(row.diffs.iter().map(|d| {
                let mut m = serde_json::Map::new();
                m.insert("column".to_string(), serde_json::json!(d.column));
                m.insert("left".to_string(), serde_json::json!(d.left));
                m.insert("right".to_string(), serde_json::json!(d.right));
                m.insert("delta".to_string(), match d.delta {
                    Some(v) => serde_json::json!(v),
                    None => serde_json::Value::Null,
                });
                m.insert("within_tolerance".to_string(), serde_json::json!(d.within_tolerance));
                serde_json::Value::Object(m)
            }).collect::<Vec<_>>())
        };

        let explain_json = match &row.match_explain {
            Some(e) => serde_json::json!({
                "mode": e.mode,
                "left_key_raw": e.left_key_raw,
                "right_key_raw": e.right_key_raw,
                "left_key_norm": e.left_key_norm,
                "right_key_norm": e.right_key_norm,
            }),
            None => serde_json::Value::Null,
        };

        let candidates_json = match &row.candidates {
            Some(cands) => serde_json::json!(cands.iter().map(|c| {
                serde_json::json!({
                    "right_key_raw": c.right_key_raw,
                    "right_row_index": c.right_row_index,
                })
            }).collect::<Vec<_>>()),
            None => serde_json::Value::Null,
        };

        let left_json = match &row.left {
            Some(vals) => serde_json::json!(vals),
            None => serde_json::Value::Null,
        };
        let right_json = match &row.right {
            Some(vals) => serde_json::json!(vals),
            None => serde_json::Value::Null,
        };

        serde_json::json!({
            "status": row.status.as_str(),
            "key": row.key,
            "left": left_json,
            "right": right_json,
            "diffs": diffs_json,
            "match_explain": explain_json,
            "candidates": candidates_json,
        })
    }).collect();

    // Build top-level object
    let summary_json = serde_json::json!({
        "left_rows": result.summary.left_rows,
        "right_rows": result.summary.right_rows,
        "matched": result.summary.matched,
        "only_left": result.summary.only_left,
        "only_right": result.summary.only_right,
        "diff": result.summary.diff,
        "diff_outside_tolerance": result.summary.diff_outside_tolerance,
        "ambiguous": result.summary.ambiguous,
        "tolerance": options.tolerance,
        "key": key_name,
        "match": match_str,
        "key_transform": kt_str,
    });

    let top = serde_json::json!({
        "contract_version": DIFF_CONTRACT_VERSION,
        "summary": summary_json,
        "results": results_json,
    });

    let mut bytes = serde_json::to_vec_pretty(&top).map_err(|e| CliError::io(e.to_string()))?;
    bytes.push(b'\n');
    Ok(bytes)
}

fn format_diff_csv(
    result: &diff::DiffResult,
    options: &diff::DiffOptions,
) -> Result<Vec<u8>, CliError> {
    let match_str = match options.match_mode {
        diff::MatchMode::Exact => "exact",
        diff::MatchMode::Contains => "contains",
    };

    let mut writer = csv::WriterBuilder::new().from_writer(Vec::new());

    // Header
    writer.write_record(&[
        "status", "key", "column", "left_value", "right_value",
        "delta", "within_tolerance", "match_mode", "match_explain",
    ]).map_err(|e| CliError::io(e.to_string()))?;

    for row in &result.results {
        if row.status == diff::RowStatus::Diff && !row.diffs.is_empty() {
            // One CSV row per column diff
            for d in &row.diffs {
                let explain = match &row.match_explain {
                    Some(e) => format!("{} left={:?} right={:?}", e.mode, e.left_key_raw, e.right_key_raw),
                    None => String::new(),
                };
                writer.write_record(&[
                    row.status.as_str(),
                    &row.key,
                    &d.column,
                    &d.left,
                    &d.right,
                    &d.delta.map(|v| format!("{}", v)).unwrap_or_default(),
                    &d.within_tolerance.to_string(),
                    match_str,
                    &explain,
                ]).map_err(|e| CliError::io(e.to_string()))?;
            }
        } else {
            // One row for the key
            let explain = match &row.match_explain {
                Some(e) => format!("{} left={:?} right={:?}", e.mode, e.left_key_raw, e.right_key_raw),
                None => String::new(),
            };
            writer.write_record(&[
                row.status.as_str(),
                &row.key,
                "",
                "",
                "",
                "",
                "",
                match_str,
                &explain,
            ]).map_err(|e| CliError::io(e.to_string()))?;
        }
    }

    writer.into_inner().map_err(|e| CliError::io(e.to_string()))
}

fn write_ambiguous_csv(path: &PathBuf, ambiguous_keys: &[diff::AmbiguousKey]) -> Result<(), CliError> {
    let mut writer = csv::WriterBuilder::new().from_writer(Vec::new());

    writer.write_record(&[
        "left_key", "candidate_count", "candidate_keys",
    ]).map_err(|e| CliError::io(e.to_string()))?;

    for ak in ambiguous_keys {
        let candidate_keys: Vec<&str> = ak.candidates.iter()
            .map(|c| c.right_key_raw.as_str())
            .collect();
        writer.write_record(&[
            &ak.key,
            &ak.candidates.len().to_string(),
            &candidate_keys.join("|"),
        ]).map_err(|e| CliError::io(e.to_string()))?;
    }

    let bytes = writer.into_inner().map_err(|e| CliError::io(e.to_string()))?;
    std::fs::write(path, &bytes)
        .map_err(|e| CliError::io(format!("{}: {}", path.display(), e)))?;

    Ok(())
}

// ============================================================================
// ai doctor
// ============================================================================

fn cmd_ai_doctor(json: bool, test: bool) -> Result<(), CliError> {
    use visigrid_config::settings::Settings;
    use visigrid_config::ai::{self, ResolvedAIConfig, AIConfigStatus};

    // Use the single source of truth
    let config = ResolvedAIConfig::load();
    let settings = Settings::load();
    let ai_settings = &settings.ai;

    let enabled = config.provider.is_enabled();
    let model_configured = !ai_settings.model.is_empty();
    let model_effective = if enabled {
        config.model.clone()
    } else {
        "(none)".to_string()
    };
    let keychain_available = ai::keychain_available();

    // Map AIConfigStatus to AIDoctorStatus
    let (status, blocking_reason) = match config.status {
        AIConfigStatus::Disabled => (AIDoctorStatus::Disabled, Some("provider=none".to_string())),
        AIConfigStatus::Ready => (AIDoctorStatus::Ready, None),
        AIConfigStatus::NotImplemented => (AIDoctorStatus::Ready, Some("provider not yet implemented".to_string())),
        AIConfigStatus::MissingKey => (AIDoctorStatus::Misconfigured, Some("missing_api_key".to_string())),
        AIConfigStatus::Error => (AIDoctorStatus::Misconfigured, config.blocking_reason.clone()),
    };

    // Context policy from resolved config
    let context_policy = if config.privacy_mode {
        "minimal_values_only"
    } else {
        "values_and_formulas"
    };

    // Build diagnostics from resolved config
    let diag = AIDoctorReport {
        enabled,
        provider: config.provider_name().to_string(),
        model_configured,
        model_effective,
        privacy_mode: config.privacy_mode,
        context_policy: context_policy.to_string(),
        allow_proposals: config.allow_proposals,
        key_present: config.api_key.is_some(),
        key_source: config.key_source,
        keychain_available,
        endpoint: config.endpoint.clone(),
        status,
        blocking_reason,
        test_skipped: !test,
        test_result: None,
    };

    // Run config validation if requested
    let diag = if test {
        let result = config.validate_config();
        let mut d = diag;
        d.test_skipped = false;
        d.test_result = Some(result.as_str().to_string());
        d
    } else {
        diag
    };

    // Output
    if json {
        let json_output = serde_json::json!({
            "schema_version": 1,
            "status": diag.status.as_str(),
            "blocking_reason": diag.blocking_reason,
            "enabled": diag.enabled,
            "provider": diag.provider,
            "model_configured": diag.model_configured,
            "model_effective": diag.model_effective,
            "privacy_mode": diag.privacy_mode,
            "context_policy": diag.context_policy,
            "allow_proposals": diag.allow_proposals,
            "key": if diag.key_present { "present" } else { "missing" },
            "key_source": diag.key_source.as_str(),
            "keychain": if diag.keychain_available { "ok" } else { "unavailable" },
            "endpoint": diag.endpoint,
            "test": if diag.test_skipped { "skipped" } else {
                diag.test_result.as_deref().unwrap_or("unknown")
            },
        });
        println!("{}", serde_json::to_string_pretty(&json_output).unwrap());
    } else {
        println!("AI Doctor");
        println!("---------");
        println!("status:          {}", diag.status.as_str());
        if let Some(reason) = &diag.blocking_reason {
            println!("blocking_reason: {}", reason);
        }
        println!("provider:        {}", diag.provider);
        println!("model_configured:{}", diag.model_configured);
        println!("model_effective: {}", diag.model_effective);
        println!("privacy_mode:    {}", diag.privacy_mode);
        println!("context_policy:  {}", diag.context_policy);
        println!("allow_proposals: {}", diag.allow_proposals);
        println!("key:             {}", if diag.key_present { "present" } else { "missing" });
        println!("key_source:      {}", diag.key_source.as_str());
        println!("keychain:        {}", if diag.keychain_available { "ok" } else { "unavailable" });
        if let Some(endpoint) = &diag.endpoint {
            println!("endpoint:        {}", endpoint);
        }
        if diag.test_skipped {
            println!("test:            skipped (use --test)");
        } else if let Some(result) = &diag.test_result {
            println!("test:            {}", result);
        }

        // Actionable fix suggestions
        if let Some(reason) = &diag.blocking_reason {
            println!();
            match reason.as_str() {
                "provider=none" => {
                    println!("AI is disabled. To enable:");
                    println!("  Set ai.provider in ~/.config/visigrid/settings.json");
                }
                "missing_api_key" => {
                    println!("Fix: set {} or store key in keychain",
                        format!("VISIGRID_{}_KEY", diag.provider.to_uppercase()));
                }
                _ => {}
            }
        }
    }

    // Determine exit code based on status
    match diag.status {
        AIDoctorStatus::Disabled => {
            Err(CliError { code: EXIT_AI_DISABLED, message: "AI is disabled".to_string(), hint: None })
        }
        AIDoctorStatus::Misconfigured => {
            let reason = diag.blocking_reason.unwrap_or_else(|| "unknown".to_string());
            Err(CliError { code: EXIT_AI_MISSING_KEY, message: format!("AI misconfigured: {}", reason), hint: None })
        }
        AIDoctorStatus::Ready => Ok(()),
    }
}

struct AIDoctorReport {
    enabled: bool,
    provider: String,
    model_configured: bool,
    model_effective: String,
    privacy_mode: bool,
    context_policy: String,
    allow_proposals: bool,
    key_present: bool,
    key_source: visigrid_config::ai::KeySource,
    keychain_available: bool,
    endpoint: Option<String>,
    status: AIDoctorStatus,
    blocking_reason: Option<String>,
    test_skipped: bool,
    test_result: Option<String>,
}

#[derive(Clone, Copy)]
enum AIDoctorStatus {
    Disabled,
    Misconfigured,
    Ready,
}

impl AIDoctorStatus {
    fn as_str(&self) -> &'static str {
        match self {
            AIDoctorStatus::Disabled => "disabled",
            AIDoctorStatus::Misconfigured => "misconfigured",
            AIDoctorStatus::Ready => "ready",
        }
    }
}

fn cmd_replay(
    script: PathBuf,
    verify: bool,
    output: Option<PathBuf>,
    format: Option<String>,
    fingerprint_only: bool,
    quiet: bool,
) -> Result<(), CliError> {
    // Execute the script
    let result = replay::execute_script(&script)?;

    // Handle --fingerprint flag
    if fingerprint_only {
        if result.has_nondeterministic {
            // Warn about nondeterministic functions but still print fingerprint
            eprintln!("warning: script contains nondeterministic functions: {}",
                result.nondeterministic_found.join(", "));
            eprintln!("warning: fingerprint will vary between runs");
        }
        println!("{}", result.fingerprint.to_string());
        return Ok(());
    }

    // Fail early if --verify is used with nondeterministic functions
    if verify && result.has_nondeterministic {
        return Err(CliError::eval(format!(
            "cannot verify: script contains nondeterministic functions ({})",
            result.nondeterministic_found.join(", ")
        )).with_hint("remove NOW(), TODAY(), RAND(), RANDBETWEEN() from formulas, or run without --verify"));
    }

    // Print result summary (unless quiet)
    if !quiet {
        // Print notes for hashed-only operations
        for note in &result.hashed_only_notes {
            eprintln!("note: hashed (not applied): {}", note);
        }

        eprintln!("Replayed {} operations", result.operations);
        eprintln!("Fingerprint: {}", result.fingerprint.to_string());

        if result.has_nondeterministic {
            eprintln!("Warning: nondeterministic functions used: {}",
                result.nondeterministic_found.join(", "));
        }

        if let Some(ref expected) = result.expected_fingerprint {
            if result.has_nondeterministic {
                eprintln!("Verification: SKIP (nondeterministic functions present)");
            } else if result.verified {
                eprintln!("Verification: PASS (matches expected)");
            } else {
                eprintln!("Verification: FAIL");
                eprintln!("  Expected: {}", expected.to_string());
                eprintln!("  Got:      {}", result.fingerprint.to_string());
            }
        } else {
            eprintln!("Verification: SKIP (no expected fingerprint in script)");
        }
    }

    // Check verification failure
    if verify && !result.verified {
        return Err(CliError::eval("fingerprint verification failed")
            .with_hint("the script or its source data may have been modified since the fingerprint was recorded"));
    }

    // Export output if requested
    if let Some(output_path) = output {
        let is_stdout = output_path.as_os_str() == "-";

        // Infer format from extension if not specified (default csv for stdout)
        let fmt = format.unwrap_or_else(|| {
            if is_stdout {
                "csv".to_string()
            } else {
                output_path
                    .extension()
                    .and_then(|e| e.to_str())
                    .map(|s| s.to_lowercase())
                    .unwrap_or_else(|| "csv".to_string())
            }
        });

        if is_stdout {
            let bytes = replay::export_to_bytes(&result.workbook, &fmt)?;
            io::stdout()
                .write_all(&bytes)
                .map_err(|e| CliError::io(e.to_string()))?;
        } else {
            replay::export_workbook(&result.workbook, &output_path, &fmt)?;
            if !quiet {
                eprintln!("Wrote output to: {}", output_path.display());
            }
        }
    }

    Ok(())
}

// ============================================================================
// Session commands
// ============================================================================

fn cmd_sessions(json: bool) -> Result<(), CliError> {
    let sessions = session::list_sessions()
        .map_err(|e| CliError::io(format!("failed to list sessions: {}", e)))?;

    if sessions.is_empty() {
        if json {
            println!("[]");
        } else {
            eprintln!("No running VisiGrid sessions found.");
        }
        return Ok(());
    }

    if json {
        let output = serde_json::to_string_pretty(&sessions)
            .map_err(|e| CliError::io(e.to_string()))?;
        println!("{}", output);
    } else {
        // Table format
        println!("{:<12} {:>6} {:>8} {:<24} {}",
            "SESSION", "PORT", "PID", "CREATED", "WORKBOOK");
        println!("{}", "-".repeat(80));

        for s in &sessions {
            let short_id = &s.session_id.to_string()[..8];
            let created = s.created_at.format("%Y-%m-%d %H:%M:%S");
            let workbook = s.workbook_path.as_ref()
                .and_then(|p| p.file_name())
                .and_then(|n| n.to_str())
                .unwrap_or(&s.workbook_title);

            println!("{:<12} {:>6} {:>8} {:<24} {}",
                short_id,
                s.port,
                s.pid,
                created,
                workbook);
        }

        eprintln!();
        eprintln!("{} session(s) found", sessions.len());
    }

    Ok(())
}

fn cmd_attach(session_id: Option<String>) -> Result<(), CliError> {
    let discovery = resolve_session(session_id.as_deref())?;
    let token = get_session_token()?;

    let client = session::SessionClient::connect(&discovery, &token)
        .map_err(CliError::session)?;

    println!("Connected to session {}", discovery.session_id);
    println!("  Revision:     {}", client.revision());
    println!("  Capabilities: {}", client.capabilities().join(", "));
    println!("  Workbook:     {}", discovery.workbook_title);
    if let Some(ref path) = discovery.workbook_path {
        println!("  Path:         {}", path.display());
    }

    Ok(())
}

fn cmd_apply(
    ops_arg: String,
    session_id: Option<String>,
    atomic: bool,
    expected_revision: Option<u64>,
    wait: bool,
    wait_timeout: u64,
) -> Result<(), CliError> {
    use std::time::{Duration, Instant};

    // Safety guard: --wait without idempotency protection is a footgun
    if wait && !atomic && expected_revision.is_none() {
        return Err(CliError {
            code: exit_codes::EXIT_USAGE,
            message: "--wait requires --atomic or --expected-revision for safety".to_string(),
            hint: Some("retrying non-atomic ops without revision check can cause double-apply".to_string()),
        });
    }

    let discovery = resolve_session(session_id.as_deref())?;
    let token = get_session_token()?;

    // Read ops from file or stdin (before connecting, so we don't hold connection while reading)
    let ops_json = if ops_arg == "-" {
        let mut buf = String::new();
        io::stdin().read_to_string(&mut buf)
            .map_err(|e| CliError::io(format!("failed to read stdin: {}", e)))?;
        buf
    } else {
        std::fs::read_to_string(&ops_arg)
            .map_err(|e| CliError::io(format!("failed to read {}: {}", ops_arg, e)))?
    };

    // Parse ops - support both JSONL (one op per line) and JSON array
    let ops: Vec<session::Op> = if ops_json.trim_start().starts_with('[') {
        // JSON array format
        serde_json::from_str(&ops_json)
            .map_err(|e| CliError::parse(format!("failed to parse ops JSON: {}", e)))?
    } else {
        // JSONL format (one op per line)
        ops_json
            .lines()
            .filter(|line| !line.trim().is_empty())
            .enumerate()
            .map(|(i, line)| {
                serde_json::from_str(line)
                    .map_err(|e| CliError::parse(format!("line {}: {}", i + 1, e)))
            })
            .collect::<Result<Vec<_>, _>>()?
    };

    if ops.is_empty() {
        eprintln!("No operations to apply");
        return Ok(());
    }

    eprintln!("Applying {} operation(s)...", ops.len());

    let deadline = if wait {
        Some(Instant::now() + Duration::from_secs(wait_timeout))
    } else {
        None
    };

    // Connect once; reuse for retries (saves connection slots)
    let mut client = session::SessionClient::connect(&discovery, &token)
        .map_err(CliError::session)?;

    // Retry loop for writer conflicts
    loop {
        let result = client.apply_ops(ops.clone(), atomic, expected_revision);

        match result {
            Ok(result) => {
                if let Some(ref err) = result.error {
                    eprintln!("Error at op {}: [{}] {}", err.op_index, err.code, err.message);
                    if let Some(ref hint) = err.suggestion {
                        eprintln!("  Suggestion: {}", hint);
                    }
                    eprintln!("Applied: {}/{}", result.applied, result.total);
                    eprintln!("Revision: {}", result.revision);
                    // Partial apply = exit 24 (EXIT_SESSION_PARTIAL)
                    return Err(CliError {
                        code: exit_codes::EXIT_SESSION_PARTIAL,
                        message: "operation failed".to_string(),
                        hint: None,
                    });
                }

                println!("Applied: {}/{}", result.applied, result.total);
                println!("Revision: {}", result.revision);
                return Ok(());
            }
            Err(session::SessionError::ServerError { code, message, retry_after_ms }) if code == "writer_conflict" => {
                if let Some(deadline) = deadline {
                    if Instant::now() >= deadline {
                        eprintln!("error: writer conflict (timeout after {}s)", wait_timeout);
                        return Err(CliError {
                            code: exit_codes::EXIT_SESSION_CONFLICT,
                            message: format!("writer conflict: {}", message),
                            hint: Some("another client holds the writer lease".to_string()),
                        });
                    }

                    // Adaptive backoff: use server hint, clamp to [50ms, 2000ms], add jitter
                    let base_ms = retry_after_ms.unwrap_or(1000).clamp(50, 2000);
                    let jitter = (base_ms as f64 * 0.1 * rand_jitter()) as u64;
                    let sleep_ms = base_ms + jitter;

                    eprintln!("Writer conflict, retrying in {}ms...", sleep_ms);
                    std::thread::sleep(Duration::from_millis(sleep_ms));
                    continue;
                } else {
                    // No --wait, fail immediately
                    return Err(CliError::session(session::SessionError::ServerError {
                        code,
                        message,
                        retry_after_ms,
                    }));
                }
            }
            Err(session::SessionError::ConnectionClosed) if wait => {
                // Connection dropped; reconnect and retry
                eprintln!("Connection lost, reconnecting...");
                client = session::SessionClient::connect(&discovery, &token)
                    .map_err(CliError::session)?;
                continue;
            }
            Err(e) => {
                return Err(CliError::session(e));
            }
        }
    }
}

/// Simple jitter factor in range [-1.0, 1.0] using timestamp entropy.
/// Not cryptographic, just enough to spread out retries.
fn rand_jitter() -> f64 {
    use std::time::SystemTime;
    let nanos = SystemTime::now()
        .duration_since(SystemTime::UNIX_EPOCH)
        .map(|d| d.subsec_nanos())
        .unwrap_or(0);
    // Map to [-1.0, 1.0]
    ((nanos % 2001) as f64 / 1000.0) - 1.0
}

fn cmd_inspect(
    range: String,
    session_id: Option<String>,
    sheet: usize,
    json: bool,
) -> Result<(), CliError> {
    use visigrid_protocol::InspectResult;

    let discovery = resolve_session(session_id.as_deref())?;
    let token = get_session_token()?;

    let mut client = session::SessionClient::connect(&discovery, &token)
        .map_err(CliError::session)?;

    // Parse the range string into an inspect target
    let result = if range.eq_ignore_ascii_case("workbook") {
        client.inspect_workbook()
    } else if let Some((start, end)) = range.split_once(':') {
        // Range like "A1:B2"
        let (start_col, start_row) = parse_cell_ref(start)
            .ok_or_else(|| CliError::args(format!("invalid cell reference: {}", start)))?;
        let (end_col, end_row) = parse_cell_ref(end)
            .ok_or_else(|| CliError::args(format!("invalid cell reference: {}", end)))?;
        client.inspect_range(sheet, start_row, start_col, end_row, end_col)
    } else {
        // Single cell like "A1"
        let (col, row) = parse_cell_ref(&range)
            .ok_or_else(|| CliError::args(format!("invalid cell reference: {}", range)))?;
        client.inspect_cell(sheet, row, col)
    }.map_err(CliError::session)?;

    if json {
        let output = serde_json::to_string_pretty(&result)
            .map_err(|e| CliError::io(e.to_string()))?;
        println!("{}", output);
    } else {
        // Human-readable table format
        let short_id = &discovery.session_id.to_string()[..8];
        println!("Session {}  Revision {}", short_id, result.revision);

        match result.result {
            InspectResult::Cell(info) => {
                // Single cell format: "A1 = value (type)"
                let cell_type = if info.formula.is_some() {
                    "formula"
                } else if info.display.parse::<f64>().is_ok() {
                    "number"
                } else if info.display.is_empty() {
                    "empty"
                } else {
                    "text"
                };

                println!("{} = {}  ({})", range.to_uppercase(), info.display, cell_type);

                if let Some(formula) = &info.formula {
                    println!("Formula: {}", formula);
                }
            }
            InspectResult::Range { cells } => {
                println!("Range {}\n", range.to_uppercase());

                if cells.is_empty() {
                    println!("(empty range)");
                } else {
                    // Simple column display - one cell per line with truncation
                    for (i, cell) in cells.iter().enumerate() {
                        let display = if cell.display.len() > 40 {
                            format!("{}…", &cell.display[..39])
                        } else {
                            cell.display.clone()
                        };

                        let formula_marker = if cell.formula.is_some() { " [f]" } else { "" };
                        println!("  [{}] {}{}", i, display, formula_marker);
                    }
                }
            }
            InspectResult::Workbook(info) => {
                println!("\nWorkbook: {}", info.title);
                println!("  Sheets:       {}", info.sheet_count);
                println!("  Active sheet: {}", info.active_sheet);
            }
        }
    }

    Ok(())
}

fn cmd_stats(session_id: Option<String>, json: bool) -> Result<(), CliError> {
    let discovery = resolve_session(session_id.as_deref())?;
    let token = get_session_token()?;

    let mut client = session::SessionClient::connect(&discovery, &token)
        .map_err(CliError::session)?;

    let stats = client.stats()
        .map_err(CliError::session)?;

    if json {
        let output = serde_json::to_string_pretty(&stats)
            .map_err(|e| CliError::io(e.to_string()))?;
        println!("{}", output);
    } else {
        // Human-readable table format
        println!("Session Statistics");
        println!("------------------");
        println!("Active connections:    {}", stats.active_connections);
        println!("Writer conflicts:      {}", stats.writer_conflict_count);
        println!("Dropped events:        {}", stats.dropped_events_total);
        println!("Refused (limit):       {}", stats.connections_refused_limit);
        println!("Parse failures:        {}", stats.connections_closed_parse_failures);
        println!("Oversize messages:     {}", stats.connections_closed_oversize);
    }

    Ok(())
}

fn cmd_view(
    session_id: Option<String>,
    range: String,
    sheet: usize,
    follow: bool,
    col_width: usize,
) -> Result<(), CliError> {
    use std::time::Duration;
    use visigrid_protocol::InspectResult;

    // Parse range (e.g., "A1:J20")
    let (start, end) = range.split_once(':')
        .ok_or_else(|| CliError::args(format!("invalid range '{}', expected format like A1:J20", range)))?;

    let (start_col, start_row) = parse_cell_ref(start)
        .ok_or_else(|| CliError::args(format!("invalid cell reference: {}", start)))?;
    let (end_col, end_row) = parse_cell_ref(end)
        .ok_or_else(|| CliError::args(format!("invalid cell reference: {}", end)))?;

    let discovery = resolve_session(session_id.as_deref())?;
    let token = get_session_token()?;

    // Single connection, reused for follow mode
    let mut client = session::SessionClient::connect(&discovery, &token)
        .map_err(CliError::session)?;

    let session_id_str = discovery.session_id.to_string();
    let short_id = &session_id_str[..8.min(session_id_str.len())];
    let mut last_revision: Option<u64> = None;

    loop {
        // Fetch range data
        let result = client.inspect_range(sheet, start_row, start_col, end_row, end_col)
            .map_err(CliError::session)?;

        // Skip redraw if revision unchanged (in follow mode)
        if follow {
            if last_revision == Some(result.revision) {
                std::thread::sleep(Duration::from_millis(500));
                continue;
            }
            last_revision = Some(result.revision);

            // Clear screen for follow mode
            print!("\x1B[2J\x1B[H");
        }

        // Print header
        println!(
            "Session: {}  Sheet: {}  Range: {}  Revision: {}",
            short_id, sheet, range, result.revision
        );
        println!("{}", "─".repeat(60));

        // Print grid
        match result.result {
            InspectResult::Range { cells } => {
                // Cells are returned in row-major order
                print_grid_from_cells(&cells, start_row, start_col, end_row, end_col, col_width);
            }
            InspectResult::Cell(info) => {
                // Single cell - just print it
                println!("{}: {}", range.to_uppercase(), info.display);
            }
            InspectResult::Workbook(_) => {
                return Err(CliError::args("view requires a cell range, not 'workbook'".to_string()));
            }
        }

        if follow {
            println!();
            println!("(following - press Ctrl+C to stop)");
            std::thread::sleep(Duration::from_millis(500));
        } else {
            break;
        }
    }

    Ok(())
}

/// Print a grid of cells in table format.
/// Cells are assumed to be in row-major order.
fn print_grid_from_cells(
    cells: &[visigrid_protocol::CellInfo],
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    col_width: usize,
) {
    let num_cols = end_col - start_col + 1;

    // Build a map of (row, col) -> display value from flat array
    let mut grid: std::collections::HashMap<(usize, usize), &str> = std::collections::HashMap::new();
    for (i, cell) in cells.iter().enumerate() {
        let row = start_row + i / num_cols;
        let col = start_col + i % num_cols;
        grid.insert((row, col), &cell.display);
    }

    // Print column headers
    print!("{:>5} ", ""); // Row number column
    for col in start_col..=end_col {
        let col_name = col_to_letter(col);
        print!("{:^width$}", col_name, width = col_width);
    }
    println!();

    // Print separator
    print!("{:─>5}─", "");
    for _ in start_col..=end_col {
        print!("{:─>width$}", "", width = col_width);
    }
    println!();

    // Print rows
    for row in start_row..=end_row {
        print!("{:>5} ", row + 1); // 1-indexed row numbers
        for col in start_col..=end_col {
            let value = grid.get(&(row, col)).map(|s| *s).unwrap_or("");
            let display = truncate_display(value, col_width);
            print!("{:>width$}", display, width = col_width);
        }
        println!();
    }
}

/// Truncate a string to fit within width, adding ".." if truncated.
/// Handles UTF-8 safely by finding char boundaries.
fn truncate_display(s: &str, width: usize) -> String {
    if width < 3 {
        return s.chars().next().map(|c| c.to_string()).unwrap_or_default();
    }

    let char_count = s.chars().count();
    if char_count <= width {
        return s.to_string();
    }

    // Truncate to width - 2 chars, add ".."
    let truncated: String = s.chars().take(width - 2).collect();
    format!("{}..", truncated)
}

/// Convert column index to letter (0 -> A, 1 -> B, 26 -> AA, etc.)
fn col_to_letter(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;
    loop {
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    result
}

/// Resolve session by ID (prefix match), or auto-select if only one session.
fn resolve_session(session_id: Option<&str>) -> Result<session::DiscoveryFile, CliError> {
    let sessions = session::list_sessions()
        .map_err(|e| CliError::io(format!("failed to list sessions: {}", e)))?;

    if sessions.is_empty() {
        return Err(CliError::io("no running VisiGrid sessions found")
            .with_hint("start VisiGrid GUI and enable session server"));
    }

    match session_id {
        Some(id) => {
            session::find_session(id)
                .map_err(|e| CliError::args(e.to_string()))?
                .ok_or_else(|| CliError::args(format!("session '{}' not found", id))
                    .with_hint("use 'visigrid sessions' to list available sessions"))
        }
        None => {
            if sessions.len() == 1 {
                Ok(sessions.into_iter().next().unwrap())
            } else {
                Err(CliError::args(format!("{} sessions found; specify --session", sessions.len()))
                    .with_hint("use 'visigrid sessions' to list available sessions"))
            }
        }
    }
}

/// Get session token from environment variable.
fn get_session_token() -> Result<String, CliError> {
    std::env::var("VISIGRID_SESSION_TOKEN")
        .map_err(|_| CliError::args("VISIGRID_SESSION_TOKEN environment variable not set")
            .with_hint("copy the token from VisiGrid GUI session panel and set: export VISIGRID_SESSION_TOKEN=xxx"))
}

// =============================================================================
// Sheet commands (Phase 2A: Agent-ready headless workflows)
// =============================================================================

/// Build a .sheet file from a Lua script.
fn cmd_sheet_apply(
    output: PathBuf,
    lua_path: PathBuf,
    verify: Option<String>,
    dry_run: bool,
    json: bool,
) -> Result<(), CliError> {
    use visigrid_io::native::save_workbook;

    // Execute the build script
    let result = sheet_ops::execute_build_script(&lua_path, verify.as_deref())?;

    // Check verification if requested
    if let Some(verified) = result.verified {
        if !verified {
            let expected = verify.as_deref().unwrap_or("(unknown)");
            let computed = result.fingerprint.to_string();

            if json {
                let output_json = serde_json::json!({
                    "ok": false,
                    "error": "fingerprint_mismatch",
                    "expected": expected,
                    "computed": computed,
                    "semantic_ops": result.semantic_ops,
                    "style_ops": result.style_ops,
                    "cells_changed": result.cells_changed,
                });
                println!("{}", serde_json::to_string_pretty(&output_json).unwrap());
            } else {
                eprintln!("Fingerprint mismatch");
                eprintln!("  Expected: {}", expected);
                eprintln!("  Computed: {}", computed);
            }
            return Err(CliError { code: EXIT_ERROR, message: "fingerprint mismatch".to_string(), hint: None });
        }
    }

    // Write output (unless dry-run)
    if !dry_run {
        // Atomic write: write to temp file first, then rename
        let temp_path = output.with_extension("sheet.tmp");

        visigrid_io::native::save_workbook_with_metadata(&result.workbook, &result.metadata, &temp_path)
            .map_err(|e| CliError::io(format!("failed to write temp file: {}", e)))?;

        std::fs::rename(&temp_path, &output)
            .map_err(|e| CliError::io(format!("failed to rename to output: {}", e)))?;
    }

    // Output result
    if json {
        let output_json = serde_json::json!({
            "ok": true,
            "fingerprint": result.fingerprint.to_string(),
            "semantic_ops": result.semantic_ops,
            "style_ops": result.style_ops,
            "cells_changed": result.cells_changed,
            "dry_run": dry_run,
            "output": if dry_run { None } else { Some(output.display().to_string()) },
        });
        println!("{}", serde_json::to_string_pretty(&output_json).unwrap());
    } else {
        if dry_run {
            println!("(dry run - file not written)");
        } else {
            println!("Wrote {}", output.display());
        }
        println!("Fingerprint:  {}", result.fingerprint.to_string());
        println!("Semantic ops: {}", result.semantic_ops);
        println!("Style ops:    {}", result.style_ops);
        println!("Cells:        {}", result.cells_changed);
    }

    Ok(())
}

/// Inspect cells/ranges in a .sheet file.
fn cmd_sheet_inspect(
    file: PathBuf,
    target: Option<String>,
    workbook_mode: bool,
    include_style: bool,
    json: bool,
) -> Result<(), CliError> {
    use visigrid_io::native::load_workbook;

    let mut workbook = load_workbook(&file)
        .map_err(|e| CliError::io(format!("failed to load {}: {}", file.display(), e)))?;

    // Rebuild dependency graph and recalculate all formulas
    // This ensures formula cells have computed values when inspected
    workbook.rebuild_dep_graph();
    workbook.recompute_full_ordered();

    if workbook_mode || target.is_none() {
        // Workbook metadata
        let fingerprint = sheet_ops::compute_sheet_fingerprint(&workbook);
        let sheet = workbook.sheet(0);
        let cell_count = sheet.map(|s| {
            s.cells_iter().filter(|(_, c)| !c.value.raw_display().is_empty()).count()
        }).unwrap_or(0);

        let result = sheet_ops::WorkbookInspectResult {
            fingerprint: fingerprint.to_string(),
            sheet_count: workbook.sheet_count(),
            cell_count,
        };

        if json {
            println!("{}", serde_json::to_string_pretty(&result).unwrap());
        } else {
            println!("File:        {}", file.display());
            println!("Fingerprint: {}", result.fingerprint);
            println!("Sheets:      {}", result.sheet_count);
            println!("Cells:       {}", result.cell_count);
        }
    } else {
        let target_str = target.unwrap();
        let sheet = workbook.sheet(0)
            .ok_or_else(|| CliError::io("no sheets in workbook"))?;

        // Parse target
        let parsed = sheet_ops::parse_cell_ref(&target_str)
            .map(|(r, c)| (r, c, r, c))
            .or_else(|| {
                // Try range
                if let Some((start, end)) = target_str.split_once(':') {
                    let (sr, sc) = sheet_ops::parse_cell_ref(start)?;
                    let (er, ec) = sheet_ops::parse_cell_ref(end)?;
                    Some((sr, sc, er, ec))
                } else {
                    None
                }
            })
            .ok_or_else(|| CliError::args(format!("invalid target: {}", target_str)))?;

        let (start_row, start_col, end_row, end_col) = parsed;

        if start_row == end_row && start_col == end_col {
            // Single cell
            let raw = sheet.get_raw(start_row, start_col);
            let display = sheet.get_display(start_row, start_col);
            let format = sheet.get_format(start_row, start_col);

            let value_type = if raw.starts_with('=') {
                "formula"
            } else if display.parse::<f64>().is_ok() {
                "number"
            } else if display.is_empty() {
                "empty"
            } else {
                "text"
            };

            let formula = if raw.starts_with('=') { Some(raw.clone()) } else { None };

            let format_info = if include_style {
                let nf_str = match &format.number_format {
                    visigrid_engine::cell::NumberFormat::General => None,
                    nf => Some(format!("{:?}", nf)),
                };
                Some(sheet_ops::CellFormatInfo {
                    bold: format.bold,
                    italic: format.italic,
                    underline: format.underline,
                    number_format: nf_str,
                })
            } else {
                None
            };

            let result = sheet_ops::CellInspectResult {
                cell: target_str.to_uppercase(),
                value: display,
                formula,
                value_type: value_type.to_string(),
                format: format_info,
            };

            if json {
                println!("{}", serde_json::to_string_pretty(&result).unwrap());
            } else {
                println!("{} = {}  ({})", result.cell, result.value, result.value_type);
                if let Some(f) = &result.formula {
                    println!("Formula: {}", f);
                }
                if include_style {
                    if format.bold { println!("Style: bold"); }
                    if format.italic { println!("Style: italic"); }
                    if format.underline { println!("Style: underline"); }
                }
            }
        } else {
            // Range
            let mut cells = Vec::new();
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    let raw = sheet.get_raw(row, col);
                    let display = sheet.get_display(row, col);

                    let value_type = if raw.starts_with('=') {
                        "formula"
                    } else if display.parse::<f64>().is_ok() {
                        "number"
                    } else if display.is_empty() {
                        "empty"
                    } else {
                        "text"
                    };

                    let formula = if raw.starts_with('=') { Some(raw.clone()) } else { None };

                    cells.push(sheet_ops::CellInspectResult {
                        cell: sheet_ops::format_cell_ref(row, col),
                        value: display,
                        formula,
                        value_type: value_type.to_string(),
                        format: None, // Style not included for ranges by default
                    });
                }
            }

            let result = sheet_ops::RangeInspectResult {
                range: target_str.to_uppercase(),
                cells,
            };

            if json {
                println!("{}", serde_json::to_string_pretty(&result).unwrap());
            } else {
                println!("Range: {}", result.range);
                for cell in &result.cells {
                    let formula_marker = if cell.formula.is_some() { " [f]" } else { "" };
                    println!("  {} = {}{}", cell.cell, cell.value, formula_marker);
                }
            }
        }
    }

    Ok(())
}

/// Verify a .sheet file's fingerprint.
fn cmd_sheet_verify(file: PathBuf, expected: String) -> Result<(), CliError> {
    use visigrid_io::native::{load_workbook, load_cell_metadata};

    let workbook = load_workbook(&file)
        .map_err(|e| CliError::io(format!("failed to load {}: {}", file.display(), e)))?;

    let metadata = load_cell_metadata(&file)
        .map_err(|e| CliError::io(format!("failed to load metadata: {}", e)))?;

    let computed = sheet_ops::compute_sheet_fingerprint_with_meta(&workbook, &metadata);
    let expected_fp = replay::ReplayFingerprint::parse(&expected)
        .ok_or_else(|| CliError::args(format!("invalid fingerprint format: {}", expected)))?;

    if computed == expected_fp {
        println!("Verification: PASS");
        println!("Fingerprint:  {}", computed.to_string());
        Ok(())
    } else {
        eprintln!("Verification: FAIL");
        eprintln!("  Expected: {}", expected);
        eprintln!("  Computed: {}", computed.to_string());
        Err(CliError { code: EXIT_ERROR, message: "fingerprint mismatch".to_string(), hint: None })
    }
}

/// Compute and print a .sheet file's fingerprint.
fn cmd_sheet_fingerprint(file: PathBuf, json: bool) -> Result<(), CliError> {
    use visigrid_io::native::{load_workbook, load_cell_metadata};

    let workbook = load_workbook(&file)
        .map_err(|e| CliError::io(format!("failed to load {}: {}", file.display(), e)))?;

    let metadata = load_cell_metadata(&file)
        .map_err(|e| CliError::io(format!("failed to load metadata: {}", e)))?;

    let fingerprint = sheet_ops::compute_sheet_fingerprint_with_meta(&workbook, &metadata);

    if json {
        let output = serde_json::json!({
            "file": file.display().to_string(),
            "fingerprint": fingerprint.to_string(),
            "ops": fingerprint.len,
        });
        println!("{}", serde_json::to_string_pretty(&output).unwrap());
    } else {
        println!("{}", fingerprint.to_string());
    }

    Ok(())
}
