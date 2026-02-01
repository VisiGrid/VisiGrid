// VisiGrid CLI - headless spreadsheet operations
// See docs/cli-v1.md for specification

mod diff;
mod replay;

use std::collections::HashMap;
use std::io::{self, Read, Write};
use std::path::PathBuf;
use std::process::ExitCode;

use clap::{Parser, Subcommand, ValueEnum};

// Exit codes per spec
pub const EXIT_SUCCESS: u8 = 0;
pub const EXIT_EVAL_ERROR: u8 = 1;
pub const EXIT_ARGS_ERROR: u8 = 2;
pub const EXIT_IO_ERROR: u8 = 3;
pub const EXIT_PARSE_ERROR: u8 = 4;
pub const EXIT_FORMAT_ERROR: u8 = 5;

// AI-specific exit codes
pub const EXIT_AI_DISABLED: u8 = 10;      // AI disabled (provider=none) - not an error
pub const EXIT_AI_MISSING_KEY: u8 = 11;   // Provider configured but key missing
pub const EXIT_AI_KEYCHAIN_ERR: u8 = 12;  // Keychain error

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
  visigrid convert report.xlsx -t csv -o - | head -5")]
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

    /// Reconcile two datasets by key (exit 0 = match, exit 1 = diffs found)
    #[command(after_help = "\
Examples:
  visigrid diff old.csv new.csv --key id
  visigrid diff old.csv new.csv --key name --tolerance 0.01
  visigrid diff old.csv new.csv --key sku --out csv --output diffs.csv
  visigrid diff old.csv new.csv --key id --compare price,quantity
  visigrid diff old.csv new.csv --key name --match contains")]
    Diff {
        /// Left dataset file
        left: PathBuf,

        /// Right dataset file
        right: PathBuf,

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

        /// Quiet mode - suppress stderr summary and warnings
        #[arg(long, short = 'q')]
        quiet: bool,

        /// Export ambiguous matches to CSV file (written before exit, even on --on-ambiguous error)
        #[arg(long)]
        save_ambiguous: Option<PathBuf>,
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

fn long_version() -> &'static str {
    if cfg!(debug_assertions) {
        concat!(
            env!("CARGO_PKG_VERSION"),
            "\nengine: visigrid-engine ",
            env!("CARGO_PKG_VERSION"),
            "\nbuild:  debug",
        )
    } else {
        concat!(
            env!("CARGO_PKG_VERSION"),
            "\nengine: visigrid-engine ",
            env!("CARGO_PKG_VERSION"),
            "\nbuild:  release",
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
        }) => cmd_convert(input, from, to, output, sheet, delimiter, headers),
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
            quiet,
            save_ambiguous,
        }) => cmd_diff(
            left, right, key, r#match, key_transform, compare, tolerance,
            on_ambiguous, out, output, summary, no_headers, header_row, delimiter, quiet,
            save_ambiguous,
        ),
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
) -> Result<(), CliError> {

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

    // Write output
    let output_bytes = write_format(&sheet, to, delimiter, headers)?;

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
) -> Result<Vec<u8>, CliError> {
    match format {
        Format::Csv => write_csv(sheet, delimiter as u8),
        Format::Tsv => write_csv(sheet, b'\t'),
        Format::Json => write_json(sheet, headers),
        Format::Lines => write_lines(sheet),
        Format::Xlsx => Err(CliError::format("xlsx export not yet implemented")
            .with_hint("use -t csv or -t json instead")),
        Format::Sheet => Err(CliError::format("sheet format cannot be written to stdout")
            .with_hint("use -o output.sheet to write to a file")),
    }
}

fn write_csv(sheet: &visigrid_engine::sheet::Sheet, delimiter: u8) -> Result<Vec<u8>, CliError> {
    let mut writer = csv::WriterBuilder::new()
        .delimiter(delimiter)
        .from_writer(Vec::new());

    let (rows, cols) = get_data_bounds(sheet);

    for row in 0..rows {
        let mut record: Vec<String> = Vec::new();
        for col in 0..cols {
            record.push(sheet.get_display(row, col));
        }
        writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
    }

    writer.into_inner().map_err(|e| CliError::io(e.to_string()))
}

fn write_json(sheet: &visigrid_engine::sheet::Sheet, headers: bool) -> Result<Vec<u8>, CliError> {
    let (rows, cols) = get_data_bounds(sheet);

    if headers && rows > 0 {
        // Array of objects
        let mut header_names: Vec<String> = Vec::new();
        for col in 0..cols {
            let name = sheet.get_display(0, col);
            // Sanitize: lowercase, spaces to _, strip invalid chars
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
        for row in 1..rows {
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
    } else {
        // Array of arrays
        let mut rows_vec: Vec<Vec<serde_json::Value>> = Vec::new();
        for row in 0..rows {
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

fn write_lines(sheet: &visigrid_engine::sheet::Sheet) -> Result<Vec<u8>, CliError> {
    let mut output = Vec::new();
    let (rows, _) = get_data_bounds(sheet);

    for row in 0..rows {
        let value = sheet.get_display(row, 0);
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

// Diff-specific exit codes (per cli-diff.md spec)
const EXIT_DIFF_DUPLICATE: u8 = 3;
const EXIT_DIFF_AMBIGUOUS: u8 = 4;
const EXIT_DIFF_PARSE: u8 = 5;

#[allow(clippy::too_many_arguments)]
fn cmd_diff(
    left_path: PathBuf,
    right_path: PathBuf,
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
    _delimiter: char,
    quiet: bool,
    save_ambiguous: Option<PathBuf>,
) -> Result<(), CliError> {
    // Load both files
    let left_format = infer_format(&left_path)?;
    let right_format = infer_format(&right_path)?;
    let left_sheet = read_file(&left_path, left_format, ',')?;
    let right_sheet = read_file(&right_path, right_format, ',')?;

    let (left_bounds_rows, left_bounds_cols) = get_data_bounds(&left_sheet);
    let (right_bounds_rows, right_bounds_cols) = get_data_bounds(&right_sheet);

    if left_bounds_rows == 0 {
        return Err(CliError { code: EXIT_DIFF_PARSE, message: format!("{}: file is empty or has no data rows", left_path.display()), hint: None });
    }
    if right_bounds_rows == 0 {
        return Err(CliError { code: EXIT_DIFF_PARSE, message: format!("{}: file is empty or has no data rows", right_path.display()), hint: None });
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
        eprintln!("left:  {} rows ({})", s.left_rows, left_path.display());
        eprintln!("right: {} rows ({})", s.right_rows, right_path.display());
        eprintln!("matched: {}", s.matched);
        eprintln!("only_left: {}", s.only_left);
        eprintln!("only_right: {}", s.only_right);
        eprintln!("value_diff: {}", s.diff);
        if s.ambiguous > 0 {
            eprintln!("ambiguous: {}", s.ambiguous);
        }
    }

    // Exit 1 when differences are found (like standard diff)
    let s = &result.summary;
    if s.only_left > 0 || s.only_right > 0 || s.diff > 0 {
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

fn format_diff_json(
    result: &diff::DiffResult,
    options: &diff::DiffOptions,
    headers: &[String],
    summary_mode: &DiffSummaryMode,
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
        "ambiguous": result.summary.ambiguous,
        "tolerance": options.tolerance,
        "key": key_name,
        "match": match_str,
        "key_transform": kt_str,
    });

    let top = match summary_mode {
        DiffSummaryMode::Json => serde_json::json!({
            "summary": summary_json,
            "results": results_json,
        }),
        _ => serde_json::json!({
            "summary": summary_json,
            "results": results_json,
        }),
    };

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
