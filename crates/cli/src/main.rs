// VisiGrid CLI - headless spreadsheet operations
// See docs/cli-v1.md for specification

use std::io::{self, Read, Write};
use std::path::PathBuf;
use std::process::ExitCode;

use clap::{Parser, Subcommand, ValueEnum};

// Exit codes per spec
const EXIT_SUCCESS: u8 = 0;
const EXIT_EVAL_ERROR: u8 = 1;
const EXIT_ARGS_ERROR: u8 = 2;
const EXIT_IO_ERROR: u8 = 3;
const EXIT_PARSE_ERROR: u8 = 4;
const EXIT_FORMAT_ERROR: u8 = 5;

#[derive(Parser)]
#[command(name = "visigrid-cli")]
#[command(about = "Fast, native spreadsheet (CLI mode, headless)")]
#[command(version)]
#[command(subcommand_required = false)]
struct Cli {
    #[command(subcommand)]
    command: Option<Commands>,
}

#[derive(Subcommand)]
enum Commands {
    /// Evaluate a spreadsheet formula against data read from stdin
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
    };

    match result {
        Ok(()) => ExitCode::from(EXIT_SUCCESS),
        Err(CliError { code, message }) => {
            eprintln!("error: {}", message);
            ExitCode::from(code)
        }
    }
}

struct CliError {
    code: u8,
    message: String,
}

impl CliError {
    fn args(msg: impl Into<String>) -> Self {
        Self { code: EXIT_ARGS_ERROR, message: msg.into() }
    }

    fn io(msg: impl Into<String>) -> Self {
        Self { code: EXIT_IO_ERROR, message: msg.into() }
    }

    fn parse(msg: impl Into<String>) -> Self {
        Self { code: EXIT_PARSE_ERROR, message: msg.into() }
    }

    fn format(msg: impl Into<String>) -> Self {
        Self { code: EXIT_FORMAT_ERROR, message: msg.into() }
    }

    fn eval(msg: impl Into<String>) -> Self {
        Self { code: EXIT_EVAL_ERROR, message: msg.into() }
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
        (None, None) => return Err(CliError::args("stdin requires --from")),
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
        Some("xlsx") | Some("xls") => Ok(Format::Xlsx),
        Some("sheet") => Ok(Format::Sheet),
        _ => Err(CliError::args("cannot infer format, use --from")),
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
        return Err(CliError::parse("empty input"));
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
    use visigrid_engine::sheet::Sheet;

    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .from_reader(content.as_bytes());

    let mut sheet = Sheet::new(1000, 26);

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
    use visigrid_engine::sheet::Sheet;

    let value: serde_json::Value = serde_json::from_str(content)
        .map_err(|e| CliError::parse(format!("JSON parse error: {}", e)))?;

    let mut sheet = Sheet::new(1000, 26);

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
    use visigrid_engine::sheet::Sheet;

    let lines: Vec<&str> = content.lines().collect();
    if lines.is_empty() {
        return Err(CliError::parse("empty input"));
    }

    let mut sheet = Sheet::new(1000, 26);
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
        Format::Xlsx => Err(CliError::format("xlsx export not yet implemented")),
        Format::Sheet => Err(CliError::format("sheet export to stdout not supported, use --output")),
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

        serde_json::to_vec_pretty(&objects).map_err(|e| CliError::io(e.to_string()))
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

        serde_json::to_vec_pretty(&rows_vec).map_err(|e| CliError::io(e.to_string()))
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
        return Err(CliError::eval(format!("formula evaluation failed: {}", result)));
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
