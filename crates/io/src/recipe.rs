//! Import recipes: a source, an ordered list of steps, and a typed result.
//!
//! One executor serves every entry point — the desktop app, `vgrid run` and
//! MCP — so a recipe behaves the same wherever it runs. The rules it follows
//! come from the Data Import and Shaping Plan:
//!
//! - The source is read once into a [`Snapshot`]. Preview, validation and the
//!   published result all come from that snapshot, so a file changing between
//!   preview and publish cannot change what is published.
//! - Steps name only the columns the user acted on. Each one says what happens
//!   when a column it names is missing: fail (the default), skip the step, or
//!   treat the column as blank. Nothing guesses a substitute.
//! - Values are never lost silently. A value that does not fit its declared
//!   type is a counted, located error, and by default it fails the run.
//! - A failed run still returns its output for inspection, with `ok == false`.
//!   Callers must not publish it; the previous result stays in place.
//! - Fields starting with `=` stay text. A recipe never evaluates formulas.

use std::collections::{BTreeMap, HashSet};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Instant;

use serde::{Deserialize, Serialize};
use visigrid_engine::cell::{interchange_number, DateStyle, NumberFormat};
use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};

use crate::csv::sniff_delimiter;
use crate::csv_import::{decode, keep_as_text, parse_date, parse_decimal_comma, parse_number, ColumnRule, DateOrder, Encoding};

/// The recipe format version this build reads and writes.
pub const RECIPE_VERSION: u32 = 1;

/// Located errors kept in a report; the count is always exact.
const MAX_REPORTED_ERRORS: usize = 1000;

// ============================================================================
// Recipe file
// ============================================================================

/// A saved import: where the data comes from and what to do to it.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Recipe {
    pub version: u32,
    pub source: Source,
    #[serde(default, rename = "step", skip_serializing_if = "Vec::is_empty")]
    pub steps: Vec<Step>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "lowercase")]
pub enum Source {
    Csv(CsvSource),
}

/// A delimited text file, and how to read it.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct CsvSource {
    /// Relative paths resolve against the recipe file's folder.
    pub path: String,
    /// `","`, `";"`, `"tab"`, `"|"`; absent: detected from the file.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delimiter: Option<String>,
    /// `utf-8`, `windows-1252`, `utf-16`; absent: detected.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub encoding: Option<String>,
    /// The line holding the column names, counting from 1. Lines above it
    /// (titles, export dates) are skipped. 0: no header; columns are named by
    /// letter.
    #[serde(default = "one", skip_serializing_if = "is_one")]
    pub header_row: usize,
    /// Read 1.234,56 as 1234.56.
    #[serde(default, skip_serializing_if = "is_false")]
    pub decimal_comma: bool,
    /// The column names when the recipe was saved. A run reports columns that
    /// have since gone missing or appeared.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
}

fn one() -> usize {
    1
}
fn is_one(n: &usize) -> bool {
    *n == 1
}
fn is_false(b: &bool) -> bool {
    !*b
}

/// What a step does when a column it names is not there.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Missing {
    /// Fail the run (the default; what CI wants).
    #[default]
    Fail,
    /// Skip this step and say so.
    Skip,
    /// Treat the column as present and empty.
    Blank,
}

fn is_default_missing(m: &Missing) -> bool {
    *m == Missing::Fail
}

/// What happens to a value that does not fit its declared type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum OnError {
    /// Count and locate it, and fail the run (the default).
    #[default]
    Fail,
    /// Keep the original text in the cell, and report it.
    KeepText,
    /// Leave the cell empty, and report it.
    Blank,
}

fn is_default_on_error(e: &OnError) -> bool {
    *e == OnError::Fail
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum FilterOp {
    #[serde(rename = "=")]
    Eq,
    #[serde(rename = "!=")]
    Ne,
    #[serde(rename = "<")]
    Lt,
    #[serde(rename = "<=")]
    Le,
    #[serde(rename = ">")]
    Gt,
    #[serde(rename = ">=")]
    Ge,
    #[serde(rename = "contains")]
    Contains,
    #[serde(rename = "not_contains")]
    NotContains,
    #[serde(rename = "starts_with")]
    StartsWith,
    #[serde(rename = "empty")]
    Empty,
    #[serde(rename = "not_empty")]
    NotEmpty,
}

impl FilterOp {
    fn label(self) -> &'static str {
        match self {
            FilterOp::Eq => "is",
            FilterOp::Ne => "is not",
            FilterOp::Lt => "<",
            FilterOp::Le => "≤",
            FilterOp::Gt => ">",
            FilterOp::Ge => "≥",
            FilterOp::Contains => "contains",
            FilterOp::NotContains => "does not contain",
            FilterOp::StartsWith => "starts with",
            FilterOp::Empty => "is empty",
            FilterOp::NotEmpty => "is not empty",
        }
    }
}

/// One step. In the file each is a `[[step]]` table with an `op` key.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "op", rename_all = "snake_case", deny_unknown_fields)]
pub enum Step {
    /// Keep these columns, in this order; drop the rest.
    Select {
        columns: Vec<String>,
        #[serde(default, skip_serializing_if = "is_default_missing")]
        missing: Missing,
    },
    /// Drop these columns.
    Remove {
        columns: Vec<String>,
        #[serde(default, skip_serializing_if = "is_default_missing")]
        missing: Missing,
    },
    /// Old name → new name.
    Rename {
        columns: BTreeMap<String, String>,
        #[serde(default, skip_serializing_if = "is_default_missing")]
        missing: Missing,
    },
    /// Declare column types: `text`, `number`, `date:ymd|dmy|mdy`, `auto`.
    /// Every value is checked.
    Types {
        columns: BTreeMap<String, String>,
        #[serde(default, skip_serializing_if = "is_default_on_error")]
        on_error: OnError,
        #[serde(default, skip_serializing_if = "is_default_missing")]
        missing: Missing,
    },
    /// Remove spaces at both ends; no columns named: every column.
    Trim {
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        columns: Vec<String>,
        #[serde(default, skip_serializing_if = "is_default_missing")]
        missing: Missing,
    },
    /// Keep the rows where the condition holds.
    Filter {
        column: String,
        #[serde(rename = "is")]
        op: FilterOp,
        #[serde(default, skip_serializing_if = "String::is_empty")]
        value: String,
        #[serde(default, skip_serializing_if = "is_default_missing")]
        missing: Missing,
    },
    /// Keep the first of each set of duplicate rows, comparing these columns
    /// (none named: the whole row).
    Dedupe {
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        columns: Vec<String>,
        #[serde(default, skip_serializing_if = "is_default_missing")]
        missing: Missing,
    },
}

impl Step {
    /// The step in plain words, for the step list and the CLI.
    pub fn describe(&self) -> String {
        let list = |c: &[String]| c.join(", ");
        match self {
            Step::Select { columns, .. } => format!("Keep columns {}", list(columns)),
            Step::Remove { columns, .. } => format!("Remove columns {}", list(columns)),
            Step::Rename { columns, .. } => {
                let pairs: Vec<String> = columns.iter().map(|(a, b)| format!("{a} → {b}")).collect();
                format!("Rename {}", pairs.join(", "))
            }
            Step::Types { columns, .. } => {
                let pairs: Vec<String> = columns.iter().map(|(c, t)| format!("{c}: {t}")).collect();
                format!("Set types {}", pairs.join(", "))
            }
            Step::Trim { columns, .. } if columns.is_empty() => "Trim spaces in every column".into(),
            Step::Trim { columns, .. } => format!("Trim spaces in {}", list(columns)),
            Step::Filter { column, op, value, .. } => match op {
                FilterOp::Empty | FilterOp::NotEmpty => format!("Keep rows where {column} {}", op.label()),
                _ => format!("Keep rows where {column} {} {value}", op.label()),
            },
            Step::Dedupe { columns, .. } if columns.is_empty() => "Remove duplicate rows".into(),
            Step::Dedupe { columns, .. } => format!("Remove duplicates by {}", list(columns)),
        }
    }

    fn missing(&self) -> Missing {
        match self {
            Step::Select { missing, .. }
            | Step::Remove { missing, .. }
            | Step::Rename { missing, .. }
            | Step::Types { missing, .. }
            | Step::Trim { missing, .. }
            | Step::Filter { missing, .. }
            | Step::Dedupe { missing, .. } => *missing,
        }
    }

    /// Every column the step names.
    fn named_columns(&self) -> Vec<&str> {
        match self {
            Step::Select { columns, .. } | Step::Remove { columns, .. } | Step::Trim { columns, .. } | Step::Dedupe { columns, .. } => {
                columns.iter().map(String::as_str).collect()
            }
            Step::Rename { columns, .. } | Step::Types { columns, .. } => columns.keys().map(String::as_str).collect(),
            Step::Filter { column, .. } => vec![column.as_str()],
        }
    }
}

impl Recipe {
    pub fn from_toml(text: &str) -> Result<Recipe, String> {
        let recipe: Recipe = toml::from_str(text).map_err(|e| e.to_string())?;
        if recipe.version != RECIPE_VERSION {
            return Err(format!(
                "recipe version {} is not supported (this VisiGrid reads version {RECIPE_VERSION})",
                recipe.version
            ));
        }
        for step in &recipe.steps {
            if let Step::Types { columns, .. } = step {
                for (column, ty) in columns {
                    parse_type(ty).map_err(|e| format!("column {column}: {e}"))?;
                }
            }
        }
        Ok(recipe)
    }

    pub fn to_toml(&self) -> String {
        toml::to_string(self).expect("a recipe always serializes")
    }

    pub fn load(path: &Path) -> Result<Recipe, String> {
        let text = std::fs::read_to_string(path).map_err(|e| format!("{}: {e}", path.display()))?;
        Recipe::from_toml(&text).map_err(|e| format!("{}: {e}", path.display()))
    }

    /// The source file: `over` if given, else the recipe's path, relative
    /// paths resolved against `recipe_dir`.
    pub fn source_path(&self, recipe_dir: &Path, over: Option<&Path>) -> PathBuf {
        if let Some(p) = over {
            return p.to_path_buf();
        }
        let Source::Csv(src) = &self.source;
        let p = Path::new(&src.path);
        if p.is_absolute() {
            p.to_path_buf()
        } else {
            recipe_dir.join(p)
        }
    }
}

/// `text`, `number`, `auto`, `date:ymd|dmy|mdy` (or `date`, meaning YMD).
pub fn parse_type(s: &str) -> Result<ColumnRule, String> {
    Ok(match s.trim().to_ascii_lowercase().as_str() {
        "auto" => ColumnRule::Auto,
        "text" => ColumnRule::Text,
        "number" => ColumnRule::Number,
        "date" | "date:ymd" => ColumnRule::Date(DateOrder::Ymd),
        "date:dmy" => ColumnRule::Date(DateOrder::Dmy),
        "date:mdy" => ColumnRule::Date(DateOrder::Mdy),
        other => return Err(format!("unknown type {other:?} (use text, number, auto, date:ymd, date:dmy or date:mdy)")),
    })
}

// ============================================================================
// Snapshot: the source, read once
// ============================================================================

/// The source file's bytes, read once. Preview, validation and publishing all
/// use the same snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    pub path: PathBuf,
    pub bytes: Arc<Vec<u8>>,
    /// blake3 of the bytes (hex, 16 characters).
    pub hash: String,
}

impl Snapshot {
    pub fn read(path: &Path) -> Result<Snapshot, String> {
        let bytes = std::fs::read(path).map_err(|e| format!("cannot read source {}: {e}", path.display()))?;
        Ok(Snapshot::from_bytes(path, bytes))
    }

    pub fn from_bytes(path: &Path, bytes: Vec<u8>) -> Snapshot {
        let hash = blake3::hash(&bytes).to_hex()[..16].to_string();
        Snapshot { path: path.to_path_buf(), bytes: Arc::new(bytes), hash }
    }
}

// ============================================================================
// Running
// ============================================================================

/// A value that did not fit, and where it came from.
#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct CellError {
    /// Line in the source file, counting from 1.
    pub line: usize,
    pub column: String,
    pub value: String,
    pub reason: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct StepReport {
    pub index: usize,
    pub description: String,
    pub rows_in: usize,
    pub rows_out: usize,
    pub millis: u128,
    /// Skipped because a column was missing and the step said skip.
    pub skipped: bool,
    /// What the step did that is worth saying: "removed 3 duplicate rows".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note: Option<String>,
}

/// Columns that changed since the recipe was saved.
#[derive(Debug, Clone, Default, Serialize)]
pub struct Drift {
    pub missing: Vec<String>,
    pub new: Vec<String>,
    /// (old, new) pairs at the same position, offered for the user to
    /// confirm. Never applied automatically.
    pub possibly_renamed: Vec<(String, String)>,
}

impl Drift {
    pub fn is_empty(&self) -> bool {
        self.missing.is_empty() && self.new.is_empty()
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct RunReport {
    /// False when the result must not be published.
    pub ok: bool,
    pub source: String,
    pub snapshot: String,
    pub source_rows: usize,
    pub rows: usize,
    pub columns: usize,
    pub drift: Drift,
    pub steps: Vec<StepReport>,
    /// Why the run failed, in plain words; empty when ok.
    pub failures: Vec<String>,
    /// Values that did not fit their type; at most MAX_REPORTED_ERRORS kept.
    pub errors: Vec<CellError>,
    pub error_count: usize,
}

/// One output column: its name and the type its values were checked against.
#[derive(Debug, Clone, PartialEq)]
pub struct OutColumn {
    pub name: String,
    pub rule: ColumnRule,
}

/// The shaped table, as text plus a type per column.
#[derive(Debug, Clone)]
pub struct RecipeOutput {
    pub columns: Vec<OutColumn>,
    pub rows: Vec<Vec<String>>,
    pub decimal_comma: bool,
}

#[derive(Debug)]
pub struct RunResult {
    pub output: RecipeOutput,
    pub report: RunReport,
}

/// The working table while steps run.
struct Frame {
    columns: Vec<OutColumn>,
    rows: Vec<Vec<String>>,
    /// Source line of each row (for error reports).
    lines: Vec<usize>,
    decimal_comma: bool,
}

impl Frame {
    fn find(&self, name: &str) -> Option<usize> {
        // Exact first, then case-insensitive (headers vary in case between exports)
        self.columns
            .iter()
            .position(|c| c.name == name)
            .or_else(|| self.columns.iter().position(|c| c.name.eq_ignore_ascii_case(name)))
    }

    fn add_blank(&mut self, name: &str) -> usize {
        self.columns.push(OutColumn { name: name.to_string(), rule: ColumnRule::Auto });
        for row in &mut self.rows {
            row.push(String::new());
        }
        self.columns.len() - 1
    }

    fn keep_rows(&mut self, keep: &[bool]) {
        let mut i = 0;
        self.rows.retain(|_| {
            i += 1;
            keep[i - 1]
        });
        let mut j = 0;
        self.lines.retain(|_| {
            j += 1;
            keep[j - 1]
        });
    }
}

/// Run a recipe against a snapshot of its source.
pub fn run(recipe: &Recipe, snapshot: &Snapshot) -> RunResult {
    let Source::Csv(src) = &recipe.source;
    let mut report = RunReport {
        ok: true,
        source: snapshot.path.display().to_string(),
        snapshot: snapshot.hash.clone(),
        source_rows: 0,
        rows: 0,
        columns: 0,
        drift: Drift::default(),
        steps: Vec::new(),
        failures: Vec::new(),
        errors: Vec::new(),
        error_count: 0,
    };

    let mut frame = match read_csv(src, snapshot) {
        Ok(frame) => frame,
        Err(e) => {
            report.ok = false;
            report.failures.push(e);
            return RunResult {
                output: RecipeOutput { columns: Vec::new(), rows: Vec::new(), decimal_comma: src.decimal_comma },
                report,
            };
        }
    };
    report.source_rows = frame.rows.len();
    report.drift = drift(&src.columns, &frame.columns);

    for (index, step) in recipe.steps.iter().enumerate() {
        let started = Instant::now();
        let rows_in = frame.rows.len();
        let mut skipped = false;
        let mut note = None;

        // Columns this step names that are not there
        let absent: Vec<String> = step
            .named_columns()
            .into_iter()
            .filter(|c| frame.find(c).is_none())
            .map(str::to_string)
            .collect();
        if !absent.is_empty() {
            match step.missing() {
                Missing::Fail => {
                    report.ok = false;
                    report.failures.push(format!(
                        "step {} ({}): column{} {} not found in the source",
                        index + 1,
                        step.describe(),
                        if absent.len() == 1 { "" } else { "s" },
                        absent.join(", ")
                    ));
                    skipped = true;
                }
                Missing::Skip => {
                    skipped = true;
                    note = Some(format!("skipped: {} not found", absent.join(", ")));
                }
                Missing::Blank => {
                    for name in &absent {
                        frame.add_blank(name);
                    }
                    note = Some(format!("{} not found; treated as empty", absent.join(", ")));
                }
            }
        }

        if !skipped {
            let step_note = apply(step, &mut frame, &mut report);
            if step_note.is_some() {
                note = step_note;
            }
        }

        report.steps.push(StepReport {
            index: index + 1,
            description: step.describe(),
            rows_in,
            rows_out: frame.rows.len(),
            millis: started.elapsed().as_millis(),
            skipped,
            note,
        });
    }

    if frame.rows.len() + 1 > NUM_ROWS {
        report.ok = false;
        report.failures.push(format!(
            "the result has {} rows; a sheet holds {} including the header row",
            frame.rows.len(),
            NUM_ROWS
        ));
    }
    if frame.columns.len() > NUM_COLS {
        report.ok = false;
        report.failures.push(format!("the result has {} columns; a sheet holds {}", frame.columns.len(), NUM_COLS));
    }

    report.rows = frame.rows.len();
    report.columns = frame.columns.len();
    RunResult {
        output: RecipeOutput { columns: frame.columns, rows: frame.rows, decimal_comma: src.decimal_comma },
        report,
    }
}

fn read_csv(src: &CsvSource, snapshot: &Snapshot) -> Result<Frame, String> {
    let encoding = match &src.encoding {
        Some(e) => Some(Encoding::parse(e).ok_or_else(|| format!("unknown encoding {e:?}"))?),
        None => None,
    };
    let (content, _) = decode(&snapshot.bytes, encoding);
    let delimiter = match src.delimiter.as_deref() {
        None => sniff_delimiter(&content),
        Some("tab") | Some("\t") => b'\t',
        Some(d) if d.len() == 1 => d.as_bytes()[0],
        Some(d) => return Err(format!("delimiter {d:?} must be one character or \"tab\"")),
    };
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(content.as_bytes());

    let mut names: Vec<String> = Vec::new();
    let mut rows = Vec::new();
    let mut lines = Vec::new();
    let mut record = csv::StringRecord::new();
    let mut n = 0usize;
    while reader.read_record(&mut record).map_err(|e| format!("cannot read the source: {e}"))? {
        n += 1;
        let line = record.position().map_or(n, |p| p.line() as usize);
        if src.header_row > 0 && n < src.header_row {
            continue; // title lines above the header
        }
        if src.header_row > 0 && n == src.header_row {
            names = record.iter().map(|f| f.trim().to_string()).collect();
            continue;
        }
        rows.push(record.iter().map(str::to_string).collect::<Vec<_>>());
        lines.push(line);
    }
    if src.header_row > 0 && n < src.header_row {
        return Err(format!("the source has {n} lines; the header is set to line {}", src.header_row));
    }

    // Name every column, including any wider than the header row
    let width = rows.iter().map(Vec::len).max().unwrap_or(0).max(names.len());
    for i in 0..width {
        match names.get_mut(i) {
            Some(name) if !name.is_empty() => {}
            Some(name) => *name = crate::csv_import::col_label(i),
            None => names.push(crate::csv_import::col_label(i)),
        }
    }
    for row in &mut rows {
        row.resize(width, String::new());
    }
    Ok(Frame {
        columns: names.into_iter().map(|name| OutColumn { name, rule: ColumnRule::Auto }).collect(),
        rows,
        lines,
        decimal_comma: src.decimal_comma,
    })
}

fn drift(saved: &[String], now: &[OutColumn]) -> Drift {
    if saved.is_empty() {
        return Drift::default();
    }
    let has = |list: &[&str], name: &str| list.iter().any(|n| n.eq_ignore_ascii_case(name));
    let now_names: Vec<&str> = now.iter().map(|c| c.name.as_str()).collect();
    let saved_names: Vec<&str> = saved.iter().map(String::as_str).collect();
    let missing: Vec<String> = saved.iter().filter(|n| !has(&now_names, n)).cloned().collect();
    let new: Vec<String> = now_names.iter().filter(|n| !has(&saved_names, n)).map(|n| n.to_string()).collect();
    // A missing and a new column at the same position may be a rename: offer
    // it, never apply it
    let possibly_renamed = missing
        .iter()
        .filter_map(|m| {
            let pos = saved.iter().position(|s| s == m)?;
            let candidate = now_names.get(pos)?;
            new.iter().any(|n| n == candidate).then(|| (m.clone(), candidate.to_string()))
        })
        .collect();
    Drift { missing, new, possibly_renamed }
}

/// Apply one step whose columns are all present. Returns a note worth showing.
fn apply(step: &Step, frame: &mut Frame, report: &mut RunReport) -> Option<String> {
    match step {
        Step::Select { columns, .. } => {
            let idx: Vec<usize> = columns.iter().filter_map(|c| frame.find(c)).collect();
            let dropped = frame.columns.len() - idx.len();
            frame.columns = idx.iter().map(|&i| frame.columns[i].clone()).collect();
            for row in &mut frame.rows {
                *row = idx.iter().map(|&i| std::mem::take(&mut row[i])).collect();
            }
            (dropped > 0).then(|| format!("dropped {dropped} other column{}", plural(dropped)))
        }
        Step::Remove { columns, .. } => {
            let drop: HashSet<usize> = columns.iter().filter_map(|c| frame.find(c)).collect();
            let keep: Vec<usize> = (0..frame.columns.len()).filter(|i| !drop.contains(i)).collect();
            frame.columns = keep.iter().map(|&i| frame.columns[i].clone()).collect();
            for row in &mut frame.rows {
                *row = keep.iter().map(|&i| std::mem::take(&mut row[i])).collect();
            }
            None
        }
        Step::Rename { columns, .. } => {
            for (from, to) in columns {
                if let Some(i) = frame.find(from) {
                    frame.columns[i].name = to.clone();
                }
            }
            let names: Vec<&str> = frame.columns.iter().map(|c| c.name.as_str()).collect();
            let unique: HashSet<String> = names.iter().map(|n| n.to_lowercase()).collect();
            if unique.len() != names.len() {
                report.ok = false;
                report.failures.push(format!("{}: two columns now have the same name", step.describe()));
            }
            None
        }
        Step::Types { columns, on_error, .. } => {
            let mut bad = 0usize;
            for (name, ty) in columns {
                let Some(i) = frame.find(name) else { continue };
                let rule = parse_type(ty).unwrap_or(ColumnRule::Auto);
                frame.columns[i].rule = rule;
                for (r, row) in frame.rows.iter_mut().enumerate() {
                    let value = &row[i];
                    if value.trim().is_empty() {
                        continue;
                    }
                    let fits = match rule {
                        ColumnRule::Number => parse_number(value, frame.decimal_comma).is_some(),
                        ColumnRule::Date(order) => parse_date(value, order).is_some(),
                        _ => true,
                    };
                    if fits {
                        continue;
                    }
                    bad += 1;
                    report.error_count += 1;
                    if report.errors.len() < MAX_REPORTED_ERRORS {
                        report.errors.push(CellError {
                            line: frame.lines[r],
                            column: frame.columns[i].name.clone(),
                            value: value.clone(),
                            reason: match rule {
                                ColumnRule::Number => "not a number".into(),
                                ColumnRule::Date(o) => format!("not a {} date", o.label()),
                                _ => String::new(),
                            },
                        });
                    }
                    if *on_error == OnError::Blank {
                        row[i].clear();
                    }
                }
            }
            if bad > 0 && *on_error == OnError::Fail {
                report.ok = false;
                report.failures.push(format!(
                    "{bad} value{} did not fit {}",
                    plural(bad),
                    if columns.len() == 1 { "its column's type" } else { "their columns' types" }
                ));
            }
            (bad > 0).then(|| format!("{bad} value{} did not fit", plural(bad)))
        }
        Step::Trim { columns, .. } => {
            let idx: Vec<usize> = if columns.is_empty() {
                (0..frame.columns.len()).collect()
            } else {
                columns.iter().filter_map(|c| frame.find(c)).collect()
            };
            let mut changed = 0usize;
            for row in &mut frame.rows {
                for &i in &idx {
                    let t = row[i].trim();
                    if t.len() != row[i].len() {
                        row[i] = t.to_string();
                        changed += 1;
                    }
                }
            }
            (changed > 0).then(|| format!("trimmed {changed} value{}", plural(changed)))
        }
        Step::Filter { column, op, value, .. } => {
            let i = frame.find(column)?;
            let keep: Vec<bool> = frame.rows.iter().map(|row| matches(&row[i], *op, value)).collect();
            let removed = keep.iter().filter(|k| !**k).count();
            frame.keep_rows(&keep);
            Some(format!("removed {removed} row{}", plural(removed)))
        }
        Step::Dedupe { columns, .. } => {
            let idx: Vec<usize> = if columns.is_empty() {
                (0..frame.columns.len()).collect()
            } else {
                columns.iter().filter_map(|c| frame.find(c)).collect()
            };
            let mut seen = HashSet::new();
            let keep: Vec<bool> = frame
                .rows
                .iter()
                .map(|row| seen.insert(idx.iter().map(|&i| row[i].as_str()).collect::<Vec<_>>().join("\u{1f}")))
                .collect();
            let removed = keep.iter().filter(|k| !**k).count();
            frame.keep_rows(&keep);
            Some(format!("removed {removed} duplicate row{}", plural(removed)))
        }
    }
}

/// Numbers compare as numbers when both sides are numbers; otherwise text,
/// ignoring case.
fn matches(cell: &str, op: FilterOp, value: &str) -> bool {
    let c = cell.trim();
    let v = value.trim();
    let num = |s: &str| parse_number(s, false);
    let ord = match (num(c), num(v)) {
        (Some(a), Some(b)) => a.partial_cmp(&b),
        _ => Some(c.to_lowercase().cmp(&v.to_lowercase())),
    };
    use std::cmp::Ordering::*;
    match op {
        FilterOp::Eq => ord == Some(Equal),
        FilterOp::Ne => ord != Some(Equal),
        FilterOp::Lt => ord == Some(Less),
        FilterOp::Le => matches!(ord, Some(Less | Equal)),
        FilterOp::Gt => ord == Some(Greater),
        FilterOp::Ge => matches!(ord, Some(Greater | Equal)),
        FilterOp::Contains => c.to_lowercase().contains(&v.to_lowercase()),
        FilterOp::NotContains => !c.to_lowercase().contains(&v.to_lowercase()),
        FilterOp::StartsWith => c.to_lowercase().starts_with(&v.to_lowercase()),
        FilterOp::Empty => c.is_empty(),
        FilterOp::NotEmpty => !c.is_empty(),
    }
}

fn plural(n: usize) -> &'static str {
    if n == 1 {
        ""
    } else {
        "s"
    }
}

impl RecipeOutput {
    /// The result as a sheet: names in row 1, typed values below. Auto
    /// columns follow the CSV importer's safe defaults (IDs stay text,
    /// nothing becomes a date, `=` stays text); declared columns use their
    /// type.
    pub fn to_sheet(&self) -> Sheet {
        let mut sheet = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
        for (c, col) in self.columns.iter().enumerate() {
            sheet.set_text(0, c, &col.name);
        }
        for (r, row) in self.rows.iter().enumerate() {
            for (c, value) in row.iter().enumerate() {
                if value.is_empty() {
                    continue;
                }
                let (dr, rule) = (r + 1, self.columns[c].rule);
                match rule {
                    ColumnRule::Text | ColumnRule::Skip => sheet.set_text(dr, c, value),
                    ColumnRule::Number => match parse_number(value, self.decimal_comma) {
                        Some(n) => sheet.set_value_deferred(dr, c, &interchange_number(n)),
                        None => sheet.set_text(dr, c, value), // reported; kept as text
                    },
                    ColumnRule::Date(order) => match parse_date(value, order) {
                        Some(serial) => {
                            sheet.set_value_deferred(dr, c, &interchange_number(serial));
                            sheet.set_number_format(dr, c, NumberFormat::Date { style: DateStyle::Iso });
                        }
                        None => sheet.set_text(dr, c, value),
                    },
                    ColumnRule::Auto => {
                        if keep_as_text(value, false).is_some() {
                            sheet.set_text(dr, c, value);
                        } else if self.decimal_comma {
                            match parse_decimal_comma(value) {
                                Some(n) => sheet.set_value_deferred(dr, c, &interchange_number(n)),
                                None => sheet.set_value_deferred(dr, c, value),
                            }
                        } else {
                            sheet.set_value_deferred(dr, c, value);
                        }
                    }
                }
            }
        }
        sheet.rows = (self.rows.len() + 1).max(1000);
        sheet.cols = self.columns.len().max(26);
        sheet
    }
}

impl RunReport {
    /// A plain-text summary for the CLI and logs.
    pub fn summary(&self) -> String {
        let mut out = String::new();
        let status = if self.ok { "ok" } else { "FAILED, nothing published" };
        out.push_str(&format!(
            "{} ({}): {} source rows -> {} rows x {} columns, {status}\n",
            self.source, self.snapshot, self.source_rows, self.rows, self.columns
        ));
        if !self.drift.is_empty() {
            if !self.drift.missing.is_empty() {
                out.push_str(&format!("  columns missing since the recipe was saved: {}\n", self.drift.missing.join(", ")));
            }
            if !self.drift.new.is_empty() {
                out.push_str(&format!("  new columns: {}\n", self.drift.new.join(", ")));
            }
            for (old, new) in &self.drift.possibly_renamed {
                out.push_str(&format!("  possibly renamed: {old} -> {new} (confirm in the recipe)\n"));
            }
        }
        for s in &self.steps {
            out.push_str(&format!(
                "  {:>2}. {} [{} -> {} rows, {} ms]{}{}\n",
                s.index,
                s.description,
                s.rows_in,
                s.rows_out,
                s.millis,
                if s.skipped { " SKIPPED" } else { "" },
                s.note.as_ref().map(|n| format!(": {n}")).unwrap_or_default()
            ));
        }
        for f in &self.failures {
            out.push_str(&format!("  error: {f}\n"));
        }
        for e in self.errors.iter().take(20) {
            out.push_str(&format!("    line {}, {}: {:?} {}\n", e.line, e.column, e.value, e.reason));
        }
        if self.error_count > 20 {
            out.push_str(&format!("    … and {} more\n", self.error_count - 20));
        }
        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use visigrid_engine::formula::eval::Value;

    /// A vendor export of the kind the prototype targets: two title lines
    /// above the header, padded values, a duplicate row, an ID with leading
    /// zeros.
    const EXPORT: &str = "Acme Supply — order export\nGenerated 2026-09-30\nOrder ID,Customer,Amount,Order Date,Notes\n00042, Alpha Co ,120.50,2026-09-01,rush\n00043,Beta LLC,0,2026-09-02,\n00042, Alpha Co ,120.50,2026-09-01,rush\n00044,Gamma Inc,75,2026-09-03,=HYPERLINK(\"http://x\")\n";

    const RECIPE: &str = r#"
version = 1

[source]
kind = "csv"
path = "export.csv"
header_row = 3
columns = ["Order ID", "Customer", "Amount", "Order Date", "Notes"]

# Only the columns the report uses
[[step]]
op = "select"
columns = ["Order ID", "Customer", "Amount", "Order Date"]

[[step]]
op = "rename"
columns = { "Order ID" = "order_id" }

[[step]]
op = "trim"

[[step]]
op = "types"
columns = { order_id = "text", Amount = "number", "Order Date" = "date:ymd" }

[[step]]
op = "filter"
column = "Amount"
is = ">"
value = "0"

[[step]]
op = "dedupe"
"#;

    fn snap(text: &str) -> Snapshot {
        Snapshot::from_bytes(Path::new("export.csv"), text.as_bytes().to_vec())
    }

    fn recipe() -> Recipe {
        Recipe::from_toml(RECIPE).unwrap()
    }

    #[test]
    fn the_recipe_round_trips_and_describes_itself() {
        let r = recipe();
        assert_eq!(Recipe::from_toml(&r.to_toml()).unwrap(), r);
        let steps: Vec<String> = r.steps.iter().map(Step::describe).collect();
        assert_eq!(steps[0], "Keep columns Order ID, Customer, Amount, Order Date");
        assert_eq!(steps[4], "Keep rows where Amount > 0");
        assert_eq!(steps[5], "Remove duplicate rows");
    }

    #[test]
    fn a_clean_export_replays() {
        let res = run(&recipe(), &snap(EXPORT));
        assert!(res.report.ok, "{}", res.report.summary());
        assert!(res.report.drift.is_empty());
        let names: Vec<&str> = res.output.columns.iter().map(|c| c.name.as_str()).collect();
        assert_eq!(names, ["order_id", "Customer", "Amount", "Order Date"]);
        // Title lines skipped, zero amount filtered, duplicate removed, spaces trimmed
        assert_eq!(res.output.rows, vec![
            vec!["00042", "Alpha Co", "120.50", "2026-09-01"],
            vec!["00044", "Gamma Inc", "75", "2026-09-03"],
        ]);
        let notes: Vec<Option<String>> = res.report.steps.iter().map(|s| s.note.clone()).collect();
        assert_eq!(notes[4].as_deref(), Some("removed 1 row"));
        assert_eq!(notes[5].as_deref(), Some("removed 1 duplicate row"));
        assert_eq!(res.report.steps[0].rows_in, 4);
    }

    #[test]
    fn output_is_typed_as_declared() {
        let res = run(&recipe(), &snap(EXPORT));
        let sheet = res.output.to_sheet();
        assert_eq!(sheet.get_display(0, 0), "order_id");
        assert!(matches!(sheet.get_computed_value(1, 0), Value::Text(_)), "IDs stay text");
        assert_eq!(sheet.get_display(1, 0), "00042");
        assert_eq!(sheet.get_computed_value(1, 2), Value::Number(120.5));
        assert!(matches!(sheet.get_computed_value(1, 3), Value::Number(_)), "a declared date is a date serial");
        assert_eq!(sheet.get_formatted_display(1, 3), "2026-09-01");
    }

    #[test]
    fn a_renamed_column_fails_with_a_plain_reason_and_a_suggestion() {
        let renamed = EXPORT.replace("Order ID,Customer,Amount", "Order Number,Customer,Amount");
        let res = run(&recipe(), &snap(&renamed));
        assert!(!res.report.ok);
        assert_eq!(res.report.drift.missing, ["Order ID"]);
        assert_eq!(res.report.drift.new, ["Order Number"]);
        assert_eq!(res.report.drift.possibly_renamed, [("Order ID".to_string(), "Order Number".to_string())]);
        assert!(res.report.failures[0].contains("Order ID not found"), "{:?}", res.report.failures);
        // The suggestion is never applied
        assert!(res.output.columns.iter().all(|c| c.name != "order_id"));
    }

    #[test]
    fn an_invalid_date_is_located_and_fails_the_run() {
        let bad = EXPORT.replace("2026-09-03", "2026-09-31");
        let res = run(&recipe(), &snap(&bad));
        assert!(!res.report.ok);
        assert_eq!(res.report.error_count, 1);
        let e = &res.report.errors[0];
        assert_eq!((e.line, e.column.as_str(), e.value.as_str()), (7, "Order Date", "2026-09-31"));
        assert_eq!(e.reason, "not a YYYY-MM-DD date");
    }

    #[test]
    fn on_error_keep_text_publishes_and_still_reports() {
        let bad = EXPORT.replace("2026-09-03", "soon");
        let mut r = recipe();
        if let Step::Types { on_error, .. } = &mut r.steps[3] {
            *on_error = OnError::KeepText;
        }
        let res = run(&r, &snap(&bad));
        assert!(res.report.ok);
        assert_eq!(res.report.error_count, 1);
        assert_eq!(res.output.rows[1][3], "soon");
    }

    #[test]
    fn reordered_columns_and_rows_still_replay_by_name() {
        let reordered = "Acme Supply — order export\nGenerated 2026-10-31\nNotes,Amount,Order Date,Customer,Order ID\n,75,2026-10-03,Gamma Inc,00044\nrush,120.50,2026-10-01, Alpha Co ,00042\n";
        let res = run(&recipe(), &snap(reordered));
        assert!(res.report.ok, "{}", res.report.summary());
        assert_eq!(res.output.rows[0], ["00044", "Gamma Inc", "75", "2026-10-03"]);
        assert_eq!(res.output.rows[1], ["00042", "Alpha Co", "120.50", "2026-10-01"]);
    }

    #[test]
    fn missing_policy_skip_and_blank() {
        let no_notes = r#"
version = 1
[source]
kind = "csv"
path = "x.csv"
[[step]]
op = "trim"
columns = ["Notes"]
missing = "skip"
[[step]]
op = "select"
columns = ["a", "Region"]
missing = "blank"
"#;
        let r = Recipe::from_toml(no_notes).unwrap();
        let res = run(&r, &snap("a,b\n1,2\n"));
        assert!(res.report.ok, "{}", res.report.summary());
        assert!(res.report.steps[0].skipped);
        assert_eq!(res.output.rows, vec![vec!["1", ""]]);
        assert_eq!(res.output.columns[1].name, "Region");
    }

    #[test]
    fn formulas_and_ids_stay_text_in_auto_columns() {
        let r = Recipe::from_toml("version = 1\n[source]\nkind = \"csv\"\npath = \"x.csv\"\n").unwrap();
        let res = run(&r, &snap("id,note\n007,=1+1\n"));
        let sheet = res.output.to_sheet();
        assert_eq!(sheet.get_display(1, 0), "007");
        assert_eq!(sheet.get_display(1, 1), "=1+1");
        assert!(matches!(sheet.get_computed_value(1, 1), Value::Text(_)));
    }

    #[test]
    fn bad_recipes_are_refused_with_reasons() {
        let unknown_op = "version = 1\n[source]\nkind = \"csv\"\npath = \"x\"\n[[step]]\nop = \"explode\"\n";
        assert!(Recipe::from_toml(unknown_op).is_err());
        let bad_type = "version = 1\n[source]\nkind = \"csv\"\npath = \"x\"\n[[step]]\nop = \"types\"\ncolumns = { a = \"float\" }\n";
        assert!(Recipe::from_toml(bad_type).unwrap_err().contains("unknown type"));
        let future = "version = 2\n[source]\nkind = \"csv\"\npath = \"x\"\n";
        assert!(Recipe::from_toml(future).unwrap_err().contains("not supported"));
    }

    #[test]
    fn relative_source_paths_resolve_against_the_recipe() {
        let r = recipe();
        assert_eq!(r.source_path(Path::new("/data/reports"), None), PathBuf::from("/data/reports/export.csv"));
        assert_eq!(r.source_path(Path::new("/data"), Some(Path::new("/tmp/next.csv"))), PathBuf::from("/tmp/next.csv"));
    }
}
