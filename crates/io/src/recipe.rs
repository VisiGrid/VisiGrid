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
    /// An Apache Parquet file: column names and types come from its schema.
    Parquet(ParquetSource),
    /// One table of a local DuckDB database, opened read-only.
    Duckdb(DuckdbSource),
    /// One sheet of an Excel workbook (.xlsx, .xlsm, .xls). Formulas are
    /// never evaluated: the values Excel last calculated are read.
    Xlsx(XlsxSource),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct XlsxSource {
    /// Relative paths resolve against the recipe file's folder.
    pub path: String,
    /// The sheet's name; empty means the first sheet.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub sheet: String,
    /// The row holding the column names, counting from 1; rows above it
    /// are skipped. 0: no header; columns are named by letter.
    #[serde(default = "one", skip_serializing_if = "is_one")]
    pub header_row: usize,
    /// The column names when the recipe was saved.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
    /// With a pattern path: append the same sheet of every matching workbook.
    #[serde(default, skip_serializing_if = "is_false")]
    pub combine: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct ParquetSource {
    /// Relative paths resolve against the recipe file's folder.
    pub path: String,
    /// The column names when the recipe was saved.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
    /// With a pattern path: append every matching file.
    #[serde(default, skip_serializing_if = "is_false")]
    pub combine: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct DuckdbSource {
    /// Relative paths resolve against the recipe file's folder.
    pub path: String,
    /// `schema.table`, or a table name that is unique in the database.
    pub table: String,
    /// The column names when the recipe was saved.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
}

impl Source {
    pub fn path(&self) -> &str {
        match self {
            Source::Csv(s) => &s.path,
            Source::Parquet(s) => &s.path,
            Source::Duckdb(s) => &s.path,
            Source::Xlsx(s) => &s.path,
        }
    }

    pub fn set_path(&mut self, path: String) {
        match self {
            Source::Csv(s) => s.path = path,
            Source::Parquet(s) => s.path = path,
            Source::Duckdb(s) => s.path = path,
            Source::Xlsx(s) => s.path = path,
        }
    }

    /// The column names saved with the recipe, checked for drift.
    pub fn columns(&self) -> &[String] {
        match self {
            Source::Csv(s) => &s.columns,
            Source::Parquet(s) => &s.columns,
            Source::Duckdb(s) => &s.columns,
            Source::Xlsx(s) => &s.columns,
        }
    }

    pub fn columns_mut(&mut self) -> &mut Vec<String> {
        match self {
            Source::Csv(s) => &mut s.columns,
            Source::Parquet(s) => &mut s.columns,
            Source::Duckdb(s) => &mut s.columns,
            Source::Xlsx(s) => &mut s.columns,
        }
    }

    /// Whether every matching file is appended (Append folder).
    pub fn combine(&self) -> bool {
        match self {
            Source::Csv(s) => s.combine,
            Source::Parquet(s) => s.combine,
            Source::Xlsx(s) => s.combine,
            Source::Duckdb(_) => false,
        }
    }

    /// Turn appending every matching file on or off. DuckDB has no such mode.
    pub fn set_combine(&mut self, on: bool) {
        match self {
            Source::Csv(s) => s.combine = on,
            Source::Parquet(s) => s.combine = on,
            Source::Xlsx(s) => s.combine = on,
            Source::Duckdb(_) => {}
        }
    }

    /// "CSV", "Parquet", "DuckDB"
    pub fn label(&self) -> &'static str {
        match self {
            Source::Csv(_) => "CSV",
            Source::Parquet(_) => "Parquet",
            Source::Duckdb(_) => "DuckDB",
            Source::Xlsx(_) => "Excel",
        }
    }

    /// A new source of the kind a file's extension names, reading `path`:
    /// `.parquet`, `.duckdb`/`.db` (with `table`), anything else as CSV.
    pub fn for_file(path: String, table: Option<String>) -> Source {
        let lower = path.to_lowercase();
        if [".xlsx", ".xlsm", ".xls"].iter().any(|e| lower.ends_with(e)) {
            Source::Xlsx(XlsxSource { path, sheet: table.unwrap_or_default(), header_row: 1, columns: Vec::new(), combine: false })
        } else if lower.ends_with(".parquet") {
            Source::Parquet(ParquetSource { path, columns: Vec::new(), combine: false })
        } else if lower.ends_with(".duckdb") || table.is_some() {
            Source::Duckdb(DuckdbSource { path, table: table.unwrap_or_default(), columns: Vec::new() })
        } else {
            Source::Csv(CsvSource { path, delimiter: None, encoding: None, header_row: 1, decimal_comma: false, columns: Vec::new(), combine: false })
        }
    }
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
    /// With a pattern path: read every matching file and append them, rather
    /// than the newest one. A "Source file" column says where each row came from.
    #[serde(default, skip_serializing_if = "is_false")]
    pub combine: bool,
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
    /// One row per distinct combination of the `by` columns, with totals.
    /// Groups appear in the order they first occur. No `by` columns: one
    /// row totalling everything.
    Group {
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        by: Vec<String>,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        totals: Vec<Total>,
        #[serde(default, skip_serializing_if = "is_default_missing")]
        missing: Missing,
    },
    /// "Unpivot other columns": keep these columns, and turn every other
    /// column into rows of (name, value). A column added to next month's
    /// file is unpivoted too.
    Unpivot {
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        keep: Vec<String>,
        /// The new column holding the unpivoted columns' names.
        #[serde(default = "attribute", skip_serializing_if = "is_attribute")]
        names_to: String,
        /// The new column holding their values.
        #[serde(default = "value_name", skip_serializing_if = "is_value_name")]
        values_to: String,
        /// Leave out empty values (Power Query does the same).
        #[serde(default = "yes", skip_serializing_if = "is_true")]
        drop_empty: bool,
        #[serde(default, skip_serializing_if = "is_default_missing")]
        missing: Missing,
    },
}

fn attribute() -> String {
    "Attribute".into()
}
fn is_attribute(s: &String) -> bool {
    s == "Attribute"
}
fn value_name() -> String {
    "Value".into()
}
fn is_value_name(s: &String) -> bool {
    s == "Value"
}
fn yes() -> bool {
    true
}
fn is_true(b: &bool) -> bool {
    *b
}

/// One total in a Group step: `{ fn = "sum", column = "Amount", as = "Total" }`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Total {
    #[serde(rename = "fn")]
    pub func: TotalFn,
    /// The column it totals; not used by `count_rows`.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub column: String,
    /// The result column's name.
    #[serde(rename = "as")]
    pub name: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TotalFn {
    Sum,
    Average,
    Min,
    Max,
    /// Values that aren't empty.
    Count,
    /// Rows in the group, empty or not.
    CountRows,
    /// Distinct values that aren't empty.
    Distinct,
    First,
    Last,
}

impl TotalFn {
    pub const ALL: [TotalFn; 9] = [
        TotalFn::Sum,
        TotalFn::Count,
        TotalFn::CountRows,
        TotalFn::Average,
        TotalFn::Min,
        TotalFn::Max,
        TotalFn::Distinct,
        TotalFn::First,
        TotalFn::Last,
    ];

    pub fn label(self) -> &'static str {
        match self {
            TotalFn::Sum => "sum",
            TotalFn::Average => "average",
            TotalFn::Min => "min",
            TotalFn::Max => "max",
            TotalFn::Count => "count",
            TotalFn::CountRows => "count rows",
            TotalFn::Distinct => "distinct count",
            TotalFn::First => "first",
            TotalFn::Last => "last",
        }
    }

    /// A default result name: `Total Amount`, `Rows`.
    pub fn default_name(self, column: &str) -> String {
        match self {
            TotalFn::CountRows => "Rows".into(),
            TotalFn::Sum => format!("Total {column}"),
            TotalFn::Average => format!("Average {column}"),
            TotalFn::Min => format!("Min {column}"),
            TotalFn::Max => format!("Max {column}"),
            TotalFn::Count => format!("Count of {column}"),
            TotalFn::Distinct => format!("Distinct {column}"),
            TotalFn::First => format!("First {column}"),
            TotalFn::Last => format!("Last {column}"),
        }
    }
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
            Step::Group { by, totals, .. } => {
                let totals: Vec<String> = totals
                    .iter()
                    .map(|t| match t.func {
                        TotalFn::CountRows => format!("{} = count of rows", t.name),
                        f => format!("{} = {} of {}", t.name, f.label(), t.column),
                    })
                    .collect();
                let head = if by.is_empty() { "Total all rows".to_string() } else { format!("Group by {}", list(by)) };
                if totals.is_empty() { head } else { format!("{head}: {}", totals.join(", ")) }
            }
            Step::Unpivot { keep, names_to, values_to, .. } if keep.is_empty() => {
                format!("Unpivot every column into {names_to}, {values_to}")
            }
            Step::Unpivot { keep, names_to, values_to, .. } => {
                format!("Unpivot all but {} into {names_to}, {values_to}", list(keep))
            }
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
            | Step::Dedupe { missing, .. }
            | Step::Group { missing, .. }
            | Step::Unpivot { missing, .. } => *missing,
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
            Step::Group { by, totals, .. } => by
                .iter()
                .map(String::as_str)
                .chain(totals.iter().filter(|t| !t.column.is_empty()).map(|t| t.column.as_str()))
                .collect(),
            Step::Unpivot { keep, .. } => keep.iter().map(String::as_str).collect(),
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

    /// This recipe as `existing` (the file's current text) rewritten with
    /// only what changed: comments, key order, inline tables and spacing
    /// survive wherever the structure is the same. A step added or removed
    /// rewrites the step list. Falls back to a plain rewrite if the result
    /// would read back as anything other than this recipe.
    pub fn to_toml_preserving(&self, existing: &str) -> String {
        let fresh = self.to_toml();
        let (Ok(mut old), Ok(new)) = (existing.parse::<toml_edit::DocumentMut>(), fresh.parse::<toml_edit::DocumentMut>()) else {
            return fresh;
        };
        merge_toml_table(old.as_table_mut(), new.as_table());
        let merged = old.to_string();
        match Recipe::from_toml(&merged) {
            Ok(r) if r == *self => merged,
            _ => fresh,
        }
    }

    pub fn load(path: &Path) -> Result<Recipe, String> {
        let bytes = read_regular_file(path, MAX_RECIPE_BYTES, "recipe")?;
        let text = String::from_utf8(bytes).map_err(|_| format!("{}: not UTF-8 text", path.display()))?;
        Recipe::from_toml(&text).map_err(|e| format!("{}: {e}", path.display()))
    }

    /// The source renamed column `old` to `new`: point every step that reads
    /// the source's `old` at `new`, and record `new` in the saved column
    /// list. Steps after one that renames `old` away, or that creates a new
    /// column called `old`, refer to that column and are left alone. Returns
    /// whether anything changed.
    pub fn rename_source_column(&mut self, old: &str, new: &str) -> bool {
        let same = |a: &str| a.eq_ignore_ascii_case(old);
        let mut changed = false;
        for c in self.source.columns_mut() {
            if same(c) {
                *c = new.to_string();
                changed = true;
            }
        }
        for step in &mut self.steps {
            let swap = |names: &mut Vec<String>, changed: &mut bool| {
                for n in names.iter_mut() {
                    if same(n) {
                        *n = new.to_string();
                        *changed = true;
                    }
                }
            };
            let swap_keys = |map: &mut BTreeMap<String, String>, changed: &mut bool| {
                if let Some(key) = map.keys().find(|k| same(k)).cloned() {
                    let v = map.remove(&key).unwrap();
                    map.insert(new.to_string(), v);
                    *changed = true;
                }
            };
            match step {
                Step::Select { columns, .. } => {
                    swap(columns, &mut changed);
                }
                Step::Trim { columns, .. } | Step::Dedupe { columns, .. } => swap(columns, &mut changed),
                Step::Types { columns, .. } => swap_keys(columns, &mut changed),
                Step::Filter { column, .. } => {
                    if same(column) {
                        *column = new.to_string();
                        changed = true;
                    }
                }
                Step::Remove { columns, .. } => {
                    let gone = columns.iter().any(|c| same(c));
                    swap(columns, &mut changed);
                    if gone {
                        break;
                    }
                }
                Step::Rename { columns, .. } => {
                    let renamed_away = columns.keys().any(|k| same(k));
                    let created = columns.values().any(|v| same(v));
                    swap_keys(columns, &mut changed);
                    if renamed_away || created {
                        break;
                    }
                }
                // After grouping or unpivoting the source's columns are
                // gone; later steps name the new ones
                Step::Group { by, totals, .. } => {
                    swap(by, &mut changed);
                    for t in totals.iter_mut().filter(|t| same(&t.column)) {
                        t.column = new.to_string();
                        changed = true;
                    }
                    break;
                }
                Step::Unpivot { keep, .. } => {
                    swap(keep, &mut changed);
                    break;
                }
            }
        }
        changed
    }

    /// Set what step `step` (counting from 1) does with values that do not
    /// fit their type. Only a Set types step has this.
    pub fn set_on_error(&mut self, step: usize, value: OnError) -> Result<(), String> {
        match self.steps.get_mut(step.wrapping_sub(1)) {
            Some(Step::Types { on_error, .. }) => {
                *on_error = value;
                Ok(())
            }
            Some(other) => Err(format!("step {step} ({}) does not check types", other.describe())),
            None => Err(format!("the recipe has no step {step}")),
        }
    }

    /// Save to `path`, replacing it whole (through a temporary file, so a
    /// failed write leaves the old recipe).
    pub fn save(&self, path: &Path) -> Result<(), String> {
        let dir = path.parent().filter(|d| !d.as_os_str().is_empty()).unwrap_or(Path::new("."));
        let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("recipe.toml");
        let tmp = dir.join(format!(".{name}.{}.partial", std::process::id()));
        // An existing recipe keeps its comments and layout: only what changed
        // is rewritten
        let text = match read_regular_file(path, MAX_RECIPE_BYTES, "recipe").ok().and_then(|b| String::from_utf8(b).ok()) {
            Some(existing) => self.to_toml_preserving(&existing),
            None => self.to_toml(),
        };
        std::fs::write(&tmp, text).map_err(|e| format!("{}: {e}", tmp.display()))?;
        std::fs::rename(&tmp, path).map_err(|e| {
            let _ = std::fs::remove_file(&tmp);
            format!("{}: {e}", path.display())
        })
    }

    /// The files a run reads: one, or with `combine` every file the pattern
    /// matches, in name order (skipping downloads in progress, and waiting
    /// while any changed in the last moments).
    pub fn resolve_sources(&self, recipe_dir: &Path, over: Option<&Path>) -> Result<Vec<PathBuf>, String> {
        if over.is_some() || !self.source.combine() || !self.source_is_pattern() {
            return self.resolve_source(recipe_dir, over).map(|p| vec![p]);
        }
        let path = self.source_path(recipe_dir, None);
        check_local(&path, "source")?;
        let pattern = path.file_name().and_then(|n| n.to_str()).unwrap_or("").to_string();
        let dir = path.parent().filter(|d| !d.as_os_str().is_empty()).unwrap_or(Path::new("."));
        let entries = std::fs::read_dir(dir).map_err(|e| format!("{}: {e}", dir.display()))?;
        let mut files = Vec::new();
        for entry in entries.flatten() {
            let name = entry.file_name().to_string_lossy().into_owned();
            if !wildcard_match(&pattern, &name) || is_partial_download(&name) {
                continue;
            }
            let Ok(meta) = entry.metadata() else { continue };
            if !meta.is_file() {
                continue;
            }
            if let Ok(age) = std::time::SystemTime::now().duration_since(meta.modified().unwrap_or(std::time::UNIX_EPOCH)) {
                if age < SETTLE_TIME {
                    return Err(format!("{name} is still being written (it changed a moment ago); refresh again in a few seconds"));
                }
            }
            files.push((name, entry.path()));
        }
        if files.is_empty() {
            return Err(format!("no file in {} matches {pattern}", dir.display()));
        }
        files.sort();
        Ok(files.into_iter().map(|(_, p)| p).collect())
    }

    /// Read the run's source once: the file, or every file it appends.
    pub fn read_snapshot(&self, recipe_dir: &Path, over: Option<&Path>) -> Result<Snapshot, String> {
        Snapshot::read_all(&self.resolve_sources(recipe_dir, over)?)
    }

    /// Whether the source names a pattern (`export-*.csv`) rather than a file.
    pub fn source_is_pattern(&self) -> bool {
        is_pattern(Path::new(self.source.path()).file_name().and_then(|n| n.to_str()).unwrap_or(""))
    }

    /// The file a run reads: `over` if given; else the recipe's path, or for
    /// a pattern (`*` and `?` in the file name) the most recently modified
    /// file it matches, so next month's export is picked up by itself.
    pub fn resolve_source(&self, recipe_dir: &Path, over: Option<&Path>) -> Result<PathBuf, String> {
        let path = self.source_path(recipe_dir, over);
        check_local(&path, "source")?;
        if over.is_some() {
            return Ok(path);
        }
        let Some(pattern) = path.file_name().and_then(|n| n.to_str()).filter(|n| is_pattern(n)) else {
            return Ok(path);
        };
        let dir = path.parent().filter(|d| !d.as_os_str().is_empty()).unwrap_or(Path::new("."));
        let entries = std::fs::read_dir(dir).map_err(|e| format!("{}: {e}", dir.display()))?;
        let mut best: Option<(std::time::SystemTime, String, PathBuf)> = None;
        for entry in entries.flatten() {
            let name = entry.file_name().to_string_lossy().into_owned();
            if !wildcard_match(pattern, &name) || is_partial_download(&name) {
                continue;
            }
            let Ok(meta) = entry.metadata() else { continue };
            if !meta.is_file() {
                continue;
            }
            let modified = meta.modified().unwrap_or(std::time::UNIX_EPOCH);
            // Newest first; the same time goes to the later name (…-10 over …-09)
            if best.as_ref().map_or(true, |(t, n, _)| (modified, &name) > (*t, n)) {
                best = Some((modified, name, entry.path()));
            }
        }
        let (modified, name, path) = best.ok_or_else(|| format!("no file in {} matches {pattern}", dir.display()))?;
        // A file that changed a moment ago may still be downloading or
        // being written: reading it now could load half an export
        if let Ok(age) = std::time::SystemTime::now().duration_since(modified) {
            if age < SETTLE_TIME {
                return Err(format!("{name} is still being written (it changed a moment ago); refresh again in a few seconds"));
            }
        }
        Ok(path)
    }

    /// The source file: `over` if given, else the recipe's path, relative
    /// paths resolved against `recipe_dir`.
    pub fn source_path(&self, recipe_dir: &Path, over: Option<&Path>) -> PathBuf {
        if let Some(p) = over {
            return p.to_path_buf();
        }
        let p = Path::new(self.source.path());
        if p.is_absolute() {
            p.to_path_buf()
        } else {
            recipe_dir.join(p)
        }
    }
}

/// How long the newest file a pattern matches must have been unchanged.
pub const SETTLE_TIME: std::time::Duration = std::time::Duration::from_secs(2);

/// Names browsers and tools use while a download or write is in progress,
/// and hidden files. A pattern never picks these.
fn is_partial_download(name: &str) -> bool {
    let lower = name.to_lowercase();
    lower.starts_with('.')
        || lower.starts_with("~$")
        || lower.ends_with('~')
        || [".crdownload", ".part", ".partial", ".download", ".tmp", ".temp", ".opdownload"].iter().any(|s| lower.ends_with(s))
}

fn is_pattern(name: &str) -> bool {
    name.contains('*') || name.contains('?')
}

/// `*` any run of characters, `?` one character; case-insensitive, since
/// exports land on case-insensitive disks as often as not.
pub fn wildcard_match(pattern: &str, name: &str) -> bool {
    let p: Vec<char> = pattern.to_lowercase().chars().collect();
    let n: Vec<char> = name.to_lowercase().chars().collect();
    let (mut pi, mut ni) = (0, 0);
    let (mut star, mut mark) = (None, 0);
    while ni < n.len() {
        if pi < p.len() && (p[pi] == '?' || p[pi] == n[ni]) {
            pi += 1;
            ni += 1;
        } else if pi < p.len() && p[pi] == '*' {
            star = Some(pi);
            mark = ni;
            pi += 1;
        } else if let Some(s) = star {
            pi = s + 1;
            mark += 1;
            ni = mark;
        } else {
            return false;
        }
    }
    p[pi..].iter().all(|c| *c == '*')
}

/// A pattern for a dated export's name: each run of digits becomes `*`
/// (`export-2026-09.csv` -> `export-*-*.csv`). None without digits.
pub fn suggest_pattern(file_name: &str) -> Option<String> {
    let mut out = String::new();
    let mut in_digits = false;
    let mut any = false;
    for c in file_name.chars() {
        if c.is_ascii_digit() {
            if !in_digits {
                out.push('*');
                any = true;
            }
            in_digits = true;
        } else {
            out.push(c);
            in_digits = false;
        }
    }
    any.then_some(out)
}

// ============================================================================
// Rewriting a recipe file without losing its comments
// ============================================================================

fn merge_toml_table(old: &mut toml_edit::Table, new: &toml_edit::Table) {
    let gone: Vec<String> = old.iter().map(|(k, _)| k.to_string()).filter(|k| !new.contains_key(k)).collect();
    for k in gone {
        old.remove(&k);
    }
    for (k, item) in new.iter() {
        match old.get_mut(k) {
            Some(existing) => merge_toml_item(existing, item),
            None => {
                old.insert(k, item.clone());
            }
        }
    }
}

fn merge_toml_item(old: &mut toml_edit::Item, new: &toml_edit::Item) {
    use toml_edit::{Item, Value};
    match new {
        Item::Table(nt) => {
            if let Item::Table(ot) = old {
                return merge_toml_table(ot, nt);
            }
            // The file wrote it inline ({ a = "b" }): keep it inline
            if let Item::Value(ov @ Value::InlineTable(_)) = old {
                return merge_toml_value(ov, &Value::InlineTable(nt.clone().into_inline_table()));
            }
        }
        Item::ArrayOfTables(na) => {
            if let Item::ArrayOfTables(oa) = old {
                if oa.len() == na.len() {
                    for (o, n) in oa.iter_mut().zip(na.iter()) {
                        merge_toml_table(o, n);
                    }
                    return;
                }
            }
        }
        Item::Value(nv) => {
            if let Item::Value(ov) = old {
                return merge_toml_value(ov, nv);
            }
        }
        Item::None => {}
    }
    *old = new.clone();
}

fn merge_toml_value(old: &mut toml_edit::Value, new: &toml_edit::Value) {
    use toml_edit::Value;
    if same_toml_value(old, new) {
        return;
    }
    if let (Value::InlineTable(oi), Value::InlineTable(ni)) = (&mut *old, new) {
        let gone: Vec<String> = oi.iter().map(|(k, _)| k.to_string()).filter(|k| !ni.contains_key(k)).collect();
        let reshaped = !gone.is_empty() || ni.iter().any(|(k, _)| !oi.contains_key(k));
        for k in gone {
            oi.remove(&k);
        }
        for (k, v) in ni.iter() {
            match oi.get_mut(k) {
                Some(existing) => merge_toml_value(existing, v),
                None => {
                    oi.insert(k, v.clone());
                }
            }
        }
        // Keys came or went: tidy the separators (an inline table can't
        // hold comments, so nothing of the user's is lost)
        if reshaped {
            oi.fmt();
        }
        return;
    }
    // A changed value keeps the comments and spacing around it
    let decor = old.decor().clone();
    *old = new.clone();
    *old.decor_mut() = decor;
}

/// Equal as TOML values, however they are written ("a" and 'a').
fn same_toml_value(a: &toml_edit::Value, b: &toml_edit::Value) -> bool {
    let plain = |v: &toml_edit::Value| {
        let mut v = v.clone();
        v.decor_mut().clear();
        toml::from_str::<toml::Table>(&format!("v = {v}")).ok().and_then(|mut t| t.remove("v"))
    };
    plain(a).is_some() && plain(a) == plain(b)
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
    /// blake3 of the bytes (hex, 16 characters); of every file's, appended.
    pub hash: String,
    /// The other files of an appended source, in name order after `path`.
    pub more: Vec<Snapshot>,
}

/// The largest source file a recipe reads. The whole file is held in memory
/// (once as bytes, once parsed), so this also bounds what a run costs.
pub const MAX_SOURCE_BYTES: u64 = 256 * 1024 * 1024;
/// The largest recipe file read. Recipes are a few kilobytes of TOML.
pub const MAX_RECIPE_BYTES: u64 = 1024 * 1024;

/// A UNC or device path (`\\server\share`, `//server/share`, `\\?\…`).
/// Recipes refuse these before touching the filesystem: on Windows merely
/// opening one sends the user's network credentials to that server, and a
/// recipe or a shared workbook can name any path.
pub fn is_network_path(path: &Path) -> bool {
    let s = path.as_os_str().to_string_lossy();
    s.starts_with("\\\\") || s.starts_with("//") || s.starts_with("\\/") || s.starts_with("/\\")
}

/// Refuse a network path, with the reason.
pub fn check_local(path: &Path, what: &str) -> Result<(), String> {
    if is_network_path(path) {
        return Err(format!(
            "{what} {} is on a network share; recipes only read local files (copy it to this computer first)",
            path.display()
        ));
    }
    Ok(())
}

/// Read a whole regular file, refusing anything else (a pipe, a device, a
/// directory) and anything over `cap` bytes. The checks are made on the
/// opened file, so a path swapped after checking is still caught.
pub fn read_regular_file(path: &Path, cap: u64, what: &str) -> Result<Vec<u8>, String> {
    use std::io::Read;
    check_local(path, what)?;
    let shown = path.display();
    // Not following through to open a FIFO for reading blocks until a writer
    // appears; check the type before opening, then again on the handle
    let meta = std::fs::metadata(path).map_err(|e| format!("cannot read {what} {shown}: {e}"))?;
    if !meta.is_file() {
        return Err(format!("{what} {shown} is not a regular file"));
    }
    let file = std::fs::File::open(path).map_err(|e| format!("cannot read {what} {shown}: {e}"))?;
    let meta = file.metadata().map_err(|e| format!("cannot read {what} {shown}: {e}"))?;
    if !meta.is_file() {
        return Err(format!("{what} {shown} is not a regular file"));
    }
    let too_big = |n: u64| format!("{what} {shown} is {} MB; recipes read files up to {} MB", n / (1024 * 1024), cap / (1024 * 1024));
    if meta.len() > cap {
        return Err(too_big(meta.len()));
    }
    let mut bytes = Vec::with_capacity(meta.len() as usize);
    file.take(cap + 1).read_to_end(&mut bytes).map_err(|e| format!("cannot read {what} {shown}: {e}"))?;
    if bytes.len() as u64 > cap {
        return Err(too_big(bytes.len() as u64));
    }
    Ok(bytes)
}

impl Snapshot {
    /// Read a source: a local regular file of at most [`MAX_SOURCE_BYTES`].
    pub fn read(path: &Path) -> Result<Snapshot, String> {
        let bytes = read_regular_file(path, MAX_SOURCE_BYTES, "source")?;
        Ok(Snapshot::from_bytes(path, bytes))
    }

    pub fn from_bytes(path: &Path, bytes: Vec<u8>) -> Snapshot {
        let hash = blake3::hash(&bytes).to_hex()[..16].to_string();
        Snapshot { path: path.to_path_buf(), bytes: Arc::new(bytes), hash, more: Vec::new() }
    }

    /// Several files read once, for an appended source. The first is the
    /// snapshot's own; the hash covers them all.
    pub fn read_all(paths: &[PathBuf]) -> Result<Snapshot, String> {
        let (first, rest) = paths.split_first().ok_or("no files to read")?;
        let mut snapshot = Snapshot::read(first)?;
        for p in rest {
            snapshot.more.push(Snapshot::read(p)?);
        }
        if !snapshot.more.is_empty() {
            let mut h = blake3::Hasher::new();
            for s in std::iter::once(&snapshot).chain(snapshot.more.iter()) {
                h.update(s.hash.as_bytes());
            }
            snapshot.hash = h.finalize().to_hex()[..16].to_string();
        }
        Ok(snapshot)
    }

    /// How many files the snapshot holds.
    pub fn file_count(&self) -> usize {
        1 + self.more.len()
    }
}

// ============================================================================
// Running
// ============================================================================

/// A value that did not fit, and where it came from.
#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct CellError {
    /// The step that checked it, counting from 1.
    pub step: usize,
    /// The file it is in, when several files were appended.
    #[serde(skip_serializing_if = "String::is_empty")]
    pub file: String,
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
    /// This step failed the run.
    #[serde(skip_serializing_if = "is_false")]
    pub failed: bool,
    /// Columns it names that the source does not have, when that failed it.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub missing: Vec<String>,
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
    /// Worth knowing, not a failure: appended files that lack a column the
    /// others have (their rows are empty there).
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub warnings: Vec<String>,
}

/// One output column: its name and the type its values were checked against.
#[derive(Debug, Clone, PartialEq)]
pub struct OutColumn {
    pub name: String,
    pub rule: ColumnRule,
    /// How a typed source declared the column, where the rule alone would
    /// lose it. A Set types step on the column clears it.
    pub kind: ValueKind,
}

/// Typed-source columns written back with their type: numbers as the
/// source stated them, date-times and times as such (their values are ISO
/// text in the recipe, so filters compare them in order).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize)]
pub enum ValueKind {
    #[default]
    Plain,
    Number,
    DateTime,
    Time,
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
    /// The file each row came from: an index into `file_names`. Empty
    /// unless several files were appended.
    files: Vec<u32>,
    file_names: Vec<String>,
    decimal_comma: bool,
}

/// How a column name a step uses matches the table.
#[derive(Debug, Clone, PartialEq)]
enum Resolve {
    One(usize),
    None,
    /// More than one column has this name: never guess which.
    Many(usize),
}

impl Frame {
    /// Exact name first, then ignoring case (headers vary in case between
    /// exports). Two matches at the same level are ambiguous.
    fn resolve(&self, name: &str) -> Resolve {
        let exact: Vec<usize> = (0..self.columns.len()).filter(|&i| self.columns[i].name == name).collect();
        let hits = if exact.is_empty() {
            (0..self.columns.len()).filter(|&i| self.columns[i].name.eq_ignore_ascii_case(name)).collect()
        } else {
            exact
        };
        match hits.len() {
            0 => Resolve::None,
            1 => Resolve::One(hits[0]),
            n => Resolve::Many(n),
        }
    }

    /// The column for a name that has already been checked to be unique.
    fn find(&self, name: &str) -> Option<usize> {
        match self.resolve(name) {
            Resolve::One(i) => Some(i),
            _ => None,
        }
    }

    fn add_blank(&mut self, name: &str) -> usize {
        self.columns.push(OutColumn { name: name.to_string(), rule: ColumnRule::Auto, kind: ValueKind::Plain });
        for row in &mut self.rows {
            row.push(String::new());
        }
        self.columns.len() - 1
    }

    /// The file a row came from, when several were appended.
    fn file_of(&self, row: usize) -> String {
        self.files.get(row).and_then(|&f| self.file_names.get(f as usize)).cloned().unwrap_or_default()
    }

    fn keep_rows(&mut self, keep: &[bool]) {
        if !self.files.is_empty() {
            let mut k = 0;
            self.files.retain(|_| {
                k += 1;
                keep[k - 1]
            });
        }
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
    let decimal_comma = matches!(&recipe.source, Source::Csv(s) if s.decimal_comma);
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
        warnings: Vec::new(),
    };

    let read_one = |snap: &Snapshot| match &recipe.source {
        Source::Csv(src) => read_csv(src, snap),
        Source::Parquet(_) => read_parquet(snap),
        Source::Duckdb(src) => read_duckdb(src, snap),
        Source::Xlsx(src) => read_xlsx(src, snap),
    };
    let read = if snapshot.more.is_empty() {
        read_one(snapshot)
    } else {
        // Append folder: every file read the same way, then stacked
        let mut parts = Vec::with_capacity(snapshot.file_count());
        let mut failed = None;
        for snap in std::iter::once(snapshot).chain(snapshot.more.iter()) {
            let name = snap.path.file_name().and_then(|n| n.to_str()).unwrap_or("").to_string();
            match read_one(snap) {
                Ok(frame) => parts.push((name, frame)),
                Err(e) => {
                    failed = Some(format!("{name}: {e}"));
                    break;
                }
            }
        }
        match failed {
            Some(e) => Err(e),
            None => Ok(append_frames(parts, &mut report.warnings)),
        }
    };
    if !snapshot.more.is_empty() {
        report.source = format!(
            "{} files in {}",
            snapshot.file_count(),
            snapshot.path.parent().map(|p| p.display().to_string()).unwrap_or_default()
        );
    }
    let mut frame = match read {
        Ok(frame) => frame,
        Err(e) => {
            report.ok = false;
            report.failures.push(e);
            return RunResult {
                output: RecipeOutput { columns: Vec::new(), rows: Vec::new(), decimal_comma },
                report,
            };
        }
    };
    report.source_rows = frame.rows.len();
    report.drift = drift(recipe.source.columns(), &frame.columns);

    for (index, step) in recipe.steps.iter().enumerate() {
        let started = Instant::now();
        let rows_in = frame.rows.len();
        let mut skipped = false;
        let mut note = None;
        let mut missing = Vec::new();
        let (failures_before, errors_before) = (report.failures.len(), report.errors.len());

        // Columns this step names that are not there
        let absent: Vec<String> = step
            .named_columns()
            .into_iter()
            .filter(|c| frame.resolve(c) == Resolve::None)
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
                    missing = absent.clone();
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

        // A name matching several columns, or two names in one step reaching
        // the same column, would silently pick or lose data: always fail
        if !skipped {
            if let Err(reason) = check_bindings(step, &frame) {
                report.ok = false;
                report.failures.push(format!("step {} ({}): {reason}", index + 1, step.describe()));
                skipped = true;
            }
        }

        if !skipped {
            let step_note = apply(step, &mut frame, &mut report);
            if step_note.is_some() {
                note = step_note;
            }
        }

        for e in &mut report.errors[errors_before..] {
            e.step = index + 1;
        }
        report.steps.push(StepReport {
            index: index + 1,
            description: step.describe(),
            rows_in,
            rows_out: frame.rows.len(),
            millis: started.elapsed().as_millis(),
            // A step that failed the run is not "skipped": it failed
            skipped: skipped && missing.is_empty() && report.failures.len() == failures_before,
            note,
            failed: report.failures.len() > failures_before,
            missing,
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
        output: RecipeOutput { columns: frame.columns, rows: frame.rows, decimal_comma },
        report,
    }
}

/// What a source's settings resolve to for one snapshot: the delimiter and
/// encoding actually used, and the file's first lines, for the builder.
#[derive(Debug, Clone)]
pub struct SourceInfo {
    pub delimiter: u8,
    pub encoding: Encoding,
    /// The first lines of the file, as a text editor shows them.
    pub first_lines: Vec<String>,
}

/// Resolve `src` against `snapshot` without running anything.
pub fn source_info(src: &CsvSource, snapshot: &Snapshot) -> SourceInfo {
    let encoding = src.encoding.as_deref().and_then(Encoding::parse);
    let (content, encoding) = decode(&snapshot.bytes, encoding);
    let skip = src.header_row.saturating_sub(1);
    let body: Vec<&str> = content.lines().skip(skip).take(20).collect();
    let delimiter = match src.delimiter.as_deref() {
        Some("tab") | Some("\t") => b'\t',
        Some(d) if d.len() == 1 => d.as_bytes()[0],
        _ => sniff_delimiter(&body.join("\n")),
    };
    let first_lines = content.lines().take(12).map(|l| l.chars().take(160).collect()).collect();
    SourceInfo { delimiter, encoding, first_lines }
}

/// The line that most likely holds the column names (counting from 1):
/// the first line with more than one field whose field count the next
/// lines repeat. Title and "generated on" lines above a table have fewer.
/// 1 when nothing stands out.
pub fn guess_header_row(snapshot: &Snapshot) -> usize {
    let (content, _) = decode(&snapshot.bytes, None);
    let lines: Vec<&str> = content.lines().take(40).collect();
    let delimiter = sniff_delimiter(&lines.join("\n"));
    let count = |line: &str| {
        csv::ReaderBuilder::new()
            .delimiter(delimiter)
            .has_headers(false)
            .flexible(true)
            .from_reader(line.as_bytes())
            .records()
            .next()
            .and_then(Result::ok)
            .map_or(0, |r| r.len())
    };
    let counts: Vec<usize> = lines.iter().map(|l| count(l)).collect();
    for i in 0..counts.len().min(10) {
        let n = counts[i];
        let next: Vec<usize> = counts[i + 1..].iter().copied().filter(|c| *c > 0).take(3).collect();
        if n > 1 && !next.is_empty() && next.iter().all(|c| *c == n) {
            return i + 1;
        }
    }
    1
}

// ============================================================================
// Typed sources: Parquet and DuckDB
// ============================================================================

/// The snapshot's bytes as a file, for readers that need a path. The
/// snapshot stays the only copy of the source a run reads.
fn snapshot_file(snapshot: &Snapshot, ext: &str) -> Result<(tempfile::TempDir, PathBuf), String> {
    let dir = tempfile::tempdir().map_err(|e| e.to_string())?;
    let path = dir.path().join(format!("source.{ext}"));
    std::fs::write(&path, snapshot.bytes.as_slice()).map_err(|e| e.to_string())?;
    Ok((dir, path))
}

fn read_parquet(snapshot: &Snapshot) -> Result<Frame, String> {
    let (_dir, path) = snapshot_file(snapshot, "parquet")?;
    frame_from_import(crate::parquet::import(&path)?)
}

fn read_duckdb(src: &DuckdbSource, snapshot: &Snapshot) -> Result<Frame, String> {
    if src.table.trim().is_empty() {
        return Err("the recipe names no DuckDB table; set source.table".into());
    }
    let (_dir, path) = snapshot_file(snapshot, "duckdb")?;
    let db = crate::duckdb::Database::open(&path)?;
    let index = db.resolve(&src.table)?;
    frame_from_import(db.read_table(index, crate::parquet::MAX_ROWS - 1, None, true)?)
}

/// What a typed cell is, for choosing its column's kind.
#[derive(Clone, Copy, PartialEq)]
enum CellKind {
    Number,
    Date,
    DateTime,
    Time,
    Text,
}

/// A typed import as recipe text: numbers as the importer stated them,
/// dates as YYYY-MM-DD (a Date column), date-times and times as ISO text.
/// A column whose cells are all one kind keeps it; a mixed column (huge
/// integers the importer kept as text, say) becomes text.
fn frame_from_import(import: crate::parquet::ParquetImport) -> Result<Frame, String> {
    if import.truncated() {
        return Err(format!(
            "the source has {} rows and {} columns; a sheet holds {} rows (including the header) and {} columns",
            import.total_rows,
            import.total_cols,
            crate::parquet::MAX_ROWS,
            crate::parquet::MAX_COLS
        ));
    }
    Ok(frame_from_sheet(&import.sheet, Some(0), import.rows_loaded, import.cols_loaded))
}

/// A sheet's typed cells as a recipe frame: names from `header` (a row
/// index; None names columns by letter), then `count` records below it,
/// `width` columns wide. Lines in reports are the sheet's row numbers.
fn frame_from_sheet(sheet: &Sheet, header: Option<usize>, count: usize, width: usize) -> Frame {
    let first = header.map_or(0, |h| h + 1);
    let mut columns = Vec::with_capacity(width);
    let mut rows: Vec<Vec<String>> = vec![Vec::with_capacity(width); count];
    for col in 0..width {
        let mut kind: Option<CellKind> = None;
        let mut mixed = false;
        let mut values = Vec::with_capacity(count);
        for row in first..first + count {
            let (text, k) = typed_cell_text(sheet, row, col);
            if let Some(k) = k {
                match kind {
                    None => kind = Some(k),
                    Some(prev) if prev != k => mixed = true,
                    _ => {}
                }
            }
            values.push(text);
        }
        let (rule, value_kind) = match (kind, mixed) {
            (_, true) | (Some(CellKind::Text), _) => (ColumnRule::Text, ValueKind::Plain),
            (Some(CellKind::Number), _) => (ColumnRule::Number, ValueKind::Number),
            (Some(CellKind::Date), _) => (ColumnRule::Date(DateOrder::Ymd), ValueKind::Plain),
            (Some(CellKind::DateTime), _) => (ColumnRule::Text, ValueKind::DateTime),
            (Some(CellKind::Time), _) => (ColumnRule::Text, ValueKind::Time),
            (None, _) => (ColumnRule::Auto, ValueKind::Plain),
        };
        let name = header.map(|h| sheet.get_raw(h, col).trim().to_string()).filter(|n| !n.is_empty());
        columns.push(OutColumn { name: name.unwrap_or_else(|| crate::csv_import::col_label(col)), rule, kind: value_kind });
        for (row, value) in values.into_iter().enumerate() {
            rows[row].push(value);
        }
    }
    // A record's row number in the sheet (or the file, for Parquet)
    let lines = (first + 1..first + count + 1).collect();
    Frame { columns, rows, lines, files: Vec::new(), file_names: Vec::new(), decimal_comma: false }
}

/// An Excel workbook's sheet names, in order, without importing it.
pub fn xlsx_sheet_names(path: &Path) -> Result<Vec<String>, String> {
    use calamine::Reader as _;
    let workbook = calamine::open_workbook_auto(path).map_err(|e| e.to_string())?;
    Ok(workbook.sheet_names().to_vec())
}

/// The row of `sheet` (counting from 1) that most likely holds the column
/// names: the same rule as for CSV, applied to the sheet's rows. 1 when
/// nothing stands out.
pub fn guess_xlsx_header_row(snapshot: &Snapshot, sheet: &str) -> usize {
    let probe = XlsxSource { path: String::new(), sheet: sheet.to_string(), header_row: 0, columns: Vec::new(), combine: false };
    let Ok(frame) = read_xlsx(&probe, snapshot) else { return 1 };
    // A header names every column, so it fills as many cells as the widest
    // row near the top; title rows above it don't. Unlike CSV, empty data
    // cells leave nothing to count, so rows below may fill fewer.
    let counts: Vec<usize> = frame.rows.iter().take(40).map(|r| r.iter().filter(|v| !v.trim().is_empty()).count()).collect();
    let widest = counts.iter().copied().max().unwrap_or(0);
    if widest < 2 {
        return 1;
    }
    counts.iter().take(10).position(|c| *c == widest).map_or(1, |i| i + 1)
}

fn read_xlsx(src: &XlsxSource, snapshot: &Snapshot) -> Result<Frame, String> {
    // The reader picks the format from the extension
    let ext = snapshot.path.extension().and_then(|e| e.to_str()).unwrap_or("xlsx").to_lowercase();
    let (_dir, path) = snapshot_file(snapshot, &ext)?;
    // Values only: the results Excel saved, never a recalculation here
    let options = crate::xlsx::ImportOptions { values_only: true, ..Default::default() };
    let (wb, result) = crate::xlsx::import_with_options(&path, &options)?;
    if result.truncated {
        return Err("the workbook is larger than VisiGrid imports; save the sheet you need on its own".into());
    }
    let names: Vec<String> = (0..wb.sheet_count()).filter_map(|i| wb.sheet(i).map(|s| s.name.clone())).collect();
    let index = if src.sheet.trim().is_empty() {
        0
    } else {
        names
            .iter()
            .position(|n| n.eq_ignore_ascii_case(src.sheet.trim()))
            .ok_or_else(|| format!("the workbook has no sheet {:?}; its sheets are {}", src.sheet, names.join(", ")))?
    };
    let sheet = wb.sheet(index).ok_or("the workbook has no sheets")?;
    let (mut last_row, mut last_col) = (None::<usize>, 0usize);
    for ((row, col), _) in sheet.cells_iter() {
        if !sheet.get_raw(row, col).is_empty() {
            last_row = Some(last_row.map_or(row, |r| r.max(row)));
            last_col = last_col.max(col + 1);
        }
    }
    let header = src.header_row.checked_sub(1);
    let first = header.map_or(0, |h| h + 1);
    if let Some(h) = header {
        if last_row.is_none_or(|r| r < h) || (0..last_col).all(|c| sheet.get_raw(h, c).trim().is_empty()) {
            return Err(format!("row {} of {} is empty; set the header to the row that holds the column names", src.header_row, sheet.name));
        }
    }
    let count = last_row.map_or(0, |r| (r + 1).saturating_sub(first));
    Ok(frame_from_sheet(sheet, header, count, last_col))
}

fn typed_cell_text(sheet: &Sheet, row: usize, col: usize) -> (String, Option<CellKind>) {
    use visigrid_engine::cell::ValueRef;
    let Some(cell) = sheet.get_cell_opt(row, col) else { return (String::new(), None) };
    match cell.value() {
        ValueRef::Empty => (String::new(), None),
        ValueRef::Text(t) => (t.to_string(), Some(CellKind::Text)),
        ValueRef::Formula { source, .. } => (source.to_string(), Some(CellKind::Text)),
        ValueRef::Number(n) => match date_kind(&sheet.get_format(row, col).number_format) {
            Some(CellKind::Date) => (serial_to_iso(n, false), Some(CellKind::Date)),
            Some(CellKind::DateTime) => (serial_to_iso(n, true), Some(CellKind::DateTime)),
            Some(CellKind::Time) => (fraction_to_time(n), Some(CellKind::Time)),
            _ => (number_text(n), Some(CellKind::Number)),
        },
    }
}

/// Whether a number format shows a date, a date-time or a time. Excel keeps
/// most date formats as format codes (`yyyy-mm-dd`, `m/d/yyyy h:mm`,
/// `h:mm`): read them as Excel does, ignoring quoted text, escapes and
/// [colour] sections. `m` alone is month or minute; it decides nothing.
fn date_kind(format: &NumberFormat) -> Option<CellKind> {
    let code = match format {
        NumberFormat::Date { .. } => return Some(CellKind::Date),
        NumberFormat::DateTime => return Some(CellKind::DateTime),
        NumberFormat::Time => return Some(CellKind::Time),
        NumberFormat::Custom(code) => code.to_lowercase(),
        _ => return None,
    };
    // The first section is the one positive numbers use
    let mut plain = String::new();
    let (mut quoted, mut bracket, mut escaped) = (false, false, false);
    for c in code.chars() {
        match c {
            _ if escaped => escaped = false,
            '\\' => escaped = true,
            '"' => quoted = !quoted,
            _ if quoted => {}
            '[' => bracket = true,
            // [h], [mm], [ss]: elapsed time
            ']' => bracket = false,
            ';' if !bracket => break,
            _ if bracket => {
                if "hms".contains(c) {
                    plain.push(c);
                }
            }
            _ => plain.push(c),
        }
    }
    let date = plain.contains('y') || plain.contains('d');
    let time = plain.contains('h') || plain.contains('s') || plain.contains("am/pm") || plain.contains("a/p");
    match (date, time) {
        (true, true) => Some(CellKind::DateTime),
        (true, false) => Some(CellKind::Date),
        (false, true) => Some(CellKind::Time),
        _ => None,
    }
}

/// 1204 rather than 1204.0; the shortest round-trip form otherwise.
fn number_text(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 9_007_199_254_740_992.0 {
        format!("{}", n as i64)
    } else {
        format!("{n:?}")
    }
}

fn serial_epoch() -> chrono::NaiveDateTime {
    chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0, 0, 0).unwrap()
}

/// A date serial as `2026-09-01`, or with `time` as `2026-09-01 14:02:00`
/// (with microseconds when there are any).
fn serial_to_iso(serial: f64, time: bool) -> String {
    let micros = (serial * 86_400_000_000.0).round() as i64;
    let at = serial_epoch() + chrono::Duration::microseconds(micros);
    if !time {
        at.format("%Y-%m-%d").to_string()
    } else if micros % 1_000_000 == 0 {
        at.format("%Y-%m-%d %H:%M:%S").to_string()
    } else {
        at.format("%Y-%m-%d %H:%M:%S%.6f").to_string()
    }
}

fn fraction_to_time(fraction: f64) -> String {
    let micros = (fraction * 86_400_000_000.0).round() as i64;
    let t = chrono::NaiveTime::MIN + chrono::Duration::microseconds(micros);
    if micros % 1_000_000 == 0 {
        t.format("%H:%M:%S").to_string()
    } else {
        t.format("%H:%M:%S%.6f").to_string()
    }
}

fn parse_iso_datetime(s: &str) -> Option<f64> {
    let t = s.trim();
    let at = chrono::NaiveDateTime::parse_from_str(t, "%Y-%m-%d %H:%M:%S%.f")
        .or_else(|_| chrono::NaiveDateTime::parse_from_str(t, "%Y-%m-%dT%H:%M:%S%.f"))
        .ok()?;
    let micros = (at - serial_epoch()).num_microseconds()?;
    Some(micros as f64 / 86_400_000_000.0)
}

fn parse_iso_time(s: &str) -> Option<f64> {
    let t = chrono::NaiveTime::parse_from_str(s.trim(), "%H:%M:%S%.f").ok()?;
    let micros = (t - chrono::NaiveTime::MIN).num_microseconds()?;
    Some(micros as f64 / 86_400_000_000.0)
}

fn read_csv(src: &CsvSource, snapshot: &Snapshot) -> Result<Frame, String> {
    let encoding = match &src.encoding {
        Some(e) => Some(Encoding::parse(e).ok_or_else(|| format!("unknown encoding {e:?}"))?),
        None => None,
    };
    let (content, _) = decode(&snapshot.bytes, encoding);

    // header_row counts physical lines, as a text editor shows them. Cut the
    // lines above it off before both detecting the delimiter and parsing, so
    // blank or oddly delimited title lines cannot shift which line is the
    // header (the CSV reader skips blank lines; counting its records would).
    let skip = src.header_row.saturating_sub(1);
    let mut offset = 0usize;
    for _ in 0..skip {
        match content[offset..].find('\n') {
            Some(i) => offset += i + 1,
            None => {
                return Err(format!(
                    "the source has fewer than {} lines; the header is set to line {}",
                    src.header_row, src.header_row
                ))
            }
        }
    }
    let body = &content[offset..];
    if src.header_row > 0 && body.lines().next().map_or(true, |l| l.trim().is_empty()) {
        return Err(format!(
            "line {} is empty; set the header to the line that holds the column names",
            src.header_row
        ));
    }

    let delimiter = match src.delimiter.as_deref() {
        None => {
            let sample: Vec<&str> = body.lines().take(20).collect();
            sniff_delimiter(&sample.join("\n"))
        }
        Some("tab") | Some("\t") => b'\t',
        Some(d) if d.len() == 1 => d.as_bytes()[0],
        Some(d) => return Err(format!("delimiter {d:?} must be one character or \"tab\"")),
    };
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(body.as_bytes());

    let mut names: Vec<String> = Vec::new();
    let mut rows = Vec::new();
    let mut lines = Vec::new();
    let mut record = csv::StringRecord::new();
    let mut n = 0usize;
    while reader.read_record(&mut record).map_err(|e| format!("cannot read the source: {e}"))? {
        n += 1;
        // Line in the whole file, counting the skipped title lines
        let line = record.position().map_or(n, |p| p.line() as usize) + skip;
        if src.header_row > 0 && n == 1 {
            names = record.iter().map(|f| f.trim().to_string()).collect();
            continue;
        }
        rows.push(record.iter().map(str::to_string).collect::<Vec<_>>());
        lines.push(line);
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
        columns: names.into_iter().map(|name| OutColumn { name, rule: ColumnRule::Auto, kind: ValueKind::Plain }).collect(),
        rows,
        lines,
        files: Vec::new(),
        file_names: Vec::new(),
        decimal_comma: src.decimal_comma,
    })
}

/// Every name the step uses must reach exactly one column, and no column
/// twice.
fn check_bindings(step: &Step, frame: &Frame) -> Result<(), String> {
    // A Group may total one column several ways (sum and average): only its
    // `by` columns must be distinct; every name must still be unambiguous
    let names: Vec<&str> = match step {
        Step::Group { by, totals, .. } => {
            for t in totals.iter().filter(|t| !t.column.is_empty()) {
                if let Resolve::Many(n) = frame.resolve(&t.column) {
                    return Err(format!("{n} columns are named {}; rename them in the source so each name is unique", t.column));
                }
            }
            by.iter().map(String::as_str).collect()
        }
        _ => step.named_columns(),
    };
    let mut seen: Vec<(usize, &str)> = Vec::new();
    for name in names {
        match frame.resolve(name) {
            Resolve::Many(n) => {
                return Err(format!("{n} columns are named {name}; rename them in the source so each name is unique"));
            }
            Resolve::One(i) => {
                if seen.iter().any(|(j, _)| *j == i) {
                    // Names only reach the same column through case (a, A)
                    return Err(format!("column {name} is listed twice"));
                }
                seen.push((i, name));
            }
            Resolve::None => {} // handled by the missing-column policy
        }
    }
    Ok(())
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
                *row = idx.iter().map(|&i| row[i].clone()).collect();
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
            // Resolve every source column before renaming any, so mappings
            // such as A→B, B→A do not interfere
            let targets: Vec<(usize, String)> =
                columns.iter().filter_map(|(from, to)| Some((frame.find(from)?, to.clone()))).collect();
            for (i, to) in targets {
                frame.columns[i].name = to;
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
                // A declared type replaces what the source said
                frame.columns[i].kind = ValueKind::Plain;
                let origins: Vec<String> = if frame.files.is_empty() {
                    Vec::new()
                } else {
                    (0..frame.rows.len()).map(|r| frame.file_of(r)).collect()
                };
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
                            step: 0, // set by run()
                            file: origins.get(r).cloned().unwrap_or_default(),
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
            let (rule, dc) = (frame.columns[i].rule, frame.decimal_comma);
            if let Err(reason) = check_filter_value(*op, value, rule) {
                report.ok = false;
                report.failures.push(format!("{}: {reason}", step.describe()));
                return None;
            }
            let keep: Vec<bool> = frame.rows.iter().map(|row| matches(&row[i], *op, value, rule, dc)).collect();
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
            // Compare the fields as a tuple; joining them with a separator
            // would let different rows collide
            let keep: Vec<bool> = {
                let mut seen: HashSet<Vec<&str>> = HashSet::new();
                frame.rows.iter().map(|row| seen.insert(idx.iter().map(|&i| row[i].as_str()).collect())).collect()
            };
            let removed = keep.iter().filter(|k| !**k).count();
            frame.keep_rows(&keep);
            Some(format!("removed {removed} duplicate row{}", plural(removed)))
        }
        Step::Group { by, totals, .. } => group(by, totals, frame, report),
        Step::Unpivot { keep, names_to, values_to, drop_empty, .. } => {
            unpivot(keep, names_to, values_to, *drop_empty, frame, report)
        }
    }
}

/// The column appended files get, naming the file each row came from.
pub const SOURCE_FILE_COLUMN: &str = "Source file";

/// Stack the frames of appended files. Columns line up by name (as steps
/// resolve them: exact, then ignoring case), in the order they first
/// appear; a file without a column has empty cells there, and says so in
/// `warnings`. A column whose type differs between files becomes untyped.
fn append_frames(parts: Vec<(String, Frame)>, warnings: &mut Vec<String>) -> Frame {
    let mut columns: Vec<OutColumn> = Vec::new();
    let position = |columns: &[OutColumn], name: &str| {
        columns.iter().position(|c| c.name == name).or_else(|| columns.iter().position(|c| c.name.eq_ignore_ascii_case(name)))
    };
    for (_, frame) in &parts {
        for c in &frame.columns {
            match position(&columns, &c.name) {
                Some(i) if columns[i].rule != c.rule || columns[i].kind != c.kind => {
                    columns[i].rule = ColumnRule::Auto;
                    columns[i].kind = ValueKind::Plain;
                }
                Some(_) => {}
                None => columns.push(c.clone()),
            }
        }
    }
    let file_names: Vec<String> = parts.iter().map(|(n, _)| n.clone()).collect();
    let decimal_comma = parts.first().is_some_and(|(_, f)| f.decimal_comma);
    let mut rows = Vec::new();
    let mut lines = Vec::new();
    let mut files = Vec::new();
    for (f, (name, frame)) in parts.into_iter().enumerate() {
        let map: Vec<Option<usize>> = columns.iter().map(|c| position(&frame.columns, &c.name)).collect();
        let absent: Vec<&str> = columns.iter().zip(&map).filter(|(_, m)| m.is_none()).map(|(c, _)| c.name.as_str()).collect();
        if !absent.is_empty() && !frame.rows.is_empty() {
            warnings.push(format!("{name} has no column {}; its rows are empty there", absent.join(", ")));
        }
        for (row, line) in frame.rows.into_iter().zip(frame.lines) {
            let mut out: Vec<String> = map.iter().map(|m| m.map_or(String::new(), |i| row[i].clone())).collect();
            out.push(name.clone());
            rows.push(out);
            lines.push(line);
            files.push(f as u32);
        }
    }
    columns.push(OutColumn { name: SOURCE_FILE_COLUMN.into(), rule: ColumnRule::Text, kind: ValueKind::Plain });
    Frame { columns, rows, lines, files, file_names, decimal_comma }
}

/// Fail the run with a step-level reason. Returns None for `apply`.
fn fail(report: &mut RunReport, reason: String) -> Option<String> {
    report.ok = false;
    report.failures.push(reason);
    None
}

fn group(by: &[String], totals: &[Total], frame: &mut Frame, report: &mut RunReport) -> Option<String> {
    // Result names must be unique, and a total needs a column to total
    let mut names: Vec<String> = by.iter().filter_map(|b| frame.find(b)).map(|i| frame.columns[i].name.clone()).collect();
    for t in totals {
        if t.name.trim().is_empty() {
            return fail(report, "every total needs a name (as = \"…\")".into());
        }
        if t.func != TotalFn::CountRows && t.column.trim().is_empty() {
            return fail(report, format!("total {} needs a column to {}", t.name, t.func.label()));
        }
        if names.iter().any(|n| n.eq_ignore_ascii_case(&t.name)) {
            return fail(report, format!("two result columns are named {}", t.name));
        }
        names.push(t.name.clone());
    }
    let by_idx: Vec<usize> = by.iter().filter_map(|b| frame.find(b)).collect();
    let total_idx: Vec<Option<usize>> = totals.iter().map(|t| frame.find(&t.column)).collect();

    // Groups in the order they first occur
    let mut order: Vec<Vec<usize>> = Vec::new();
    let mut index: std::collections::HashMap<Vec<&str>, usize> = std::collections::HashMap::new();
    for (r, row) in frame.rows.iter().enumerate() {
        let key: Vec<&str> = by_idx.iter().map(|&i| row[i].as_str()).collect();
        let g = *index.entry(key).or_insert_with(|| {
            order.push(Vec::new());
            order.len() - 1
        });
        order[g].push(r);
    }
    // No rows at all: one empty group only when totalling everything
    if order.is_empty() && by_idx.is_empty() {
        order.push(Vec::new());
    }

    let dc = frame.decimal_comma;
    let mut out_rows: Vec<Vec<String>> = Vec::with_capacity(order.len());
    let mut errors = 0usize;
    for members in &order {
        let mut row: Vec<String> = by_idx.iter().map(|&i| members.first().map_or(String::new(), |&r| frame.rows[r][i].clone())).collect();
        for (t, col) in totals.iter().zip(&total_idx) {
            let values: Vec<(usize, &str)> = match col {
                Some(i) => members.iter().map(|&r| (r, frame.rows[r][*i].as_str())).collect(),
                None => members.iter().map(|&r| (r, "")).collect(),
            };
            let rule = col.map_or(ColumnRule::Auto, |i| frame.columns[i].rule);
            let filled = || values.iter().filter(|(_, v)| !v.trim().is_empty());
            let cell = match t.func {
                TotalFn::CountRows => members.len().to_string(),
                TotalFn::Count => filled().count().to_string(),
                TotalFn::Distinct => filled().map(|(_, v)| *v).collect::<HashSet<&str>>().len().to_string(),
                TotalFn::First => values.first().map_or(String::new(), |(_, v)| v.to_string()),
                TotalFn::Last => values.last().map_or(String::new(), |(_, v)| v.to_string()),
                TotalFn::Sum | TotalFn::Average => {
                    let mut sum = 0.0;
                    let mut n = 0usize;
                    for (r, v) in filled() {
                        match parse_number(v, dc) {
                            Some(x) => {
                                sum += x;
                                n += 1;
                            }
                            None => {
                                errors += 1;
                                report.error_count += 1;
                                if report.errors.len() < MAX_REPORTED_ERRORS {
                                    report.errors.push(CellError {
                                        step: 0,
                                        file: frame.file_of(*r),
                                        line: frame.lines[*r],
                                        column: frame.columns[col.unwrap()].name.clone(),
                                        value: v.to_string(),
                                        reason: format!("not a number, so it can't be part of a {}", t.func.label()),
                                    });
                                }
                            }
                        }
                    }
                    match (t.func, n) {
                        (TotalFn::Average, 0) => String::new(),
                        (TotalFn::Average, n) => number_text(round_total(sum / n as f64)),
                        _ => number_text(round_total(sum)),
                    }
                }
                TotalFn::Min | TotalFn::Max => {
                    // Compared as the column's type: numbers as numbers,
                    // dates as dates, everything else (ISO date-times
                    // included) as text
                    let key = |v: &str| -> Option<f64> {
                        match rule {
                            ColumnRule::Date(order) => parse_date(v, order),
                            ColumnRule::Number => parse_number(v, dc),
                            ColumnRule::Auto if frame.columns[col.unwrap()].kind == ValueKind::Number => parse_number(v, dc),
                            _ => None,
                        }
                    };
                    let mut best: Option<&str> = None;
                    for (_, v) in filled() {
                        let v: &str = v;
                        let better = match best {
                            None => true,
                            Some(b) => {
                                let ord = match (key(v), key(b)) {
                                    (Some(x), Some(y)) => x.partial_cmp(&y).unwrap_or(std::cmp::Ordering::Equal),
                                    _ => v.cmp(b),
                                };
                                if t.func == TotalFn::Min { ord.is_lt() } else { ord.is_gt() }
                            }
                        };
                        if better {
                            best = Some(v);
                        }
                    }
                    best.unwrap_or("").to_string()
                }
            };
            row.push(cell);
        }
        out_rows.push(row);
    }
    if errors > 0 {
        report.ok = false;
        report.failures.push(format!("{errors} value{} could not be totalled", plural(errors)));
    }

    let mut columns: Vec<OutColumn> = by_idx.iter().map(|&i| frame.columns[i].clone()).collect();
    for (t, col) in totals.iter().zip(&total_idx) {
        let (rule, kind) = match (t.func, col) {
            (TotalFn::Min | TotalFn::Max | TotalFn::First | TotalFn::Last, Some(i)) => (frame.columns[*i].rule, frame.columns[*i].kind),
            (TotalFn::Min | TotalFn::Max | TotalFn::First | TotalFn::Last, None) => (ColumnRule::Auto, ValueKind::Plain),
            _ => (ColumnRule::Number, ValueKind::Number),
        };
        columns.push(OutColumn { name: t.name.clone(), rule, kind });
    }
    let rows_in = frame.rows.len();
    frame.lines = order.iter().map(|m| m.first().map_or(0, |&r| frame.lines[r])).collect();
    if !frame.files.is_empty() {
        frame.files = order.iter().map(|m| m.first().map_or(0, |&r| frame.files[r])).collect();
    }
    frame.columns = columns;
    frame.rows = out_rows;
    Some(format!("{rows_in} row{} into {} group{}", plural(rows_in), frame.rows.len(), plural(frame.rows.len())))
}

/// Sums of decimals pick up binary noise (0.1 + 0.2): round to 9 places,
/// past anything a currency or quantity export carries.
fn round_total(x: f64) -> f64 {
    let r = (x * 1e9).round() / 1e9;
    if r == 0.0 { 0.0 } else { r }
}

fn unpivot(keep: &[String], names_to: &str, values_to: &str, drop_empty: bool, frame: &mut Frame, report: &mut RunReport) -> Option<String> {
    let keep_idx: Vec<usize> = keep.iter().filter_map(|k| frame.find(k)).collect();
    for name in [names_to, values_to] {
        if name.trim().is_empty() {
            return fail(report, "the unpivoted columns need names (names_to, values_to)".into());
        }
        if keep_idx.iter().any(|&i| frame.columns[i].name.eq_ignore_ascii_case(name)) {
            return fail(report, format!("a kept column is already named {name}"));
        }
    }
    if names_to.eq_ignore_ascii_case(values_to) {
        return fail(report, "names_to and values_to must differ".into());
    }
    let others: Vec<usize> = (0..frame.columns.len()).filter(|i| !keep_idx.contains(i)).collect();
    if others.is_empty() {
        return Some("nothing to unpivot: every column is kept".into());
    }
    // The values keep their type when every unpivoted column agrees
    let first = &frame.columns[others[0]];
    let same = others.iter().all(|&i| frame.columns[i].rule == first.rule && frame.columns[i].kind == first.kind);
    let (rule, kind) = if same { (first.rule, first.kind) } else { (ColumnRule::Auto, ValueKind::Plain) };

    let mut rows = Vec::new();
    let mut lines = Vec::new();
    let mut files = Vec::new();
    let mut dropped = 0usize;
    for (r, row) in frame.rows.iter().enumerate() {
        for &j in &others {
            if drop_empty && row[j].trim().is_empty() {
                dropped += 1;
                continue;
            }
            let mut out: Vec<String> = keep_idx.iter().map(|&i| row[i].clone()).collect();
            out.push(frame.columns[j].name.clone());
            out.push(row[j].clone());
            rows.push(out);
            lines.push(frame.lines[r]);
            if !frame.files.is_empty() {
                files.push(frame.files[r]);
            }
        }
    }
    let mut columns: Vec<OutColumn> = keep_idx.iter().map(|&i| frame.columns[i].clone()).collect();
    columns.push(OutColumn { name: names_to.to_string(), rule: ColumnRule::Text, kind: ValueKind::Plain });
    columns.push(OutColumn { name: values_to.to_string(), rule, kind });
    let n = others.len();
    frame.columns = columns;
    frame.rows = rows;
    frame.lines = lines;
    frame.files = files;
    Some(format!(
        "{n} column{} into rows{}",
        plural(n),
        if dropped > 0 { format!("; {dropped} empty value{} left out", plural(dropped)) } else { String::new() }
    ))
}

/// A number as recipes write them: a point for decimals, no thousands
/// separators, whatever the source's locale (`1.5`, `-1200`, `2e3`). One
/// format, so a recipe means the same thing everywhere.
fn recipe_number(v: &str) -> Option<f64> {
    let t = v.trim();
    if t.contains(',') {
        return None;
    }
    visigrid_engine::cell::parse_finite(t)
}

/// A filter value must be readable as the column's type. `1,50` in a numeric
/// comparison could mean 1.5 or 150, so it is refused rather than guessed.
fn check_filter_value(op: FilterOp, value: &str, rule: ColumnRule) -> Result<(), String> {
    let comparison = matches!(op, FilterOp::Eq | FilterOp::Ne | FilterOp::Lt | FilterOp::Le | FilterOp::Gt | FilterOp::Ge);
    if !comparison {
        return Ok(());
    }
    let v = value.trim();
    let looks_numeric = parse_number(v, false).is_some() || parse_number(v, true).is_some();
    let number_hint = || {
        format!("{v:?} is not a number as recipes write them; use a point for decimals and no thousands separators, like 1.5")
    };
    match rule {
        ColumnRule::Number if recipe_number(v).is_none() => Err(number_hint()),
        ColumnRule::Date(order) if parse_date(v, order).is_none() => {
            Err(format!("{v:?} is not a {} date", order.label()))
        }
        ColumnRule::Auto if recipe_number(v).is_none() && looks_numeric => Err(number_hint()),
        _ => Ok(()),
    }
}

/// Compare as the column is typed. Number columns compare numerically (the
/// cell read with the source's decimal mark, the filter value in the recipe
/// format); Date columns compare as dates; Text columns compare as text, so
/// 001 is not 1. Auto columns compare as numbers only when both sides are
/// numbers and neither is ID-like (001), else as text. Text comparisons
/// ignore case. A cell that is not a number (or date) never satisfies a
/// numeric comparison.
fn matches(cell: &str, op: FilterOp, value: &str, rule: ColumnRule, decimal_comma: bool) -> bool {
    let c = cell.trim();
    let v = value.trim();
    match op {
        FilterOp::Contains => return c.to_lowercase().contains(&v.to_lowercase()),
        FilterOp::NotContains => return !c.to_lowercase().contains(&v.to_lowercase()),
        FilterOp::StartsWith => return c.to_lowercase().starts_with(&v.to_lowercase()),
        FilterOp::Empty => return c.is_empty(),
        FilterOp::NotEmpty => return !c.is_empty(),
        _ => {}
    }
    let text = || Some(c.to_lowercase().cmp(&v.to_lowercase()));
    let ord = match rule {
        ColumnRule::Number => match (parse_number(c, decimal_comma), recipe_number(v)) {
            (Some(a), Some(b)) => a.partial_cmp(&b),
            _ => None,
        },
        ColumnRule::Date(order) => match (parse_date(c, order), parse_date(v, order)) {
            (Some(a), Some(b)) => a.partial_cmp(&b),
            _ => None,
        },
        ColumnRule::Text | ColumnRule::Skip => text(),
        ColumnRule::Auto => {
            let id_like = keep_as_text(c, false).is_some() || keep_as_text(v, false).is_some();
            match (parse_number(c, decimal_comma), recipe_number(v)) {
                (Some(a), Some(b)) if !id_like => a.partial_cmp(&b),
                _ => text(),
            }
        }
    };
    use std::cmp::Ordering::*;
    match op {
        FilterOp::Eq => ord == Some(Equal),
        FilterOp::Ne => ord != Some(Equal),
        FilterOp::Lt => ord == Some(Less),
        FilterOp::Le => matches!(ord, Some(Less | Equal)),
        FilterOp::Gt => ord == Some(Greater),
        FilterOp::Ge => matches!(ord, Some(Greater | Equal)),
        _ => unreachable!("handled above"),
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
                self.write_value(&mut sheet, r + 1, c, c, value);
            }
        }
        sheet.rows = (self.rows.len() + 1).max(1000);
        sheet.cols = self.columns.len().max(26);
        sheet
    }
}

impl RecipeOutput {
    /// No columns, no rows: what a run that never read its source produces.
    pub fn empty() -> RecipeOutput {
        RecipeOutput { columns: Vec::new(), rows: Vec::new(), decimal_comma: false }
    }

    /// Write one value of column `col` at (`row`, `at_col`), typed the way
    /// [`to_sheet`](Self::to_sheet) types it. An empty value writes nothing.
    pub(crate) fn write_value(&self, sheet: &mut Sheet, row: usize, at_col: usize, col: usize, value: &str) {
        if value.is_empty() {
            return;
        }
        let (dr, c) = (row, at_col);
        match self.columns[col].kind {
            ValueKind::Plain => {}
            // As the Parquet importer writes it, so a recipe changes nothing
            // about a value it doesn't touch
            ValueKind::Number => return sheet.set_value_deferred(dr, c, value),
            ValueKind::DateTime => {
                return match parse_iso_datetime(value) {
                    Some(serial) => {
                        sheet.set_value_deferred(dr, c, &interchange_number(serial));
                        sheet.set_number_format(dr, c, NumberFormat::DateTime);
                    }
                    None => sheet.set_text(dr, c, value),
                }
            }
            ValueKind::Time => {
                return match parse_iso_time(value) {
                    Some(fraction) => {
                        sheet.set_value_deferred(dr, c, &interchange_number(fraction));
                        sheet.set_number_format(dr, c, NumberFormat::Time);
                    }
                    None => sheet.set_text(dr, c, value),
                }
            }
        }
        match self.columns[col].rule {
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

impl RunReport {
    /// A run that could not read its source at all.
    pub fn unreadable(source: &Path, error: String) -> RunReport {
        RunReport {
            ok: false,
            source: source.display().to_string(),
            snapshot: String::new(),
            source_rows: 0,
            rows: 0,
            columns: 0,
            drift: Drift::default(),
            steps: Vec::new(),
            failures: vec![error],
            errors: Vec::new(),
            error_count: 0,
            warnings: Vec::new(),
        }
    }

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
                if s.failed { " FAILED" } else if s.skipped { " SKIPPED" } else { "" },
                s.note.as_ref().map(|n| format!(": {n}")).unwrap_or_default()
            ));
        }
        for f in &self.failures {
            out.push_str(&format!("  error: {f}\n"));
        }
        for w in &self.warnings {
            out.push_str(&format!("  note: {w}\n"));
        }
        for e in self.errors.iter().take(20) {
            let file = if e.file.is_empty() { String::new() } else { format!("{} ", e.file) };
            out.push_str(&format!("    {file}line {}, {}: {:?} {}\n", e.line, e.column, e.value, e.reason));
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

    // ---- regressions from the first review (2026-10-02) ----

    fn recipe_with(steps: &str, source_extra: &str) -> Recipe {
        Recipe::from_toml(&format!("version = 1\n[source]\nkind = \"csv\"\npath = \"x.csv\"\n{source_extra}\n{steps}")).unwrap()
    }

    #[test]
    fn filters_follow_the_decimal_comma_and_declared_types() {
        // 1,50 is 1.5 with a decimal comma: not greater than 10
        let r = recipe_with("[[step]]\nop = \"filter\"\ncolumn = \"Amount\"\nis = \">\"\nvalue = \"10\"\n", "delimiter = \";\"\ndecimal_comma = true");
        let res = run(&r, &snap("id;Amount\n1;1,50\n2;12,00\n"));
        assert!(res.report.ok, "{}", res.report.summary());
        assert_eq!(res.output.rows, vec![vec!["2", "12,00"]]);
        // A Text column compares as text: 001 is not 1
        let r = recipe_with("[[step]]\nop = \"types\"\ncolumns = { code = \"text\" }\n[[step]]\nop = \"filter\"\ncolumn = \"code\"\nis = \"=\"\nvalue = \"1\"\n", "");
        let res = run(&r, &snap("code\n001\n1\n"));
        assert_eq!(res.output.rows, vec![vec!["1"]]);
        // Auto keeps ID-like values as text when comparing too
        let r = recipe_with("[[step]]\nop = \"filter\"\ncolumn = \"code\"\nis = \"=\"\nvalue = \"1\"\n", "");
        assert_eq!(run(&r, &snap("code\n001\n1\n")).output.rows, vec![vec!["1"]]);
        // A Number column: text never satisfies a numeric comparison
        let r = recipe_with("[[step]]\nop = \"types\"\ncolumns = { n = \"number\" }\non_error = \"keep_text\"\n[[step]]\nop = \"filter\"\ncolumn = \"n\"\nis = \"<\"\nvalue = \"5\"\n", "");
        assert_eq!(run(&r, &snap("n\nabc\n3\n")).output.rows, vec![vec!["3"]]);
    }

    #[test]
    fn duplicate_source_headers_are_never_guessed() {
        let r = recipe_with("[[step]]\nop = \"select\"\ncolumns = [\"Amount\"]\n", "");
        let res = run(&r, &snap("Amount,Amount\n10,999\n"));
        assert!(!res.report.ok);
        assert!(res.report.failures[0].contains("2 columns are named Amount"), "{:?}", res.report.failures);
    }

    #[test]
    fn a_column_listed_twice_is_refused() {
        let r = recipe_with("[[step]]\nop = \"select\"\ncolumns = [\"a\", \"a\"]\n", "");
        let res = run(&r, &snap("a,b\nx,y\n"));
        assert!(!res.report.ok);
        assert!(res.report.failures[0].contains("listed twice"), "{:?}", res.report.failures);
        // Two spellings of one column are the same column
        let r = recipe_with("[[step]]\nop = \"select\"\ncolumns = [\"a\", \"A\"]\n", "");
        assert!(run(&r, &snap("a,b\nx,y\n")).report.failures[0].contains("listed twice"));
    }

    #[test]
    fn renames_resolve_together_so_swaps_work() {
        let r = recipe_with("[[step]]\nop = \"rename\"\ncolumns = { A = \"B\", B = \"A\" }\n", "");
        let res = run(&r, &snap("A,B\n1,2\n"));
        assert!(res.report.ok, "{}", res.report.summary());
        let names: Vec<&str> = res.output.columns.iter().map(|c| c.name.as_str()).collect();
        assert_eq!(names, ["B", "A"]);
        assert_eq!(res.output.rows, vec![vec!["1", "2"]]);
        // Renaming onto an existing name is a failure, not a silent merge
        let r = recipe_with("[[step]]\nop = \"rename\"\ncolumns = { A = \"B\" }\n", "");
        assert!(!run(&r, &snap("A,B\n1,2\n")).report.ok);
    }

    #[test]
    fn dedupe_never_merges_distinct_rows() {
        let r = recipe_with("[[step]]\nop = \"dedupe\"\n", "");
        let data = "a,b\n\"x\u{1f}y\",z\nx,\"y\u{1f}z\"\nx,\"y\u{1f}z\"\n";
        let res = run(&r, &snap(data));
        assert_eq!(res.output.rows.len(), 2, "the two distinct rows stay; the real duplicate goes");
        assert_eq!(res.report.steps[0].note.as_deref(), Some("removed 1 duplicate row"));
    }

    #[test]
    fn the_delimiter_is_detected_from_the_header_line() {
        let r = recipe_with("[[step]]\nop = \"select\"\ncolumns = [\"Amount\"]\n", "header_row = 2");
        let res = run(&r, &snap("Report, generated 2026-10-01, all regions\nRegion\tAmount\nNorth\t10\n"));
        assert!(res.report.ok, "{}", res.report.summary());
        assert_eq!(res.output.rows, vec![vec!["10"]]);
    }

    // ---- regressions from the second review (2026-10-02) ----

    #[test]
    fn filter_numbers_have_one_format_and_ambiguous_ones_are_refused() {
        // With a decimal comma, "1,50" could mean 1.5 or 150: refused, not guessed
        let r = recipe_with("[[step]]\nop = \"filter\"\ncolumn = \"Amount\"\nis = \">\"\nvalue = \"1,50\"\n", "delimiter = \";\"\ndecimal_comma = true");
        let res = run(&r, &snap("id;Amount\n1;1,50\n2;2,00\n"));
        assert!(!res.report.ok);
        assert!(res.report.failures[0].contains("use a point for decimals"), "{:?}", res.report.failures);
        // Written the recipe way, it works against decimal-comma cells
        let r = recipe_with("[[step]]\nop = \"filter\"\ncolumn = \"Amount\"\nis = \">\"\nvalue = \"1.5\"\n", "delimiter = \";\"\ndecimal_comma = true");
        let res = run(&r, &snap("id;Amount\n1;1,50\n2;2,00\n"));
        assert!(res.report.ok, "{}", res.report.summary());
        assert_eq!(res.output.rows, vec![vec!["2", "2,00"]]);
        // A Number column refuses a value that is not a recipe number at all
        let r = recipe_with("[[step]]\nop = \"types\"\ncolumns = { n = \"number\" }\n[[step]]\nop = \"filter\"\ncolumn = \"n\"\nis = \">\"\nvalue = \"lots\"\n", "");
        assert!(!run(&r, &snap("n\n1\n")).report.ok);
        // Text-like values in an Auto column are fine
        let r = recipe_with("[[step]]\nop = \"filter\"\ncolumn = \"name\"\nis = \"=\"\nvalue = \"Smith, John\"\n", "");
        let res = run(&r, &snap("name\n\"Smith, John\"\nDoe\n"));
        assert!(res.report.ok);
        assert_eq!(res.output.rows, vec![vec!["Smith, John"]]);
    }

    #[test]
    fn id_comparisons_are_symmetric() {
        let eq = |value: &str| {
            let r = recipe_with(&format!("[[step]]\nop = \"filter\"\ncolumn = \"code\"\nis = \"=\"\nvalue = \"{value}\"\n"), "");
            run(&r, &snap("code\n001\n1\n")).output.rows
        };
        assert_eq!(eq("1"), vec![vec!["1"]]);
        assert_eq!(eq("001"), vec![vec!["001"]]);
    }

    #[test]
    fn header_row_counts_physical_lines_including_blank_ones() {
        let src = "Monthly export\n\nid,amount\n1,10\n2,20\n";
        let r = recipe_with("[[step]]\nop = \"select\"\ncolumns = [\"id\", \"amount\"]\n", "header_row = 3");
        let res = run(&r, &snap(src));
        assert!(res.report.ok, "{}", res.report.summary());
        assert_eq!(res.output.rows, vec![vec!["1", "10"], vec!["2", "20"]]);
        // The same with an explicit delimiter
        let r = recipe_with("[[step]]\nop = \"select\"\ncolumns = [\"id\"]\n", "header_row = 3\ndelimiter = \",\"");
        assert!(run(&r, &snap(src)).report.ok);
        // Error lines count the skipped title lines too
        let r = recipe_with("[[step]]\nop = \"types\"\ncolumns = { amount = \"number\" }\n", "header_row = 3");
        let res = run(&r, &snap("Monthly export\n\nid,amount\n1,10\n2,x\n"));
        assert_eq!(res.report.errors[0].line, 5);
        // A header pointing at a blank line is an error, never the next row
        let r = recipe_with("", "header_row = 2");
        let res = run(&r, &snap(src));
        assert!(!res.report.ok);
        assert!(res.report.failures[0].contains("line 2 is empty"), "{:?}", res.report.failures);
        // Past the end of the file
        let r = recipe_with("", "header_row = 40");
        assert!(run(&r, &snap(src)).report.failures[0].contains("fewer than 40 lines"));
    }

    #[test]
    fn failed_steps_name_their_missing_columns_and_errors_their_step() {
        let text = r#"
version = 1
[source]
kind = "csv"
path = "x.csv"
columns = ["Order ID", "Amount"]
[[step]]
op = "rename"
columns = { "Order ID" = "order_id" }
[[step]]
op = "types"
columns = { Amount = "number" }
"#;
        let r = Recipe::from_toml(text).unwrap();
        let res = run(&r, &snap("Order Number,Amount\n1,x\n"));
        assert!(!res.report.ok);
        let s = &res.report.steps;
        assert!(s[0].failed && !s[0].skipped);
        assert_eq!(s[0].missing, ["Order ID"]);
        assert!(s[1].failed);
        assert_eq!(res.report.errors[0].step, 2);
        assert!(res.report.summary().contains("FAILED"));
    }

    #[test]
    fn rename_source_column_stops_where_the_column_is_renamed_away() {
        let text = r#"
version = 1
[source]
kind = "csv"
path = "x.csv"
columns = ["Order ID", "Amount"]
[[step]]
op = "trim"
columns = ["Order ID"]
[[step]]
op = "rename"
columns = { "Order ID" = "order_id" }
[[step]]
op = "filter"
column = "Order ID"
is = "not_empty"
"#;
        let mut r = Recipe::from_toml(text).unwrap();
        assert!(r.rename_source_column("order id", "Order Number"));
        assert_eq!(r.source.columns(), ["Order Number", "Amount"]);
        assert_eq!(r.steps[0], Step::Trim { columns: vec!["Order Number".into()], missing: Missing::Fail });
        assert!(matches!(&r.steps[1], Step::Rename { columns, .. } if columns.get("Order Number").map(String::as_str) == Some("order_id")));
        // After the rename, "Order ID" would be some other column: untouched
        assert!(matches!(&r.steps[2], Step::Filter { column, .. } if column == "Order ID"));
        assert!(!r.rename_source_column("Nope", "X"));
    }

    #[test]
    fn set_on_error_only_on_types_steps_and_save_round_trips() {
        let text = r#"
version = 1
[source]
kind = "csv"
path = "x.csv"
[[step]]
op = "trim"
[[step]]
op = "types"
columns = { Amount = "number" }
"#;
        let mut r = Recipe::from_toml(text).unwrap();
        assert!(r.set_on_error(1, OnError::KeepText).is_err());
        assert!(r.set_on_error(9, OnError::KeepText).is_err());
        r.set_on_error(2, OnError::KeepText).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("x.recipe.toml");
        r.save(&path).unwrap();
        assert_eq!(Recipe::load(&path).unwrap(), r);
        assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
    }

    #[test]
    fn guesses_the_header_below_title_lines_and_reports_the_source() {
        let s = snap("Acme export\nGenerated 2026-09-30\n\nID,Name,Amount\n1,a,2\n2,b,3\n3,c,4\n");
        assert_eq!(guess_header_row(&s), 4);
        assert_eq!(guess_header_row(&snap("a;b\n1;2\n")), 1);
        let src = CsvSource { path: "x.csv".into(), delimiter: None, encoding: None, header_row: 4, decimal_comma: false, columns: vec![], combine: false };
        let info = source_info(&src, &s);
        assert_eq!(info.delimiter, b',');
        assert_eq!(info.first_lines[0], "Acme export");
    }

    #[test]
    fn patterns_pick_the_newest_matching_export() {
        assert!(wildcard_match("export-*-*.csv", "Export-2026-10.CSV"));
        assert!(!wildcard_match("export-*-*.csv", "export-2026.csv"));
        assert!(wildcard_match("a?c*", "abcdef"));
        assert_eq!(suggest_pattern("export-2026-09.csv").as_deref(), Some("export-*-*.csv"));
        assert_eq!(suggest_pattern("orders.csv"), None);

        let dir = tempfile::tempdir().unwrap();
        let write = |name: &str, secs: u64| {
            let p = dir.path().join(name);
            std::fs::write(&p, "a\n1\n").unwrap();
            let t = std::time::UNIX_EPOCH + std::time::Duration::from_secs(secs);
            std::fs::File::options().write(true).open(&p).unwrap().set_modified(t).unwrap();
        };
        write("export-2026-08.csv", 1_000);
        write("export-2026-10.csv", 3_000);
        write("export-2026-09.csv", 2_000);
        write("other-2026-11.csv", 9_000);
        let text = r#"
version = 1
[source]
kind = "csv"
path = "export-*-*.csv"
"#;
        let r = Recipe::from_toml(text).unwrap();
        assert!(r.source_is_pattern());
        assert_eq!(r.resolve_source(dir.path(), None).unwrap(), dir.path().join("export-2026-10.csv"));
        let over = dir.path().join("export-2026-08.csv");
        assert_eq!(r.resolve_source(dir.path(), Some(&over)).unwrap(), over);
        let empty = tempfile::tempdir().unwrap();
        assert!(r.resolve_source(empty.path(), None).unwrap_err().contains("no file in"));
    }

    #[test]
    fn sources_must_be_local_regular_files_within_the_cap() {
        let dir = tempfile::tempdir().unwrap();
        // A directory
        assert!(Snapshot::read(dir.path()).unwrap_err().contains("not a regular file"));
        // Over the cap
        let big = dir.path().join("big.csv");
        std::fs::write(&big, vec![b'a'; 2048]).unwrap();
        assert!(read_regular_file(&big, 1024, "source").unwrap_err().contains("up to"));
        assert_eq!(read_regular_file(&big, 4096, "source").unwrap().len(), 2048);
        // Network paths are refused without touching the filesystem
        for p in [r"\\server\share\x.csv", "//server/share/x.csv", r"\\?\UNC\server\x.csv"] {
            assert!(is_network_path(Path::new(p)), "{p}");
            assert!(Snapshot::read(Path::new(p)).unwrap_err().contains("network share"));
        }
        assert!(!is_network_path(Path::new("/data/x.csv")));
        assert!(!is_network_path(Path::new(r"C:\data\x.csv")));
        let text = "version = 1\n[source]\nkind = \"csv\"\npath = \"//evil/share/x.csv\"\n";
        let r = Recipe::from_toml(text).unwrap();
        assert!(r.resolve_source(dir.path(), None).unwrap_err().contains("network share"));
    }

    #[cfg(unix)]
    #[test]
    fn a_pipe_or_device_is_refused_without_blocking() {
        let dir = tempfile::tempdir().unwrap();
        let fifo = dir.path().join("pipe.csv");
        assert!(std::process::Command::new("mkfifo").arg(&fifo).status().unwrap().success());
        assert!(Snapshot::read(&fifo).unwrap_err().contains("not a regular file"));
        assert!(Recipe::load(&fifo).unwrap_err().contains("not a regular file"));
        assert!(Snapshot::read(Path::new("/dev/zero")).unwrap_err().contains("not a regular file"));
    }

    /// A sheet with text IDs, numbers, dates and date-times, written out by
    /// the real Parquet and DuckDB exporters.
    fn typed_sheet() -> Sheet {
        use visigrid_engine::cell::NumberFormat;
        let mut sheet = Sheet::new(SheetId(1), 100, 10);
        for (c, name) in ["ID", "Amount", "Day", "At"].iter().enumerate() {
            sheet.set_text(0, c, name);
        }
        let rows = [("007", "10.5", 46266.0, 46266.5), ("008", "0", 46267.0, 46267.25), ("009", "1204", 46268.0, 46268.75)];
        for (i, (id, amount, day, at)) in rows.iter().enumerate() {
            let r = i + 1;
            sheet.set_text(r, 0, id);
            sheet.set_value_deferred(r, 1, amount);
            sheet.set_value_deferred(r, 2, &day.to_string());
            sheet.set_number_format(r, 2, NumberFormat::Date { style: DateStyle::Iso });
            sheet.set_value_deferred(r, 3, &at.to_string());
            sheet.set_number_format(r, 3, NumberFormat::DateTime);
        }
        sheet
    }

    fn typed_plan(sheet: &Sheet) -> crate::parquet_export::Plan<'_> {
        let columns = ["ID", "Amount", "Day", "At"]
            .iter()
            .enumerate()
            .map(|(index, name)| crate::parquet_export::Column { index, name: name.to_string(), as_text: false })
            .collect();
        crate::parquet_export::analyze(sheet, vec![1, 2, 3], columns).unwrap()
    }

    fn check_typed(res: &RunResult) {
        assert!(res.report.ok, "{}", res.report.summary());
        let out = &res.output;
        // Amount > 0 dropped the zero row
        assert_eq!(out.rows, vec![
            vec!["007", "10.5", "2026-09-01", "2026-09-01 12:00:00"],
            vec!["009", "1204", "2026-09-03", "2026-09-03 18:00:00"],
        ]);
        assert_eq!(out.columns[0].rule, ColumnRule::Text);
        assert_eq!((out.columns[1].rule, out.columns[1].kind), (ColumnRule::Number, ValueKind::Number));
        assert_eq!(out.columns[2].rule, ColumnRule::Date(DateOrder::Ymd));
        assert_eq!(out.columns[3].kind, ValueKind::DateTime);
        let sheet = out.to_sheet();
        assert_eq!(sheet.get_raw(1, 0), "007");
        assert_eq!(sheet.get_display(2, 1), "1204");
        assert!(matches!(sheet.get_format(1, 3).number_format, visigrid_engine::cell::NumberFormat::DateTime));
        assert_eq!(sheet.get_formatted_display(1, 2), "2026-09-01");
    }

    const TYPED_STEPS: &str = "[[step]]\nop = \"filter\"\ncolumn = \"Amount\"\nis = \">\"\nvalue = \"0\"\n";

    #[test]
    fn a_parquet_source_keeps_its_column_types() {
        let sheet = typed_sheet();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("orders.parquet");
        typed_plan(&sheet).write_path(&path).unwrap();
        let text = format!("version = 1\n[source]\nkind = \"parquet\"\npath = \"orders.parquet\"\n{TYPED_STEPS}");
        let r = Recipe::from_toml(&text).unwrap();
        assert_eq!(r.source.label(), "Parquet");
        let src = r.resolve_source(dir.path(), None).unwrap();
        check_typed(&run(&r, &Snapshot::read(&src).unwrap()));
        // A declared type replaces the source's
        let typed = format!("{text}[[step]]\nop = \"types\"\ncolumns = {{ At = \"text\" }}\n");
        let res = run(&Recipe::from_toml(&typed).unwrap(), &Snapshot::read(&src).unwrap());
        assert_eq!(res.output.columns[3].kind, ValueKind::Plain);
        assert_eq!(res.output.to_sheet().get_raw(1, 3), "2026-09-01 12:00:00");
    }

    #[test]
    fn a_duckdb_source_reads_one_table() {
        let sheet = typed_sheet();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("orders.duckdb");
        crate::duckdb::export(&typed_plan(&sheet), &path).unwrap();
        let text = format!("version = 1\n[source]\nkind = \"duckdb\"\npath = \"orders.duckdb\"\ntable = \"main.data\"\n{TYPED_STEPS}");
        let r = Recipe::from_toml(&text).unwrap();
        let src = r.resolve_source(dir.path(), None).unwrap();
        check_typed(&run(&r, &Snapshot::read(&src).unwrap()));
        // An unknown table fails the run with the tables there are
        let wrong = text.replace("main.data", "nope");
        let res = run(&Recipe::from_toml(&wrong).unwrap(), &Snapshot::read(&src).unwrap());
        assert!(!res.report.ok);
        assert!(res.report.failures[0].contains("main.data"), "{:?}", res.report.failures);
    }

    #[test]
    fn patterns_skip_downloads_in_progress_and_files_still_changing() {
        let dir = tempfile::tempdir().unwrap();
        let write = |name: &str, secs_ago: u64| {
            let p = dir.path().join(name);
            std::fs::write(&p, "a\n1\n").unwrap();
            let t = std::time::SystemTime::now() - std::time::Duration::from_secs(secs_ago);
            std::fs::File::options().write(true).open(&p).unwrap().set_modified(t).unwrap();
        };
        write("export-09.csv", 3600);
        write("export-10.csv.crdownload", 10);
        write("export-10.csv.part", 10);
        write(".export-10.csv", 10);
        let text = "version = 1\n[source]\nkind = \"csv\"\npath = \"export-*.csv*\"\n";
        let r = Recipe::from_toml(text).unwrap();
        assert_eq!(r.resolve_source(dir.path(), None).unwrap(), dir.path().join("export-09.csv"));
        // The finished file lands but is still changing: wait, don't read half
        write("export-10.csv", 0);
        assert!(r.resolve_source(dir.path(), None).unwrap_err().contains("still being written"));
        write("export-10.csv", 5);
        assert_eq!(r.resolve_source(dir.path(), None).unwrap(), dir.path().join("export-10.csv"));
    }

    #[test]
    fn an_excel_source_reads_one_sheet_below_its_title_rows() {
        use visigrid_engine::cell::NumberFormat;
        let mut wb = visigrid_engine::workbook::Workbook::new();
        let s = wb.sheet_mut(0).unwrap();
        s.set_text(0, 0, "Acme order export");
        s.set_text(1, 0, "Generated 2026-09-30");
        for (c, name) in ["ID", "Amount", "Day", "Double"].iter().enumerate() {
            s.set_text(2, c, name);
        }
        for (i, (id, amount, day)) in [("007", "10.5", 46266.0), ("008", "0", 46267.0), ("009", "1204", 46268.0)].iter().enumerate() {
            let r = 3 + i;
            s.set_text(r, 0, id);
            s.set_value(r, 1, amount);
            s.set_value(r, 2, &day.to_string());
            s.set_number_format(r, 2, NumberFormat::Date { style: DateStyle::Iso });
            s.set_value(r, 3, &format!("=B{}*2", r + 1));
        }
        let other = wb.add_sheet_named("Notes").unwrap();
        wb.sheet_mut(other).unwrap().set_text(0, 0, "unrelated");
        wb.rebuild_dep_graph();
        wb.recompute_full_ordered();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("orders.xlsx");
        crate::xlsx::export(&wb, &path, None).unwrap();
        // Saved results that differ from what the formulas compute: a recipe
        // reads what Excel saved and never recalculates
        let saved = dir.path().join("saved.xlsx");
        {
            use std::io::{Read, Write};
            let mut input = zip::ZipArchive::new(std::fs::File::open(&path).unwrap()).unwrap();
            let mut output = zip::ZipWriter::new(std::fs::File::create(&saved).unwrap());
            for i in 0..input.len() {
                let mut entry = input.by_index(i).unwrap();
                let mut bytes = Vec::new();
                entry.read_to_end(&mut bytes).unwrap();
                let name = entry.name().to_string();
                if name == "xl/worksheets/sheet1.xml" {
                    // Whatever result the exporter saved (0 before #100, the
                    // computed value after), replace it with one no
                    // recalculation would produce
                    let xml = String::from_utf8(bytes).unwrap();
                    let mut out = String::new();
                    let mut rest = xml.as_str();
                    let mut n = 0;
                    while let Some(at) = rest.find("</f><v>") {
                        let start = at + "</f><v>".len();
                        let end = start + rest[start..].find("</v>").unwrap();
                        out.push_str(&rest[..start]);
                        out.push_str("99");
                        rest = &rest[end..];
                        n += 1;
                    }
                    out.push_str(rest);
                    assert_eq!(n, 3, "{xml}");
                    bytes = out.into_bytes();
                }
                output.start_file(name, zip::write::SimpleFileOptions::default()).unwrap();
                output.write_all(&bytes).unwrap();
            }
            output.finish().unwrap();
        }
        std::fs::rename(&saved, &path).unwrap();

        assert_eq!(xlsx_sheet_names(&path).unwrap()[0], wb.sheet(0).unwrap().name);
        let snap = Snapshot::read(&path).unwrap();
        assert_eq!(guess_xlsx_header_row(&snap, ""), 3);
        // Rows with empty cells below the header don't hide it
        let mut sparse = visigrid_engine::workbook::Workbook::new();
        let sh = sparse.sheet_mut(0).unwrap();
        sh.set_text(0, 0, "Report");
        for (c, v) in ["A", "B", "C"].iter().enumerate() {
            sh.set_text(1, c, v);
        }
        sh.set_text(2, 0, "1");
        sh.set_text(2, 2, "x");
        sh.set_text(3, 0, "2");
        let sparse_path = dir.path().join("sparse.xlsx");
        crate::xlsx::export(&sparse, &sparse_path, None).unwrap();
        assert_eq!(guess_xlsx_header_row(&Snapshot::read(&sparse_path).unwrap(), ""), 2);

        let text = "version = 1\n[source]\nkind = \"xlsx\"\npath = \"orders.xlsx\"\nheader_row = 3\n[[step]]\nop = \"filter\"\ncolumn = \"Amount\"\nis = \">\"\nvalue = \"0\"\n";
        let r = Recipe::from_toml(text).unwrap();
        assert_eq!(r.source.label(), "Excel");
        let res = run(&r, &snap);
        assert!(res.report.ok, "{}", res.report.summary());
        assert_eq!(res.output.columns.iter().map(|c| c.name.as_str()).collect::<Vec<_>>(), ["ID", "Amount", "Day", "Double"]);
        assert_eq!(res.output.rows[0], vec!["007", "10.5", "2026-09-01", "99"]);
        assert_eq!(res.output.rows.len(), 2);
        assert_eq!(res.output.columns[2].rule, ColumnRule::Date(DateOrder::Ymd));
        // Report lines are the sheet's row numbers: the first record is row 4
        let dates = text.replace("value = \"0\"", "value = \"0\"\n[[step]]\nop = \"types\"\ncolumns = { ID = \"date\" }");
        let failing = run(&Recipe::from_toml(&dates).unwrap(), &snap);
        assert_eq!(failing.report.errors[0].line, 4, "{}", failing.report.summary());

        // Another sheet by name, and an unknown one
        let notes = text.replace("header_row = 3\n", "sheet = \"Notes\"\nheader_row = 0\n").replace("[[step]]\nop = \"filter\"\ncolumn = \"Amount\"\nis = \">\"\nvalue = \"0\"\n", "");
        let res = run(&Recipe::from_toml(&notes).unwrap(), &snap);
        assert_eq!(res.output.rows, vec![vec!["unrelated"]]);
        let missing = run(&Recipe::from_toml(&notes.replace("Notes", "Nope")).unwrap(), &snap);
        assert!(missing.report.failures[0].contains("its sheets are"), "{:?}", missing.report.failures);
    }

    #[test]
    fn excel_date_format_codes_are_recognised() {
        let k = |code: &str| date_kind(&visigrid_engine::cell::NumberFormat::Custom(code.into()));
        assert!(matches!(k("yyyy-mm-dd"), Some(CellKind::Date)));
        assert!(matches!(k("d-mmm-yy"), Some(CellKind::Date)));
        assert!(matches!(k("m/d/yyyy h:mm"), Some(CellKind::DateTime)));
        assert!(matches!(k("h:mm AM/PM"), Some(CellKind::Time)));
        assert!(matches!(k("[h]:mm:ss"), Some(CellKind::Time)));
        assert!(k("#,##0.00").is_none());
        assert!(k("0.00\" days\"").is_none());
        assert!(k("[Red]0.00;[Blue]-0.00").is_none());
        assert!(k("mm").is_none());
    }

    #[test]
    fn group_by_with_totals_in_first_seen_order() {
        let text = r#"
version = 1
[source]
kind = "csv"
path = "x.csv"
[[step]]
op = "types"
columns = { Amount = "number", Day = "date" }
[[step]]
op = "group"
by = ["Customer"]
totals = [
  { fn = "sum", column = "Amount", as = "Total" },
  { fn = "count_rows", as = "Orders" },
  { fn = "average", column = "Amount", as = "Avg" },
  { fn = "max", column = "Day", as = "Last order" },
  { fn = "distinct", column = "Rep", as = "Reps" },
]
"#;
        let r = Recipe::from_toml(text).unwrap();
        let csv = "Customer,Amount,Day,Rep\nBeta,0.1,2026-09-02,KM\nAlpha,10,2026-09-01,JT\nBeta,0.2,2026-09-10,KM\nAlpha,5.5,2026-09-03,AL\n";
        let res = run(&r, &snap(csv));
        assert!(res.report.ok, "{}", res.report.summary());
        assert_eq!(res.output.rows, vec![
            vec!["Beta", "0.3", "2", "0.15", "2026-09-10", "1"],
            vec!["Alpha", "15.5", "2", "7.75", "2026-09-03", "2"],
        ]);
        assert_eq!(res.output.columns[1].rule, ColumnRule::Number);
        assert_eq!(res.output.columns[4].rule, ColumnRule::Date(DateOrder::Ymd));
        assert_eq!(res.report.steps[1].note.as_deref(), Some("4 rows into 2 groups"));
        assert!(Step::describe(&r.steps[1]).starts_with("Group by Customer: Total = sum of Amount"));
        // A value that isn't a number fails the run, located
        let bad = run(&r, &snap("Customer,Amount,Day,Rep\nA,x,2026-09-01,K\n"));
        assert!(!bad.report.ok);
        // (the types step catches it first; a Group on an untyped column too)
        let untyped = text.replace("columns = { Amount = \"number\", Day = \"date\" }", "columns = { Day = \"date\" }");
        let bad = run(&Recipe::from_toml(&untyped).unwrap(), &snap("Customer,Amount,Day,Rep\nA,x,2026-09-01,K\n"));
        assert!(!bad.report.ok);
        assert_eq!(bad.report.errors[0].line, 2);
        assert!(bad.report.errors[0].reason.contains("sum"), "{:?}", bad.report.errors);
    }

    #[test]
    fn group_without_by_totals_everything_and_names_are_checked() {
        let base = "version = 1\n[source]\nkind = \"csv\"\npath = \"x.csv\"\n[[step]]\nop = \"group\"\n";
        let r = Recipe::from_toml(&format!("{base}totals = [{{ fn = \"count_rows\", as = \"Rows\" }}]\n")).unwrap();
        assert_eq!(run(&r, &snap("a\n1\n2\n")).output.rows, vec![vec!["2"]]);
        let dup = Recipe::from_toml(&format!("{base}by = [\"a\"]\ntotals = [{{ fn = \"count_rows\", as = \"A\" }}]\n")).unwrap();
        assert!(run(&dup, &snap("a\n1\n")).report.failures[0].contains("named A"));
        let nocol = Recipe::from_toml(&format!("{base}totals = [{{ fn = \"sum\", as = \"S\" }}]\n")).unwrap();
        assert!(run(&nocol, &snap("a\n1\n")).report.failures[0].contains("needs a column"));
    }

    #[test]
    fn unpivot_other_columns_picks_up_new_columns() {
        let text = r#"
version = 1
[source]
kind = "csv"
path = "x.csv"
[[step]]
op = "unpivot"
keep = ["Region"]
names_to = "Month"
values_to = "Sales"
"#;
        let r = Recipe::from_toml(text).unwrap();
        let res = run(&r, &snap("Region,Jan,Feb\nWest,1,\nEast,3,4\n"));
        assert!(res.report.ok, "{}", res.report.summary());
        assert_eq!(res.output.columns.iter().map(|c| c.name.as_str()).collect::<Vec<_>>(), ["Region", "Month", "Sales"]);
        assert_eq!(res.output.rows, vec![vec!["West", "Jan", "1"], vec!["East", "Jan", "3"], vec!["East", "Feb", "4"]]);
        assert!(res.report.steps[0].note.as_deref().unwrap().contains("1 empty value left out"));
        // Next month's file has a new column: unpivoted with no recipe change
        let res = run(&r, &snap("Region,Jan,Feb,Mar\nWest,1,2,3\n"));
        assert_eq!(res.output.rows.len(), 3);
        assert_eq!(res.output.rows[2], vec!["West", "Mar", "3"]);
        // A kept column can't share the new columns' names
        let clash = Recipe::from_toml(&text.replace("names_to = \"Month\"", "names_to = \"Region\"")).unwrap();
        assert!(!run(&clash, &snap("Region,Jan\nW,1\n")).report.ok);
        // Round trip keeps defaults out of the file
        let plain = Recipe::from_toml("version = 1\n[source]\nkind = \"csv\"\npath = \"x.csv\"\n[[step]]\nop = \"unpivot\"\n").unwrap();
        assert!(!plain.to_toml().contains("names_to"));
    }

    #[test]
    fn append_folder_stacks_every_matching_file_with_its_name() {
        let dir = tempfile::tempdir().unwrap();
        let write = |name: &str, body: &str| {
            let p = dir.path().join(name);
            std::fs::write(&p, body).unwrap();
            let t = std::time::SystemTime::now() - std::time::Duration::from_secs(60);
            std::fs::File::options().write(true).open(&p).unwrap().set_modified(t).unwrap();
        };
        write("sales-02.csv", "Region,Amount\nEast,2\n");
        write("sales-01.csv", "Region,Amount,Rep\nWest,1,KM\n");
        write("sales-03.csv.part", "Region,Amount\nNorth,9\n");
        write("other.csv", "Region,Amount\nSouth,5\n");
        let text = "version = 1\n[source]\nkind = \"csv\"\npath = \"sales-*.csv\"\ncombine = true\n[[step]]\nop = \"types\"\ncolumns = { Amount = \"number\" }\n";
        let r = Recipe::from_toml(text).unwrap();
        assert!(r.source.combine());
        let files = r.resolve_sources(dir.path(), None).unwrap();
        assert_eq!(files, vec![dir.path().join("sales-01.csv"), dir.path().join("sales-02.csv")]);
        let snapshot = r.read_snapshot(dir.path(), None).unwrap();
        let res = run(&r, &snapshot);
        assert!(res.report.ok, "{}", res.report.summary());
        assert_eq!(res.output.columns.iter().map(|c| c.name.as_str()).collect::<Vec<_>>(), ["Region", "Amount", "Rep", "Source file"]);
        assert_eq!(res.output.rows, vec![vec!["West", "1", "KM", "sales-01.csv"], vec!["East", "2", "", "sales-02.csv"]]);
        assert!(res.report.warnings[0].contains("sales-02.csv has no column Rep"), "{:?}", res.report.warnings);
        assert!(res.report.source.starts_with("2 files in"));
        // Errors say which file
        write("sales-04.csv", "Region,Amount\nSouth,lots\n");
        let res = run(&r, &r.read_snapshot(dir.path(), None).unwrap());
        assert!(!res.report.ok);
        assert_eq!((res.report.errors[0].file.as_str(), res.report.errors[0].line), ("sales-04.csv", 2));
        assert!(res.report.summary().contains("sales-04.csv line 2"));
        // Without combine the same pattern reads only the newest file
        let one = Recipe::from_toml(&text.replace("combine = true\n", "")).unwrap();
        assert_eq!(one.resolve_sources(dir.path(), None).unwrap().len(), 1);
    }

    #[test]
    fn fixes_keep_the_recipe_files_comments_and_layout() {
        let text = r#"# Monthly orders, from the vendor portal
version = 1

[source]
kind = "csv"
path = "export-*-*.csv"   # newest export
header_row = 3
columns = ["Order ID", "Customer", "Amount"]

# Keep only what accounting needs
[[step]]
op = "select"
columns = ["Order ID", "Customer", "Amount"]

[[step]]
op = "rename"
columns = { "Order ID" = "order_id" }   # our key

[[step]]
op = "types"
columns = { Amount = "number" }
"#;
        let mut r = Recipe::from_toml(text).unwrap();
        assert!(r.rename_source_column("Order ID", "Order Number"));
        r.set_on_error(3, OnError::KeepText).unwrap();
        let out = r.to_toml_preserving(text);
        assert_eq!(Recipe::from_toml(&out).unwrap(), r);
        for kept in ["# Monthly orders, from the vendor portal", "# newest export", "# Keep only what accounting needs", "# our key"] {
            assert!(out.contains(kept), "lost {kept:?}:\n{out}");
        }
        // Inline tables stay inline, and only the changed values changed
        assert!(out.contains(r#"columns = { "Order Number" = "order_id" }"#), "{out}");
        assert!(!out.contains(" ,"), "{out}");
        assert!(out.contains("on_error = \"keep_text\""), "{out}");
        assert!(out.contains(r#"columns = ["Order Number", "Customer", "Amount"]"#), "{out}");

        // A key renamed inside a multi-key inline table: tidy separators
        let two = "version = 1\n[source]\nkind = \"csv\"\npath = \"x.csv\"\n[[step]]\nop = \"types\"\ncolumns = { ID = \"text\", Amount = \"number\" }   # keep zeros\n";
        let mut r2 = Recipe::from_toml(two).unwrap();
        r2.rename_source_column("ID", "Order ID");
        let out2 = r2.to_toml_preserving(two);
        assert!(out2.contains("# keep zeros") && !out2.contains(" ,"), "{out2}");

        // Unchanged: byte for byte
        let same = Recipe::from_toml(text).unwrap();
        assert_eq!(same.to_toml_preserving(text), text);

        // A removed step rewrites the step list, but the header comment stays
        let mut fewer = Recipe::from_toml(text).unwrap();
        fewer.steps.pop();
        let out = fewer.to_toml_preserving(text);
        assert_eq!(Recipe::from_toml(&out).unwrap(), fewer);
        assert!(out.contains("# Monthly orders"));

        // And save() uses it
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("orders.recipe.toml");
        std::fs::write(&path, text).unwrap();
        r.save(&path).unwrap();
        assert!(std::fs::read_to_string(&path).unwrap().contains("# our key"));
    }
}
