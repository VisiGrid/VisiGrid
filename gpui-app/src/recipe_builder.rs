//! The recipe builder (Mode::RecipeBuilder): source settings, the step list,
//! a step's settings, and a preview of the result after any step.
//!
//! Every change re-runs the recipe against the snapshot read when the builder
//! opened, so the counts are for the whole file and match what Save and
//! Refresh will do. The view is `views/recipe_builder_view.rs`.

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use gpui::*;
use visigrid_engine::table::TableId;
use visigrid_io::csv::{CsvOptions, Encoding};
use visigrid_io::csv_import::{ColumnRule, DateOrder};
use visigrid_io::recipe::{
    self, CsvSource, FilterOp, Missing, OnError, OutColumn, Recipe, RunReport, Snapshot, Source, SourceInfo, Step,
    RECIPE_VERSION,
};

use crate::app::Spreadsheet;
use crate::mode::Mode;

/// Rows the preview shows; counts are always for the whole file.
pub const PREVIEW_ROWS: usize = 200;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Pane {
    Source,
    Steps,
    Editor,
}

/// Source settings in the left column, in focus order.
pub const CSV_ROWS: [&str; 6] = ["File", "Each refresh reads", "Delimiter", "Encoding", "Header line", "Decimal mark"];
const PARQUET_ROWS: [&str; 2] = ["File", "Each refresh reads"];
const DUCKDB_ROWS: [&str; 3] = ["File", "Each refresh reads", "Table"];
const XLSX_ROWS: [&str; 4] = ["File", "Each refresh reads", "Sheet", "Header row"];

/// The kinds of step "Add step" offers, in menu order.
pub const ADD_KINDS: [(&str, &str); 7] = [
    ("Keep columns", "choose which, in order"),
    ("Remove columns", ""),
    ("Rename columns", ""),
    ("Set types", "checked on every run"),
    ("Trim spaces", ""),
    ("Filter rows", "column, condition, value"),
    ("Remove duplicates", "whole row or by key"),
];

/// One row of the selected step's settings.
#[derive(Clone, Debug, PartialEq)]
pub enum EditorRow {
    /// A column: checked (keep/remove/trim/dedupe), its new name (rename) or
    /// its type (types). `present` is false for a name the source lacks.
    Column { name: String, present: bool },
    FilterColumn,
    FilterOp,
    FilterValue,
    OnError,
    Missing,
}

pub struct Preview {
    pub columns: Vec<OutColumn>,
    pub rows: Vec<Vec<String>>,
    pub total_rows: usize,
}

pub struct RecipeBuilder {
    /// Where the recipe is saved; None until the first save.
    pub recipe_path: Option<PathBuf>,
    pub recipe: Recipe,
    pub source_path: PathBuf,
    pub snapshot: Result<Snapshot, String>,
    pub info: Option<SourceInfo>,
    /// Editing the recipe of this Table: the primary action saves and
    /// refreshes it.
    pub link_table: Option<TableId>,
    /// Preview after this step; None: straight from the source.
    pub selected: Option<usize>,
    pub pane: Pane,
    pub source_focus: usize,
    pub editor_focus: usize,
    /// The text of a name/value field is selected: typing replaces it.
    pub text_selected: bool,
    pub add_menu: Option<usize>,
    pub dirty: bool,
    pub confirm_discard: bool,
    pub file_columns: Vec<String>,
    /// A DuckDB source's tables, for the Table setting.
    pub tables: Vec<String>,
    /// Columns going into the selected step.
    pub step_columns: Vec<String>,
    pub full: Option<RunReport>,
    pub preview: Option<Preview>,
    pub error: Option<String>,
    pub steps_scroll: ScrollHandle,
    pub editor_scroll: ScrollHandle,
}

/// The CSV settings, when the source is a CSV.
fn csv_mut(r: &mut Recipe) -> Option<&mut CsvSource> {
    match &mut r.source {
        Source::Csv(s) => Some(s),
        _ => None,
    }
}

pub fn csv(r: &Recipe) -> Option<&CsvSource> {
    match &r.source {
        Source::Csv(s) => Some(s),
        _ => None,
    }
}

/// The tables of a DuckDB file, for the Table setting.
fn duckdb_tables(path: &Path) -> Vec<String> {
    visigrid_io::duckdb::Database::open(path)
        .map(|db| db.tables().iter().map(|t| t.name.clone()).collect())
        .unwrap_or_default()
}

/// `types` values in the order Space cycles them.
const TYPE_CYCLE: [&str; 6] = ["auto", "text", "number", "date:ymd", "date:dmy", "date:mdy"];

pub fn type_label(t: &str) -> String {
    match recipe::parse_type(t) {
        Ok(ColumnRule::Date(o)) => format!("Date · {}", o.label()),
        Ok(rule) => rule.label(),
        Err(_) => t.to_string(),
    }
}

pub fn rule_label(rule: ColumnRule) -> String {
    match rule {
        ColumnRule::Date(DateOrder::Ymd) => "Date · YYYY-MM-DD".into(),
        ColumnRule::Date(o) => format!("Date · {}", o.label()),
        r => r.label(),
    }
}

const FILTER_OPS: [FilterOp; 11] = [
    FilterOp::Eq,
    FilterOp::Ne,
    FilterOp::Gt,
    FilterOp::Ge,
    FilterOp::Lt,
    FilterOp::Le,
    FilterOp::Contains,
    FilterOp::NotContains,
    FilterOp::StartsWith,
    FilterOp::NotEmpty,
    FilterOp::Empty,
];

pub fn filter_op_label(op: FilterOp) -> &'static str {
    match op {
        FilterOp::Eq => "is",
        FilterOp::Ne => "is not",
        FilterOp::Lt => "is less than",
        FilterOp::Le => "is at most",
        FilterOp::Gt => "is greater than",
        FilterOp::Ge => "is at least",
        FilterOp::Contains => "contains",
        FilterOp::NotContains => "does not contain",
        FilterOp::StartsWith => "starts with",
        FilterOp::Empty => "is empty",
        FilterOp::NotEmpty => "is not empty",
    }
}

pub fn missing_label(m: Missing) -> &'static str {
    match m {
        Missing::Fail => "Fail the run",
        Missing::Skip => "Skip this step",
        Missing::Blank => "Treat it as empty",
    }
}

pub fn on_error_label(e: OnError) -> &'static str {
    match e {
        OnError::Fail => "Fail the run",
        OnError::KeepText => "Keep the text",
        OnError::Blank => "Leave the cell empty",
    }
}

fn cycle<T: Copy + PartialEq>(all: &[T], current: T, back: bool) -> T {
    let i = all.iter().position(|x| *x == current).unwrap_or(0);
    let n = all.len();
    all[if back { (i + n - 1) % n } else { (i + 1) % n }]
}

pub fn step_missing(step: &Step) -> Missing {
    match step {
        Step::Select { missing, .. }
        | Step::Remove { missing, .. }
        | Step::Rename { missing, .. }
        | Step::Types { missing, .. }
        | Step::Trim { missing, .. }
        | Step::Filter { missing, .. }
        | Step::Dedupe { missing, .. } => *missing,
    }
}

fn step_missing_mut(step: &mut Step) -> &mut Missing {
    match step {
        Step::Select { missing, .. }
        | Step::Remove { missing, .. }
        | Step::Rename { missing, .. }
        | Step::Types { missing, .. }
        | Step::Trim { missing, .. }
        | Step::Filter { missing, .. }
        | Step::Dedupe { missing, .. } => missing,
    }
}

/// "Keep columns", "Set types", ...: the kind of a step, for its header.
pub fn step_kind(step: &Step) -> &'static str {
    match step {
        Step::Select { .. } => ADD_KINDS[0].0,
        Step::Remove { .. } => ADD_KINDS[1].0,
        Step::Rename { .. } => ADD_KINDS[2].0,
        Step::Types { .. } => ADD_KINDS[3].0,
        Step::Trim { .. } => ADD_KINDS[4].0,
        Step::Filter { .. } => ADD_KINDS[5].0,
        Step::Dedupe { .. } => ADD_KINDS[6].0,
    }
}

fn delimiter_value(d: &Option<String>) -> Option<u8> {
    match d.as_deref() {
        None => None,
        Some("tab") | Some("\t") => Some(b'\t'),
        Some(s) => s.bytes().next(),
    }
}

pub fn delimiter_name(d: u8) -> &'static str {
    crate::csv_import_ui::delimiter_name(d)
}

/// The shell-quoted `vgrid` line that runs this recipe unattended.
pub fn cli_line(recipe_path: Option<&Path>, source_path: &Path) -> String {
    let name = recipe_path
        .and_then(|p| p.file_name())
        .and_then(|n| n.to_str())
        .map(str::to_string)
        .unwrap_or_else(|| default_recipe_name(source_path));
    let stem = name.strip_suffix(".recipe.toml").unwrap_or(&name);
    let quote = |s: &str| {
        if s.chars().all(|c| c.is_ascii_alphanumeric() || "._-/".contains(c)) {
            s.to_string()
        } else {
            format!("'{}'", s.replace('\'', "'\\''"))
        }
    };
    // Never suggest writing over the file the recipe reads
    let mut out = format!("{stem}.csv");
    if source_path.file_name().and_then(|n| n.to_str()).is_some_and(|n| n.eq_ignore_ascii_case(&out)) {
        out = format!("{stem}-clean.csv");
    }
    format!("vgrid recipe run {} -o {}", quote(&name), quote(&out))
}

/// `export-2026-09.csv` -> `export-2026-09.recipe.toml`.
pub fn default_recipe_name(source_path: &Path) -> String {
    let stem = source_path.file_stem().and_then(|s| s.to_str()).unwrap_or("import");
    format!("{stem}.recipe.toml")
}

/// The source path as stored: relative when the source sits in the recipe's
/// folder (so the two can move together), else absolute.
fn stored_source_path(source: &Path, recipe_dir: Option<&Path>) -> String {
    if let (Some(dir), Some(parent), Some(name)) = (recipe_dir, source.parent(), source.file_name()) {
        if parent == dir {
            return name.to_string_lossy().into_owned();
        }
    }
    source.display().to_string()
}

impl RecipeBuilder {
    fn new(recipe: Recipe, recipe_path: Option<PathBuf>, source_path: PathBuf, link_table: Option<TableId>) -> Self {
        let snapshot = Snapshot::read(&source_path);
        let mut b = Self {
            recipe_path,
            recipe,
            source_path,
            snapshot,
            info: None,
            link_table,
            selected: None,
            pane: Pane::Steps,
            source_focus: 0,
            editor_focus: 0,
            text_selected: false,
            add_menu: None,
            dirty: false,
            confirm_discard: false,
            file_columns: Vec::new(),
            tables: Vec::new(),
            step_columns: Vec::new(),
            full: None,
            preview: None,
            error: None,
            steps_scroll: ScrollHandle::new(),
            editor_scroll: ScrollHandle::new(),
        };
        b.selected = b.recipe.steps.len().checked_sub(1);
        b.refresh_tables();
        b.recompute();
        b
    }

    /// Re-list a DuckDB source's tables (after opening or choosing a file).
    pub fn refresh_tables(&mut self) {
        self.tables = match &self.recipe.source {
            Source::Duckdb(_) => duckdb_tables(&self.source_path),
            Source::Xlsx(_) => recipe::xlsx_sheet_names(&self.source_path).unwrap_or_default(),
            _ => Vec::new(),
        };
    }

    /// The source settings this kind of source has, in focus order.
    pub fn source_rows(&self) -> &'static [&'static str] {
        match &self.recipe.source {
            Source::Csv(_) => &CSV_ROWS,
            Source::Parquet(_) => &PARQUET_ROWS,
            Source::Duckdb(_) => &DUCKDB_ROWS,
            Source::Xlsx(_) => &XLSX_ROWS,
        }
    }

    /// Re-run everything the builder shows. Cheap for the monthly exports
    /// recipes are for; the whole file is read so counts are exact.
    pub fn recompute(&mut self) {
        let Ok(snapshot) = &self.snapshot else {
            self.info = None;
            self.full = None;
            self.preview = None;
            self.file_columns.clear();
            self.step_columns.clear();
            return;
        };
        self.info = csv(&self.recipe).map(|src| recipe::source_info(src, snapshot));
        let upto = |n: usize| {
            let mut r = self.recipe.clone();
            r.steps.truncate(n);
            recipe::run(&r, snapshot)
        };
        let raw = upto(0);
        self.file_columns = raw.output.columns.iter().map(|c| c.name.clone()).collect();
        let shown = match self.selected {
            None => raw,
            Some(i) => {
                self.step_columns = upto(i).output.columns.iter().map(|c| c.name.clone()).collect();
                upto(i + 1)
            }
        };
        if self.selected.is_none() {
            self.step_columns = self.file_columns.clone();
        }
        let total = shown.output.rows.len();
        self.preview = Some(Preview {
            columns: shown.output.columns,
            rows: shown.output.rows.into_iter().take(PREVIEW_ROWS).collect(),
            total_rows: total,
        });
        self.full = Some(recipe::run(&self.recipe, snapshot).report);
    }

    fn changed(&mut self) {
        self.dirty = true;
        self.confirm_discard = false;
        self.recompute();
    }

    pub fn step(&self) -> Option<&Step> {
        self.selected.and_then(|i| self.recipe.steps.get(i))
    }

    fn step_mut(&mut self) -> Option<&mut Step> {
        let i = self.selected?;
        self.recipe.steps.get_mut(i)
    }

    /// The rows of the selected step's settings.
    pub fn editor_rows(&self) -> Vec<EditorRow> {
        let Some(step) = self.step() else { return Vec::new() };
        let mut rows = Vec::new();
        let columns = |named: &[String]| {
            let mut out: Vec<EditorRow> =
                self.step_columns.iter().map(|n| EditorRow::Column { name: n.clone(), present: true }).collect();
            for n in named {
                if !self.step_columns.iter().any(|c| c.eq_ignore_ascii_case(n)) {
                    out.push(EditorRow::Column { name: n.clone(), present: false });
                }
            }
            out
        };
        match step {
            Step::Select { columns: c, .. }
            | Step::Remove { columns: c, .. }
            | Step::Trim { columns: c, .. }
            | Step::Dedupe { columns: c, .. } => rows.extend(columns(c)),
            Step::Rename { columns: m, .. } => rows.extend(columns(&m.keys().cloned().collect::<Vec<_>>())),
            Step::Types { columns: m, .. } => {
                rows.extend(columns(&m.keys().cloned().collect::<Vec<_>>()));
                rows.push(EditorRow::OnError);
            }
            Step::Filter { op, .. } => {
                rows.push(EditorRow::FilterColumn);
                rows.push(EditorRow::FilterOp);
                if !matches!(op, FilterOp::Empty | FilterOp::NotEmpty) {
                    rows.push(EditorRow::FilterValue);
                }
            }
        }
        rows.push(EditorRow::Missing);
        rows
    }

    /// Whether a column row is checked (keep/remove/trim/dedupe).
    pub fn column_checked(&self, name: &str) -> bool {
        match self.step() {
            Some(Step::Select { columns, .. })
            | Some(Step::Remove { columns, .. })
            | Some(Step::Trim { columns, .. })
            | Some(Step::Dedupe { columns, .. }) => columns.iter().any(|c| c.eq_ignore_ascii_case(name)),
            _ => false,
        }
    }

    /// Space/Enter/click on an editor row; `back` cycles the other way.
    pub fn activate_row(&mut self, index: usize, back: bool) {
        let rows = self.editor_rows();
        let Some(row) = rows.get(index).cloned() else { return };
        self.editor_focus = index;
        let step_columns = self.step_columns.clone();
        let Some(step) = self.step_mut() else { return };
        match (row, step) {
            (EditorRow::Missing, step) => {
                let m = step_missing_mut(step);
                *m = cycle(&[Missing::Fail, Missing::Skip, Missing::Blank], *m, back);
            }
            (EditorRow::OnError, Step::Types { on_error, .. }) => {
                *on_error = cycle(&[OnError::Fail, OnError::KeepText, OnError::Blank], *on_error, back);
            }
            (EditorRow::FilterColumn, Step::Filter { column, .. }) => {
                if !step_columns.is_empty() {
                    let i = step_columns.iter().position(|c| c.eq_ignore_ascii_case(column));
                    let n = step_columns.len();
                    let next = match (i, back) {
                        (None, _) => 0,
                        (Some(i), false) => (i + 1) % n,
                        (Some(i), true) => (i + n - 1) % n,
                    };
                    *column = step_columns[next].clone();
                }
            }
            (EditorRow::FilterOp, Step::Filter { op, .. }) => *op = cycle(&FILTER_OPS, *op, back),
            (EditorRow::FilterValue, _) => {
                self.text_selected = true;
                return;
            }
            (EditorRow::Column { name, .. }, Step::Types { columns, .. }) => {
                let current = columns.get(&name).cloned().unwrap_or_else(|| "auto".into());
                let current = TYPE_CYCLE.iter().position(|t| *t == current).map_or("auto", |i| TYPE_CYCLE[i]);
                let next = cycle(&TYPE_CYCLE, current, back);
                if next == "auto" {
                    columns.remove(&name);
                } else {
                    columns.insert(name, next.into());
                }
            }
            (EditorRow::Column { .. }, Step::Rename { .. }) => {
                self.text_selected = true;
                return;
            }
            (EditorRow::Column { name, present }, step) => {
                let list = match step {
                    Step::Select { columns, .. }
                    | Step::Remove { columns, .. }
                    | Step::Trim { columns, .. }
                    | Step::Dedupe { columns, .. } => columns,
                    _ => return,
                };
                if let Some(i) = list.iter().position(|c| c.eq_ignore_ascii_case(&name)) {
                    list.remove(i);
                } else if present {
                    list.push(name);
                }
            }
            _ => return,
        }
        self.changed();
    }

    /// The text a row edits: a column's new name, or the filter value.
    pub fn row_text(&self, row: &EditorRow) -> Option<String> {
        match (row, self.step()?) {
            (EditorRow::Column { name, .. }, Step::Rename { columns, .. }) => {
                Some(columns.iter().find(|(k, _)| k.eq_ignore_ascii_case(name)).map(|(_, v)| v.clone()).unwrap_or_default())
            }
            (EditorRow::FilterValue, Step::Filter { value, .. }) => Some(value.clone()),
            _ => None,
        }
    }

    fn set_row_text(&mut self, row: &EditorRow, text: String) {
        match (row, self.step_mut()) {
            (EditorRow::Column { name, .. }, Some(Step::Rename { columns, .. })) => {
                let key = columns.keys().find(|k| k.eq_ignore_ascii_case(name)).cloned().unwrap_or_else(|| name.clone());
                if text.is_empty() {
                    columns.remove(&key);
                } else {
                    columns.insert(key, text);
                }
            }
            (EditorRow::FilterValue, Some(Step::Filter { value, .. })) => *value = text,
            _ => return,
        }
        self.changed();
    }

    /// Insert a step of kind `kind` (an index into ADD_KINDS) after the
    /// selected one, with settings that change nothing until edited where
    /// that is possible.
    pub fn add_step(&mut self, kind: usize) {
        let cols = self.step_columns_after_selected();
        let step = match kind {
            0 => Step::Select { columns: cols.clone(), missing: Missing::Fail },
            1 => Step::Remove { columns: Vec::new(), missing: Missing::Fail },
            2 => Step::Rename { columns: BTreeMap::new(), missing: Missing::Fail },
            3 => Step::Types { columns: BTreeMap::new(), on_error: OnError::Fail, missing: Missing::Fail },
            4 => Step::Trim { columns: Vec::new(), missing: Missing::Fail },
            5 => Step::Filter {
                column: cols.first().cloned().unwrap_or_default(),
                op: FilterOp::NotEmpty,
                value: String::new(),
                missing: Missing::Fail,
            },
            _ => Step::Dedupe { columns: Vec::new(), missing: Missing::Fail },
        };
        let at = self.selected.map_or(0, |i| i + 1);
        self.recipe.steps.insert(at, step);
        self.selected = Some(at);
        self.reveal_selected();
        self.add_menu = None;
        self.pane = Pane::Editor;
        self.editor_focus = 0;
        self.text_selected = false;
        self.changed();
    }

    /// Columns coming out of the selected step (what a new step after it sees).
    fn step_columns_after_selected(&self) -> Vec<String> {
        match &self.preview {
            Some(p) => p.columns.iter().map(|c| c.name.clone()).collect(),
            None => self.file_columns.clone(),
        }
    }

    pub fn remove_step(&mut self) {
        let Some(i) = self.selected else { return };
        if i < self.recipe.steps.len() {
            self.recipe.steps.remove(i);
            self.selected = if self.recipe.steps.is_empty() { None } else { Some(i.saturating_sub(1).min(self.recipe.steps.len() - 1)) };
            self.pane = Pane::Steps;
            self.changed();
        }
    }

    pub fn move_step(&mut self, up: bool) {
        let Some(i) = self.selected else { return };
        let j = if up { i.checked_sub(1) } else { Some(i + 1).filter(|j| *j < self.recipe.steps.len()) };
        if let Some(j) = j {
            self.recipe.steps.swap(i, j);
            self.selected = Some(j);
            self.reveal_selected();
            self.changed();
        }
    }

    /// Keep the selected step (row 0 is the source) in view.
    pub fn reveal_selected(&self) {
        self.steps_scroll.scroll_to_item(self.selected.map_or(0, |i| i + 1));
    }

    /// Open the add-step menu and scroll it into view (it follows the
    /// source row, every step and the "+ Add step" button).
    pub fn open_add_menu(&mut self) {
        self.add_menu = Some(0);
        self.steps_scroll.scroll_to_item(self.recipe.steps.len() + 2);
    }

    pub fn select(&mut self, step: Option<usize>) {
        self.selected = step.filter(|i| *i < self.recipe.steps.len());
        self.reveal_selected();
        self.editor_focus = 0;
        self.text_selected = false;
        self.recompute();
    }

    /// Left/Right/Space on a source setting.
    pub fn change_source(&mut self, row: usize, back: bool) {
        let name = self.source_rows().get(row).copied().unwrap_or("");
        if name == "Each refresh reads" {
            return self.toggle_pattern();
        }
        if let Source::Xlsx(src) = &mut self.recipe.source {
            match name {
                "Sheet" if !self.tables.is_empty() => {
                    let i = self.tables.iter().position(|t| t.eq_ignore_ascii_case(&src.sheet)).unwrap_or(0);
                    let n = self.tables.len();
                    src.sheet = self.tables[if back { (i + n - 1) % n } else { (i + 1) % n }].clone();
                    src.columns.clear();
                }
                "Header row" => {
                    src.header_row = if back { src.header_row.saturating_sub(1) } else { (src.header_row + 1).min(50) };
                }
                _ => return,
            }
            self.changed();
            return;
        }
        if name == "Table" {
            if let Source::Duckdb(src) = &mut self.recipe.source {
                if self.tables.is_empty() {
                    return;
                }
                let i = self.tables.iter().position(|t| t.eq_ignore_ascii_case(&src.table));
                let n = self.tables.len();
                let next = match (i, back) {
                    (None, _) => 0,
                    (Some(i), false) => (i + 1) % n,
                    (Some(i), true) => (i + n - 1) % n,
                };
                src.table = self.tables[next].clone();
                // The saved column list was for the other table
                src.columns.clear();
                self.changed();
            }
            return;
        }
        let Some(src) = csv_mut(&mut self.recipe) else { return };
        match name {
            "Delimiter" => {
                let all: [Option<&str>; 5] = [None, Some(","), Some(";"), Some("tab"), Some("|")];
                let current = all.iter().position(|d| delimiter_value(&d.map(String::from)) == delimiter_value(&src.delimiter)).unwrap_or(0);
                src.delimiter = cycle(&[0usize, 1, 2, 3, 4], current, back).pipe(|i| all[i].map(String::from));
            }
            "Encoding" => {
                let all: [Option<&str>; 4] = [None, Some("utf-8"), Some("windows-1252"), Some("utf-16")];
                let current = all
                    .iter()
                    .position(|e| e.and_then(Encoding::parse) == src.encoding.as_deref().and_then(Encoding::parse) && e.is_some() == src.encoding.is_some())
                    .unwrap_or(0);
                src.encoding = cycle(&[0usize, 1, 2, 3], current, back).pipe(|i| all[i].map(String::from));
            }
            "Header line" => {
                src.header_row = if back { src.header_row.saturating_sub(1) } else { (src.header_row + 1).min(50) };
            }
            "Decimal mark" => src.decimal_comma = !src.decimal_comma,
            _ => return,
        }
        self.changed();
    }

    /// Switch between reading this file and the newest file like it
    /// (`export-2026-09.csv` <-> `export-*-*.csv`, in the same folder).
    pub fn toggle_pattern(&mut self) {
        let stored = self.recipe.source.path().to_string();
        let current = Path::new(&stored);
        let name = if self.recipe.source_is_pattern() {
            self.source_path.file_name().and_then(|n| n.to_str()).map(str::to_string)
        } else {
            self.source_path.file_name().and_then(|n| n.to_str()).and_then(recipe::suggest_pattern)
        };
        let Some(name) = name else {
            self.error = Some("This file's name has no date or number to match next month's by.".into());
            return;
        };
        self.recipe.source.set_path(current.with_file_name(name).display().to_string());
        self.changed();
    }

    /// What a source setting shows, and a hint below it.
    pub fn source_value(&self, row: usize) -> (String, String) {
        let info = self.info.as_ref();
        let name = self.source_rows().get(row).copied().unwrap_or("");
        if let Source::Xlsx(src) = &self.recipe.source {
            match name {
                "Sheet" => {
                    let shown = if src.sheet.is_empty() { self.tables.first().cloned().unwrap_or_else(|| "First sheet".into()) } else { src.sheet.clone() };
                    let hint = match self.tables.len() {
                        0 => "Can't list this workbook's sheets.".to_string(),
                        1 => "The only sheet in this workbook.".to_string(),
                        n => format!("{n} sheets in this workbook. Saved by name, so reordering them is safe."),
                    };
                    return (shown, hint);
                }
                "Header row" => {
                    return if src.header_row == 0 {
                        ("None".into(), "No header; columns are named by letter.".into())
                    } else if src.header_row == 1 {
                        ("Row 1".into(), String::new())
                    } else {
                        (format!("Row {}", src.header_row), format!("Rows 1–{} skipped.", src.header_row - 1))
                    };
                }
                _ => {}
            }
        }
        if let Source::Duckdb(src) = &self.recipe.source {
            if name == "Table" {
                let hint = match self.tables.len() {
                    0 => "Can't list this database's tables.".to_string(),
                    1 => "The only table in this database.".to_string(),
                    n => format!("{n} tables in this database."),
                };
                let table = if src.table.is_empty() { "None chosen".into() } else { src.table.clone() };
                return (table, hint);
            }
        }
        let path_shown = self.recipe.source.path().to_string();
        let csv_default = CsvSource { path: String::new(), delimiter: None, encoding: None, header_row: 1, decimal_comma: false, columns: Vec::new() };
        let src = csv(&self.recipe).unwrap_or(&csv_default);
        match name {
            "File" => {
                let dir = self.source_path.parent().map(|p| p.display().to_string()).unwrap_or_default();
                let home = dirs::home_dir().map(|h| h.display().to_string()).unwrap_or_default();
                let dir = match dir.strip_prefix(&home) {
                    Some(rest) if !home.is_empty() => format!("~{rest}"),
                    _ => dir,
                };
                (self.source_path.file_name().and_then(|n| n.to_str()).unwrap_or("").to_string(), dir)
            }
            "Each refresh reads" => {
                let file = self.source_path.file_name().and_then(|n| n.to_str()).unwrap_or("");
                if self.recipe.source_is_pattern() {
                    let pattern = Path::new(&path_shown).file_name().and_then(|n| n.to_str()).unwrap_or("").to_string();
                    ("Newest match".into(), format!("The newest {pattern}, so next month's export is picked up by itself. Now: {file}."))
                } else {
                    match recipe::suggest_pattern(file) {
                        Some(p) => ("This file only".into(), format!("Or the newest file like {p}.")),
                        None => ("This file only".into(), String::new()),
                    }
                }
            }
            "Delimiter" => match delimiter_value(&src.delimiter) {
                None => (format!("Detected · {}", info.map_or("Comma", |i| delimiter_name(i.delimiter))), "Read from the header line down.".into()),
                Some(d) => (delimiter_name(d).into(), String::new()),
            },
            "Encoding" => match src.encoding.as_deref() {
                None => (format!("Detected · {}", info.map_or("UTF-8", |i| i.encoding.label())), String::new()),
                Some(e) => (Encoding::parse(e).map_or(e, |e| e.label()).to_string(), String::new()),
            },
            "Header line" => {
                if src.header_row == 0 {
                    ("None".into(), "No header; columns are named by letter.".into())
                } else {
                    let above: Vec<String> = info
                        .map(|i| i.first_lines.iter().take(src.header_row - 1).map(|l| format!("\"{}\"", l.chars().take(40).collect::<String>())).collect())
                        .unwrap_or_default();
                    let hint = match above.len() {
                        0 => String::new(),
                        1 => format!("Line 1 skipped: {}", above[0]),
                        n => format!("Lines 1–{n} skipped: {}", above.join(", ")),
                    };
                    (format!("Line {}", src.header_row), hint)
                }
            }
            "Decimal mark" => {
                if src.decimal_comma {
                    ("Comma · 1.234,56".into(), String::new())
                } else {
                    ("Point · 1,234.56".into(), String::new())
                }
            }
            _ => (String::new(), String::new()),
        }
    }

    /// Typing into the focused text field (a new name, a filter value).
    pub fn type_text(&mut self, key: &Keystroke, paste: Option<String>) -> bool {
        let rows = self.editor_rows();
        let Some(row) = rows.get(self.editor_focus).cloned() else { return false };
        let Some(mut text) = self.row_text(&row) else { return false };
        let mut selected = self.text_selected;
        let changed = if let Some(p) = paste {
            crate::ui::text_input::handle_input_paste(&mut text, &mut selected, &p);
            true
        } else {
            matches!(
                crate::ui::text_input::handle_input_key(
                    &mut text,
                    &mut selected,
                    key.key.as_str(),
                    key.key_char.as_deref(),
                    key.modifiers.control || key.modifiers.platform || key.modifiers.alt,
                ),
                crate::ui::text_input::InputAction::Changed
            )
        };
        self.text_selected = selected;
        if changed {
            self.set_row_text(&row, text.chars().take(200).collect());
        }
        changed
    }
}

trait Pipe: Sized {
    fn pipe<R>(self, f: impl FnOnce(Self) -> R) -> R {
        f(self)
    }
}
impl<T> Pipe for T {}

// ============================================================================
// Spreadsheet: opening, keys, saving
// ============================================================================

impl Spreadsheet {
    /// Edit an existing recipe; `link` is the Table it feeds, if any.
    pub fn open_recipe_builder(&mut self, recipe_path: &Path, link: Option<TableId>, cx: &mut Context<Self>) {
        let recipe = match Recipe::load(recipe_path) {
            Ok(r) => r,
            Err(e) => {
                self.status_message = Some(format!("Couldn't open the recipe: {e}"));
                cx.notify();
                return;
            }
        };
        // The builder previews the source: the same approval as a run
        if !crate::recipe_trust::is_approved(recipe_path, &recipe) {
            self.ask_to_approve(crate::recipe_ui::ConfirmThen::Edit(link), recipe_path.to_path_buf(), recipe, cx);
            return;
        }
        let dir = recipe_path.parent().unwrap_or(Path::new("."));
        // A pattern opens on the file it matches now
        let source_path = recipe.resolve_source(dir, None).unwrap_or_else(|_| recipe.source_path(dir, None));
        self.recipe_blocked = None;
        self.recipe_builder = Some(RecipeBuilder::new(recipe, Some(recipe_path.to_path_buf()), source_path, link));
        self.mode = Mode::RecipeBuilder;
        cx.notify();
    }

    /// Start a recipe from a CSV file, with the import dialog's settings if
    /// it was opened with some.
    pub fn new_recipe_from_file(&mut self, csv_path: &Path, options: Option<&CsvOptions>, cx: &mut Context<Self>) {
        let csv_path = std::path::absolute(csv_path).unwrap_or_else(|_| csv_path.to_path_buf());
        // Parquet and DuckDB carry their own column names and types: no
        // source settings to guess (a DuckDB recipe starts on its first table)
        let lower = csv_path.to_string_lossy().to_lowercase();
        let excel = [".xlsx", ".xlsm", ".xls"].iter().any(|e| lower.ends_with(e));
        if lower.ends_with(".parquet") || lower.ends_with(".duckdb") || excel {
            // DuckDB starts on its first table; Excel on its first sheet, kept
            // by name, with the header row guessed below any title rows
            let table = if excel {
                recipe::xlsx_sheet_names(&csv_path).ok().and_then(|names| names.into_iter().next())
            } else {
                lower.ends_with(".duckdb").then(|| duckdb_tables(&csv_path).into_iter().next().unwrap_or_default())
            };
            let mut source = Source::for_file(csv_path.display().to_string(), table);
            if let (Source::Xlsx(src), Ok(snap)) = (&mut source, Snapshot::read(&csv_path)) {
                src.header_row = recipe::guess_xlsx_header_row(&snap, &src.sheet);
            }
            let recipe = Recipe { version: RECIPE_VERSION, source, steps: Vec::new() };
            self.recipe_builder = Some(RecipeBuilder::new(recipe, None, csv_path, None));
            if let Some(b) = self.recipe_builder.as_mut() {
                b.dirty = true;
            }
            self.mode = Mode::RecipeBuilder;
            cx.notify();
            return;
        }
        let mut src = CsvSource {
            path: csv_path.display().to_string(),
            delimiter: None,
            encoding: None,
            header_row: 1,
            decimal_comma: false,
            columns: Vec::new(),
        };
        let mut steps = Vec::new();
        if let Some(o) = options {
            src.delimiter = o.delimiter.map(|d| if d == b'\t' { "tab".into() } else { (d as char).to_string() });
            src.encoding = o.encoding.map(|e| match e {
                Encoding::Utf8 => "utf-8".into(),
                Encoding::Windows1252 => "windows-1252".into(),
                _ => "utf-16".into(),
            });
            src.decimal_comma = o.decimal_comma;
            // The import dialog has no header-line setting: guess it, so an
            // export with title lines above its table starts out right
            src.header_row = if o.no_header {
                0
            } else {
                Snapshot::read(&csv_path).map_or(1, |snap| recipe::guess_header_row(&snap))
            };
            let types: BTreeMap<String, String> = o
                .columns
                .iter()
                .filter_map(|(name, rule)| {
                    let t = match rule {
                        ColumnRule::Text => "text",
                        ColumnRule::Number => "number",
                        ColumnRule::Date(DateOrder::Ymd) => "date:ymd",
                        ColumnRule::Date(DateOrder::Dmy) => "date:dmy",
                        ColumnRule::Date(DateOrder::Mdy) => "date:mdy",
                        _ => return None,
                    };
                    Some((name.clone(), t.to_string()))
                })
                .collect();
            let skipped: Vec<String> = o.columns.iter().filter(|(_, r)| *r == ColumnRule::Skip).map(|(n, _)| n.clone()).collect();
            if !skipped.is_empty() {
                steps.push(Step::Remove { columns: skipped, missing: Missing::Fail });
            }
            if !types.is_empty() {
                steps.push(Step::Types { columns: types, on_error: OnError::Fail, missing: Missing::Fail });
            }
        } else if let Ok(snap) = Snapshot::read(&csv_path) {
            src.header_row = recipe::guess_header_row(&snap);
        }
        let recipe = Recipe { version: RECIPE_VERSION, source: Source::Csv(src), steps };
        self.dismiss_csv_banner(cx);
        self.recipe_builder = Some(RecipeBuilder::new(recipe, None, csv_path, None));
        if let Some(b) = self.recipe_builder.as_mut() {
            b.dirty = true;
        }
        self.mode = Mode::RecipeBuilder;
        cx.notify();
    }

    /// Palette "New Import Recipe…": choose the file first.
    pub fn new_recipe_prompt(&mut self, cx: &mut Context<Self>) {
        let future = cx.prompt_for_paths(PathPromptOptions {
            files: true,
            directories: false,
            multiple: false,
            prompt: Some("Choose the file the recipe reads".into()),
        });
        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(paths))) = future.await {
                if let Some(path) = paths.first().cloned() {
                    let _ = this.update(cx, |this, cx| this.new_recipe_from_file(&path, None, cx));
                }
            }
        })
        .detach();
    }

    /// "Choose file…" in the builder: read another file with the same recipe.
    pub fn recipe_builder_choose_file(&mut self, cx: &mut Context<Self>) {
        let future = cx.prompt_for_paths(PathPromptOptions {
            files: true,
            directories: false,
            multiple: false,
            prompt: Some("Choose the file the recipe reads".into()),
        });
        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(paths))) = future.await {
                if let Some(path) = paths.first().cloned() {
                    let _ = this.update(cx, |this, cx| {
                        if let Some(b) = this.recipe_builder.as_mut() {
                            let recipe_dir = b.recipe_path.as_ref().and_then(|p| p.parent().map(Path::to_path_buf));
                            b.recipe.source.set_path(stored_source_path(&path, recipe_dir.as_deref()));
                            b.snapshot = Snapshot::read(&path);
                            b.source_path = path;
                            b.refresh_tables();
                            b.changed();
                        }
                        cx.notify();
                    });
                }
            }
        })
        .detach();
    }

    pub fn close_recipe_builder(&mut self, cx: &mut Context<Self>) {
        self.recipe_builder = None;
        if self.mode == Mode::RecipeBuilder {
            self.mode = Mode::Navigation;
        }
        cx.notify();
    }

    /// Save, then run: open the result as a new Table, or refresh the
    /// linked one. `then_run` false: save only.
    pub fn recipe_builder_save(&mut self, then_run: bool, cx: &mut Context<Self>) {
        let Some(b) = self.recipe_builder.as_ref() else { return };
        match b.recipe_path.clone() {
            Some(path) => self.recipe_builder_write(path, then_run, cx),
            None => {
                let dir = b.source_path.parent().map(Path::to_path_buf).unwrap_or_else(|| PathBuf::from("."));
                let name = default_recipe_name(&b.source_path);
                let future = cx.prompt_for_new_path(&dir, Some(&name));
                cx.spawn(async move |this, cx| {
                    if let Ok(Ok(Some(path))) = future.await {
                        let _ = this.update(cx, |this, cx| this.recipe_builder_write(path, then_run, cx));
                    }
                })
                .detach();
            }
        }
    }

    fn recipe_builder_write(&mut self, path: PathBuf, then_run: bool, cx: &mut Context<Self>) {
        let Some(b) = self.recipe_builder.as_mut() else { return };
        // A recipe not yet named stores its source relative to where it lands
        if b.recipe_path.as_ref() != Some(&path) {
            let stored = stored_source_path(&b.source_path, path.parent());
            let new_path = if b.recipe.source_is_pattern() {
                // Keep the pattern; only where it is looked for changes
                let pattern = Path::new(b.recipe.source.path()).file_name().map(|n| n.to_os_string());
                match pattern {
                    Some(p) => Path::new(&stored).with_file_name(p).display().to_string(),
                    None => stored,
                }
            } else {
                stored
            };
            b.recipe.source.set_path(new_path);
        }
        // The columns drift is checked against next time
        if !b.file_columns.is_empty() {
            *b.recipe.source.columns_mut() = b.file_columns.clone();
        }
        if let Err(e) = b.recipe.save(&path) {
            b.error = Some(format!("Couldn't save the recipe: {e}"));
            cx.notify();
            return;
        }
        // The user built this recipe and chose its source here
        if let Err(e) = crate::recipe_trust::approve(&path, &b.recipe) {
            b.error = Some(format!("Saved, but couldn't remember the approval: {e}"));
        }
        b.recipe_path = Some(path.clone());
        b.dirty = false;
        b.error = None;
        let link = b.link_table;
        let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("recipe").to_string();
        if !then_run {
            self.status_message = Some(format!("Saved {name}"));
            cx.notify();
            return;
        }
        self.close_recipe_builder(cx);
        match link {
            Some(id) if self.wb(cx).table(id).is_some() => {
                self.select_table_for_refresh(id, cx);
                self.refresh_recipe_table(cx);
            }
            _ => self.open_recipe(&path, cx),
        }
    }

    /// Put the cursor in Table `id`, so Refresh acts on it.
    fn select_table_for_refresh(&mut self, id: TableId, cx: &mut Context<Self>) {
        let Some((sheet_id, t)) = self.wb(cx).table(id).map(|(s, t)| (s, t.clone())) else { return };
        if let Some(index) = self.wb(cx).sheet_index_by_id(sheet_id) {
            if index != self.sheet_index(cx) {
                self.goto_sheet(index, cx);
            }
        }
        if let Some(row) = self.row_view.data_to_view(t.range.start_row) {
            self.view_state.select_cell(row, t.range.start_col);
        }
    }

    /// Keys while the builder is open. Returns true when handled.
    pub(crate) fn recipe_builder_key(&mut self, key: &Keystroke, cx: &mut Context<Self>) {
        let command = key.modifiers.control || key.modifiers.platform;
        let Some(b) = self.recipe_builder.as_mut() else { return };
        b.error = None;

        // Add-step menu
        if let Some(i) = b.add_menu {
            match key.key.as_str() {
                "escape" => b.add_menu = None,
                "up" => b.add_menu = Some(i.saturating_sub(1)),
                "down" => b.add_menu = Some((i + 1).min(ADD_KINDS.len() - 1)),
                "enter" | "space" => b.add_step(i),
                k if k.len() == 1 && ('1'..='7').contains(&k.chars().next().unwrap()) => {
                    b.add_step(k.parse::<usize>().unwrap() - 1)
                }
                _ => {}
            }
            cx.notify();
            return;
        }

        if command && key.key == "s" {
            self.recipe_builder_save(false, cx);
            return;
        }
        if command && key.key == "enter" {
            if b.full.as_ref().is_some_and(|r| r.ok) {
                self.recipe_builder_save(true, cx);
            } else {
                b.error = Some("Fix the steps that fail before loading; Ctrl+S saves the recipe as it is.".into());
                cx.notify();
            }
            return;
        }
        if key.key == "escape" {
            if b.pane == Pane::Editor {
                b.pane = Pane::Steps;
                b.text_selected = false;
            } else if b.dirty && !b.confirm_discard {
                b.confirm_discard = true;
            } else {
                self.close_recipe_builder(cx);
                return;
            }
            cx.notify();
            return;
        }
        if key.key == "tab" {
            let order = if b.selected.is_some() { vec![Pane::Source, Pane::Steps, Pane::Editor] } else { vec![Pane::Source, Pane::Steps] };
            let i = order.iter().position(|p| *p == b.pane).unwrap_or(1);
            let n = order.len();
            b.pane = order[if key.modifiers.shift { (i + n - 1) % n } else { (i + 1) % n }];
            b.text_selected = b.pane == Pane::Editor;
            cx.notify();
            return;
        }

        match b.pane {
            Pane::Source => match key.key.as_str() {
                "up" => b.source_focus = b.source_focus.saturating_sub(1),
                "down" => b.source_focus = (b.source_focus + 1).min(b.source_rows().len() - 1),
                "enter" | "space" if b.source_focus == 0 => {
                    self.recipe_builder_choose_file(cx);
                    return;
                }
                "left" => b.change_source(b.source_focus, true),
                "right" | "space" | "enter" => b.change_source(b.source_focus, false),
                _ => {}
            },
            Pane::Steps => {
                let n = b.recipe.steps.len();
                match key.key.as_str() {
                    "up" if command => b.move_step(true),
                    "down" if command => b.move_step(false),
                    "up" => b.select(match b.selected {
                        Some(0) | None => None,
                        Some(i) => Some(i - 1),
                    }),
                    "down" => b.select(match b.selected {
                        None if n > 0 => Some(0),
                        None => None,
                        Some(i) => Some((i + 1).min(n.saturating_sub(1))),
                    }),
                    "home" => b.select(None),
                    "end" => b.select(n.checked_sub(1)),
                    "enter" if b.selected.is_some() => {
                        b.pane = Pane::Editor;
                        b.editor_focus = 0;
                    }
                    "a" | "+" => b.open_add_menu(),
                    "delete" | "backspace" => b.remove_step(),
                    _ => {}
                }
            }
            Pane::Editor => {
                let rows = b.editor_rows().len();
                let focus_text = b.editor_rows().get(b.editor_focus).and_then(|r| b.row_text(r)).is_some();
                match key.key.as_str() {
                    "up" => {
                        b.editor_focus = b.editor_focus.saturating_sub(1);
                        b.text_selected = true;
                    }
                    "down" => {
                        b.editor_focus = (b.editor_focus + 1).min(rows.saturating_sub(1));
                        b.text_selected = true;
                    }
                    "enter" => {
                        if focus_text {
                            b.editor_focus = (b.editor_focus + 1).min(rows.saturating_sub(1));
                            b.text_selected = true;
                        } else {
                            b.activate_row(b.editor_focus, false);
                        }
                    }
                    "a" if command && focus_text => b.text_selected = true,
                    "v" if command && focus_text => {
                        if let Some(text) = cx.read_from_clipboard().and_then(|i| i.text()) {
                            let text = text.lines().next().unwrap_or("").to_string();
                            b.type_text(key, Some(text));
                        }
                    }
                    _ if focus_text => {
                        b.type_text(key, None);
                    }
                    "left" => b.activate_row(b.editor_focus, true),
                    "right" | "space" => b.activate_row(b.editor_focus, false),
                    _ => {}
                }
            }
        }
        cx.notify();
    }

    /// Capture keys while the builder is open, before Spreadsheet bindings.
    pub(crate) fn intercept_recipe_builder_keys(window: &mut Window, cx: &mut Context<Self>) -> Subscription {
        let this = cx.entity().downgrade();
        let handle = window.window_handle();
        cx.intercept_keystrokes(move |event, window, cx| {
            if window.window_handle() != handle {
                return;
            }
            let Some(this) = this.upgrade() else { return };
            let handled = this.update(cx, |this, cx| {
                if this.mode != Mode::RecipeBuilder {
                    return false;
                }
                this.recipe_builder_key(&event.keystroke, cx);
                true
            });
            if handled {
                cx.stop_propagation();
            }
        })
    }
}

#[cfg(test)]
mod tests {
    use super::{cli_line, stored_source_path, EditorRow, RecipeBuilder};
    use gpui::{Keystroke, Modifiers};
    use std::path::Path;
    use visigrid_io::recipe::{CsvSource, FilterOp, Missing, Recipe, Source, Step, RECIPE_VERSION};

    fn builder(csv: &str, steps: Vec<Step>) -> RecipeBuilder {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("orders.csv");
        std::fs::write(&path, csv).unwrap();
        let recipe = Recipe {
            version: RECIPE_VERSION,
            source: Source::Csv(CsvSource {
                path: path.display().to_string(),
                delimiter: None,
                encoding: None,
                header_row: 1,
                decimal_comma: false,
                columns: vec![],
            }),
            steps,
        };
        let b = RecipeBuilder::new(recipe, None, path, None);
        std::mem::forget(dir); // the snapshot is already in memory
        b
    }

    #[test]
    fn previews_after_the_selected_step_with_whole_file_counts() {
        let mut b = builder(
            "ID,Amount\n1,5\n2,0\n3,7\n",
            vec![Step::Filter { column: "Amount".into(), op: FilterOp::Gt, value: "0".into(), missing: Missing::Fail }],
        );
        assert_eq!(b.selected, Some(0));
        assert_eq!(b.preview.as_ref().unwrap().total_rows, 2);
        b.select(None);
        assert_eq!(b.preview.as_ref().unwrap().total_rows, 3);
        assert_eq!(b.file_columns, ["ID", "Amount"]);
    }

    #[test]
    fn adding_editing_and_removing_steps() {
        let mut b = builder("ID,Amount,Notes\n1,5,x\n", vec![]);
        b.add_step(0); // Keep columns: all, changes nothing
        assert_eq!(b.preview.as_ref().unwrap().columns.len(), 3);
        // Uncheck Notes
        let notes = b.editor_rows().iter().position(|r| matches!(r, EditorRow::Column { name, .. } if name == "Notes")).unwrap();
        b.activate_row(notes, false);
        assert_eq!(b.preview.as_ref().unwrap().columns.len(), 2);
        assert!(b.dirty);

        b.add_step(3); // Set types after it
        let amount = b.editor_rows().iter().position(|r| matches!(r, EditorRow::Column { name, .. } if name == "Amount")).unwrap();
        b.activate_row(amount, false); // auto -> text
        b.activate_row(amount, false); // -> number
        assert!(matches!(b.step(), Some(Step::Types { columns, .. }) if columns["Amount"] == "number"));
        b.move_step(true);
        assert_eq!(b.selected, Some(0));
        b.remove_step();
        assert_eq!(b.recipe.steps.len(), 1);
    }

    #[test]
    fn rename_and_filter_text_fields() {
        let mut b = builder("ID,Amount\n1,5\n", vec![]);
        b.add_step(2);
        b.editor_focus = 0;
        b.text_selected = true;
        for ch in ["o", "r", "d"] {
            let key = Keystroke { key: ch.into(), key_char: Some(ch.into()), modifiers: Modifiers::default(), ..Default::default() };
            b.type_text(&key, None);
        }
        assert!(matches!(b.step(), Some(Step::Rename { columns, .. }) if columns["ID"] == "ord"));
        assert_eq!(b.preview.as_ref().unwrap().columns[0].name, "ord");
        b.add_step(5);
        assert!(matches!(b.step(), Some(Step::Filter { column, .. }) if column == "ord"));
    }

    #[test]
    fn toggles_between_this_file_and_the_newest_like_it() {
        let mut b = builder("ID\n1\n", vec![]);
        // builder() names the file orders.csv: no digits, nothing to match by
        b.toggle_pattern();
        assert!(b.error.is_some());
        let dir = tempfile::tempdir().unwrap();
        let file = dir.path().join("export-2026-09.csv");
        std::fs::write(&file, "ID\n1\n").unwrap();
        b.source_path = file.clone();
        b.recipe.source = Source::Csv(CsvSource { path: file.display().to_string(), delimiter: None, encoding: None, header_row: 1, decimal_comma: false, columns: vec![] });
        b.change_source(1, false);
        assert!(b.recipe.source_is_pattern());
        assert_eq!(b.source_value(1).0, "Newest match");
        assert!(b.source_value(1).1.contains("export-*-*.csv"));
        b.change_source(1, false);
        assert!(!b.recipe.source_is_pattern());
        assert_eq!(b.recipe.source.path(), file.display().to_string());
    }

    #[test]
    fn source_settings_and_stored_paths() {
        let mut b = builder("Title\nID;Amount\n1;5\n2;6\n", vec![]);
        b.change_source(4, false); // header line 2
        assert_eq!(b.file_columns, ["ID", "Amount"]);
        assert!(b.source_value(4).1.contains("Title"));
        assert_eq!(stored_source_path(Path::new("/d/x.csv"), Some(Path::new("/d"))), "x.csv");
        assert_eq!(stored_source_path(Path::new("/e/x.csv"), Some(Path::new("/d"))), "/e/x.csv");
        assert_eq!(cli_line(Some(Path::new("/d/my orders.recipe.toml")), Path::new("x.csv")), "vgrid recipe run 'my orders.recipe.toml' -o 'my orders.csv'");
        assert_eq!(cli_line(None, Path::new("/d/export-09.csv")), "vgrid recipe run export-09.recipe.toml -o export-09-clean.csv");
    }
}
