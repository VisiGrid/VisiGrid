// CSV import: options, per-column rules and the report of what was decided.
//
// Spec: Obsidian `Projects/Visi/VisiGrid CSV Import Spec.md`.
//
// Two rules hold for every option combination:
// - A value is never changed silently. Anything that cannot be read as a number
//   without changing it stays text unless the user chose a type for its column,
//   and even then a value that does not fit the chosen type stays text and is
//   counted, rather than becoming an error or a wrong number.
// - Dates are never guessed. A column becomes dates only when it is given
//   `ColumnRule::Date` with an explicit day/month order.

use std::path::Path;

use visigrid_engine::cell::{date_to_serial, DateStyle, NumberFormat};
use visigrid_engine::sheet::{Sheet, SheetId, NUM_ROWS};

use crate::csv::sniff_delimiter;

/// How many raw records the report keeps for a preview.
pub const SAMPLE_ROWS: usize = 50;

/// Text encodings the importer reads.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Encoding {
    Utf8,
    Utf16Le,
    Utf16Be,
    Windows1252,
}

impl Encoding {
    pub fn label(self) -> &'static str {
        match self {
            Encoding::Utf8 => "UTF-8",
            Encoding::Utf16Le | Encoding::Utf16Be => "UTF-16",
            Encoding::Windows1252 => "Windows-1252",
        }
    }

    /// Parse a CLI value: utf-8, utf-16, windows-1252 (also latin-1, cp1252).
    pub fn parse(s: &str) -> Option<Encoding> {
        match s.to_ascii_lowercase().replace('_', "-").as_str() {
            "utf-8" | "utf8" => Some(Encoding::Utf8),
            "utf-16" | "utf16" | "utf-16le" => Some(Encoding::Utf16Le),
            "utf-16be" => Some(Encoding::Utf16Be),
            "windows-1252" | "cp1252" | "latin-1" | "latin1" | "iso-8859-1" => Some(Encoding::Windows1252),
            _ => None,
        }
    }
}

/// Day/month/year order for a date column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateOrder {
    Ymd,
    Dmy,
    Mdy,
}

impl DateOrder {
    pub fn label(self) -> &'static str {
        match self {
            DateOrder::Ymd => "YYYY-MM-DD",
            DateOrder::Dmy => "DD/MM/YYYY",
            DateOrder::Mdy => "MM/DD/YYYY",
        }
    }

    pub fn parse(s: &str) -> Option<DateOrder> {
        match s.to_ascii_uppercase().as_str() {
            "YMD" | "YYYY-MM-DD" | "ISO" => Some(DateOrder::Ymd),
            "DMY" | "DD/MM/YYYY" => Some(DateOrder::Dmy),
            "MDY" | "MM/DD/YYYY" => Some(DateOrder::Mdy),
            _ => None,
        }
    }
}

/// What a column becomes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ColumnRule {
    /// Numbers where that changes nothing; text otherwise. Never dates.
    #[default]
    Auto,
    Text,
    /// Numbers, including ones Auto keeps as text (007 becomes 7). A value that
    /// is not a number stays text and is counted.
    Number,
    Date(DateOrder),
    /// Not imported; later columns move left.
    Skip,
}

impl ColumnRule {
    pub fn label(self) -> String {
        match self {
            ColumnRule::Auto => "Auto".into(),
            ColumnRule::Text => "Text".into(),
            ColumnRule::Number => "Number".into(),
            ColumnRule::Date(o) => format!("Date {}", o.label()),
            ColumnRule::Skip => "Skip".into(),
        }
    }
}

/// How a CSV is read.
#[derive(Debug, Clone, Default)]
pub struct CsvOptions {
    /// None: sniffed from the first lines.
    pub delimiter: Option<u8>,
    /// None: detected (byte-order mark, then UTF-8, then Windows-1252).
    pub encoding: Option<Encoding>,
    /// The first row holds column names: the names that `columns` and the
    /// report use, exempt from column rules. On: every row is data and columns
    /// are named by letter.
    pub no_header: bool,
    /// Read 1.234,56 as 1234.56 (Germany, France, Brazil…).
    pub decimal_comma: bool,
    /// Run fields that start with `=` as formulas. Off by default: a CSV is
    /// data, and a formula in one is the classic injection vector (OWASP; a
    /// 2025 USENIX WOOT study found every spreadsheet tested would exfiltrate
    /// through one). `vgrid fill` refuses them outright.
    pub evaluate_formulas: bool,
    /// Per-column rules, matched by header name (case-insensitive) or by column
    /// letter. Columns not named are Auto.
    pub columns: Vec<(String, ColumnRule)>,
    /// Where the first field lands on the sheet (row, column). The CLI's
    /// `calc --into` uses it; everything else imports at A1.
    pub origin: (usize, usize),
}

impl CsvOptions {
    fn rule_for(&self, index: usize, name: &str) -> ColumnRule {
        let letter = col_label(index);
        self.columns
            .iter()
            .find(|(key, _)| key.trim().eq_ignore_ascii_case(name.trim()) || key.trim().eq_ignore_ascii_case(&letter))
            .map(|(_, rule)| *rule)
            .unwrap_or_default()
    }

    /// The same options as `vgrid` flags, for "Same import from the command
    /// line". Only what differs from the defaults is written.
    pub fn cli_flags(&self) -> Vec<String> {
        let mut flags = Vec::new();
        let group = |rule: fn(&ColumnRule) -> bool| -> Vec<String> {
            self.columns.iter().filter(|(_, r)| rule(r)).map(|(n, _)| quote_flag_value(n)).collect()
        };
        let text = group(|r| *r == ColumnRule::Text);
        if !text.is_empty() {
            flags.push(format!("--text {}", text.join(",")));
        }
        let number = group(|r| *r == ColumnRule::Number);
        if !number.is_empty() {
            flags.push(format!("--number {}", number.join(",")));
        }
        for (name, rule) in &self.columns {
            if let ColumnRule::Date(order) = rule {
                let o = match order {
                    DateOrder::Ymd => "ymd",
                    DateOrder::Dmy => "dmy",
                    DateOrder::Mdy => "mdy",
                };
                flags.push(format!("--date {}={}", quote_flag_value(name), o));
            }
        }
        let skip = group(|r| *r == ColumnRule::Skip);
        if !skip.is_empty() {
            flags.push(format!("--skip {}", skip.join(",")));
        }
        if let Some(d) = self.delimiter {
            flags.push(format!(
                "--delimiter {}",
                match d {
                    b'\t' => "$'\\t'".to_string(),
                    b';' => "';'".to_string(),
                    b'|' => "'|'".to_string(),
                    other => (other as char).to_string(),
                }
            ));
        }
        if let Some(e) = self.encoding {
            flags.push(format!("--encoding {}", e.label().to_ascii_lowercase()));
        }
        if self.no_header {
            flags.push("--no-header".into());
        }
        if self.decimal_comma {
            flags.push("--decimal-comma".into());
        }
        if self.evaluate_formulas {
            flags.push("--formulas".into());
        }
        flags
    }
}

fn quote_flag_value(s: &str) -> String {
    if s.chars().all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '-' || c == '.') {
        s.to_string()
    } else {
        format!("'{}'", s.replace('\'', "'\\''"))
    }
}

/// What Auto (or a chosen rule) did to one column.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Resolved {
    Empty,
    Text,
    Number,
    Date,
    /// Some numbers, some text.
    Mixed,
    Skipped,
}

impl Resolved {
    pub fn label(&self) -> &'static str {
        match self {
            Resolved::Empty => "Empty",
            Resolved::Text => "Text",
            Resolved::Number => "Number",
            Resolved::Date => "Date",
            Resolved::Mixed => "Mixed",
            Resolved::Skipped => "Skipped",
        }
    }
}

/// One source column in the report.
#[derive(Debug, Clone)]
pub struct ColumnDecision {
    /// Header name, or the column letter without a header.
    pub name: String,
    /// Index in the file (0-based), before skipped columns are removed.
    pub source_index: usize,
    pub rule: ColumnRule,
    pub resolved: Resolved,
    /// Why Auto kept text, in a few words ("leading zeros"), or None.
    pub reason: Option<&'static str>,
    /// Values kept as text because a number would have changed them.
    pub kept_as_text: usize,
    /// Values a chosen Number or Date rule could not read; kept as text.
    pub unreadable: usize,
    /// Values that choosing Number would change (drop zeros, lose digits).
    pub changed_if_number: usize,
    /// Every non-empty value parses as YYYY-MM-DD: offer Date, never apply it.
    pub looks_like_iso_dates: bool,
}

/// A CSV import, with what was decided and what did not fit.
#[derive(Debug)]
pub struct CsvImport {
    pub sheet: Sheet,
    pub delimiter: u8,
    pub encoding: Encoding,
    pub columns: Vec<ColumnDecision>,
    /// Values kept as text across all Auto columns.
    pub kept_as_text: usize,
    /// Fields starting with `=` kept as text, and where they are (sheet row, col).
    pub formulas_as_text: usize,
    pub formula_cells: Vec<(usize, usize)>,
    /// Header labels of columns where values were kept as text.
    pub text_columns: Vec<String>,
    /// Records in the file, and records placed on the sheet.
    pub rows_in_file: usize,
    pub rows_loaded: usize,
    /// The first records as read, for a preview.
    pub sample: Vec<Vec<String>>,
    /// Fingerprint of every field starting with `=` (position and text),
    /// evaluated or not. Evaluating needs the user to have approved exactly
    /// these formulas: a file that changes after review is asked about again.
    /// Comparable within one process only.
    pub formulas_digest: u64,
}

impl CsvImport {
    /// One line for the status bar or banner body, or None when nothing needs saying.
    pub fn message(&self) -> Option<String> {
        let mut parts = Vec::new();
        if self.rows_loaded < self.rows_in_file {
            parts.push(format!(
                "Loaded the first {} of {} rows — the sheet holds {}; the rest were not imported",
                self.rows_loaded, self.rows_in_file, NUM_ROWS
            ));
        }
        if self.kept_as_text > 0 {
            let cols = match self.text_columns.len() {
                0 => String::new(),
                n if n <= 3 => format!(" ({})", self.text_columns.join(", ")),
                n => format!(" ({} and {} more)", self.text_columns[..3].join(", "), n - 3),
            };
            parts.push(format!(
                "Kept {} value{} as text so leading zeros and long IDs stay exact{}",
                self.kept_as_text,
                plural(self.kept_as_text),
                cols
            ));
        }
        let unreadable: usize = self.columns.iter().map(|c| c.unreadable).sum();
        if unreadable > 0 {
            let names: Vec<&str> = self.columns.iter().filter(|c| c.unreadable > 0).map(|c| c.name.as_str()).collect();
            parts.push(format!(
                "{} value{} in {} did not fit the chosen type and stayed as text",
                unreadable,
                plural(unreadable),
                names.join(", ")
            ));
        }
        if self.formulas_as_text > 0 {
            parts.push(format!("Left {} formula{} as text", self.formulas_as_text, plural(self.formulas_as_text)));
        }
        (!parts.is_empty()).then(|| parts.join(". ") + ".")
    }

    /// The short form for the status bar: "comma · UTF-8 · 3 text columns".
    pub fn summary_line(&self) -> String {
        let delim = match self.delimiter {
            b',' => "comma".to_string(),
            b';' => "semicolon".to_string(),
            b'\t' => "tab".to_string(),
            b'|' => "pipe".to_string(),
            other => format!("'{}'", other as char),
        };
        let text_cols = self.columns.iter().filter(|c| c.resolved == Resolved::Text || c.resolved == Resolved::Mixed).count();
        let mut s = format!("{} · {}", delim, self.encoding.label());
        if text_cols > 0 {
            s.push_str(&format!(" · {} text column{}", text_cols, plural(text_cols)));
        }
        s
    }
}

fn plural(n: usize) -> &'static str {
    if n == 1 {
        ""
    } else {
        "s"
    }
}

/// Read a CSV file.
pub fn import_report(path: &Path, options: &CsvOptions) -> Result<CsvImport, String> {
    let bytes = std::fs::read(path).map_err(|e| e.to_string())?;
    let (content, encoding) = decode(&bytes, options.encoding);
    import_str(&content, encoding, options, usize::MAX)
}

/// Import CSV text that is already decoded (stdin, the clipboard).
pub fn import_text(content: &str, options: &CsvOptions) -> Result<CsvImport, String> {
    import_str(content, Encoding::Utf8, options, usize::MAX)
}

/// Read only enough of a file for a preview (the import dialog). Same rules as
/// a full import, applied to the first `max_rows` records.
pub fn preview(path: &Path, options: &CsvOptions, max_rows: usize) -> Result<CsvImport, String> {
    let bytes = std::fs::read(path).map_err(|e| e.to_string())?;
    let (content, encoding) = decode(&bytes, options.encoding);
    import_str(&content, encoding, options, max_rows)
}

/// Preview from bytes already read (the import dialog reads the start of a file
/// once and re-previews from it on every change). A trailing partial record is
/// harmless: only the first `max_rows` records are used.
pub fn preview_bytes(bytes: &[u8], options: &CsvOptions, max_rows: usize) -> Result<CsvImport, String> {
    let (content, encoding) = decode(bytes, options.encoding);
    import_str(&content, encoding, options, max_rows)
}

/// The first `limit` bytes of a file, cut back to the last line break so a
/// multi-byte character is never split (which would make UTF-8 look invalid).
pub fn read_head(path: &Path, limit: usize) -> Result<Vec<u8>, String> {
    use std::io::Read;
    let mut file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut bytes = Vec::with_capacity(limit.min(1 << 20));
    file.by_ref().take(limit as u64).read_to_end(&mut bytes).map_err(|e| e.to_string())?;
    if bytes.len() == limit {
        if let Some(end) = bytes.iter().rposition(|&b| b == b'\n') {
            let utf16 = bytes.starts_with(&[0xFF, 0xFE]) || bytes.starts_with(&[0xFE, 0xFF]);
            // UTF-16 LE puts the newline's zero byte after it
            let keep = if utf16 && bytes.starts_with(&[0xFF, 0xFE]) { end + 2 } else { end + 1 };
            bytes.truncate(keep.min(bytes.len()));
            if utf16 && bytes.len() % 2 == 1 {
                bytes.pop();
            }
        }
    }
    Ok(bytes)
}

/// Decode bytes: an explicit encoding, else a byte-order mark, else UTF-8, else
/// Windows-1252 (Excel's "CSV" on Western Windows). A UTF-8 byte-order mark is
/// removed rather than left on the first header name.
pub fn decode(bytes: &[u8], encoding: Option<Encoding>) -> (String, Encoding) {
    let detected = encoding.unwrap_or_else(|| {
        if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
            Encoding::Utf8
        } else if bytes.starts_with(&[0xFF, 0xFE]) {
            Encoding::Utf16Le
        } else if bytes.starts_with(&[0xFE, 0xFF]) {
            Encoding::Utf16Be
        } else if std::str::from_utf8(bytes).is_ok() {
            Encoding::Utf8
        } else {
            Encoding::Windows1252
        }
    });
    let text = match detected {
        Encoding::Utf8 => encoding_rs::UTF_8.decode_with_bom_removal(bytes).0.into_owned(),
        Encoding::Utf16Le => encoding_rs::UTF_16LE.decode_with_bom_removal(bytes).0.into_owned(),
        Encoding::Utf16Be => encoding_rs::UTF_16BE.decode_with_bom_removal(bytes).0.into_owned(),
        Encoding::Windows1252 => encoding_rs::WINDOWS_1252.decode(bytes).0.into_owned(),
    };
    (text, detected)
}

/// Why Auto keeps a field as text.
#[derive(Debug, PartialEq, Clone, Copy)]
pub(crate) enum KeepText {
    LeadingZeros,
    LongDigits,
    ENotationId,
    Formula,
}

impl KeepText {
    fn reason(self) -> &'static str {
        match self {
            KeepText::LeadingZeros => "leading zeros",
            KeepText::LongDigits => "more than 15 digits",
            KeepText::ENotationId => "IDs like 2310009E13",
            KeepText::Formula => "starts with =",
        }
    }
}

/// Whether Auto must keep a field as text. Everything else is typed as if it
/// had been entered by hand (numbers, 5%, $1,234, (500)) — and nothing becomes
/// a date, so gene names, part numbers and 1-2 stay as written.
pub(crate) fn keep_as_text(field: &str, evaluate_formulas: bool) -> Option<KeepText> {
    let t = field.trim();
    if t.starts_with('=') {
        return (!evaluate_formulas).then_some(KeepText::Formula);
    }
    let digits = t.strip_prefix(['+', '-']).unwrap_or(t);
    if !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit()) {
        // 007, 00123: a number would drop the zeros (ZIP codes, SKUs, accounts).
        if digits.len() > 1 && digits.starts_with('0') {
            return Some(KeepText::LeadingZeros);
        }
        // More than 15 digits: past what a double holds exactly (cards, barcodes).
        if digits.len() > 15 {
            return Some(KeepText::LongDigits);
        }
        return None;
    }
    // 12E4, 2310009E13: digits, E, digits, nothing else. Programs write real
    // scientific notation with a sign or a point (1e-05, 1.5E+07); this shape
    // is an ID, and reading it as 120000 destroys it.
    if let Some((m, e)) = t.split_once(['e', 'E']) {
        let all_digits = |x: &str| !x.is_empty() && x.bytes().all(|b| b.is_ascii_digit());
        if all_digits(m) && all_digits(e) {
            return Some(KeepText::ENotationId);
        }
    }
    None
}

/// A number read with a decimal comma: 1.234,56 → 1234.56, 1,5 → 1.5,
/// -1.234 → -1234. None if the text is not shaped like one.
pub(crate) fn parse_decimal_comma(s: &str) -> Option<f64> {
    let t = s.trim();
    let (neg, body) = match t.strip_prefix('-') {
        Some(rest) => (true, rest),
        None => (false, t.strip_prefix('+').unwrap_or(t)),
    };
    let (int_part, frac) = match body.split_once(',') {
        Some((i, f)) if !f.is_empty() && f.bytes().all(|b| b.is_ascii_digit()) => (i, Some(f)),
        Some(_) => return None,
        None => (body, None),
    };
    if int_part.is_empty() {
        return None;
    }
    // Thousands separators are dots in groups of three after the first group.
    let groups: Vec<&str> = int_part.split('.').collect();
    if groups.len() > 1 {
        if groups[0].is_empty() || groups[0].len() > 3 || groups[1..].iter().any(|g| g.len() != 3) {
            return None;
        }
    }
    if !groups.iter().all(|g| g.bytes().all(|b| b.is_ascii_digit())) {
        return None;
    }
    let mut n: f64 = format!("{}.{}", groups.concat(), frac.unwrap_or("0")).parse().ok()?;
    if neg {
        n = -n;
    }
    Some(n)
}

/// A number a Number column accepts: plain, signed, decimal, E notation, with
/// thousands separators; respecting the decimal mark. Leading zeros are fine
/// here — the user chose Number.
pub(crate) fn parse_number(s: &str, decimal_comma: bool) -> Option<f64> {
    if decimal_comma {
        return parse_decimal_comma(s);
    }
    let t = s.trim().replace(',', "");
    visigrid_engine::cell::parse_finite(&t)
}

/// A date in the given order, with -, / or . between parts; ISO is always
/// accepted. Returns the date serial.
pub(crate) fn parse_date(s: &str, order: DateOrder) -> Option<f64> {
    let t = s.trim();
    let parts: Vec<&str> = t.split(['-', '/', '.']).collect();
    if parts.len() != 3 || parts.iter().any(|p| p.is_empty() || !p.bytes().all(|b| b.is_ascii_digit())) {
        return None;
    }
    let n = |i: usize| parts[i].parse::<u32>().ok();
    let (y, m, d) = if parts[0].len() == 4 {
        (n(0)?, n(1)?, n(2)?)
    } else {
        match order {
            DateOrder::Ymd => return None,
            DateOrder::Dmy => (n(2)?, n(1)?, n(0)?),
            DateOrder::Mdy => (n(2)?, n(0)?, n(1)?),
        }
    };
    let y = if parts[0].len() != 4 && y < 100 { y + 2000 } else { y };
    if !(1..=12).contains(&m) || d == 0 || d > days_in_month(y as i32, m) || !(1900..=9999).contains(&y) {
        return None;
    }
    Some(date_to_serial(y as i32, m, d))
}

fn days_in_month(y: i32, m: u32) -> u32 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        _ if (y % 4 == 0 && y % 100 != 0) || y % 400 == 0 => 29,
        _ => 28,
    }
}

/// Import already-decoded text. `max_rows` limits the records read (a preview);
/// usize::MAX reads them all.
pub(crate) fn import_str(content: &str, encoding: Encoding, options: &CsvOptions, max_rows: usize) -> Result<CsvImport, String> {
    let delimiter = options.delimiter.unwrap_or_else(|| sniff_delimiter(content));
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(content.as_bytes());

    let mut sheet = Sheet::new(SheetId(1), NUM_ROWS, visigrid_engine::sheet::NUM_COLS);
    if max_rows == usize::MAX {
        sheet.reserve_cells(crate::csv::estimate_cells(content, delimiter));
    }
    let mut max_row = 0usize;
    let mut max_col = 0usize;
    let mut decisions: Vec<ColumnDecision> = Vec::new();
    // Per column: numbers seen, texts seen, ISO-date-shaped non-empty values seen.
    let mut seen: Vec<(usize, usize, usize, usize)> = Vec::new();
    let mut dest_of: Vec<Option<usize>> = Vec::new();
    let (mut formulas_as_text, mut kept_as_text) = (0usize, 0usize);
    let mut formula_cells = Vec::new();
    let mut rows_in_file = 0usize;
    let mut sample = Vec::new();

    let ensure = |decisions: &mut Vec<ColumnDecision>, seen: &mut Vec<(usize, usize, usize, usize)>, dest_of: &mut Vec<Option<usize>>, upto: usize, header: &[String]| {
        while decisions.len() <= upto {
            let i = decisions.len();
            let name = header.get(i).filter(|h| !h.is_empty()).cloned().unwrap_or_else(|| col_label(i));
            let rule = options.rule_for(i, &name);
            decisions.push(ColumnDecision {
                name,
                source_index: i,
                rule,
                resolved: if rule == ColumnRule::Skip { Resolved::Skipped } else { Resolved::Empty },
                reason: None,
                kept_as_text: 0,
                unreadable: 0,
                changed_if_number: 0,
                looks_like_iso_dates: false,
            });
            seen.push((0, 0, 0, 0));
            let skipped_before = decisions[..i].iter().filter(|d| d.rule == ColumnRule::Skip).count();
            dest_of.push((rule != ColumnRule::Skip).then_some(i - skipped_before));
        }
    };

    let mut header: Vec<String> = Vec::new();
    let (orow, ocol) = options.origin;
    let mut formulas = std::collections::hash_map::DefaultHasher::new();
    for (row_idx, result) in reader.records().enumerate() {
        if row_idx >= max_rows {
            break;
        }
        let record = result.map_err(|e| e.to_string())?;
        rows_in_file = row_idx + 1;
        if sample.len() < SAMPLE_ROWS {
            sample.push(record.iter().map(|f| f.to_string()).collect());
        }
        if row_idx >= NUM_ROWS {
            continue;
        }
        let is_header = row_idx == 0 && !options.no_header;
        if is_header {
            header = record.iter().map(|f| f.trim().to_string()).collect();
        }
        if !record.is_empty() {
            ensure(&mut decisions, &mut seen, &mut dest_of, record.len() - 1, &header);
        }
        for (col_idx, field) in record.iter().enumerate() {
            // The same test keep_as_text uses (trimmed), so a formula with
            // leading spaces is fingerprinted like any other
            if field.trim().starts_with('=') {
                use std::hash::Hash;
                (row_idx, col_idx, field).hash(&mut formulas);
            }
            let Some(dest) = dest_of[col_idx] else { continue };
            if field.is_empty() {
                continue;
            }
            max_col = max_col.max(dest + ocol);
            if is_header {
                // Header names take no column rule (a Number column's name is
                // not a number) and are not counted. They are typed by Auto like
                // any value, so a file with no header row still opens with its
                // first row of numbers as numbers.
                match keep_as_text(field, options.evaluate_formulas) {
                    Some(_) => sheet.set_text(row_idx + orow, dest + ocol, field),
                    None => sheet.set_value_deferred(row_idx + orow, dest + ocol, field),
                }
                continue;
            }
            let d = &mut decisions[col_idx];
            let s = &mut seen[col_idx];
            if parse_date(field, DateOrder::Ymd).is_some() {
                s.2 += 1;
            }
            s.3 += 1;
            // What Number would do to this field (for the dialog's warning).
            let auto_keep = keep_as_text(field, true);
            if matches!(auto_keep, Some(KeepText::LeadingZeros | KeepText::LongDigits | KeepText::ENotationId))
                || parse_number(field, options.decimal_comma).is_none()
            {
                d.changed_if_number += 1;
            }
            match d.rule {
                ColumnRule::Skip => {}
                ColumnRule::Text => {
                    sheet.set_text(row_idx + orow, dest + ocol, field);
                    s.1 += 1;
                }
                ColumnRule::Number => match parse_number(field, options.decimal_comma) {
                    Some(n) => {
                        sheet.set_value_deferred(row_idx + orow, dest + ocol, &visigrid_engine::cell::interchange_number(n));
                        s.0 += 1;
                    }
                    None => {
                        sheet.set_text(row_idx + orow, dest + ocol, field);
                        d.unreadable += 1;
                        s.1 += 1;
                    }
                },
                ColumnRule::Date(order) => match parse_date(field, order) {
                    Some(serial) => {
                        sheet.set_value_deferred(row_idx + orow, dest + ocol, &visigrid_engine::cell::interchange_number(serial));
                        sheet.set_number_format(row_idx + orow, dest + ocol, NumberFormat::Date { style: DateStyle::Iso });
                        s.0 += 1;
                    }
                    None => {
                        sheet.set_text(row_idx + orow, dest + ocol, field);
                        d.unreadable += 1;
                        s.1 += 1;
                    }
                },
                ColumnRule::Auto => match keep_as_text(field, options.evaluate_formulas) {
                    Some(KeepText::Formula) => {
                        sheet.set_text(row_idx + orow, dest + ocol, field);
                        formulas_as_text += 1;
                        formula_cells.push((row_idx + orow, dest + ocol));
                        s.1 += 1;
                    }
                    Some(reason) => {
                        sheet.set_text(row_idx + orow, dest + ocol, field);
                        d.kept_as_text += 1;
                        kept_as_text += 1;
                        d.reason.get_or_insert(reason.reason());
                        s.1 += 1;
                    }
                    None => {
                        if options.decimal_comma {
                            match parse_decimal_comma(field) {
                                Some(n) => {
                                    sheet.set_value_deferred(row_idx + orow, dest + ocol, &visigrid_engine::cell::interchange_number(n));
                                    s.0 += 1;
                                }
                                None => {
                                    sheet.set_value_deferred(row_idx + orow, dest + ocol, field);
                                    s.1 += 1;
                                }
                            }
                        } else {
                            sheet.set_value_deferred(row_idx + orow, dest + ocol, field);
                            if field.starts_with('=') || matches!(visigrid_engine::cell::CellValue::from_input(field), visigrid_engine::cell::CellValue::Number(_)) {
                                s.0 += 1;
                            } else {
                                s.1 += 1;
                            }
                        }
                    }
                },
            }
        }
        max_row = row_idx + orow;
    }

    for (d, s) in decisions.iter_mut().zip(&seen) {
        if d.rule == ColumnRule::Skip {
            continue;
        }
        let (numbers, texts, isoish, nonempty) = *s;
        d.looks_like_iso_dates = nonempty > 0 && isoish == nonempty && !matches!(d.rule, ColumnRule::Date(_));
        d.resolved = match (numbers, texts) {
            (0, 0) => Resolved::Empty,
            (_, 0) if matches!(d.rule, ColumnRule::Date(_)) => Resolved::Date,
            (_, 0) => Resolved::Number,
            (0, _) => Resolved::Text,
            _ => Resolved::Mixed,
        };
        if d.looks_like_iso_dates && d.reason.is_none() {
            d.reason = Some("looks like YYYY-MM-DD dates; choose Date to convert");
        }
    }

    sheet.rows = (max_row + 1).max(1000);
    sheet.cols = (max_col + 1).max(26);

    // Evaluate in dependency order rather than in the order the rows happened
    // to be read: a formula naming a row further down the file must see it.
    let mut wb = visigrid_engine::workbook::Workbook::from_sheets(vec![sheet], 0);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();

    let text_columns = decisions.iter().filter(|d| d.kept_as_text > 0).map(|d| d.name.clone()).collect();
    let rows_loaded = rows_in_file.min(NUM_ROWS);
    Ok(CsvImport {
        sheet: wb.into_sheets().swap_remove(0),
        delimiter,
        encoding,
        columns: decisions,
        kept_as_text,
        formulas_as_text,
        formula_cells,
        text_columns,
        rows_in_file,
        rows_loaded,
        sample,
        formulas_digest: std::hash::Hasher::finish(&formulas),
    })
}

/// Records in a file and its first record, without typing any values: what
/// "file changed on disk" needs to say how it changed.
pub fn count_records(path: &Path, options: &CsvOptions) -> Result<(usize, Vec<String>), String> {
    let bytes = std::fs::read(path).map_err(|e| e.to_string())?;
    let (content, _) = decode(&bytes, options.encoding);
    let delimiter = options.delimiter.unwrap_or_else(|| sniff_delimiter(&content));
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .flexible(true)
        .from_reader(content.as_bytes());
    let mut first = Vec::new();
    let mut count = 0usize;
    let mut record = csv::StringRecord::new();
    while reader.read_record(&mut record).map_err(|e| e.to_string())? {
        if count == 0 {
            first = record.iter().map(|f| f.trim().to_string()).collect();
        }
        count += 1;
    }
    Ok((count, first))
}

/// Column letters for a 0-based index (0 -> A, 26 -> AA).
pub fn col_label(mut c: usize) -> String {
    let mut s = String::new();
    loop {
        s.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;
    use visigrid_engine::formula::eval::Value;

    fn run(csv: &str, options: &CsvOptions) -> CsvImport {
        import_str(csv, Encoding::Utf8, options, usize::MAX).unwrap()
    }

    fn cell(r: &CsvImport, row: usize, col: usize) -> (String, bool) {
        let v = r.sheet.get_computed_value(row, col);
        (r.sheet.get_display(row, col), matches!(v, Value::Text(_)))
    }

    #[test]
    fn formulas_digest_follows_the_formulas_not_the_evaluate_flag() {
        let a = "id,note\n1,=1+1\n2,=HYPERLINK(\"http://x\",\"y\")\n";
        let text = run(a, &CsvOptions::default());
        let evaluated = run(a, &CsvOptions { evaluate_formulas: true, ..Default::default() });
        assert_eq!(text.formulas_digest, evaluated.formulas_digest);
        // Same values elsewhere, one formula edited: a different digest
        let b = "id,note\n1,=1+1\n2,=HYPERLINK(\"http://evil\",\"y\")\n";
        assert_ne!(text.formulas_digest, run(b, &CsvOptions::default()).formulas_digest);
        // A formula moved to another row: different too
        let c = "id,note\n1,=HYPERLINK(\"http://x\",\"y\")\n2,=1+1\n";
        assert_ne!(text.formulas_digest, run(c, &CsvOptions::default()).formulas_digest);
        // Non-formula changes do not matter
        let d = "id,note\n9,=1+1\n8,=HYPERLINK(\"http://x\",\"y\")\n";
        assert_eq!(text.formulas_digest, run(d, &CsvOptions::default()).formulas_digest);
        // A formula behind leading spaces is still a formula: adding one changes it
        let e = "id,note\n1,=1+1\n2,=HYPERLINK(\"http://x\",\"y\")\n3, =HYPERLINK(\"http://evil\",\"y\")\n";
        assert_ne!(text.formulas_digest, run(e, &CsvOptions::default()).formulas_digest);
    }

    #[test]
    fn count_records_reads_quoted_newlines_as_one_record() {
        let dir = std::env::temp_dir().join(format!("vg-count-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("n.csv");
        std::fs::write(&path, "a;b\n1;\"two\nlines\"\n3;4\n").unwrap();
        let (n, header) = count_records(&path, &CsvOptions::default()).unwrap();
        assert_eq!(n, 3);
        assert_eq!(header, vec!["a", "b"]);
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn column_rules_by_name_and_letter() {
        let csv = "zip,sku,amount,when,junk\n00501,007,12,2024-01-05,x\n02134,0123,abc,05/01/2024,y\n";
        let options = CsvOptions {
            columns: vec![
                ("sku".into(), ColumnRule::Number),
                ("C".into(), ColumnRule::Number),
                ("When".into(), ColumnRule::Date(DateOrder::Dmy)),
                ("junk".into(), ColumnRule::Skip),
            ],
            ..Default::default()
        };
        let r = run(csv, &options);
        assert_eq!(cell(&r, 1, 0), ("00501".into(), true), "zip stays Auto: text");
        assert_eq!(cell(&r, 1, 1), ("7".into(), false), "sku chosen Number: 007 -> 7");
        assert_eq!(cell(&r, 2, 2), ("abc".into(), true), "not a number: kept as text");
        assert_eq!(r.columns[2].unreadable, 1);
        assert_eq!(r.sheet.get_formatted_display(1, 3), "2024-01-05", "ISO is accepted in any order");
        assert_eq!(r.sheet.get_formatted_display(2, 3), "2024-01-05", "05/01/2024 read day-first");
        assert!(!cell(&r, 2, 3).1, "a date is a number underneath");
        assert_eq!(r.columns[4].resolved, Resolved::Skipped);
        assert!(r.sheet.get_display(1, 4).is_empty(), "the skipped column is gone, nothing moved into column E");
        assert_eq!(r.sheet.get_display(0, 3), "when", "headers move left with their columns");
    }

    #[test]
    fn auto_reports_why_and_what_number_would_change() {
        let csv = "zip,amount,day\n00501,12,2024-01-05\n02134,3.5,2024-01-06\n";
        let r = run(csv, &CsvOptions::default());
        let zip = &r.columns[0];
        assert_eq!(zip.resolved, Resolved::Text);
        assert_eq!(zip.reason, Some("leading zeros"));
        assert_eq!(zip.changed_if_number, 2);
        assert_eq!(r.columns[1].resolved, Resolved::Number);
        assert_eq!(r.columns[1].changed_if_number, 0);
        let day = &r.columns[2];
        assert!(day.looks_like_iso_dates, "offered, never applied");
        assert_eq!(cell(&r, 1, 2), ("2024-01-05".into(), true));
        assert_eq!(r.summary_line(), "comma · UTF-8 · 2 text columns");
    }

    #[test]
    fn decimal_comma() {
        let csv = "a;b\n1.234,56;1,5\n-2.000;0,25\n";
        let options = CsvOptions { decimal_comma: true, ..Default::default() };
        let r = run(csv, &options);
        assert_eq!(r.delimiter, b';');
        let n = |row, col| match r.sheet.get_computed_value(row, col) {
            Value::Number(n) => n,
            other => panic!("({row},{col}) {other:?}"),
        };
        assert_eq!(n(1, 0), 1234.56);
        assert_eq!(n(1, 1), 1.5);
        assert_eq!(n(2, 0), -2000.0);
        assert_eq!(n(2, 1), 0.25);
        // Without it, 1,5 is text rather than 15.
        let r = run(csv, &CsvOptions::default());
        assert_eq!(cell(&r, 1, 1), ("1,5".into(), true));
    }

    #[test]
    fn no_header_types_the_first_row() {
        let r = run("007,12\n", &CsvOptions { no_header: true, ..Default::default() });
        assert_eq!(cell(&r, 0, 1), ("12".into(), false));
        assert_eq!(cell(&r, 0, 0), ("007".into(), true));
        assert_eq!(r.columns[0].name, "A");
    }

    #[test]
    fn decode_detects_and_strips_byte_order_marks() {
        let (s, e) = decode(b"\xEF\xBB\xBFzip,name\n", None);
        assert_eq!((s.as_str(), e), ("zip,name\n", Encoding::Utf8));
        let utf16: Vec<u8> = [0xFF, 0xFE].into_iter().chain("é,1\n".encode_utf16().flat_map(|u| u.to_le_bytes())).collect();
        assert_eq!(decode(&utf16, None), ("é,1\n".to_string(), Encoding::Utf16Le));
        assert_eq!(decode(b"caf\xE9\n", None), ("café\n".to_string(), Encoding::Windows1252));
        assert_eq!(decode("café".as_bytes(), Some(Encoding::Windows1252)).0, "cafÃ©");
    }

    #[test]
    fn preview_reads_only_what_it_needs() {
        let mut csv = String::from("n\n");
        for i in 0..500 {
            csv.push_str(&format!("{i}\n"));
        }
        let r = import_str(&csv, Encoding::Utf8, &CsvOptions::default(), 20).unwrap();
        assert_eq!(r.rows_in_file, 20);
        assert_eq!(r.sample.len(), 20);
    }

    #[test]
    fn cli_flags_round_trip_the_options() {
        let options = CsvOptions {
            delimiter: Some(b';'),
            decimal_comma: true,
            evaluate_formulas: true,
            columns: vec![
                ("zip".into(), ColumnRule::Text),
                ("amount".into(), ColumnRule::Number),
                ("order date".into(), ColumnRule::Date(DateOrder::Dmy)),
                ("notes".into(), ColumnRule::Skip),
            ],
            ..Default::default()
        };
        assert_eq!(
            options.cli_flags().join(" "),
            "--text zip --number amount --date 'order date'=dmy --skip notes --delimiter ';' --decimal-comma --formulas"
        );
    }
}
