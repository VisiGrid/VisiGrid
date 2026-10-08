//! Structured-reference grammar and lossless source spans. Apostrophe escapes
//! are decoded while reading atoms, never by splitting on bracket characters.
use super::eval::CellLookup;
use super::parser::{BoundExpr, Expr};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableSection {
    Data,
    Headers,
    Totals,
    All,
    ThisRow,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StructuredReference {
    pub table: Option<String>,
    pub section: TableSection,
    pub columns: Option<(String, String)>,
}

impl StructuredReference {
    pub fn body(table: String) -> Self {
        Self {
            table: Some(table),
            section: TableSection::Data,
            columns: None,
        }
    }
    pub fn format(&self) -> String {
        let prefix = self.table.as_deref().unwrap_or("");
        let column = |name: &str| format!("[{}]", escape_header(name));
        let columns = self.columns.as_ref().map(|(a, b)| {
            if a == b {
                column(a)
            } else {
                format!("{}:{}", column(a), column(b))
            }
        });
        let suffix = match (self.section, columns) {
            (TableSection::Data, Some(c)) if self.columns.as_ref().is_some_and(|(a, b)| a == b) => {
                c
            }
            (TableSection::Data, Some(c)) => format!("[{c}]"),
            (TableSection::ThisRow, Some(c)) => format!("[@{c}]"),
            (TableSection::ThisRow, None) => "[@]".into(),
            (section, cols) => {
                let label = match section {
                    TableSection::Data => "#Data",
                    TableSection::Headers => "#Headers",
                    TableSection::Totals => "#Totals",
                    TableSection::All => "#All",
                    TableSection::ThisRow => unreachable!(),
                };
                match cols {
                    Some(c) => format!("[[{label}],{c}]"),
                    None => format!("[{label}]"),
                }
            }
        };
        format!("{prefix}{suffix}")
    }
}

pub fn escape_header(name: &str) -> String {
    let mut out = String::new();
    for ch in name.chars() {
        if matches!(ch, '[' | ']' | '#' | '@' | '\'') {
            out.push('\'');
        }
        out.push(ch);
    }
    out
}

/// Byte length of one balanced bracket expression. Shared by tokenization,
/// rename and copy/fill, so escapes cannot disagree between those paths.
pub fn bracket_len(input: &str) -> Result<usize, String> {
    if !input.starts_with('[') {
        return Err("Expected structured reference bracket".into());
    }
    let mut depth = 0usize;
    let mut chars = input.char_indices();
    while let Some((i, ch)) = chars.next() {
        match ch {
            '\'' => {
                chars
                    .next()
                    .ok_or("Incomplete structured-reference escape")?;
            }
            '[' => {
                depth += 1;
                if depth > 3 {
                    return Err("Unsupported nested structured reference".into());
                }
            }
            ']' => {
                depth -= 1;
                if depth == 0 {
                    return Ok(i + 1);
                }
            }
            _ => {}
        }
    }
    Err("Unclosed structured reference".into())
}

fn atom(input: &str) -> Result<(String, bool), String> {
    if input.is_empty() {
        return Err("Empty structured column name".into());
    }
    let escaped_first = input.starts_with('\'');
    let mut out = String::new();
    let mut chars = input.chars();
    while let Some(ch) = chars.next() {
        if ch == '\'' {
            let escaped = chars
                .next()
                .ok_or("Incomplete structured-reference escape")?;
            if !matches!(escaped, '[' | ']' | '#' | '@' | '\'') {
                return Err("Unsupported structured-reference escape".into());
            }
            out.push(escaped);
        } else {
            if matches!(ch, '[' | ']') {
                return Err("Escape brackets in column names with an apostrophe".into());
            }
            out.push(ch);
        }
    }
    if out.trim() != out || out.is_empty() {
        return Err("Invalid structured column name".into());
    }
    Ok((out, escaped_first))
}

pub fn parse(table: Option<String>, input: &str) -> Result<StructuredReference, String> {
    if bracket_len(input)? != input.len() {
        return Err("Unexpected structured-reference suffix".into());
    }
    let inner = input[1..input.len() - 1].trim();
    let mut section = TableSection::Data;
    let mut rest = inner;
    if let Some(s) = rest.strip_prefix('@') {
        section = TableSection::ThisRow;
        rest = s;
    }
    let mut parts = Vec::new();
    let mut separators = Vec::new();
    if rest.starts_with('[') {
        while rest.starts_with('[') {
            let n = bracket_len(rest)?;
            parts.push(atom(&rest[1..n - 1])?);
            rest = rest[n..].trim();
            if rest.is_empty() {
                break;
            }
            let sep = rest.chars().next().unwrap();
            if !matches!(sep, ',' | ':') {
                return Err("Expected comma or colon in structured reference".into());
            }
            separators.push(sep);
            rest = rest[1..].trim_start();
            if !rest.starts_with('[') {
                return Err("Expected bracketed column name".into());
            }
        }
    } else if !rest.is_empty() {
        parts.push(atom(rest)?);
        rest = "";
    }
    if !rest.is_empty() {
        return Err("Unsupported structured reference".into());
    }
    if let Some((label, false)) = parts.first() {
        if label.starts_with('#') {
            if section == TableSection::ThisRow {
                return Err("Cannot combine row and section selectors".into());
            }
            section = match label.to_ascii_lowercase().as_str() {
                "#data" => TableSection::Data,
                "#headers" => TableSection::Headers,
                "#totals" => TableSection::Totals,
                "#all" => TableSection::All,
                "#this row" => TableSection::ThisRow,
                _ => return Err(format!("Unsupported table selector: {label}")),
            };
            parts.remove(0);
            if !parts.is_empty() && separators.remove(0) != ',' {
                return Err("Expected comma after section selector".into());
            }
        }
    }
    if parts
        .iter()
        .any(|(s, escaped)| !escaped && (s.starts_with('#') || s.starts_with('@')))
    {
        return Err("Escape a leading # or @ in a column name".into());
    }
    let columns = match parts.as_slice() {
        [] if section != TableSection::Data || inner.eq_ignore_ascii_case("#data") => None,
        [(a, _)] if separators.is_empty() => Some((a.clone(), a.clone())),
        [(a, _), (b, _)] if separators == [':'] => Some((a.clone(), b.clone())),
        _ => return Err("Unsupported structured column/section union".into()),
    };
    Ok(StructuredReference {
        table,
        section,
        columns,
    })
}

/// Resolve a single argument, retaining ordinary ASTs. Used before functions
/// consume references, and recursively for dependency extraction.
pub fn resolve<L: CellLookup>(expr: &BoundExpr, lookup: &L) -> BoundExpr {
    match expr {
        Expr::StructuredRef(reference) => {
            lookup.resolve_table_reference(reference, lookup.current_cell())
        }
        Expr::NamedRange(name) if lookup.is_table_name(name) => lookup.resolve_table_reference(
            &StructuredReference::body(name.clone()),
            lookup.current_cell(),
        ),
        Expr::NamedRange(name) => lookup.resolve_named_reference(name).unwrap_or_else(|| expr.clone()),
        _ => expr.clone(),
    }
}

pub fn resolve_tree<L: CellLookup>(expr: &BoundExpr, lookup: &L) -> BoundExpr {
    match expr {
        Expr::Function { name, args } => Expr::Function {
            name: name.clone(),
            args: args.iter().map(|a| resolve_tree(a, lookup)).collect(),
        },
        Expr::BinaryOp { op, left, right } => Expr::BinaryOp {
            op: *op,
            left: Box::new(resolve_tree(left, lookup)),
            right: Box::new(resolve_tree(right, lookup)),
        },
        _ => resolve(expr, lookup),
    }
}

/// Potential table references with byte offsets, excluding strings, quoted
/// sheet names and function/sheet prefixes. Bare names are filtered by registry.
pub fn source_references(source: &str) -> Vec<(usize, usize, StructuredReference)> {
    let mut out = Vec::new();
    let mut i = 0;
    let bytes = source.as_bytes();
    while i < bytes.len() {
        // Exponents are part of numeric literals, not a reference to a Table
        // named E (for example 1E+2). Match the formula lexer's number grammar.
        if bytes[i].is_ascii_digit() || bytes[i] == b'.' {
            while i < bytes.len() && (bytes[i].is_ascii_digit() || bytes[i] == b'.') {
                i += 1;
            }
            if matches!(bytes.get(i), Some(b'e' | b'E')) {
                let mut end = i + 1;
                if matches!(bytes.get(end), Some(b'+' | b'-')) {
                    end += 1;
                }
                if bytes.get(end).is_some_and(u8::is_ascii_digit) {
                    i = end + 1;
                    while bytes.get(i).is_some_and(u8::is_ascii_digit) {
                        i += 1;
                    }
                }
            }
            continue;
        }
        if matches!(bytes[i], b'"' | b'\'') {
            let quote = bytes[i];
            i += 1;
            while i < bytes.len() {
                if bytes[i] == quote {
                    i += 1;
                    if bytes.get(i) == Some(&quote) {
                        i += 1;
                    } else {
                        break;
                    }
                } else {
                    i += 1;
                }
            }
            continue;
        }
        let start = i;
        let mut table = None;
        if bytes[i].is_ascii_alphabetic() || bytes[i] == b'_' {
            i += 1;
            while i < bytes.len()
                && (bytes[i].is_ascii_alphanumeric() || matches!(bytes[i], b'_' | b'.'))
            {
                i += 1;
            }
            table = Some(source[start..i].to_string());
        }
        if bytes.get(i) == Some(&b'[') {
            if let Ok(n) = bracket_len(&source[i..]) {
                if let Ok(reference) = parse(table, &source[i..i + n]) {
                    out.push((start, i + n, reference));
                }
                i += n;
                continue;
            }
        } else if let Some(name) = table {
            let next = source[i..].trim_start().chars().next();
            if !matches!(next, Some('!' | '(')) {
                out.push((start, i, StructuredReference::body(name)));
            }
            continue;
        }
        i += source[i..].chars().next().map_or(1, char::len_utf8);
    }
    out
}

/// Resolve through stable column IDs/positions in the current schema. The
/// resulting rectangle is ephemeral; the stored formula stays symbolic.
pub fn resolve_region(
    table: &crate::table::DataTable,
    sheet: crate::sheet::SheetId,
    context_sheet: crate::sheet::SheetId,
    cell: Option<(usize, usize)>,
    reference: &StructuredReference,
) -> BoundExpr {
    use crate::sheet::SheetRef;
    let (c0, c1) = match &reference.columns {
        None => (table.range.start_col, table.range.end_col),
        Some((a, b)) => {
            let first = table
                .column_by_name(a)
                .and_then(|column| table.columns.iter().position(|c| c.id == column.id));
            let last = table
                .column_by_name(b)
                .and_then(|column| table.columns.iter().position(|c| c.id == column.id));
            let (Some(first), Some(last)) = (first, last) else {
                return Expr::ReferenceError("#NAME? Unknown table column".into());
            };
            if first > last {
                return Expr::ReferenceError("#REF! Reversed table column span".into());
            }
            (table.range.start_col + first, table.range.start_col + last)
        }
    };
    let (r0, r1) = match reference.section {
        TableSection::Headers => (table.range.start_row, table.range.start_row),
        TableSection::All => (table.range.start_row, table.full_range().end_row),
        TableSection::Totals => match table.totals_row() {
            Some(row) => (row, row),
            None => return Expr::ReferenceError("#REF! Table has no totals row".into()),
        },
        TableSection::Data => {
            if table.range.data_rows() == 0 {
                return Expr::EmptyRange {
                    columns: c1 - c0 + 1,
                };
            }
            (table.range.start_row + 1, table.range.end_row)
        }
        TableSection::ThisRow => {
            let Some((row, _)) = cell else {
                return Expr::ReferenceError(
                    "#VALUE! This-row reference requires a cell context".into(),
                );
            };
            if sheet != context_sheet || row <= table.range.start_row || row > table.range.end_row {
                return Expr::ReferenceError(
                    "#VALUE! This-row reference is outside the table body".into(),
                );
            }
            (row, row)
        }
    };
    if reference.section == TableSection::ThisRow && c0 == c1 {
        Expr::CellRef {
            sheet: SheetRef::Id(sheet),
            row: r0,
            col: c0,
            row_abs: true,
            col_abs: true,
        }
    } else {
        Expr::Range {
            sheet: SheetRef::Id(sheet),
            start_row: r0,
            start_col: c0,
            end_row: r1,
            end_col: c1,
            start_row_abs: true,
            start_col_abs: true,
            end_row_abs: true,
            end_col_abs: true,
        }
    }
}
