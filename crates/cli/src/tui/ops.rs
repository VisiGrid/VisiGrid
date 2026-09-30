//! Data operations behind the viewer's keys: sort order, search, frequency
//! and pivot tabs. Pure functions over `PeekData`, so they are testable
//! without a terminal.
//!
//! Every operation works on the rows peek loaded. When the preview is
//! truncated, derived tabs say so in their name rather than presenting a
//! partial answer as the file's.

use std::cmp::Ordering;
use std::collections::HashMap;

use visigrid_engine::workbook::Workbook;

use super::data::PeekData;
use crate::util;

/// Sort key: numbers (ascending by value), then text (case-insensitive),
/// with blanks always last whichever direction is chosen.
#[derive(PartialEq, PartialOrd)]
enum Key {
    Number(f64),
    Text(String),
}

fn sort_key(s: &str) -> Option<Key> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    let percent = t.ends_with('%');
    let numeric: String = t.chars().filter(|c| !matches!(c, ',' | '$' | '€' | '£' | '%' | ' ')).collect();
    let numeric = numeric.strip_prefix('(').and_then(|n| n.strip_suffix(')')).map(|n| format!("-{n}")).unwrap_or(numeric);
    match numeric.parse::<f64>() {
        Ok(n) if n.is_finite() => Some(Key::Number(if percent { n / 100.0 } else { n })),
        _ => Some(Key::Text(t.to_lowercase())),
    }
}

fn compare(a: &Option<Key>, b: &Option<Key>, descending: bool) -> Ordering {
    match (a, b) {
        (None, None) => Ordering::Equal,
        (None, _) => Ordering::Greater,
        (_, None) => Ordering::Less,
        (Some(x), Some(y)) => {
            let ord = match (x, y) {
                (Key::Number(p), Key::Number(q)) => p.partial_cmp(q).unwrap_or(Ordering::Equal),
                (Key::Number(_), Key::Text(_)) => Ordering::Less,
                (Key::Text(_), Key::Number(_)) => Ordering::Greater,
                (Key::Text(p), Key::Text(q)) => p.cmp(q),
            };
            if descending { ord.reverse() } else { ord }
        }
    }
}

/// Row order sorted by one column. Stable, so sorting by a second column
/// after a first keeps the first as the tiebreak.
pub fn sorted_order(data: &PeekData, order: &[usize], col: usize, descending: bool) -> Vec<usize> {
    let mut keyed: Vec<(usize, Option<Key>)> = order
        .iter()
        .map(|&r| (r, data.rows.get(r).and_then(|row| row.get(col)).and_then(|s| sort_key(s))))
        .collect();
    keyed.sort_by(|a, b| compare(&a.1, &b.1, descending));
    keyed.into_iter().map(|(r, _)| r).collect()
}

/// Does this cell contain the (already lowercased) query?
pub fn cell_matches(value: &str, query_lower: &str) -> bool {
    !query_lower.is_empty() && value.to_lowercase().contains(query_lower)
}

/// Next matching cell in row-major view order, strictly after (or before,
/// when `backward`) the given position, wrapping around. Returns the view
/// position (index into `order`) and column.
pub fn find(
    data: &PeekData,
    order: &[usize],
    from: (usize, usize),
    query: &str,
    backward: bool,
) -> Option<(usize, usize)> {
    let cols = data.num_cols;
    let total = order.len() * cols;
    if total == 0 || query.is_empty() {
        return None;
    }
    let q = query.to_lowercase();
    let start = from.0 * cols + from.1;
    for step in 1..=total {
        let i = if backward { (start + total - step % total) % total } else { (start + step) % total };
        let (pos, col) = (i / cols, i % cols);
        let value = data.rows.get(order[pos]).and_then(|row| row.get(col)).map(|s| s.as_str()).unwrap_or("");
        if cell_matches(value, &q) {
            return Some((pos, col));
        }
    }
    None
}

/// Count of matching cells, for the status line.
pub fn count_matches(data: &PeekData, query: &str) -> usize {
    let q = query.to_lowercase();
    data.rows.iter().flat_map(|row| row.iter()).filter(|v| cell_matches(v, &q)).count()
}

fn derived(col_names: Vec<String>, rows: Vec<Vec<String>>) -> PeekData {
    let num_cols = col_names.len();
    let col_widths = PeekData::compute_widths(&col_names, &rows, num_cols, 0);
    PeekData {
        num_rows: rows.len(),
        num_cols,
        rows,
        json_rows: None,
        raw: None,
        col_widths,
        col_names,
        has_headers: true,
        first_data_file_row: 1,
        total_rows: None,
        delimiter: 0,
    }
}

/// Distinct values of one column with their counts, most frequent first.
/// Values are compared exactly as displayed; blanks count as "(blank)".
pub fn frequency(data: &PeekData, col: usize) -> PeekData {
    let mut counts: HashMap<&str, usize> = HashMap::new();
    let mut first_seen: Vec<&str> = Vec::new();
    for row in &data.rows {
        let v = row.get(col).map(|s| s.as_str()).unwrap_or("");
        let entry = counts.entry(v).or_insert(0);
        if *entry == 0 {
            first_seen.push(v);
        }
        *entry += 1;
    }
    let total = data.rows.len().max(1) as f64;
    // Keys computed once: a high-cardinality column has one item per row.
    let mut items: Vec<(&str, usize, Option<Key>)> =
        first_seen.into_iter().map(|v| (v, counts[v], sort_key(v))).collect();
    items.sort_by(|a, b| b.1.cmp(&a.1).then_with(|| compare(&a.2, &b.2, false)));
    let name = data.col_names.get(col).cloned().unwrap_or_else(|| util::col_to_letter(col));
    let rows = items
        .into_iter()
        .map(|(v, n, _)| {
            vec![
                if v.trim().is_empty() { "(blank)".to_string() } else { v.to_string() },
                n.to_string(),
                format!("{:.1}%", n as f64 * 100.0 / total),
            ]
        })
        .collect();
    derived(vec![name, "count".into(), "percent".into()], rows)
}

/// Parse `rows=Region,Rep column=Month values=sum:Amount`. Keys may be
/// abbreviated (r, c/col, v); names may contain spaces ("rows=Order Date").
pub fn parse_pivot_spec(spec: &str) -> Result<(Vec<String>, Option<String>, Vec<String>), String> {
    let mut parts: Vec<(String, String)> = Vec::new();
    for token in spec.split_whitespace() {
        match token.split_once('=') {
            Some((k, v)) if !k.is_empty() && k.chars().all(|c| c.is_ascii_alphabetic()) => {
                parts.push((k.to_ascii_lowercase(), v.to_string()))
            }
            _ => match parts.last_mut() {
                Some((_, v)) => {
                    v.push(' ');
                    v.push_str(token);
                }
                None => return Err(format!("expected key=value, got \"{token}\" (keys: rows, column, values)")),
            },
        }
    }
    let (mut rows, mut column, mut values) = (Vec::new(), None, Vec::new());
    for (k, v) in parts {
        match k.as_str() {
            "rows" | "row" | "r" => rows.push(v),
            "column" | "col" | "c" => column = Some(v.trim().to_string()).filter(|s| !s.is_empty()),
            "values" | "value" | "v" => values.push(v),
            other => return Err(format!("unknown key \"{other}\" (keys: rows, column, values)")),
        }
    }
    Ok((rows, column, values))
}

/// Compute a pivot of the loaded rows with the engine the desktop uses.
/// Fields resolve exactly as `vgrid pivot` and the session op resolve them.
pub fn pivot(data: &PeekData, spec: &str) -> Result<(PeekData, String), String> {
    let (rows, column, values) = parse_pivot_spec(spec)?;
    // A pivot needs field names. Without a header row (CSV opened without
    // --headers), the first loaded row is the header row.
    let (names, body): (&[String], &[Vec<String>]) = match (data.has_headers, data.rows.split_first()) {
        (false, Some((first, rest))) => (first.as_slice(), rest),
        _ => (&data.col_names, &data.rows),
    };
    let last = format!("{}{}", util::col_to_letter(data.num_cols.saturating_sub(1)), body.len() + 1);
    let args = crate::pivot::PivotArgs { rows, column, values, sheet: None, range: Some(format!("A1:{last}")) };
    let mut wb = Workbook::new();
    {
        let sheet = wb.sheet_mut(0).ok_or("no sheet")?;
        for (c, name) in names.iter().enumerate() {
            sheet.set_value(0, c, name);
        }
        for (r, row) in body.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                // Loaded cells are values: "=1+1" in a CSV is text, not a formula.
                if v.starts_with('=') {
                    sheet.set_text(r + 1, c, v);
                } else if !v.is_empty() {
                    sheet.set_value(r + 1, c, v);
                }
            }
        }
    }
    let op = crate::pivot::create_op(&args, Some(0));
    let (source, definition) = visigrid_session_host::resolve_create_pivot(&op, &wb).map_err(|(_, m)| m)?;
    let label = match (&definition.values[..], &definition.column) {
        ([v], Some(c)) => format!("{} by {}", v.header(), c.header),
        ([v], None) => v.header(),
        (_, Some(c)) => format!("pivot by {}", c.header),
        _ => "pivot".to_string(),
    };
    let header_rows = if definition.column.is_some() { 2 } else { 1 };
    let (id, idx) = wb.create_pivot(source, definition)?;
    let (h, w) = wb.find_pivot(id).and_then(|(_, t)| t.extent).map(|(h, w)| (h as usize, w as usize)).unwrap_or((0, 0));
    let out = wb.sheet(idx).ok_or("pivot output missing")?;
    let cell = |r: usize, c: usize| out.get_formatted_display(r, c);
    let names: Vec<String> = (0..w).map(|c| cell(header_rows - 1, c)).collect();
    let body: Vec<Vec<String>> = (header_rows..h).map(|r| (0..w).map(|c| cell(r, c)).collect()).collect();
    Ok((derived(names, body), label))
}

/// A number to 15 significant digits, as Excel displays it, so binary
/// floating-point noise (120.25000000000001) does not show.
fn significant(n: f64) -> String {
    if !n.is_finite() || n == 0.0 {
        return if n == 0.0 { "0".into() } else { n.to_string() };
    }
    let magnitude = n.abs().log10().floor() as i32;
    if !(-6..15).contains(&magnitude) {
        return format!("{:e}", n);
    }
    let decimals = (14 - magnitude).max(0) as usize;
    let s = format!("{:.*}", decimals, n);
    let s = if s.contains('.') { s.trim_end_matches('0').trim_end_matches('.').to_string() } else { s };
    if s == "-0" { "0".into() } else { s }
}

/// Split a computed-column entry into (name, formula). `Tax =D2*0.08` names
/// the column "Tax"; `=D2*0.08` names it after the formula.
pub fn parse_column_spec(input: &str) -> Result<(Option<String>, String), String> {
    let input = input.trim();
    let eq = input.find('=').ok_or("start the formula with = (e.g. =D2*1.08, or Tax =D2*0.08)")?;
    let name = input[..eq].trim();
    let formula = input[eq..].trim().to_string();
    if formula.len() < 2 {
        return Err("the formula is empty".into());
    }
    Ok(((!name.is_empty()).then(|| name.to_string()), formula))
}

/// Replace `[Header]` with the header's column letter at `row` (1-based), so
/// `=[Amount]*1.08` can be typed without looking up letters.
fn expand_header_refs(formula: &str, names: &[String], row: usize) -> Result<String, String> {
    let mut out = String::with_capacity(formula.len());
    let mut rest = formula;
    let mut in_string = false;
    while let Some(i) = rest.find(|c| c == '[' || c == '"') {
        out.push_str(&rest[..i]);
        if rest.as_bytes()[i] == b'"' {
            in_string = !in_string;
            out.push('"');
            rest = &rest[i + 1..];
            continue;
        }
        if in_string {
            out.push('[');
            rest = &rest[i + 1..];
            continue;
        }
        let end = rest[i..].find(']').ok_or("unclosed [ in the formula")? + i;
        let want = rest[i + 1..end].trim();
        let col = names
            .iter()
            .position(|n| n.trim().eq_ignore_ascii_case(want))
            .ok_or_else(|| format!("no column headed \"{want}\" (columns: {})", names.join(", ")))?;
        out.push_str(&format!("{}{}", util::col_to_letter(col), row));
        rest = &rest[end + 1..];
    }
    out.push_str(rest);
    Ok(out)
}

/// Compute a new column from an Excel formula written for the first data
/// row, filled down like Excel: relative references move with each row,
/// `$`-anchored ones stay. Row numbers are the ones in the gutter (file row
/// numbers), so with a header row the first data row is row 2. Evaluated by
/// the desktop's engine over the loaded rows, in file order. Returns the
/// column name and one display value (and filled-down formula) per data row.
pub fn computed_column(data: &PeekData, input: &str) -> Result<(String, Vec<String>, Vec<String>), String> {
    use visigrid_engine::formula::parser::{adjust_formula_refs, parse};
    let (name, formula) = parse_column_spec(input)?;
    let base = data.first_data_file_row; // gutter number of data row 0
    let formula = expand_header_refs(&formula, &data.col_names, base)?;
    parse(&formula).map_err(|e| format!("formula: {e}"))?;
    let out_col = data.num_cols;
    let mut wb = Workbook::new();
    let mut formulas = Vec::with_capacity(data.rows.len());
    {
        let sheet = wb.sheet_mut(0).ok_or("no sheet")?;
        for (i, row) in data.rows.iter().enumerate() {
            let r = base - 1 + i;
            for (c, v) in row.iter().enumerate() {
                if !v.is_empty() {
                    sheet.set_value_deferred(r, c, v);
                }
            }
            let f = adjust_formula_refs(&formula, i as i32, 0);
            sheet.set_value_deferred(r, out_col, &f);
            formulas.push(f);
        }
    }
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let sheet = wb.sheet(0).ok_or("no sheet")?;
    // Full precision, as `vgrid calc` prints it: General display would round
    // 0.1234 to 0.12, which hides what a new column is for checking.
    let values = (0..data.rows.len())
        .map(|i| match sheet.get_computed_value(base - 1 + i, out_col) {
            visigrid_engine::formula::eval::Value::Number(n) => significant(n),
            v => v.to_text(),
        })
        .collect();
    Ok((name.unwrap_or_else(|| formula.clone()), values, formulas))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table(rows: &[&[&str]]) -> PeekData {
        let names = vec!["Region".to_string(), "Amount".to_string()];
        derived(names, rows.iter().map(|r| r.iter().map(|s| s.to_string()).collect()).collect())
    }

    #[test]
    fn sort_is_numeric_aware_stable_and_keeps_blanks_last() {
        let d = table(&[&["b", "10"], &["a", ""], &["c", "9"], &["d", "1,200"], &["e", "x"], &["f", "(5)"]]);
        let id: Vec<usize> = (0..6).collect();
        assert_eq!(sorted_order(&d, &id, 1, false), vec![5, 2, 0, 3, 4, 1]);
        assert_eq!(sorted_order(&d, &id, 1, true), vec![4, 3, 0, 2, 5, 1]);
        // Percent is a fraction; euro and pound amounts are numbers.
        let d = table(&[&["a", "5%"], &["b", "0.5"], &["c", "€10"], &["d", "€9"], &["e", "£1"]]);
        assert_eq!(sorted_order(&d, &[0, 1, 2, 3, 4], 1, false), vec![0, 1, 4, 3, 2]);
    }

    #[test]
    fn search_wraps_both_ways_in_view_order() {
        let d = table(&[&["West", "1"], &["East", "2"], &["west", "3"]]);
        let order = vec![2, 1, 0];
        assert_eq!(find(&d, &order, (0, 0), "WEST", false), Some((2, 0)));
        assert_eq!(find(&d, &order, (2, 0), "west", false), Some((0, 0)));
        assert_eq!(find(&d, &order, (0, 0), "west", true), Some((2, 0)));
        assert_eq!(find(&d, &order, (0, 0), "north", false), None);
        assert_eq!(count_matches(&d, "west"), 2);
    }

    #[test]
    fn frequency_counts_and_orders() {
        let d = table(&[&["West", "1"], &["East", "2"], &["West", "3"], &["", "4"]]);
        let f = frequency(&d, 0);
        assert_eq!(f.col_names, vec!["Region", "count", "percent"]);
        assert_eq!(f.rows[0], vec!["West", "2", "50.0%"]);
        assert_eq!(f.rows[2][0], "(blank)");
    }

    #[test]
    fn pivot_tab_uses_the_engine_and_named_fields() {
        let d = table(&[&["West", "10"], &["East", "5"], &["West", "2.5"]]);
        let (p, label) = pivot(&d, "rows=region values=sum:amount").unwrap();
        assert_eq!(label, "Sum of Amount");
        assert_eq!(p.col_names, vec!["Region", "Sum of Amount"]);
        assert_eq!(p.rows, vec![vec!["East", "5.00"], vec!["West", "12.50"], vec!["Grand Total", "17.50"]]);
        let err = |spec: &str| pivot(&d, spec).err().unwrap_or_default();
        assert!(err("rows=Month").contains("no column headed \"Month\""));
        assert!(err("Region").contains("expected key=value"));

        // No header row loaded: the first row supplies the field names.
        let mut raw = table(&[&["Region", "Amount"], &["West", "10"], &["West", "1"]]);
        raw.has_headers = false;
        raw.col_names = vec!["A".into(), "B".into()];
        let (p, _) = pivot(&raw, "rows=Region values=Amount").unwrap();
        assert_eq!(p.rows[0], vec!["West", "11"]);

        // A loaded "=1+1" is text: counted, never evaluated into a number.
        let f = table(&[&["=1+1", "3"], &["West", "4"]]);
        let (p, _) = pivot(&f, "rows=Region values=sum:Amount").unwrap();
        assert_eq!(p.rows[0], vec!["=1+1", "3"]);
    }

    #[test]
    fn pivot_spec_allows_spaces_and_abbreviations() {
        let (r, c, v) = parse_pivot_spec("r=Order Date c=Ship Mode v=sum:Unit Price,count:Order ID").unwrap();
        assert_eq!(r, vec!["Order Date"]);
        assert_eq!(c.as_deref(), Some("Ship Mode"));
        assert_eq!(v, vec!["sum:Unit Price,count:Order ID"]);
    }

    #[test]
    fn computed_columns_fill_down_like_excel() {
        let mut d = table(&[&["West", "10"], &["East", "5"], &["West", "2.5"]]);
        // A CSV with a header row: the first data row is file row 2.
        d.first_data_file_row = 2;
        let (name, v, f) = computed_column(&d, "Tax =B2*0.1").unwrap();
        assert_eq!(name, "Tax");
        assert_eq!(v, vec!["1", "0.5", "0.25"]);
        assert_eq!(f[2], "=B4*0.1");
        // Header names, anchors, text functions, errors.
        let (name, v, _) = computed_column(&d, "=[amount]/$B$3").unwrap();
        assert_eq!(name, "=B2/$B$3");
        assert_eq!(v, vec!["2", "1", "0.5"]);
        let (_, v, _) = computed_column(&d, "=IF([Amount]>4,UPPER([Region]),\"[small]\")").unwrap();
        assert_eq!(v, vec!["WEST", "EAST", "[small]"]);
        let (_, v, _) = computed_column(&d, "=B2*0.1+0.2").unwrap();
        assert_eq!(v[0], "1.2", "no float noise");
        assert_eq!(significant(1200.5 * 0.1 + 0.2), "120.25");
        assert_eq!(significant(-1.0 / 3.0), "-0.333333333333333");
        assert_eq!(significant(1e20), "1e20");
        let (_, v, _) = computed_column(&d, "=B2/0").unwrap();
        assert_eq!(v[0], "#DIV/0!");
        // Refusals.
        assert!(computed_column(&d, "B2*2").unwrap_err().contains("start the formula with ="));
        assert!(computed_column(&d, "=[Price]*2").unwrap_err().contains("no column headed \"Price\""));
        assert!(computed_column(&d, "=B2*(").unwrap_err().starts_with("formula:"));

        // Without a header row the first data row is row 1.
        let mut raw = d;
        raw.has_headers = false;
        raw.first_data_file_row = 1;
        let (_, v, _) = computed_column(&raw, "=B1*2").unwrap();
        assert_eq!(v, vec!["20", "10", "5"]);
    }
}
