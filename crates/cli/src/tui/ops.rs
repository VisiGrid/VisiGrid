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
    let numeric: String = t.chars().filter(|c| !matches!(c, ',' | '$' | '%' | ' ')).collect();
    let numeric = numeric.strip_prefix('(').and_then(|n| n.strip_suffix(')')).map(|n| format!("-{n}")).unwrap_or(numeric);
    match numeric.parse::<f64>() {
        Ok(n) if n.is_finite() => Some(Key::Number(n)),
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
    let mut items: Vec<(&str, usize)> = first_seen.into_iter().map(|v| (v, counts[v])).collect();
    items.sort_by(|a, b| b.1.cmp(&a.1).then_with(|| compare(&sort_key(a.0), &sort_key(b.0), false)));
    let name = data.col_names.get(col).cloned().unwrap_or_else(|| util::col_to_letter(col));
    let rows = items
        .into_iter()
        .map(|(v, n)| {
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
                if !v.is_empty() {
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
    }

    #[test]
    fn pivot_spec_allows_spaces_and_abbreviations() {
        let (r, c, v) = parse_pivot_spec("r=Order Date c=Ship Mode v=sum:Unit Price,count:Order ID").unwrap();
        assert_eq!(r, vec!["Order Date"]);
        assert_eq!(c.as_deref(), Some("Ship Mode"));
        assert_eq!(v, vec!["sum:Unit Price,count:Order ID"]);
    }
}
