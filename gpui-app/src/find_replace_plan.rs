//! Search uses displayed order; replacement plans retain canonical identities
//! and original sources so stale offsets can never edit a different value.
use crate::table_edit::TableCellWrite;
use std::{
    collections::{BTreeMap, BTreeSet},
    sync::Arc,
};
use visigrid_engine::{
    cell::ValueRef,
    filter::RowView,
    sheet::{Sheet, SheetId},
    workbook::Workbook,
};

const MAX_MATCHES: usize = 100_000;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MatchKind {
    Text,
    Formula,
}

#[derive(Clone, Debug)]
pub struct MatchHit {
    pub sheet: usize,
    pub sheet_id: SheetId,
    /// Canonical worksheet coordinate, never a displayed slot.
    pub row: usize,
    pub col: usize,
    /// Display-only numeric/error results cannot be replaced.
    pub kind: Option<MatchKind>,
    pub start: usize,
    pub end: usize,
    source: Arc<str>,
}

pub(super) fn visible_row(
    rows: &RowView,
    hidden: Option<&BTreeSet<usize>>,
    row: usize,
) -> Option<usize> {
    if hidden.is_some_and(|h| h.contains(&row)) {
        None
    } else {
        rows.data_to_view(row)
    }
}

/// Every folded character carries both boundaries in the original UTF-8.
/// This also handles lowercase expansions such as capital dotted I.
fn folded(text: &str, numeric: bool) -> Vec<(char, usize, usize)> {
    text.char_indices()
        .filter(|(_, c)| !numeric || !matches!(c, ',' | '$' | '%'))
        .flat_map(|(start, c)| {
            c.to_lowercase()
                .map(move |lc| (lc, start, start + c.len_utf8()))
        })
        .collect()
}

fn reference_query(query: &str) -> bool {
    use visigrid_engine::formula::parser::{parse, Expr};
    matches!(
        parse(&format!("={query}")),
        Ok(Expr::CellRef { .. } | Expr::Range { .. })
    )
}

/// Build exclusions once per formula, avoiding a prefix scan for every match.
fn formula_exclusions(source: &str, reference: bool) -> Vec<bool> {
    let mut excluded = vec![false; source.len() + 1];
    if reference {
        for (start, end, _) in visigrid_engine::formula::structured::source_references(source) {
            if source[start..end].contains('[') {
                excluded[start..end].fill(true);
            }
        }
    }
    let mut quote = None;
    let mut chars = source.char_indices().peekable();
    while let Some((i, c)) = chars.next() {
        if excluded[i] {
            continue;
        } // Do not interpret escaped header names as quotes.
        excluded[i] = quote == Some('"') || (reference && quote.is_some());
        if let Some(q) = quote {
            if c == q {
                if chars.peek().is_some_and(|(_, next)| *next == q) {
                    let (next, _) = chars.next().unwrap();
                    excluded[next] = excluded[i];
                } else {
                    quote = None;
                }
            }
        } else if matches!(c, '"' | '\'') {
            quote = Some(c);
        }
    }
    excluded
}

fn spans(source: &str, query: &str, formula: bool, reference: bool) -> Vec<(usize, usize)> {
    let numeric = query.chars().any(|c| c.is_ascii_digit()) && !(formula && reference);
    let text = folded(source, numeric);
    let needle: Vec<_> = folded(query, numeric)
        .into_iter()
        .map(|(c, _, _)| c)
        .collect();
    if needle.is_empty() || needle.len() > text.len() {
        return Vec::new();
    }
    let excluded = formula.then(|| formula_exclusions(source, reference));
    // Linear matching matters while typing into a large sheet, especially for
    // repetitive text and long queries. Keep source spans separate from folding.
    let mut prefix = vec![0; needle.len()];
    let mut matched = 0;
    for i in 1..needle.len() {
        while matched > 0 && needle[i] != needle[matched] {
            matched = prefix[matched - 1];
        }
        if needle[i] == needle[matched] {
            matched += 1;
        }
        prefix[i] = matched;
    }
    let mut found = Vec::new();
    matched = 0;
    for (i, &(ch, _, _)) in text.iter().enumerate() {
        while matched > 0 && ch != needle[matched] {
            matched = prefix[matched - 1];
        }
        if ch == needle[matched] {
            matched += 1;
        }
        if matched != needle.len() {
            continue;
        }
        let (start, end) = (text[i + 1 - needle.len()].1, text[i].2);
        let word = |c: char| c.is_alphanumeric() || matches!(c, '_' | '$' | '.');
        let boundary = !formula
            || !reference
            || (!source[..start].chars().next_back().is_some_and(word)
                && !source[end..]
                    .chars()
                    .next()
                    .is_some_and(|c| word(c) || c == '!'));
        if boundary
            && excluded.as_ref().is_none_or(|positions| !positions[start])
            && found
                .last()
                .is_none_or(|&(_, previous_end)| start >= previous_end)
        {
            found.push((start, end));
            if found.len() > MAX_MATCHES {
                break;
            }
        }
        matched = 0; // Matches replace non-overlapping occurrences.
    }
    found
}

pub(super) fn search(
    sheet: &Sheet,
    index: usize,
    rows: &RowView,
    hidden: Option<&BTreeSet<usize>>,
    hidden_cols: Option<&BTreeSet<usize>>,
    query: &str,
) -> Result<Vec<MatchHit>, String> {
    if query.is_empty() {
        return Ok(Vec::new());
    }
    let reference = reference_query(query);
    let mut hits = Vec::new();
    for ((row, col), cell) in sheet.cells_iter() {
        if visible_row(rows, hidden, row).is_none() || hidden_cols.is_some_and(|h| h.contains(&col))
            || sheet.is_merge_hidden(row, col) {
            continue;
        }
        let (kind, source): (_, Arc<str>) = match cell.value() {
            ValueRef::Empty => continue,
            ValueRef::Text(text) => (Some(MatchKind::Text), Arc::from(text)),
            ValueRef::Formula { source, .. } => (Some(MatchKind::Formula), Arc::from(source)),
            _ => (None, Arc::from(sheet.get_display(row, col))),
        };
        for (start, end) in spans(&source, query, kind == Some(MatchKind::Formula), reference) {
            if hits.len() == MAX_MATCHES {
                return Err(
                    "More than 100,000 matches. Use a more specific search before replacing."
                        .into(),
                );
            }
            hits.push(MatchHit {
                sheet: index,
                sheet_id: sheet.id,
                row,
                col,
                kind,
                start,
                end,
                source: source.clone(),
            });
        }
    }
    hits.sort_by_key(|h| (rows.data_to_view_unchecked(h.row), h.col, h.start));
    Ok(hits)
}

/// All writes are built and checked before the shared guarded edit path runs.
/// Text remains text (including leading zeros and literal formula-looking text).
pub(super) fn replacements(
    wb: &Workbook,
    index: usize,
    rows: &RowView,
    hidden: Option<&BTreeSet<usize>>,
    hidden_cols: Option<&BTreeSet<usize>>,
    hits: &[MatchHit],
    replacement: &str,
) -> Result<(Vec<TableCellWrite>, usize), String> {
    wb.ensure_writable()?;
    let sheet = wb
        .sheet(index)
        .ok_or("The search sheet no longer exists.")?;
    let mut groups = BTreeMap::<_, Vec<_>>::new();
    for hit in hits.iter().filter(|h| h.kind.is_some()) {
        if hit.sheet != index
            || hit.sheet_id != sheet.id
            || visible_row(rows, hidden, hit.row).is_none()
            || hidden_cols.is_some_and(|h| h.contains(&hit.col))
        {
            return Err("Search results changed or are hidden. Search again before replacing. Nothing was changed.".into());
        }
        let cell = sheet.get_cell(hit.row, hit.col);
        let current = match cell.value() {
            ValueRef::Text(t) if hit.kind == Some(MatchKind::Text) => t,
            ValueRef::Formula { source, .. } if hit.kind == Some(MatchKind::Formula) => source,
            _ => return Err("A search result changed type. Search again before replacing.".into()),
        };
        if current != hit.source.as_ref()
            || !current.is_char_boundary(hit.start)
            || !current.is_char_boundary(hit.end)
            || hit.start >= hit.end
            || hit.end > current.len()
        {
            return Err(
                "A search result changed. Search again before replacing. Nothing was changed."
                    .into(),
            );
        }
        if sheet.table_header_at(hit.row, hit.col).is_some() {
            return Err("Replacement includes a Table header. Edit that header directly to rename its column. Nothing was changed.".into());
        }
        groups.entry((hit.row, hit.col)).or_default().push(hit);
    }
    let mut writes = Vec::new();
    let mut count = 0;
    for ((row, col), mut hits) in groups {
        hits.sort_by_key(|h| std::cmp::Reverse(h.start));
        let original = &hits[0].source;
        let mut value = original.to_string();
        let mut last_start = original.len();
        for hit in &hits {
            if hit.end > last_start {
                return Err("Search results overlap. Search again before replacing.".into());
            }
            value.replace_range(hit.start..hit.end, replacement);
            last_start = hit.start;
        }
        if value == original.as_ref() {
            continue;
        }
        let literal = hits[0].kind == Some(MatchKind::Text) && !value.is_empty();
        let mut write = TableCellWrite::value(row, col, value);
        write.literal_text = literal;
        writes.push(write);
        count += hits.len();
    }
    Ok((writes, count))
}
