//! Table-aware formula editing. All offsets here are UTF-8 bytes; analyzer
//! offsets and syntax-color spans are converted explicitly at the boundary.
use crate::formula_context::{self, TokenType};
use std::ops::Range;
use visigrid_engine::{
    formula::{
        parser::Expr,
        structured::{self, TableSection},
    },
    sheet::SheetId,
    workbook::Workbook,
};

#[derive(Debug, Clone)]
pub struct Suggestion {
    pub label: String,
    pub detail: String,
    pub replacement: String,
    pub range: Range<usize>,
    pub opens_columns: bool,
}

fn byte(source: &str, ch: usize) -> usize {
    formula_context::char_index_to_byte_offset(source, ch)
}

/// Tokenization excludes strings and quoted sheet names, including incomplete
/// literals. A selector must contain the caret; the next token is never replaced.
pub fn suggestions(
    source: &str,
    cursor: usize,
    wb: &Workbook,
    home: SheetId,
    cell: (usize, usize),
) -> Vec<Suggestion> {
    if !source.starts_with('=') && !source.starts_with('+') {
        return vec![];
    }
    if !source.is_char_boundary(cursor) {
        return vec![];
    }
    let tokens = formula_context::tokenize_for_highlight(source);
    let token = tokens.iter().find(|(r, kind)| {
        byte(source, r.start) < cursor
            && cursor <= byte(source, r.end)
            && matches!(
                kind,
                TokenType::StructuredRef | TokenType::NamedRange | TokenType::CellRef
            )
    });
    let Some((r, kind)) = token else {
        return vec![];
    };
    let mut range = byte(source, r.start)..byte(source, r.end);
    if *kind == TokenType::StructuredRef && cursor < range.end {
        let token = &source[range.clone()];
        if token
            .find('[')
            .is_some_and(|i| structured::bracket_len(&token[i..]).is_err())
        {
            // An unfinished selector can make the lexer swallow the rest of a
            // formula. Never replace its trailing expression along with a name.
            if source[cursor..].starts_with([')', '+', '-', '*', '/', '&', ',', '=', '<', '>']) {
                range.end = cursor;
            } else {
                return vec![];
            }
        }
    }
    // Sheet!Table and external workbook references are not supported.
    if source[..range.start].trim_end().ends_with('!') {
        return vec![];
    }
    let text = &source[range.clone()];
    let caret = cursor - range.start;
    if *kind != TokenType::StructuredRef {
        let prefix = text[..caret].to_ascii_lowercase();
        return wb
            .tables()
            .filter(|(_, t)| t.name.to_ascii_lowercase().starts_with(&prefix))
            .map(|(sid, t)| Suggestion {
                label: t.name.clone(),
                detail: format!("Table · {}", wb.sheet_by_id(sid).unwrap().name),
                replacement: format!("{}[", t.name),
                range: range.clone(),
                opens_columns: true,
            })
            .collect();
    }
    let Some(open) = text.find('[') else {
        return vec![];
    };
    if caret <= open {
        return vec![];
    }
    let qualified = &text[..open];
    let target = if qualified.is_empty() {
        wb.sheet_by_id(home)
            .and_then(|s| s.table_at(cell.0, cell.1))
            .map(|t| (home, t))
    } else {
        wb.table_by_name(qualified)
    };
    let Some((sid, table)) = target else {
        return vec![];
    };
    let mut depth = 0usize;
    let mut atom_start = open + 1;
    let mut iter = text[open..caret].char_indices().peekable();
    while let Some((i, ch)) = iter.next() {
        let at = open + i;
        if ch == '\'' {
            iter.next();
            continue;
        }
        match ch {
            '[' => {
                depth += 1;
                atom_start = at + 1;
            }
            ']' => {
                depth = depth.saturating_sub(1);
                atom_start = at + 1;
            }
            '@' if at == open + 1 => atom_start = at + 1,
            _ => {}
        }
    }
    if depth == 0 || atom_start > caret {
        return vec![];
    }
    let raw_prefix = &text[atom_start..caret];
    if raw_prefix.contains(['[', ']']) && !raw_prefix.contains('\'') {
        return vec![];
    }
    let mut prefix = String::new();
    let mut chars = raw_prefix.chars();
    while let Some(ch) = chars.next() {
        if ch == '\'' {
            if let Some(ch) = chars.next() {
                prefix.push(ch);
            }
        } else {
            prefix.push(ch);
        }
    }
    let prefix = prefix.to_lowercase();
    let mut atom_end = text.len();
    let mut chars = text[caret..].char_indices();
    while let Some((i, ch)) = chars.next() {
        if ch == '\'' {
            chars.next();
            continue;
        }
        if ch == ']' {
            atom_end = caret + i;
            break;
        }
    }
    let this_row_available =
        sid == home && cell.0 > table.range.start_row && cell.0 <= table.range.end_row;
    let mut names: Vec<(String, String, bool)> = table
        .columns
        .iter()
        .map(|c| (c.name.clone(), format!("Column · {}", table.name), false))
        .collect();
    if atom_start == open + 1 || text[..atom_start].ends_with("[[") {
        for label in ["#Data", "#Headers", "#All", "#This Row"] {
            if label != "#This Row" || this_row_available {
                names.push((label.into(), "Table selector".into(), true));
            }
        }
    }
    names
        .into_iter()
        .filter(|(name, _, selector)| {
            name.to_lowercase().starts_with(&prefix)
                && (!raw_prefix.starts_with('#') || *selector)
                && (!raw_prefix.starts_with('\'') || !*selector)
        })
        .filter_map(|(label, detail, selector)| {
            let atom = if selector {
                label.clone()
            } else {
                structured::escape_header(&label)
            };
            let mut replacement = format!("{}{}{}", &text[..atom_start], atom, &text[atom_end..]);
            // Supply only the brackets still missing from a partial token.
            let mut balance = 0usize;
            let mut cs = replacement[open..].chars();
            while let Some(ch) = cs.next() {
                if ch == '\'' {
                    cs.next();
                    continue;
                }
                match ch {
                    '[' => balance += 1,
                    ']' => balance = balance.checked_sub(1)?,
                    _ => {}
                }
            }
            replacement.push_str(&"]".repeat(balance));
            let reference = structured::parse(
                (!qualified.is_empty()).then(|| qualified.to_string()),
                &replacement[open..],
            )
            .ok()?;
            if reference.section == TableSection::ThisRow && !this_row_available {
                return None;
            }
            Some(Suggestion {
                label,
                detail,
                replacement: reference.format(),
                range: range.clone(),
                opens_columns: false,
            })
        })
        .collect()
}

#[derive(Debug)]
pub struct ResolvedReference {
    pub sheet: SheetId,
    pub start: (usize, usize),
    pub end: (usize, usize),
    pub span: Range<usize>,
}

/// Resolution shares the engine's exact Table semantics; missing fields and
/// empty bodies produce no fabricated rectangle.
pub fn references(
    source: &str,
    wb: &Workbook,
    home: SheetId,
    cell: (usize, usize),
) -> Vec<ResolvedReference> {
    structured::source_references(source)
        .into_iter()
        .filter_map(|(start, end, reference)| {
            if source[..start].trim_end().ends_with('!') {
                return None;
            }
            let (sid, t) = match &reference.table {
                Some(name) => wb.table_by_name(name)?,
                None => (home, wb.sheet_by_id(home)?.table_at(cell.0, cell.1)?),
            };
            let (first, last) =
                match structured::resolve_region(t, sid, home, Some(cell), &reference) {
                    Expr::CellRef { row, col, .. } => ((row, col), (row, col)),
                    Expr::Range {
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        ..
                    } => ((start_row, start_col), (end_row, end_col)),
                    _ => return None,
                };
            Some(ResolvedReference {
                sheet: sid,
                start: first,
                end: last,
                span: source[..start].chars().count()..source[..end].chars().count(),
            })
        })
        .collect()
}

/// Visible runs of a canonical reference, bounded to the rendered viewport.
/// Hidden records never appear and sorting cannot retarget a this-row highlight.
pub fn projected_runs(
    view: &visigrid_engine::filter::RowView,
    data: Range<usize>,
    rows: impl Iterator<Item = usize>,
) -> Vec<(usize, usize)> {
    let mut out = Vec::new();
    let mut start = None;
    let mut last = 0;
    for row in rows {
        if data.contains(&view.view_to_data(row)) {
            start.get_or_insert(row);
            last = row;
        } else if let Some(first) = start.take() {
            out.push((first, last));
        }
    }
    if let Some(first) = start {
        out.push((first, last));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use visigrid_engine::{sheet::Sheet, table::TableRange};
    fn book() -> Workbook {
        let mut wb =
            Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(1), 100, 20, "Data")], 0);
        for (i, h) in [
            "Qty",
            "Amount",
            "Unit Price",
            "Tax [net]",
            "#Data",
            "@Qty",
            "O'Brien",
            "金額",
            "A1",
        ]
        .iter()
        .enumerate()
        {
            wb.set_cell_value_tracked(0, 2, i + 1, h);
        }
        wb.create_table(
            SheetId(1),
            TableRange {
                start_row: 2,
                start_col: 1,
                end_row: 6,
                end_col: 9,
            },
            "Sales",
        )
        .unwrap();
        wb.add_sheet_named("Summary").unwrap();
        wb
    }
    fn pick(
        marked: &str,
        label: &str,
        wb: &Workbook,
        home: SheetId,
        cell: (usize, usize),
    ) -> (String, usize) {
        let cursor = marked.find('|').unwrap();
        let mut source = marked.replace('|', "");
        let entries = suggestions(&source, cursor, wb, home, cell);
        let entry = entries
            .iter()
            .find(|s| s.label == label)
            .unwrap_or_else(|| panic!("Missing {label} in {marked}: {entries:?}"));
        source.replace_range(entry.range.clone(), &entry.replacement);
        (source, entry.range.start + entry.replacement.len())
    }
    #[test]
    fn table_names_open_columns_without_function_parentheses() {
        let wb = book();
        let (s, c) = pick("=SUM(sa|)", "Sales", &wb, SheetId(2), (0, 0));
        assert_eq!(s, "=SUM(Sales[)");
        assert_eq!(c, 11);
        let entries = suggestions(&s, c, &wb, SheetId(2), (0, 0));
        assert!(entries.iter().any(|s| s.label == "Amount"));
        let (s, _) = pick("=SUM(Sales[|)+A1", "Amount", &wb, SheetId(2), (0, 0));
        assert_eq!(s, "=SUM(Sales[Amount])+A1");
    }
    #[test]
    fn columns_this_row_selectors_and_spans_preserve_surrounding_formula() {
        let wb = book();
        for (source, label, expected) in [
            (
                "=SUM(Sales[Am|ount])+A1",
                "Amount",
                "=SUM(Sales[Amount])+A1",
            ),
            ("=[@Un|", "Unit Price", "=[@[Unit Price]]"),
            ("=Sales[@[Q|", "Qty", "=Sales[@[Qty]]"),
            (
                "=SUM(Sales[[#Headers],[Am|",
                "Amount",
                "=SUM(Sales[[#Headers],[Amount]]",
            ),
            (
                "=SUM(Sales[[Q|ty]:[Amount]])",
                "Qty",
                "=SUM(Sales[[Qty]:[Amount]])",
            ),
            (
                "=SUM(Sales[[Qty]:[Am|",
                "Amount",
                "=SUM(Sales[[Qty]:[Amount]]",
            ),
            ("=Sales[#H|", "#Headers", "=Sales[#Headers]"),
        ] {
            assert_eq!(pick(source, label, &wb, SheetId(1), (4, 2)).0, expected);
        }
    }
    #[test]
    fn escaped_and_unicode_headers_are_inserted_as_valid_references() {
        let wb = book();
        for label in ["Tax [net]", "#Data", "@Qty", "O'Brien", "金額", "A1"] {
            let (s, c) = pick("=\"é\"&[@|", label, &wb, SheetId(1), (4, 2));
            assert_eq!(c, s.len());
            assert!(visigrid_engine::formula::parser::parse(&s).is_ok(), "{s}");
            let refs = references(&s, &wb, SheetId(1), (4, 2));
            assert_eq!(refs.len(), 1, "{s}");
            assert_eq!(refs[0].start.0, 4);
            assert_eq!(refs[0].span.start, 5); // Char offset, not UTF-8 byte offset.
        }
    }
    #[test]
    fn strings_unknown_tables_and_invalid_row_contexts_do_not_complete() {
        let wb = book();
        for (source, home, cell) in [
            ("=\"Sales[", SheetId(1), (4, 2)),
            ("='Sales[", SheetId(1), (4, 2)),
            ("=Missing[", SheetId(1), (4, 2)),
            ("=[@", SheetId(2), (4, 2)),
            ("=Sales[@", SheetId(2), (4, 2)),
            ("=Sales[@", SheetId(1), (2, 2)),
            ("=Sheet1!Sales[", SheetId(1), (4, 2)),
            ("=Sales[Qty]", SheetId(1), (4, 2)),
        ] {
            assert!(
                suggestions(source, source.len(), &wb, home, cell).is_empty(),
                "{source}"
            );
        }
    }
    #[test]
    fn highlights_resolve_exact_sections_and_home_sheet_context() {
        let wb = book();
        let refs = references(
            "=SUM(Sales[Amount])+SUM(Sales[#Headers])+SUM(Sales[#All])",
            &wb,
            SheetId(2),
            (0, 0),
        );
        assert_eq!(refs.len(), 3);
        assert_eq!(refs[0].sheet, SheetId(1));
        assert_eq!((refs[0].start, refs[0].end), ((3, 2), (6, 2)));
        assert_eq!((refs[1].start, refs[1].end), ((2, 1), (2, 9)));
        assert_eq!((refs[2].start, refs[2].end), ((2, 1), (6, 9)));
        assert_eq!(
            references("=[@Qty]", &wb, SheetId(1), (5, 2))[0].start,
            (5, 1)
        );
        assert!(references("=Sales[@Qty]", &wb, SheetId(2), (5, 2)).is_empty());
        assert!(references(
            "=SUM(Sales[Missing])+\"Sales[Amount]\"",
            &wb,
            SheetId(2),
            (0, 0)
        )
        .is_empty());
    }
    #[test]
    fn empty_tables_have_no_body_highlight_but_keep_header_suggestions() {
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "Qty");
        wb.create_table(
            SheetId(1),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 0,
                end_col: 0,
            },
            "EmptyData",
        )
        .unwrap();
        assert!(references("=SUM(EmptyData)", &wb, SheetId(1), (3, 0)).is_empty());
        assert_eq!(
            references("=EmptyData[#Headers]", &wb, SheetId(1), (3, 0)).len(),
            1
        );
        assert!(suggestions("=EmptyData[", 11, &wb, SheetId(1), (3, 0))
            .iter()
            .any(|s| s.label == "Qty"));
    }
    #[test]
    fn projected_highlights_follow_records_through_sort_and_filter() {
        use visigrid_engine::filter::RowView;
        let mut view = RowView::new(8);
        view.restore(
            vec![0, 1, 2, 6, 4, 3, 5, 7],
            vec![true, true, true, true, false, true, true, true],
        );
        assert_eq!(
            projected_runs(&view, 3..7, view.visible_rows().iter().copied()),
            vec![(3, 6)]
        );
        assert_eq!(
            projected_runs(&view, 3..4, view.visible_rows().iter().copied()),
            vec![(5, 5)]
        );
        assert!(projected_runs(&view, 4..5, view.visible_rows().iter().copied()).is_empty());
        assert!(
            projected_runs(&view, 3..4, view.visible_rows().iter().copied().take(3)).is_empty()
        );
    }
    #[test]
    fn accepted_formula_evaluates_and_suggestions_follow_schema_changes() {
        let mut wb = book();
        wb.set_cell_value_tracked(0, 4, 1, "3");
        wb.set_cell_value_tracked(0, 4, 3, "10");
        let first = pick("=[@Q|", "Qty", &wb, SheetId(1), (4, 2)).0;
        let formula = pick(
            &format!("{first}*[@Un|"),
            "Unit Price",
            &wb,
            SheetId(1),
            (4, 2),
        )
        .0;
        wb.set_cell_value_tracked(0, 4, 2, &formula);
        assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "30");
        let id = wb.table_by_name("Sales").unwrap().1.id;
        let mut names: Vec<_> = wb
            .table(id)
            .unwrap()
            .1
            .columns
            .iter()
            .map(|c| c.name.clone())
            .collect();
        names[0] = "Count".into();
        wb.rename_table_columns(id, &names).unwrap();
        wb.rename_table(id, "Orders").unwrap();
        assert!(suggestions("=Sales[", 7, &wb, SheetId(1), (4, 2)).is_empty());
        assert_eq!(
            pick("=Orders[@Co|", "Count", &wb, SheetId(1), (4, 2)).0,
            "=Orders[@[Count]]"
        );
        assert_eq!(wb.sheet(0).unwrap().get_display(4, 2), "30");
    }

    #[test]
    fn escaped_selector_names_are_columns_and_bare_table_names_are_not_a1_refs() {
        let mut wb = book();
        let entries = suggestions("=Sales['#", 9, &wb, SheetId(1), (4, 2));
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].replacement, "Sales['#Data]");
        let id = wb.table_by_name("Sales").unwrap().1.id;
        wb.rename_table(id, "Table1").unwrap();
        let refs = crate::app::Spreadsheet::parse_table_formula_refs(
            "=SUM(Table1)",
            &wb,
            SheetId(2),
            (0, 0),
        );
        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].start, (3, 1));
        assert_eq!(refs[0].end, Some((6, 9)));
        let refs = crate::app::Spreadsheet::parse_table_formula_refs(
            "=\"é\"&SUM(Table1[Amount])",
            &wb,
            SheetId(2),
            (0, 0),
        );
        assert_eq!(refs[0].text_char_range, 9..23);
    }
}
