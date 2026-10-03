//! Conservative OOXML checkbox-filter interchange. Unsupported predicates are
//! refused as a complete set, never installed as a misleading partial filter.
use super::{
    xlsx::{ExportResult, ImportResult},
    xlsx_tables::{attr, range},
};
use quick_xml::{
    events::{BytesEnd, BytesStart, Event},
    Reader, Writer,
};
use std::collections::{BTreeSet, HashMap, HashSet};
use visigrid_engine::{
    filter::{ColumnFilter, FilterKey, NormalizedFilterKey},
    sheet::Sheet,
    table::{DataTable, TableId},
    table_view::{TableFilter, TableViewSpec},
    workbook::Workbook,
};

pub(super) fn boolean(e: &BytesStart<'_>, key: &[u8], default: bool) -> Result<bool, String> {
    match attr(e, key)?.as_deref() {
        None => Ok(default),
        Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        _ => Err("Invalid Table boolean attribute".into()),
    }
}

// ST_Xstring escapes are distinct from XML entity escapes. Decode UTF-16
// units once; an escaped underscore must not trigger a second decoding pass.
fn escaped_unit(bytes: &[u8]) -> Option<u16> {
    if bytes.len() < 7
        || bytes[0] != b'_'
        || bytes[1] != b'x'
        || bytes[6] != b'_'
        || !bytes[2..6].iter().all(u8::is_ascii_hexdigit)
    {
        return None;
    }
    u16::from_str_radix(std::str::from_utf8(&bytes[2..6]).ok()?, 16).ok()
}
fn decode_value(value: &str) -> Result<String, String> {
    let mut units = Vec::new();
    let mut rest = value;
    while !rest.is_empty() {
        if let Some(unit) = escaped_unit(rest.as_bytes()) {
            units.push(unit);
            rest = &rest[7..];
        } else {
            let ch = rest.chars().next().unwrap();
            units.extend_from_slice(ch.encode_utf16(&mut [0; 2]));
            rest = &rest[ch.len_utf8()..];
        }
    }
    String::from_utf16(&units).map_err(|_| "Invalid UTF-16 filter value".into())
}
fn encode_value(value: &str) -> String {
    let mut out = String::new();
    for (i, ch) in value.char_indices() {
        if escaped_unit(&value.as_bytes()[i..]).is_some() {
            out.push_str("_x005F_");
        } else {
            out.push(ch);
        }
    }
    out
}

#[derive(Default)]
pub(super) struct ImportedView {
    filters: Vec<(usize, Values)>,
    buttons: bool,
    sorted: bool,
    mixed_buttons: bool,
}
#[derive(Default)]
struct Values {
    blank: bool,
    values: BTreeSet<String>,
}
pub(super) struct PendingView {
    pub sheet: usize,
    pub table: TableId,
    pub view: Result<ImportedView, String>,
}

pub(super) fn parse(xml: &str, table: &DataTable) -> Result<ImportedView, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().expand_empty_elements = true;
    let mut path = Vec::<String>::new();
    let mut view = ImportedView::default();
    let mut seen_auto = false;
    let mut hidden = HashSet::new();
    let mut columns = HashSet::new();
    let mut column = None;
    let mut values = None;
    let mut count = 0;
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) => {
                let tag = String::from_utf8_lossy(e.local_name().as_ref()).into_owned();
                let parent = match path.as_slice() {
                    [root] if root == "table" => Some("table"),
                    [root, auto] if root == "table" && auto == "autoFilter" => Some("autoFilter"),
                    [root, auto, column]
                        if root == "table" && auto == "autoFilter" && column == "filterColumn" =>
                    {
                        Some("filterColumn")
                    }
                    [root, auto, column, list]
                        if root == "table"
                            && auto == "autoFilter"
                            && column == "filterColumn"
                            && list == "filters" =>
                    {
                        Some("filters")
                    }
                    _ => None,
                };
                match (parent, tag.as_str()) {
                    (Some("table"), "autoFilter") => {
                        if seen_auto
                            || range(&attr(&e, b"ref")?.ok_or("Missing filter range")?)?
                                != table.range
                        {
                            return Err("Invalid or duplicate Table filter range".into());
                        }
                        seen_auto = true;
                        view.buttons = true;
                    }
                    (Some("autoFilter"), "filterColumn") => {
                        let c = attr(&e, b"colId")?
                            .ok_or("Missing filter column")?
                            .parse::<usize>()
                            .map_err(|_| "Invalid filter column")?;
                        if c >= table.columns.len() || !columns.insert(c) {
                            return Err("Invalid or duplicate filter column".into());
                        }
                        if boolean(&e, b"hiddenButton", false)?
                            || !boolean(&e, b"showButton", true)?
                        {
                            hidden.insert(c);
                        }
                        column = Some(c);
                        values = None;
                    }
                    (Some("filterColumn"), "filters") => {
                        if values.is_some() {
                            return Err("Duplicate value-filter list".into());
                        }
                        values = Some(Values {
                            blank: boolean(&e, b"blank", false)?,
                            ..Default::default()
                        });
                    }
                    (Some("filters"), "filter") => {
                        count += 1;
                        if count > 10_000 {
                            return Err("More than 10,000 selected filter values".into());
                        }
                        let value =
                            decode_value(&attr(&e, b"val")?.ok_or("Missing filter value")?)?;
                        raw_value(&FilterKey::Text(value.clone()))?;
                        values
                            .as_mut()
                            .ok_or("Missing filter list")?
                            .values
                            .insert(value);
                    }
                    (Some("table" | "autoFilter"), "sortState") => view.sorted = true,
                    (Some("filterColumn" | "filters" | "autoFilter"), _) => {
                        return Err(format!("Unsupported Excel filter element: {tag}"));
                    }
                    _ => {}
                }
                path.push(tag);
            }
            Event::End(e) => {
                if e.local_name().as_ref() == b"filterColumn"
                    && path.len() == 3
                    && path[1] == "autoFilter"
                {
                    if let Some(v) = values.take() {
                        if !v.blank && v.values.is_empty() {
                            return Err("Empty Excel filter list".into());
                        }
                        view.filters
                            .push((column.take().ok_or("Missing filter column")?, v));
                    }
                }
                path.pop();
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !hidden.is_empty() {
        view.mixed_buttons = hidden.len() != table.columns.len();
        view.buttons = view.mixed_buttons;
    }
    Ok(view)
}

// Raw numeric serials (not formatted dates/currency) are used in OOXML filter
// values. Infer types from existing records; retain absent values as text or
// numbers so clearing a row does not silently clear the saved selection.
fn selected_keys(
    sheet: &Sheet,
    table: &DataTable,
    offset: usize,
    values: &Values,
) -> Result<HashSet<NormalizedFilterKey>, String> {
    let mut selected = HashSet::new();
    let mut matches_by_key = HashMap::new();
    let needles: HashSet<_> = values.values.iter().map(|s| s.to_lowercase()).collect();
    let numbers: HashSet<_> = needles
        .iter()
        .filter_map(|s| s.parse::<f64>().ok())
        .filter(|n| n.is_finite())
        .map(|n| FilterKey::Number(n.into()).normalized())
        .collect();
    let mut remaining = needles.clone();
    if values.blank {
        selected.insert(NormalizedFilterKey::Blank);
    }
    for row in table.range.start_row + 1..=table.range.end_row {
        let key =
            FilterKey::from_value(&sheet.get_computed_value(row, table.range.start_col + offset));
        if matches!(key, FilterKey::Blank) {
            continue;
        }
        let text = raw_value(&key)?;
        let normalized = key.normalized();
        let passes = needles.contains(&text.to_lowercase())
            || (matches!(key, FilterKey::Number(_)) && numbers.contains(&normalized));
        if matches_by_key
            .insert(normalized.clone(), passes)
            .is_some_and(|old| old != passes)
        {
            return Err("Excel distinguishes text values that VisiGrid groups together".into());
        }
        if passes {
            remaining.remove(&text.to_lowercase());
            selected.insert(normalized);
        }
    }
    for text in remaining {
        let key = if let Ok(n) = text.parse::<f64>() {
            if !n.is_finite() {
                return Err("Non-finite filter value".into());
            }
            FilterKey::Number(n.into())
        } else {
            FilterKey::Text(text)
        };
        selected.insert(key.normalized());
    }
    Ok(selected)
}

pub(super) fn finish_import(
    pending: Vec<PendingView>,
    wb: &mut Workbook,
    result: &mut ImportResult,
) {
    let mut owners = HashMap::<usize, usize>::new();
    for p in &pending {
        if p.view
            .as_ref()
            .is_ok_and(|v| !v.filters.is_empty() || !v.buttons)
        {
            *owners.entry(p.sheet).or_default() += 1;
        }
    }
    for p in pending {
        let Some((sid, table)) = wb.table(p.table) else {
            continue;
        };
        let table = table.clone();
        let name = table.name.clone();
        let range = table.range;
        let previously_hidden: Vec<_> = result
            .imported_layouts
            .get(p.sheet)
            .map(|layout| {
                layout
                    .hidden_rows
                    .iter()
                    .copied()
                    .filter(|r| *r > range.start_row && *r <= range.end_row)
                    .collect()
            })
            .unwrap_or_default();
        let has_filters = p.view.as_ref().map_or(true, |v| !v.filters.is_empty());
        // Excel hides whole rows for filters. Those bits must not become
        // manual hidden rows, which would survive Clear Filters in VisiGrid.
        if has_filters {
            if let Some(layout) = result.imported_layouts.get_mut(p.sheet) {
                layout
                    .hidden_rows
                    .retain(|r| *r <= range.start_row || *r > range.end_row);
            }
        }
        let installed: Result<(), String> = (|| {
            let view = p.view?;
            if view.sorted {
                result.warnings.push(format!("Table {name}: saved Excel sorting was not imported. Records retain their physically stored order."));
            }
            if view.mixed_buttons {
                result.warnings.push(format!("Table {name}: mixed per-column filter-button visibility is unsupported; all buttons are shown."));
            }
            if view.filters.is_empty() && view.buttons {
                return Ok(());
            }
            if owners.get(&p.sheet).copied().unwrap_or(0) > 1 {
                return Err(
                    "More than one Table on this sheet needs saved filter/button settings".into(),
                );
            }
            let sheet = wb.sheet_by_id(sid).unwrap();
            if !view.filters.is_empty() {
                if result.imported_layouts.get(p.sheet).is_some_and(|layout| {
                    layout
                        .row_heights
                        .keys()
                        .any(|r| *r > range.start_row && *r <= range.end_row)
                        || (layout.frozen_rows > range.start_row + 1
                            && layout.frozen_rows <= range.end_row)
                }) {
                    return Err("Table body has custom row heights or a freeze boundary".into());
                }
            }
            let mut spec = TableViewSpec::new(p.table);
            spec.show_filter_buttons = view.buttons;
            for (offset, values) in view.filters {
                spec.filters.push(TableFilter {
                    column: table.columns[offset].id,
                    criteria: ColumnFilter {
                        selected: Some(selected_keys(sheet, &table, offset, &values)?),
                        text_filter: None,
                    },
                });
            }
            let reveals_manual_rows = !spec.filters.is_empty()
                && previously_hidden.iter().any(|row| {
                    spec.filters.iter().all(|filter| {
                        let offset = table
                            .columns
                            .iter()
                            .position(|c| c.id == filter.column)
                            .unwrap();
                        filter.criteria.passes(&FilterKey::from_value(
                            &sheet.get_computed_value(*row, range.start_col + offset),
                        ))
                    })
                });
            wb.set_table_view_spec(sid, Some(spec))?;
            if reveals_manual_rows {
                result.warnings.push(format!("Table {name}: manually hidden or stale hidden body rows were made visible; saved filter criteria now control visibility."));
            }
            Ok(())
        })();
        if let Err(reason) = installed {
            result.warnings.push(format!("Table {name}: saved filters/button settings were not imported ({reason}). All records are shown in their stored order; hidden body rows were made visible."));
        }
    }
}

fn raw_value(key: &FilterKey) -> Result<String, String> {
    use visigrid_engine::filter::ErrorKind::*;
    Ok(match key {
        FilterKey::Blank => return Err("Blank values use the blank filter flag".into()),
        FilterKey::Number(n) if n.is_finite() => n.to_string(),
        FilterKey::Number(_) => return Err("Non-finite filter number".into()),
        FilterKey::Bool(b) => if *b { "TRUE" } else { "FALSE" }.into(),
        FilterKey::Text(s) if !s.is_empty() && !s.chars().any(char::is_control) => s.clone(),
        FilterKey::Text(_) => {
            return Err("Empty text or control characters in a filter value".into())
        }
        FilterKey::Error(e) => match e {
            Ref => "#REF!",
            Value => "#VALUE!",
            Div0 => "#DIV/0!",
            Name => "#NAME?",
            Null => "#NULL!",
            Num => "#NUM!",
            Na => "#N/A",
            Spill => "#SPILL!",
            Other => return Err("Unsupported error filter value".into()),
        }
        .into(),
    })
}
fn denormalize(key: &NormalizedFilterKey) -> FilterKey {
    match key {
        NormalizedFilterKey::Blank => FilterKey::Blank,
        NormalizedFilterKey::Number(n) => FilterKey::Number(*n),
        NormalizedFilterKey::Bool(b) => FilterKey::Bool(*b),
        NormalizedFilterKey::Text(s) => FilterKey::Text(s.clone()),
        NormalizedFilterKey::Error(e) => FilterKey::Error(*e),
    }
}

pub(super) struct ExportFilters(Vec<(usize, Values)>);
/// A complete representable set or a single loss warning. Never export only
/// half an AND expression, or hide rows with no corresponding Excel criterion.
pub(super) fn export_filters(sheet: &Sheet, table: &DataTable) -> Result<ExportFilters, String> {
    let Some(spec) = sheet.table_view_spec().filter(|s| s.table == table.id) else {
        return Ok(ExportFilters(Vec::new()));
    };
    if !spec.filters.is_empty() {
        visigrid_engine::table_view::validate_table_view_layout(sheet, table.id)?;
    }
    let mut out = Vec::new();
    let mut count = 0;
    for filter in &spec.filters {
        if filter.criteria.text_filter.is_some() {
            return Err("Text predicates are not yet supported".into());
        }
        let Some(selected) = &filter.criteria.selected else {
            continue;
        };
        if selected.is_empty() {
            return Err("An empty checkbox selection cannot be represented safely".into());
        }
        let offset = table
            .columns
            .iter()
            .position(|c| c.id == filter.column)
            .ok_or("Missing filter column")?;
        let mut values = Values {
            blank: selected.contains(&NormalizedFilterKey::Blank),
            ..Default::default()
        };
        let mut found = HashSet::new();
        let mut unselected = HashSet::new();
        for row in table.range.start_row + 1..=table.range.end_row {
            let key = FilterKey::from_value(
                &sheet.get_computed_value(row, table.range.start_col + offset),
            );
            if matches!(key, FilterKey::Blank) {
                continue;
            }
            let text = raw_value(&key)?;
            let normalized = key.normalized();
            if selected.contains(&normalized) {
                values.values.insert(text);
                found.insert(normalized);
            } else {
                unselected.insert(text.to_lowercase());
            }
        }
        for key in selected.difference(&found) {
            if !matches!(key, NormalizedFilterKey::Blank) {
                values.values.insert(raw_value(&denormalize(key))?);
            }
        }
        if values
            .values
            .iter()
            .any(|s| unselected.contains(&s.to_lowercase()))
        {
            return Err(
                "Excel cannot distinguish selected and unselected values with the same filter text"
                    .into(),
            );
        }
        count += values.values.len();
        if count > 10_000 {
            return Err("More than 10,000 selected filter values".into());
        }
        if selected_keys(sheet, &table, offset, &values)? != *selected {
            return Err("Selected value types cannot be represented unambiguously in Excel".into());
        }
        out.push((offset, values));
    }
    out.sort_by_key(|(c, _)| *c);
    Ok(ExportFilters(out))
}

pub(super) fn write_xml(
    sheet: &Sheet,
    table: &DataTable,
    writer: &mut Writer<Vec<u8>>,
) -> Result<(), String> {
    let filters = export_filters(sheet, table).unwrap_or(ExportFilters(Vec::new()));
    let buttons = sheet
        .table_view_spec()
        .filter(|s| s.table == table.id)
        .is_none_or(|s| s.show_filter_buttons);
    if !buttons && filters.0.is_empty() {
        return Ok(());
    }
    let r = table.range;
    let reference = format!(
        "{}{}:{}{}",
        super::xlsx::col_to_letter(r.start_col),
        r.start_row + 1,
        super::xlsx::col_to_letter(r.end_col),
        r.end_row + 1
    );
    let mut auto = BytesStart::new("autoFilter");
    auto.push_attribute(("ref", reference.as_str()));
    writer
        .write_event(Event::Start(auto))
        .map_err(|e| e.to_string())?;
    for col in 0..table.columns.len() {
        let values = filters.0.iter().find(|(c, _)| *c == col).map(|(_, v)| v);
        if values.is_none() && buttons {
            continue;
        }
        let mut column = BytesStart::new("filterColumn");
        let id = col.to_string();
        column.push_attribute(("colId", id.as_str()));
        if !buttons {
            column.push_attribute(("hiddenButton", "1"));
        }
        writer
            .write_event(Event::Start(column))
            .map_err(|e| e.to_string())?;
        if let Some(values) = values {
            let mut list = BytesStart::new("filters");
            if values.blank {
                list.push_attribute(("blank", "1"));
            }
            writer
                .write_event(Event::Start(list))
                .map_err(|e| e.to_string())?;
            for value in &values.values {
                let mut v = BytesStart::new("filter");
                let encoded = encode_value(value);
                v.push_attribute(("val", encoded.as_str()));
                writer
                    .write_event(Event::Empty(v))
                    .map_err(|e| e.to_string())?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("filters")))
                .map_err(|e| e.to_string())?;
        }
        writer
            .write_event(Event::End(BytesEnd::new("filterColumn")))
            .map_err(|e| e.to_string())?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("autoFilter")))
        .map_err(|e| e.to_string())
}

pub(super) fn write_hidden_rows(
    sheet: &Sheet,
    ws: &mut rust_xlsxwriter::Worksheet,
    result: &mut ExportResult,
) -> Result<(), String> {
    let Some(spec) = sheet.table_view_spec().filter(|s| !s.filters.is_empty()) else {
        return Ok(());
    };
    let table = sheet
        .tables()
        .iter()
        .find(|t| t.id == spec.table)
        .ok_or("Missing filtered Table")?;
    if export_filters(sheet, table).is_err() {
        return Ok(());
    }
    for row in table.range.start_row + 1..=table.range.end_row {
        let passes = spec.filters.iter().all(|f| {
            let col = table.columns.iter().position(|c| c.id == f.column).unwrap()
                + table.range.start_col;
            f.criteria
                .passes(&FilterKey::from_value(&sheet.get_computed_value(row, col)))
        });
        if !passes {
            ws.set_row_hidden(row as u32).map_err(|e| e.to_string())?;
            result.hidden_rows_exported += 1;
        }
    }
    Ok(())
}

/// Mark active Table filtering at worksheet level as Excel does. Touch only
/// the first child, retaining the cell XML verbatim (including cached results).
pub(super) fn mark_filter_mode(xml: &str) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    let mut in_sheet = false;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| e.to_string())?;
        match &event {
            Event::Start(e) if e.local_name().as_ref() == b"worksheet" => in_sheet = true,
            Event::Start(e) | Event::Empty(e) if in_sheet => {
                let mut out = xml[..start].to_string();
                if e.local_name().as_ref() == b"sheetPr" {
                    let mut properties = BytesStart::new("sheetPr");
                    for a in e.attributes() {
                        let a = a.map_err(|e| e.to_string())?;
                        if a.key.local_name().as_ref() != b"filterMode" {
                            properties.push_attribute(a);
                        }
                    }
                    properties.push_attribute(("filterMode", "1"));
                    let mut writer = Writer::new(Vec::new());
                    writer
                        .write_event(if matches!(event, Event::Empty(_)) {
                            Event::Empty(properties)
                        } else {
                            Event::Start(properties)
                        })
                        .map_err(|e| e.to_string())?;
                    out.push_str(
                        std::str::from_utf8(&writer.into_inner()).map_err(|e| e.to_string())?,
                    );
                    out.push_str(&xml[reader.buffer_position() as usize..]);
                } else {
                    out.push_str("<sheetPr filterMode=\"1\"/>");
                    out.push_str(&xml[start..]);
                }
                return Ok(out);
            }
            Event::Eof => return Err("Missing exported worksheet".into()),
            _ => {}
        }
    }
}

pub(super) fn has_exported_filters(sheet: &Sheet) -> bool {
    sheet
        .tables()
        .iter()
        .any(|table| export_filters(sheet, table).is_ok_and(|f| !f.0.is_empty()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn spreadsheet_string_escapes_handle_utf16_and_literal_lookalikes() {
        assert_eq!(decode_value("_xD83D__xDE00_").unwrap(), "😀");
        assert!(decode_value("_xD83D_").is_err());
        for text in [
            "_x0041_",
            "_x005F_x0041_",
            "_x+041_",
            "_xZZZZ_",
            "é 😀 & < >",
        ] {
            assert_eq!(decode_value(&encode_value(text)).unwrap(), text);
        }
    }
}
