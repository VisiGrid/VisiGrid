//! XLSX Table interoperability. Cell contents remain authoritative, including
//! calculated-column exceptions. Relationships, not part numbering, bind sources.
use super::xlsx::{ExportResult, ImportResult};
use quick_xml::{
    events::{BytesStart, BytesText, Event},
    Reader, Writer,
};
use std::{
    collections::{HashMap, HashSet},
    fs::File,
    io::{Cursor, Read, Write},
    path::Path,
};
use visigrid_engine::{
    formula::parser::parse,
    table::{DataTable, TableColumn, TableColumnId, TableId, TableRange, TableStyle},
    workbook::Workbook,
};

const MAX_PART_BYTES: u64 = 32 * 1024 * 1024;
// Bound work independently of cell-import limits, including malformed/duplicate
// definitions. Large ordinary sheets without Table links do not use this parser.
const MAX_TABLES: usize = 1024;
const MAX_TABLE_BYTES: u64 = 64 * 1024 * 1024;
struct TableBudget { attempts: usize, bytes_left: u64 }
struct PendingTable { sheet: usize, imported: ImportedTable }

pub(super) fn attr(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<String>, String> {
    for a in e.attributes() {
        let a = a.map_err(|e| e.to_string())?;
        if a.key.local_name().as_ref() == key {
            return quick_xml::escape::unescape(
                std::str::from_utf8(&a.value).map_err(|e| e.to_string())?,
            )
            .map(|s| Some(s.into_owned()))
            .map_err(|e| e.to_string());
        }
    }
    Ok(None)
}
fn part(zip: &mut zip::ZipArchive<File>, name: &str) -> Result<String, String> {
    let file = zip
        .by_name(name)
        .map_err(|e| format!("Missing XLSX part {name}: {e}"))?;
    if file.size() > MAX_PART_BYTES {
        return Err(format!("XLSX part {name} is too large"));
    }
    let mut xml = String::new();
    file.take(MAX_PART_BYTES + 1)
        .read_to_string(&mut xml)
        .map_err(|e| e.to_string())?;
    if xml.len() as u64 > MAX_PART_BYTES {
        return Err(format!("XLSX part {name} is too large"));
    }
    Ok(xml)
}
fn table_part(zip: &mut zip::ZipArchive<File>, name: &str, budget: &mut TableBudget) -> Result<String, String> {
    let exhausted = "The workbook's 64 MiB Table metadata budget was exceeded";
    if budget.bytes_left == 0 { return Err(exhausted.into()); }
    let file = zip.by_name(name).map_err(|e| format!("Missing XLSX part {name}: {e}"))?;
    if file.size() > MAX_PART_BYTES { return Err(format!("XLSX part {name} is too large")); }
    if file.size() > budget.bytes_left { return Err(exhausted.into()); }
    let limit = budget.bytes_left.min(MAX_PART_BYTES);
    let mut bytes = Vec::new();
    let read = file.take(limit + 1).read_to_end(&mut bytes);
    // Count actual decompressed bytes, even on malformed XML/UTF-8 or an I/O
    // error. A forged ZIP size must not bypass the aggregate budget.
    budget.bytes_left = budget.bytes_left.saturating_sub(bytes.len() as u64);
    read.map_err(|e| e.to_string())?;
    if bytes.len() as u64 > limit {
        return Err(if limit < MAX_PART_BYTES { exhausted.into() } else { format!("XLSX part {name} is too large") });
    }
    String::from_utf8(bytes).map_err(|e| e.to_string())
}
pub(super) fn target(source: &str, path: &str) -> Result<String, String> {
    let joined = if path.starts_with('/') {
        path.trim_start_matches('/').into()
    } else {
        format!("{}/{}", source.rsplit_once('/').map_or("", |x| x.0), path)
    };
    let mut parts = Vec::new();
    for p in joined.split('/') {
        match p {
            "" | "." => {}
            ".." => {
                parts.pop().ok_or("Invalid XLSX relationship path")?;
            }
            p => parts.push(p),
        }
    }
    Ok(parts.join("/"))
}
fn rel_path(part: &str) -> String {
    let (base, file) = part.rsplit_once('/').unwrap_or(("", part));
    format!("{base}/_rels/{file}.rels")
}
struct Rel {
    target: String,
    kind: String,
    external: bool,
}
fn relationships(xml: &str, source: &str) -> Result<HashMap<String, Rel>, String> {
    let mut r = Reader::from_str(xml);
    let mut rels = HashMap::new();
    loop {
        match r.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                let id = attr(&e, b"Id")?.ok_or("Missing relationship ID")?;
                let external = attr(&e, b"TargetMode")?.as_deref() == Some("External");
                let path = attr(&e, b"Target")?.ok_or("Missing relationship target")?;
                let rel = Rel {
                    target: if external {
                        path
                    } else {
                        target(source, &path)?
                    },
                    kind: attr(&e, b"Type")?.unwrap_or_default(),
                    external,
                };
                if rels.insert(id, rel).is_some() {
                    return Err("Duplicate relationship ID".into());
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(rels)
}
pub(super) fn range(input: &str) -> Result<TableRange, String> {
    fn cell(s: &str) -> Option<(usize, usize)> {
        let s = s.replace('$', "");
        let n = s.find(|c: char| c.is_ascii_digit())?;
        let (letters, digits) = s.split_at(n);
        if letters.is_empty() || !letters.bytes().all(|b| b.is_ascii_alphabetic()) {
            return None;
        }
        let c = letters.bytes().try_fold(0usize, |c, b| {
            c.checked_mul(26)?
                .checked_add((b.to_ascii_uppercase() - b'A' + 1) as usize)
        })?;
        Some((
            digits.parse::<usize>().ok()?.checked_sub(1)?,
            c.checked_sub(1)?,
        ))
    }
    let (a, b) = input.split_once(':').unwrap_or((input, input));
    let (r0, c0) = cell(a).ok_or("Invalid Table range")?;
    let (r1, c1) = cell(b).ok_or("Invalid Table range")?;
    let range = TableRange {
        start_row: r0,
        start_col: c0,
        end_row: r1,
        end_col: c1,
    };
    range.validate(
        visigrid_engine::sheet::NUM_ROWS,
        visigrid_engine::sheet::NUM_COLS,
    )?;
    Ok(range)
}
struct ImportedTable {
    table: DataTable,
    view: Result<super::xlsx_table_filters::ImportedView, String>,
    warnings: Vec<String>,
}
fn parse_table(xml: &str) -> Result<ImportedTable, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().expand_empty_elements = true;
    let mut column_count = None;
    let mut in_column = false;
    let mut closed = false;
    let mut table = None;
    let mut columns = Vec::<TableColumn>::new();
    let mut column_ids = HashSet::new();
    let mut style = TableStyle {
        banded_rows: false,
        excel_style: None,
        ..Default::default()
    };
    let mut warnings = Vec::new();
    let mut totals_visible = false;
    let mut totals_shown = None;
    let mut totals_columns = Vec::<visigrid_engine::table::TableTotal>::new();
    let mut in_total_formula = false;
    let mut total_formula = String::new();
    let mut in_formula = false;
    let mut formula = String::new();
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) | Event::Empty(e) => match e.local_name().as_ref() {
                b"table" => {
                    if table.is_some() {
                        return Err("Duplicate Table definition".into());
                    }
                    let name = attr(&e, b"displayName")?
                        .or(attr(&e, b"name")?)
                        .ok_or("Missing Table name")?;
                    if attr(&e, b"headerRowCount")?
                        .as_deref()
                        .is_some_and(|s| s != "1")
                    {
                        return Err(format!(
                            "Table {name} has hidden headers, which are not supported"
                        ));
                    }
                    totals_visible = match attr(&e, b"totalsRowCount")?.as_deref() {
                        None | Some("0") => false,
                        Some("1") => true,
                        _ => return Err("Invalid Table totals-row count".into()),
                    };
                    totals_shown = match attr(&e, b"totalsRowShown")?.as_deref() {
                        None => None,
                        Some("1" | "true") => Some(true),
                        Some("0" | "false") => Some(false),
                        _ => return Err("Invalid totals-row visibility".into()),
                    };
                    if attr(&e, b"tableType")?
                        .as_deref()
                        .is_some_and(|s| s != "worksheet")
                    {
                        return Err(format!("Table {name} uses an external or XML data source, which is not supported"));
                    }
                    table = Some((
                        name,
                        range(&attr(&e, b"ref")?.ok_or("Missing Table range")?)?,
                    ));
                }
                b"tableColumns" => {
                    if column_count.is_some() {
                        return Err("Duplicate Table columns".into());
                    }
                    column_count = Some(
                        attr(&e, b"count")?
                            .ok_or("Missing Table column count")?
                            .parse::<usize>()
                            .map_err(|_| "Invalid Table column count")?,
                    );
                }
                b"tableColumn" => {
                    if columns.len() >= visigrid_engine::sheet::NUM_COLS {
                        return Err("Table has more columns than the worksheet supports".into());
                    }
                    for key in [
                        b"dataDxfId".as_slice(),
                        b"headerRowDxfId",
                        b"totalsRowDxfId",
                        b"dataCellStyle",
                        b"headerRowCellStyle",
                        b"totalsRowCellStyle",
                    ] {
                        if attr(&e, key)?.is_some() {
                            warnings.push("Column style overrides are not preserved.".into());
                            break;
                        }
                    }
                    if in_column {
                        return Err("Nested Table column".into());
                    }
                    in_column = true;
                    let id: u64 = attr(&e, b"id")?
                        .ok_or("Missing column ID")?
                        .parse()
                        .map_err(|_| "Invalid column ID")?;
                    if id == 0 || id == u64::MAX || !column_ids.insert(id) {
                        return Err("Invalid or duplicate Table column ID".into());
                    }
                    totals_columns.push(visigrid_engine::table::TableTotal {
                        function: attr(&e, b"totalsRowFunction")?,
                        label: attr(&e, b"totalsRowLabel")?,
                        formula: None,
                    });
                    columns.push(TableColumn {
                        id: TableColumnId(id),
                        name: attr(&e, b"name")?.ok_or("Missing Table column name")?,
                        formula: None,
                        formula_origin: 1,
                    });
                }
                b"totalsRowFormula" => {
                    if !in_column || in_total_formula || totals_columns.last().is_some_and(|c| c.formula.is_some()) {
                        return Err("Invalid totals-row formula".into());
                    }
                    if super::xlsx_table_filters::boolean(&e, b"array", false)? {
                        return Err("Array totals-row formulas are not supported".into());
                    }
                    in_total_formula = true;
                    total_formula.clear();
                }
                b"calculatedColumnFormula" => {
                    if !in_column
                        || in_formula
                        || columns.last().is_some_and(|c| c.formula.is_some())
                    {
                        return Err("Invalid calculated-column formula".into());
                    }
                    in_formula = true;
                    formula.clear();
                    if attr(&e, b"array")?
                        .as_deref()
                        .is_some_and(|s| s == "1" || s == "true")
                    {
                        return Err("Array calculated columns are not supported".into());
                    }
                }
                b"tableStyleInfo" => {
                    style.banded_rows =
                        super::xlsx_table_filters::boolean(&e, b"showRowStripes", false)?;
                    style.banded_columns =
                        super::xlsx_table_filters::boolean(&e, b"showColumnStripes", false)?;
                    style.first_column =
                        super::xlsx_table_filters::boolean(&e, b"showFirstColumn", false)?;
                    style.last_column =
                        super::xlsx_table_filters::boolean(&e, b"showLastColumn", false)?;
                    style.excel_style = attr(&e, b"name")?.filter(|name| !name.is_empty());
                    if style
                        .excel_style
                        .as_deref()
                        .is_some_and(|name| !TableStyle::is_builtin_excel_style(name))
                    {
                        warnings.push("Custom Excel Table styles are not supported; the Table uses the default style on export.".into());
                        style.excel_style = TableStyle::default().excel_style;
                    }
                    if style.excel_style != TableStyle::default().excel_style
                        || style.banded_columns
                        || style.first_column
                        || style.last_column
                    {
                        warnings.push("Excel Table style options are retained for export; VisiGrid displays its own theme. Custom workbook theme colors are not retained.".into());
                    }
                }
                b"extLst" | b"xmlColumnPr" => {
                    warnings.push("Extended Excel Table metadata is not preserved.".into())
                }
                _ => {}
            },
            Event::Text(t) if in_total_formula => total_formula.push_str(
                &quick_xml::escape::unescape(&t.decode().map_err(|e| e.to_string())?).map_err(|e| e.to_string())?),
            Event::CData(t) if in_total_formula => total_formula.push_str(&t.decode().map_err(|e| e.to_string())?),
            Event::GeneralRef(e) if in_total_formula => total_formula.push_str(
                &quick_xml::escape::unescape(&format!("&{};", e.decode().map_err(|e| e.to_string())?)).map_err(|e| e.to_string())?),
            Event::End(e) if e.local_name().as_ref() == b"totalsRowFormula" => {
                if !in_total_formula { return Err("Invalid totals formula close".into()); }
                totals_columns.last_mut().ok_or("Missing totals column")?.formula =
                    Some(super::xlsx_functions::from_excel(&format!("={}", total_formula.trim().trim_start_matches('='))));
                in_total_formula = false;
            }
            Event::Text(t) if in_formula => formula.push_str(
                &quick_xml::escape::unescape(&t.decode().map_err(|e| e.to_string())?)
                    .map_err(|e| e.to_string())?,
            ),
            Event::CData(t) if in_formula => {
                formula.push_str(&t.decode().map_err(|e| e.to_string())?)
            }
            Event::GeneralRef(e) if in_formula => {
                let name = e.decode().map_err(|e| e.to_string())?;
                formula.push_str(
                    &quick_xml::escape::unescape(&format!("&{name};"))
                        .map_err(|e| e.to_string())?,
                );
            }
            Event::End(e) if e.local_name().as_ref() == b"calculatedColumnFormula" => {
                // Same _xlfn./_xlws./_xlpm. stripping as cell formulas (#89).
                let source = super::xlsx_functions::from_excel(&format!("={}", formula.trim().trim_start_matches('=')));
                if parse(&source).is_ok() {
                    columns.last_mut().unwrap().formula = Some(source);
                } else {
                    warnings.push(format!("Calculated-column rule for {} is unsupported; existing cells are kept without an automatic-fill rule.", columns.last().unwrap().name));
                }
                in_formula = false;
            }
            Event::End(e) if e.local_name().as_ref() == b"tableColumn" => in_column = false,
            Event::End(e) if e.local_name().as_ref() == b"table" => closed = true,
            Event::Eof => break,
            _ => {}
        }
    }
    if !closed || in_formula || in_total_formula || in_column || column_count != Some(columns.len()) {
        return Err("Incomplete or inconsistent Table metadata".into());
    }
    let (name, mut range) = table.ok_or("Missing Table definition")?;
    if totals_visible {
        if range.end_row <= range.start_row { return Err("Totals row overlaps the header".into()); }
        range.end_row -= 1;
    }
    let totals = (totals_visible || totals_columns.iter().any(|t| *t != Default::default()))
        .then_some(visigrid_engine::table::TableTotals {
            visible: totals_visible, shown: totals_shown, hidden_rows: Default::default(), columns: totals_columns,
        });
    let next_column_id = columns.iter().map(|c| c.id.0).max().unwrap_or(0) + 1;
    let table = DataTable {
        id: TableId(1),
        name,
        range,
        columns,
        next_column_id,
        style,
        source: None,
        saved_views: Vec::new(),
        totals,
    };
    table.validate(
        visigrid_engine::sheet::NUM_ROWS,
        visigrid_engine::sheet::NUM_COLS,
    )?;
    warnings.sort();
    warnings.dedup();
    let view = super::xlsx_table_filters::parse(xml, &table);
    Ok(ImportedTable {
        table,
        view,
        warnings,
    })
}

/// Tables are installed before dependency binding/recalculation. Invalid or
/// unsupported definitions leave their cells untouched and produce a warning.
pub(crate) fn import(
    path: &Path,
    wb: &mut Workbook,
    result: &mut ImportResult,
    values_only: bool,
) -> Vec<super::xlsx_table_filters::PendingView> {
    let mut pending = Vec::new();
    if let Err(error) = read_tables(path, wb, result, &mut pending) {
        result.warnings.push(format!("Excel Table metadata could not be read: {error}. Cells were kept; formulas referencing skipped Tables may show errors."));
    }
    let (definitions, details): (Vec<_>, Vec<_>) = pending.into_iter().map(|pending| {
        let ImportedTable { mut table, view, warnings } = pending.imported;
        if values_only {
            for column in &mut table.columns { column.formula = None; }
            if let Some(totals) = &mut table.totals {
                for column in &mut totals.columns { column.function = None; column.formula = None; }
            }
        }
        let name = table.name.clone();
        ((pending.sheet, table), (pending.sheet, name, view, warnings))
    }).unzip();
    let outcomes = if result.truncated {
        definitions.iter().map(|_| Err("Workbook import was truncated; Table definitions cannot safely be restored".into())).collect()
    } else { wb.restore_imported_tables(definitions) };
    let mut views = Vec::new();
    for (outcome, (sheet, name, view, warnings)) in outcomes.into_iter().zip(details) {
        match outcome {
            Ok(table) => {
                result.tables_imported += 1;
                for warning in warnings { result.warnings.push(format!("Table {name}: {warning}")); }
                views.push(super::xlsx_table_filters::PendingView { sheet, table, view });
            }
            Err(error) => skip_table(result, &wb.sheet(sheet).unwrap().name, &error),
        }
    }
    views
}
fn skip_table(result: &mut ImportResult, sheet: &str, error: &str) {
    result.tables_skipped += 1;
    result.warnings.push(format!("Excel Table on {sheet} was kept as plain cells: {error}. Formulas referencing it may show errors."));
}
fn read_tables(
    path: &Path, wb: &Workbook, result: &mut ImportResult, pending: &mut Vec<PendingTable>,
) -> Result<(), String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let Ok(mut zip) = zip::ZipArchive::new(file) else { return Ok(()); };
    if !zip.file_names().any(|p| p == "xl/workbook.xml") { return Ok(()); }
    let rels = relationships(&part(&mut zip, "xl/_rels/workbook.xml.rels")?, "xl/workbook.xml")?;
    let xml = part(&mut zip, "xl/workbook.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut seen = HashSet::new();
    let mut budget = TableBudget { attempts: 0, bytes_left: MAX_TABLE_BYTES };
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheet" => {
                let name = attr(&e, b"name")?.ok_or("Missing sheet name")?;
                let Some(si) = wb.sheets().iter().position(|s| s.name == name) else { continue; };
                let outcome = (|| {
                    let rel = rels.get(&attr(&e, b"id")?.ok_or("Missing sheet relationship")?)
                        .ok_or("Missing sheet relationship")?;
                    if !rel.kind.ends_with("/worksheet") || rel.external { return Ok(()); }
                    read_sheet_tables(&mut zip, &rel.target, si, &name, result, pending, &mut seen, &mut budget)
                })();
                if let Err(error) = outcome {
                    result.warnings.push(format!("Excel Table metadata on {name} could not be read: {error}. Cells were kept; formulas referencing skipped Tables may show errors."));
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(())
}
fn read_sheet_tables(
    zip: &mut zip::ZipArchive<File>, path: &str, sheet: usize, name: &str,
    result: &mut ImportResult, pending: &mut Vec<PendingTable>,
    seen: &mut HashSet<String>, budget: &mut TableBudget,
) -> Result<(), String> {
    let links = rel_path(path);
    match zip.by_name(&links) {
        Ok(_) => {},
        Err(zip::result::ZipError::FileNotFound) => return Ok(()),
        Err(error) => return Err(error.to_string()),
    }
    let rels = relationships(&part(zip, &links)?, path)?;
    if !rels.values().any(|rel| rel.kind.ends_with("/table")) { return Ok(()); }
    // Worksheet XML includes all cells and is not Table metadata. Stream it
    // without the metadata part-size limit, while bounding collected Table ids.
    let (ids, skipped) = (|| {
        let file = zip.by_name(path).map_err(|e| format!("Missing XLSX part {path}: {e}"))?;
        let mut reader = Reader::from_reader(std::io::BufReader::new(file));
        let mut buf = Vec::new();
        let mut ids = Vec::new();
        let mut skipped = 0usize;
        loop {
            match reader.read_event_into(&mut buf).map_err(|e| e.to_string())? {
                Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"tablePart" => {
                    if budget.attempts == MAX_TABLES { skipped += 1; }
                    else {
                        budget.attempts += 1;
                        ids.push(attr(&e, b"id")?.ok_or("Missing Table relationship")?);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok::<_, String>((ids, skipped))
    })().map_err(|error| {
        result.tables_skipped += rels.values().filter(|rel| rel.kind.ends_with("/table")).count();
        error
    })?;
    if skipped > 0 {
        result.tables_skipped += skipped;
        result.warnings.push(format!("{skipped} Excel Table definitions on {name} exceeded the workbook limit of {MAX_TABLES}. Cells were kept; formulas referencing skipped Tables may show errors."));
    }
    for id in ids {
        let outcome = (|| {
            let rel = rels.get(&id).ok_or("Missing Table relationship")?;
            if !rel.kind.ends_with("/table") || rel.external { return Err("Unsupported Table relationship".into()); }
            if !seen.insert(rel.target.clone()) { return Err("Duplicate Table part reference".into()); }
            let xml = table_part(zip, &rel.target, budget)?;
            parse_table(&xml)
        })();
        match outcome {
            Ok(imported) => pending.push(PendingTable { sheet, imported }),
            Err(error) => skip_table(result, name, &error),
        }
    }
    Ok(())
}

pub(crate) fn export_warnings(
    wb: &Workbook,
    order: super::xlsx::ExportOrder,
) -> Result<Vec<String>, String> {
    wb.validate_table_view_specs()?;
    let mut warnings = Vec::new();
    for (sheet_id, table) in wb.tables() {
        if !table.saved_views.is_empty() {
            warnings.push(format!("Table {}: named saved views are not exported to Excel. Keep a VisiGrid or full JSON copy to preserve them. The active view is exported using the supported sort/filter settings.", table.name));
        }
        if table.name.chars().count() > 255 {
            return Err(format!(
                "Table {} has a name longer than Excel's 255-character limit; no file was written.",
                table.name
            ));
        }
        if table.range.data_rows() == 0 && table.totals_row().is_none() {
            return Err(format!("Table {} has only headers. Add an empty record before exporting it to Excel; no file was written.", table.name));
        }
        if totals_exported_as_values(table) {
            warnings.push(format!(
                "Table {} has no records; its totals row is exported as values.",
                table.name
            ));
        }
        if let Some(view) = wb
            .sheet_by_id(sheet_id)
            .and_then(|s| s.table_view_spec())
            .filter(|v| v.table == table.id)
        {
            if view.sort.is_some() {
                warnings.push(match order {
                    super::xlsx::ExportOrder::Sorted => format!("Table {}: Excel export places every record in the saved sort order, including filtered-out records, and updates supported formula references. Clearing the sort in the exported copy keeps that order. Reapplying it in Excel may change text or mixed-type ordering.", table.name),
                    super::xlsx::ExportOrder::Stored => format!("Table {}: Every record stays in stored order and formulas keep their stored coordinates. The saved sort is included; use Reapply in Excel to display that order. Text and mixed-type ordering may differ from VisiGrid.", table.name),
                });
            }
            if let Err(reason) =
                super::xlsx_table_filters::export_filters(wb.sheet_by_id(sheet_id).unwrap(), table)
            {
                warnings.push(format!("Table {}: filter criteria are not exported ({reason}). All records are shown in Excel.", table.name));
            }
        }
    }
    warnings.extend(crate::xlsx_cond_formats::warnings(wb)?);
    warnings.extend(crate::xlsx_names::export_warnings(wb));
    if !wb.pivots().is_empty() {
        warnings.push("Pivot results are exported as cells. Pivot definitions and Table-source bindings are not exported to Excel.".into());
    }
    Ok(warnings)
}

pub(crate) fn write(
    sheet: &visigrid_engine::sheet::Sheet,
    ws: &mut rust_xlsxwriter::Worksheet,
    result: &mut ExportResult,
) -> Result<(), String> {
    for table in sheet.tables() {
        let columns: Vec<_> = table
            .columns
            .iter()
            .enumerate()
            .map(|(i, c)| {
                rust_xlsxwriter::TableColumn::new()
                    .set_header(&c.name)
                    .set_header_format(super::xlsx::build_excel_format(
                        &sheet.get_format(table.range.start_row, table.range.start_col + i),
                    ))
            })
            .collect();
        let arrows = sheet
            .table_view_spec()
            .filter(|v| v.table == table.id)
            .is_none_or(|v| v.show_filter_buttons);
        let xlsx_table = rust_xlsxwriter::Table::new()
            .set_name(&table.name)
            .set_total_row(table.totals_row().is_some() && !totals_exported_as_values(table))
            .set_columns(&columns)
            .set_style(rust_xlsxwriter::TableStyle::Medium2)
            .set_banded_rows(table.style.banded_rows)
            .set_autofilter(
                arrows
                    || sheet
                        .table_view_spec()
                        .is_some_and(|v| v.table == table.id && !v.filters.is_empty()),
            );
        let r = table.full_range();
        ws.add_table(
            r.start_row as u32,
            r.start_col as u16,
            r.end_row as u32,
            r.end_col as u16,
            &xlsx_table,
        )
        .map_err(|e| format!("Cannot export Table {}: {e}", table.name))?;
        result.tables_exported += 1;
    }
    Ok(())
}

/// OOXML uses the long this-row selector. Rewrite parsed references only, so
/// quoted strings, sheet names and escaped @ characters remain untouched.
/// Function names and LET/LAMBDA names then take Excel's stored form.
pub(crate) fn excel_formula(source: &str) -> String {
    use visigrid_engine::formula::structured::{source_references, TableSection};
    let mut out = source.to_string();
    for (start, end, reference) in source_references(source).into_iter().rev() {
        if reference.section != TableSection::ThisRow {
            continue;
        }
        let prefix = reference.table.as_deref().unwrap_or("");
        let replacement = if reference.columns.is_some() {
            let canonical = reference.format();
            format!("{}[[#This Row],{}", prefix, &canonical[prefix.len() + 2..])
        } else {
            format!("{prefix}[#This Row]")
        };
        out.replace_range(start..end, &replacement);
    }
    super::xlsx_functions::prefix_future_functions(&super::xlsx_functions::excel_function_names(
        &out,
    ))
}

/// Excel cannot represent a header plus a totals row with nothing between
/// them. Export keeps that footer row as the table's one data row and writes
/// its cells as values, so nothing below the table moves.
pub(crate) fn totals_exported_as_values(table: &visigrid_engine::table::DataTable) -> bool {
    table.range.data_rows() == 0 && table.totals_row().is_some()
}

/// Add calculated-column metadata without asking the writer to fill cells.
/// Its automatic fill would overwrite blank/value/formula exceptions.
pub(crate) fn finish(bytes: Vec<u8>, wb: &Workbook) -> Result<Vec<u8>, String> {
    if wb.tables().next().is_none() {
        return Ok(bytes);
    }
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| e.to_string())?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut written = HashSet::new();
    // Names here belong to the package just produced by rust_xlsxwriter;
    // imports continue resolving arbitrary worksheet relationship targets.
    let filtered_sheets: HashSet<_> = wb
        .sheets()
        .iter()
        .enumerate()
        .filter(|(_, sheet)| super::xlsx_table_filters::has_exported_filters(sheet))
        .map(|(i, _)| format!("xl/worksheets/sheet{}.xml", i + 1))
        .collect();
    for i in 0..input.len() {
        let mut entry = input.by_index(i).map_err(|e| e.to_string())?;
        if filtered_sheets.contains(entry.name()) {
            let name = entry.name().to_string();
            let mut xml = String::new();
            entry.read_to_string(&mut xml).map_err(|e| e.to_string())?;
            let xml = super::xlsx_table_filters::mark_filter_mode(&xml)?;
            output
                .start_file(
                    name,
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(|e| e.to_string())?;
            output
                .write_all(xml.as_bytes())
                .map_err(|e| e.to_string())?;
        } else if entry.name().starts_with("xl/tables/") && entry.name().ends_with(".xml") {
            let name = entry.name().to_string();
            let mut xml = String::new();
            entry.read_to_string(&mut xml).map_err(|e| e.to_string())?;
            let mut reader = Reader::from_str(&xml);
            let mut writer = Writer::new(Vec::new());
            let mut table = None;
            let mut offset = 0;
            loop {
                let event = reader.read_event().map_err(|e| e.to_string())?;
                if let Event::Start(e) | Event::Empty(e) = &event {
                    if e.local_name().as_ref() == b"table" {
                        table =
                            attr(e, b"name")?.and_then(|n| wb.table_by_name(&n).map(|(_, t)| t));
                        if table.is_some_and(totals_exported_as_values) {
                            let mut start = e.clone();
                            start.clear_attributes();
                            for attr in e.attributes() {
                                let attr = attr.map_err(|e| e.to_string())?;
                                if attr.key.as_ref() != b"totalsRowShown"
                                    && attr.key.as_ref() != b"totalsRowCount"
                                {
                                    start.push_attribute(attr);
                                }
                            }
                            writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(start) } else { Event::Start(start) }).map_err(|e| e.to_string())?;
                            continue;
                        }
                        if let Some(totals) = table.and_then(|t| t.totals.as_ref()) {
                            let mut start = e.clone();
                            start.clear_attributes();
                            for attr in e.attributes() {
                                let attr = attr.map_err(|e| e.to_string())?;
                                if attr.key.as_ref() != b"totalsRowShown" { start.push_attribute(attr); }
                            }
                            if let Some(shown) = totals.shown { start.push_attribute(("totalsRowShown", if shown { "1" } else { "0" })); }
                            writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(start) } else { Event::Start(start) }).map_err(|e| e.to_string())?;
                            continue;
                        }
                    }
                    if e.local_name().as_ref() == b"tableStyleInfo" {
                        if !matches!(event, Event::Empty(_)) {
                            return Err("Unexpected XLSX Table style element".into());
                        }
                        if let Some(t) = table {
                            let mut style = BytesStart::new("tableStyleInfo");
                            if let Some(name) = &t.style.excel_style {
                                style.push_attribute(("name", name.as_str()));
                            }
                            for (key, value) in [
                                ("showRowStripes", t.style.banded_rows),
                                ("showColumnStripes", t.style.banded_columns),
                                ("showFirstColumn", t.style.first_column),
                                ("showLastColumn", t.style.last_column),
                            ] {
                                style.push_attribute((key, if value { "1" } else { "0" }));
                            }
                            writer
                                .write_event(Event::Empty(style))
                                .map_err(|e| e.to_string())?;
                            continue;
                        }
                    }
                    if e.local_name().as_ref() == b"autoFilter" {
                        if !matches!(event, Event::Empty(_)) {
                            return Err("Unexpected XLSX Table filter element".into());
                        }
                        if let Some(t) = table {
                            let (sid, _) = wb.table(t.id).ok_or("Missing exported Table")?;
                            super::xlsx_table_filters::write_xml(
                                wb.sheet_by_id(sid).unwrap(),
                                t,
                                &mut writer,
                            )?;
                            continue;
                        }
                    }
                    if e.local_name().as_ref() == b"tableColumns" {
                        if let Some(t) = table {
                            let (sid, _) = wb.table(t.id).ok_or("Missing exported Table")?;
                            super::xlsx_table_sorts::write(
                                wb.sheet_by_id(sid).unwrap(),
                                t,
                                &mut writer,
                            )?;
                        }
                    }
                    if e.local_name().as_ref() == b"tableColumn" {
                        if let Some(t) = table {
                            let mut column = e.clone();
                            let total = t
                                .totals
                                .as_ref()
                                .filter(|_| !totals_exported_as_values(t))
                                .and_then(|totals| totals.columns.get(offset));
                            if let Some(total) = total {
                                if let Some(function) = &total.function { column.push_attribute(("totalsRowFunction", function.as_str())); }
                                if let Some(label) = &total.label { column.push_attribute(("totalsRowLabel", label.as_str())); }
                            }
                            writer.write_event(Event::Start(column.clone())).map_err(|e| e.to_string())?;
                            let calculated = t.formula_at(t.range.start_row + 1, t.range.start_col + offset)
                                .or_else(|| (t.range.data_rows() == 0).then(|| t.columns.get(offset).and_then(|c| c.formula.clone())).flatten());
                            let current_total_formula = t.totals_row().and_then(|row| {
                                let (sid, _) = wb.table(t.id)?;
                                let source = wb.sheet_by_id(sid)?.get_raw(row, t.range.start_col + offset);
                                source.starts_with('=').then_some(source)
                            });
                            let total_formula = total.filter(|t| t.formula.is_some()).and_then(|t|
                                current_total_formula.as_ref().or(t.formula.as_ref()));
                            for (tag, source) in [("calculatedColumnFormula", calculated.as_ref()),
                                ("totalsRowFormula", total_formula)] {
                                if let Some(source) = source {
                                    writer.write_event(Event::Start(BytesStart::new(tag))).map_err(|e| e.to_string())?;
                                    writer.write_event(Event::Text(BytesText::new(excel_formula(source).trim_start_matches('=')))).map_err(|e| e.to_string())?;
                                    writer.write_event(Event::End(quick_xml::events::BytesEnd::new(tag))).map_err(|e| e.to_string())?;
                                }
                            }
                            if matches!(event, Event::Empty(_)) {
                                writer.write_event(Event::End(column.to_end())).map_err(|e| e.to_string())?;
                            }
                            if calculated.is_some() { written.insert((t.id, offset)); }
                            offset += 1;
                            continue;
                        }
                        offset += 1;
                    }
                }
                if matches!(event, Event::Eof) {
                    break;
                }
                writer.write_event(event).map_err(|e| e.to_string())?;
            }
            output
                .start_file(
                    name,
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(|e| e.to_string())?;
            output
                .write_all(&writer.into_inner())
                .map_err(|e| e.to_string())?;
        } else {
            output.raw_copy_file(entry).map_err(|e| e.to_string())?;
        }
    }
    let expected = wb
        .tables()
        .map(|(_, t)| t.columns.iter().filter(|c| c.formula.is_some()).count())
        .sum::<usize>();
    if written.len() != expected {
        return Err("XLSX writer did not produce all calculated Table columns".into());
    }
    output
        .finish()
        .map(|c| c.into_inner())
        .map_err(|e| e.to_string())
}
