//! XLSX Table interoperability. Cell contents remain authoritative, including
//! calculated-column exceptions. Relationships, not part numbering, bind sources.
use super::xlsx::{ExportResult, ImportResult};
use quick_xml::{
    events::{BytesStart, BytesText, Event},
    Reader, Writer,
};
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fs::File,
    io::{Cursor, Read, Write},
    path::Path,
};
use visigrid_engine::{
    formula::parser::parse,
    table::{DataTable, TableColumn, TableColumnId, TableId, TableRange, TableStyle},
    workbook::{SavedTableSheet, Workbook},
};

const MAX_PART_BYTES: u64 = 32 * 1024 * 1024;
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
fn target(source: &str, path: &str) -> Result<String, String> {
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
                    if attr(&e, b"totalsRowCount")?
                        .as_deref()
                        .is_some_and(|s| s != "0")
                    {
                        return Err(format!(
                            "Table {name} has a totals row, which is not supported"
                        ));
                    }
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
                    for key in [
                        b"dataDxfId".as_slice(),
                        b"headerRowDxfId",
                        b"totalsRowDxfId",
                        b"dataCellStyle",
                        b"headerRowCellStyle",
                        b"totalsRowCellStyle",
                        b"totalsRowFunction",
                        b"totalsRowLabel",
                    ] {
                        if attr(&e, key)?.is_some() {
                            warnings.push("Column style overrides and dormant totals settings are not preserved.".into());
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
                    columns.push(TableColumn {
                        id: TableColumnId(id),
                        name: attr(&e, b"name")?.ok_or("Missing Table column name")?,
                        formula: None,
                        formula_origin: 1,
                    });
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
                let source = format!("={}", formula.trim().trim_start_matches('='));
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
    if !closed || in_formula || in_column || column_count != Some(columns.len()) {
        return Err("Incomplete or inconsistent Table metadata".into());
    }
    let (name, range) = table.ok_or("Missing Table definition")?;
    let next_column_id = columns.iter().map(|c| c.id.0).max().unwrap_or(0) + 1;
    let table = DataTable {
        id: TableId(1),
        name,
        range,
        columns,
        next_column_id,
        style,
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
    let mut views = Vec::new();
    if let Err(error) = read_tables(path, wb, result, values_only, &mut views) {
        result.warnings.push(format!("Excel Table metadata could not be read: {error}. Cells were kept; formulas referencing skipped Tables may show errors."));
    }
    views
}
fn read_tables(
    path: &Path,
    wb: &mut Workbook,
    result: &mut ImportResult,
    values_only: bool,
    views: &mut Vec<super::xlsx_table_filters::PendingView>,
) -> Result<(), String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let Ok(mut zip) = zip::ZipArchive::new(file) else {
        return Ok(());
    };
    if !zip.file_names().any(|p| p == "xl/workbook.xml") {
        return Ok(());
    }
    let rels = relationships(
        &part(&mut zip, "xl/_rels/workbook.xml.rels")?,
        "xl/workbook.xml",
    )?;
    let xml = part(&mut zip, "xl/workbook.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut seen = HashSet::new();
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheet" => {
                let name = attr(&e, b"name")?.ok_or("Missing sheet name")?;
                let Some(si) = wb.sheets().iter().position(|s| s.name == name) else {
                    continue;
                };
                let rel = rels
                    .get(&attr(&e, b"id")?.ok_or("Missing sheet relationship")?)
                    .ok_or("Missing sheet relationship")?;
                if !rel.kind.ends_with("/worksheet") || rel.external {
                    continue;
                }
                let sheet_xml = part(&mut zip, &rel.target)?;
                let mut sr = Reader::from_str(&sheet_xml);
                let mut ids = Vec::new();
                loop {
                    match sr.read_event().map_err(|e| e.to_string())? {
                        Event::Start(e) | Event::Empty(e)
                            if e.local_name().as_ref() == b"tablePart" =>
                        {
                            ids.push(attr(&e, b"id")?.ok_or("Missing Table relationship")?)
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                }
                if ids.is_empty() {
                    continue;
                }
                let sheet_rels = part(&mut zip, &rel_path(&rel.target))
                    .and_then(|xml| relationships(&xml, &rel.target));
                for id in ids {
                    let imported = (|| {
                        let rs = sheet_rels.as_ref().map_err(|e| e.clone())?;
                        let rel = rs.get(&id).ok_or("Missing Table relationship")?;
                        if !rel.kind.ends_with("/table") || rel.external {
                            return Err("Unsupported Table relationship".into());
                        }
                        if !seen.insert(rel.target.clone()) {
                            return Err("Duplicate Table part reference".into());
                        }
                        parse_table(&part(&mut zip, &rel.target)?)
                    })();
                    let outcome = imported.and_then(|mut imported| {
                        if result.truncated { return Err("Workbook import was truncated; Table definitions cannot safely be restored".into()); }
                        let t = &mut imported.table;
                        let sheet = wb.sheet(si).unwrap();
                        for (i, column) in t.columns.iter_mut().enumerate() {
                            if sheet.get_raw(t.range.start_row, t.range.start_col + i) != column.name { return Err(format!("Table {} has missing or inconsistent header cells", t.name)); }
                            if values_only { column.formula = None; }
                        }
                        let mut catalog = wb.saved_tables();
                        catalog.version = catalog.version.max(2);
                        t.id = TableId(catalog.next_table_id);
                        catalog.next_table_id += 1;
                        if let Some(entry) = catalog.sheets.iter_mut().find(|s| s.sheet == si) { entry.tables.push(t.clone()); }
                        else { catalog.sheets.push(SavedTableSheet { sheet: si, tables: vec![t.clone()], column_allocators: BTreeMap::new(), view: None }); }
                        wb.restore_tables(catalog)?;
                        result.tables_imported += 1;
                        for warning in imported.warnings { result.warnings.push(format!("Table {}: {warning}", t.name)); }
                        views.push(super::xlsx_table_filters::PendingView { sheet: si, table: t.id, view: imported.view });
                        Ok(())
                    });
                    if let Err(error) = outcome {
                        result.tables_skipped += 1;
                        result.warnings.push(format!("Excel Table on {name} was kept as plain cells: {error}. Formulas referencing it may show errors."));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(())
}

pub(crate) fn export_warnings(wb: &Workbook) -> Result<Vec<String>, String> {
    wb.validate_table_view_specs()?;
    let mut warnings = Vec::new();
    for (sheet_id, table) in wb.tables() {
        if table.name.chars().count() > 255 {
            return Err(format!(
                "Table {} has a name longer than Excel's 255-character limit; no file was written.",
                table.name
            ));
        }
        if table.range.data_rows() == 0 {
            return Err(format!("Table {} has only headers. Add an empty record before exporting it to Excel; no file was written.", table.name));
        }
        if let Some(view) = wb
            .sheet_by_id(sheet_id)
            .and_then(|s| s.table_view_spec())
            .filter(|v| v.table == table.id)
        {
            if view.sort.is_some() {
                warnings.push(format!("Table {}: Excel export includes every record in stored order. VisiGrid view sorting is not exported; formulas keep their stored coordinates.", table.name));
            }
            if let Err(reason) =
                super::xlsx_table_filters::export_filters(wb.sheet_by_id(sheet_id).unwrap(), table)
            {
                warnings.push(format!("Table {}: filter criteria are not exported ({reason}). All records are shown in Excel.", table.name));
            }
        }
    }
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
            .set_columns(&columns)
            .set_style(rust_xlsxwriter::TableStyle::Medium2)
            .set_banded_rows(table.style.banded_rows)
            .set_autofilter(
                arrows
                    || sheet
                        .table_view_spec()
                        .is_some_and(|v| v.table == table.id && !v.filters.is_empty()),
            );
        let r = table.range;
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
    out
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
                    if e.local_name().as_ref() == b"tableColumn" {
                        if let Some(t) = table {
                            if let Some(formula) =
                                t.formula_at(t.range.start_row + 1, t.range.start_col + offset)
                            {
                                writer
                                    .write_event(Event::Start(e.clone()))
                                    .map_err(|e| e.to_string())?;
                                writer
                                    .write_event(Event::Start(BytesStart::new(
                                        "calculatedColumnFormula",
                                    )))
                                    .map_err(|e| e.to_string())?;
                                writer
                                    .write_event(Event::Text(BytesText::new(
                                        excel_formula(&formula).trim_start_matches('='),
                                    )))
                                    .map_err(|e| e.to_string())?;
                                writer
                                    .write_event(Event::End(quick_xml::events::BytesEnd::new(
                                        "calculatedColumnFormula",
                                    )))
                                    .map_err(|e| e.to_string())?;
                                if matches!(event, Event::Empty(_)) {
                                    writer
                                        .write_event(Event::End(e.to_end()))
                                        .map_err(|e| e.to_string())?;
                                }
                                written.insert((t.id, offset));
                                offset += 1;
                                continue;
                            }
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
