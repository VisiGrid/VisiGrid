//! Write typed last-calculated results without evaluating or mutating the source.
//! The writer's string-result API guesses types, so patch its generated XML.
use quick_xml::{
    events::{BytesEnd, BytesStart, BytesText, Event},
    Reader, Writer,
};
use std::{
    collections::{HashMap, HashSet},
    fs::File,
    io::{BufRead, BufReader, Cursor, Write},
    path::Path,
};
use visigrid_engine::{cell::ValueRef, formula::eval::Value, sheet::Sheet, workbook::Workbook};

#[derive(Default)]
struct Omitted {
    missing: usize,
    unsupported: usize,
}
struct Cache {
    kind: Option<&'static str>,
    value: Option<String>,
}
fn cached(sheet: &Sheet, row: usize, col: usize, omitted: &mut Omitted) -> Cache {
    let value = if sheet
        .get_cell_opt(row, col)
        .is_some_and(|cell| cell.has_spill_error())
    {
        Some(Value::Error("#SPILL!".into()))
    } else if let Some(value) = sheet.get_spill_value(row, col) {
        Some(value.clone())
    } else {
        sheet.get_cached_value(row, col)
    };
    let (kind, value) = match value {
        Some(Value::Number(n)) if n.is_finite() => ("n", n.to_string()),
        Some(Value::Boolean(b)) => ("b", if b { "1" } else { "0" }.into()),
        Some(Value::Text(s)) => ("str", s),
        Some(Value::Empty) => ("str", String::new()),
        Some(Value::Error(e)) => {
            let code = [
                "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!",
            ]
            .into_iter()
            .find(|code| {
                e.strip_prefix(code).is_some_and(|rest| {
                    rest.is_empty()
                        || rest.starts_with(char::is_whitespace)
                        || rest.starts_with(':')
                })
            });
            if let Some(code) = code {
                ("e", code.into())
            } else if e.starts_with("Cannot convert ")
                && (e.ends_with(" to number") || e.ends_with(" to boolean"))
            {
                ("e", "#VALUE!".into())
            } else if e.starts_with("Unknown function") {
                ("e", "#NAME?".into())
            } else {
                omitted.unsupported += 1;
                return Cache {
                    kind: None,
                    value: None,
                };
            }
        }
        Some(Value::Number(_)) => {
            omitted.unsupported += 1;
            return Cache {
                kind: None,
                value: None,
            };
        }
        None => {
            omitted.missing += 1;
            return Cache {
                kind: None,
                value: None,
            };
        }
    };
    Cache {
        kind: Some(kind),
        value: Some(value),
    }
}
fn exported_formula(sheet: &Sheet, row: usize, col: usize) -> bool {
    !sheet.is_merge_hidden(row, col)
        && sheet.get_cell_opt(row, col).is_some_and(|cell| {
            cell.is_spill_receiver()
                || matches!(cell.value(), ValueRef::Formula { ast: Some(_), .. })
        })
}

fn patch<R: BufRead, W: Write>(
    source: R,
    output: W,
    sheet: &Sheet,
    omitted: &mut Omitted,
) -> Result<(), String> {
    let mut reader = Reader::from_reader(source);
    let mut writer = Writer::new(output);
    let mut buffer = Vec::new();
    let mut skip = Vec::new();
    let mut cache: Option<Cache> = None;
    loop {
        buffer.clear();
        let mut event = reader
            .read_event_into(&mut buffer)
            .map_err(|e| e.to_string())?;
        match &mut event {
            Event::Start(e) if e.local_name().as_ref() == b"c" => {
                cache = None;
                if let Some(address) = super::xlsx_tables::attr(e, b"r")? {
                    let range = super::xlsx_tables::range(&address)?;
                    if exported_formula(sheet, range.start_row, range.start_col) {
                        let result = cached(sheet, range.start_row, range.start_col, omitted);
                        let attrs = e
                            .attributes()
                            .map(|a| a.map(|a| (a.key.as_ref().to_vec(), a.value.into_owned())))
                            .collect::<Result<Vec<_>, _>>()
                            .map_err(|e| e.to_string())?;
                        e.clear_attributes();
                        for (key, value) in &attrs {
                            if key != b"t" {
                                e.push_attribute((key.as_slice(), value.as_slice()));
                            }
                        }
                        if let Some(kind) = result.kind {
                            e.push_attribute(("t", kind));
                        }
                        cache = Some(result);
                    }
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"v" && cache.is_some() => {
                reader
                    .read_to_end_into(e.name(), &mut skip)
                    .map_err(|e| e.to_string())?;
                if let Some(value) = &cache.as_ref().unwrap().value {
                    writer
                        .write_event(Event::Start(BytesStart::new("v")))
                        .map_err(|e| e.to_string())?;
                    // Includes SpreadsheetML UTF-16 escapes and literal escape protection.
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(
                            super::xlsx_comments::xml_text(value),
                        )))
                        .map_err(|e| e.to_string())?;
                    writer
                        .write_event(Event::End(BytesEnd::new("v")))
                        .map_err(|e| e.to_string())?;
                }
                continue;
            }
            Event::End(e) if e.local_name().as_ref() == b"c" => cache = None,
            Event::Eof => break,
            _ => {}
        }
        writer.write_event(event).map_err(|e| e.to_string())?;
    }
    Ok(())
}
fn omission_warnings(omitted: &Omitted) -> Vec<String> {
    if omitted.missing + omitted.unsupported == 0 {
        return Vec::new();
    }
    vec![format!("{} formula results have no usable XLSX cache ({} not calculated; {} unsupported errors or non-finite values). Their formulas are retained; recalculate in Excel. Values-only readers may show blanks.", omitted.missing + omitted.unsupported, omitted.missing, omitted.unsupported)]
}

pub(crate) fn warnings(wb: &Workbook) -> Vec<String> {
    let mut omitted = Omitted::default();
    for sheet in wb.sheets() {
        for ((r, c), _) in sheet.cells_iter() {
            if exported_formula(sheet, r, c) {
                cached(sheet, r, c, &mut omitted);
            }
        }
    }
    omission_warnings(&omitted)
}

pub(crate) fn finish(
    bytes: Vec<u8>,
    wb: &Workbook,
    warnings: &mut Vec<String>,
) -> Result<Vec<u8>, String> {
    let sheets: HashMap<_, _> = wb
        .sheets()
        .iter()
        .enumerate()
        .filter(|(_, s)| s.cells_iter().any(|((r, c), _)| exported_formula(s, r, c)))
        .map(|(i, s)| (format!("xl/worksheets/sheet{}.xml", i + 1), s))
        .collect();
    if sheets.is_empty() {
        return Ok(bytes);
    }
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| e.to_string())?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut omitted = Omitted::default();
    for i in 0..input.len() {
        let entry = input.by_index(i).map_err(|e| e.to_string())?;
        if let Some(sheet) = sheets.get(entry.name()) {
            output
                .start_file(
                    entry.name(),
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(|e| e.to_string())?;
            patch(BufReader::new(entry), &mut output, sheet, &mut omitted)?;
        } else {
            output.raw_copy_file(entry).map_err(|e| e.to_string())?;
        }
    }
    warnings.extend(omission_warnings(&omitted));
    Ok(output.finish().map_err(|e| e.to_string())?.into_inner())
}

// Calamine decodes SpreadsheetML escapes in shared/inline strings but not in
// formula and array-receiver <v> strings. Identify that exact storage type before decoding once.
// Called lazily only when imported text actually contains an escape candidate.
pub(crate) fn string_cells(path: &Path) -> Result<HashSet<(usize, usize, usize)>, String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let Ok(mut zip) = zip::ZipArchive::new(file) else {
        return Ok(HashSet::new());
    };
    let Some(workbook) = super::xlsx::read_zip_file_for_shared(&mut zip, "xl/workbook.xml") else {
        return Ok(HashSet::new());
    };
    let Some(rels) = super::xlsx::read_zip_file_for_shared(&mut zip, "xl/_rels/workbook.xml.rels")
    else {
        return Ok(HashSet::new());
    };
    let mut cells = HashSet::new();
    for (sheet, path) in super::xlsx::resolve_worksheet_paths(&workbook, &rels)
        .iter()
        .enumerate()
    {
        let entry = zip.by_name(path).map_err(|e| e.to_string())?;
        let mut reader = Reader::from_reader(BufReader::new(entry));
        let mut buffer = Vec::new();
        let mut coord = None;
        loop {
            buffer.clear();
            match reader
                .read_event_into(&mut buffer)
                .map_err(|e| e.to_string())?
            {
                Event::Start(e) if e.local_name().as_ref() == b"c" => {
                    coord = if super::xlsx_tables::attr(&e, b"t")?.as_deref() == Some("str") {
                        super::xlsx_tables::attr(&e, b"r")?
                            .and_then(|a| super::xlsx_tables::range(&a).ok())
                            .map(|r| (sheet, r.start_row, r.start_col))
                    } else {
                        None
                    };
                }
                Event::End(e) if e.local_name().as_ref() == b"c" => {
                    if let Some(c) = coord {
                        cells.insert(c);
                    }
                    coord = None;
                }
                Event::Eof => break,
                _ => {}
            }
        }
    }
    Ok(cells)
}
