//! Workbook-scoped, absolute cell/range names. Scope and expression kinds
//! absent from the engine model are reported, never flattened silently.
use quick_xml::{events::Event, Reader, Writer};
use std::{
    collections::HashMap,
    fs::File,
    io::{Cursor, Read, Write},
    path::Path,
};
use visigrid_engine::{
    formula::parser::{self, Expr},
    named_range::{NamedRange, NamedRangeTarget},
    sheet::{UnboundSheetRef, NUM_COLS, NUM_ROWS},
    workbook::Workbook,
};

const MAX_BYTES: u64 = 8 * 1024 * 1024;
const MAX_NAMES: usize = 8192;

struct Definition {
    name: String,
    formula: String,
    description: Option<String>,
    local: bool,
    hidden: bool,
}
fn read(path: &Path) -> Result<Vec<Definition>, String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let Ok(mut zip) = zip::ZipArchive::new(file) else {
        return Ok(Vec::new());
    };
    let Ok(part) = zip.by_name("xl/workbook.xml") else {
        return Ok(Vec::new());
    };
    if part.size() > MAX_BYTES {
        return Err("Workbook XML exceeds the 8 MiB defined-name limit".into());
    }
    let mut xml = String::new();
    part.take(MAX_BYTES + 1)
        .read_to_string(&mut xml)
        .map_err(|e| e.to_string())?;
    if xml.len() as u64 > MAX_BYTES {
        return Err("Workbook XML exceeds the 8 MiB defined-name limit".into());
    }
    let mut reader = Reader::from_str(&xml);
    reader.config_mut().expand_empty_elements = true;
    let mut names = Vec::new();
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) if e.local_name().as_ref() == b"definedName" => {
                if names.len() == MAX_NAMES {
                    return Err("Workbook exceeds the 8,192 defined-name limit".into());
                }
                let attr = |key| super::xlsx_tables::attr(&e, key);
                let name = attr(b"name")?.ok_or("Defined name has no name")?;
                let description = attr(b"comment")?;
                let local = attr(b"localSheetId")?.is_some();
                let hidden = matches!(attr(b"hidden")?.as_deref(), Some("1" | "true"));
                let text = reader.read_text(e.name()).map_err(|e| e.to_string())?;
                let formula = quick_xml::escape::unescape(&text)
                    .map_err(|e| e.to_string())?
                    .into_owned();
                names.push(Definition {
                    name,
                    formula,
                    description,
                    local,
                    hidden,
                });
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(names)
}
fn target(wb: &Workbook, name: &str, formula: &str) -> Result<NamedRange, String> {
    let source = format!("={}", formula.trim().trim_start_matches('='));
    let sheet_index = |reference: UnboundSheetRef| -> Result<usize, String> {
        let UnboundSheetRef::Named(name) = reference else {
            return Err("unqualified target is relative to a sheet".into());
        };
        wb.sheets()
            .iter()
            .position(|s| s.name.eq_ignore_ascii_case(&name))
            .ok_or_else(|| "target sheet is missing or external".into())
    };
    let bounded = |r: usize, c: usize| r < NUM_ROWS && c < NUM_COLS;
    match parser::parse(&source)? {
        Expr::RefError => Ok(NamedRange { name: name.into(), target: NamedRangeTarget::RefError, description: None }),
        Expr::CellRef { sheet, row, col, row_abs: true, col_abs: true } if bounded(row, col) => Ok(NamedRange::cell(name, sheet_index(sheet)?, row, col)),
        Expr::Range { sheet, start_row, start_col, end_row, end_col, start_row_abs: true, start_col_abs: true, end_row_abs: true, end_col_abs: true }
            if bounded(start_row, start_col) && bounded(end_row, end_col) && start_row <= end_row && start_col <= end_col =>
            Ok(NamedRange::range(name, sheet_index(sheet)?, start_row, start_col, end_row, end_col)),
        _ => Err("only absolute cells and rectangular ranges are supported; relative, formula, constant, union and whole-row/column names are not imported".into()),
    }
}
pub(crate) fn import(path: &Path, wb: &mut Workbook, warnings: &mut Vec<String>) {
    let names = match read(path) {
        Ok(names) => names,
        Err(e) => {
            warnings.push(format!("Defined names were not imported: {e}"));
            return;
        }
    };
    let mut counts = HashMap::new();
    for n in &names {
        *counts.entry(n.name.to_lowercase()).or_insert(0) += 1;
    }
    for n in names {
        // These are worksheet settings, not user-defined engine names.
        if matches!(
            n.name.as_str(),
            "_xlnm.Print_Area" | "_xlnm.Print_Titles" | "_xlnm._FilterDatabase"
        ) {
            continue;
        }
        let reason = if n.name.chars().count() > 255
            || n.name.to_ascii_lowercase().starts_with("_xlnm.")
        {
            Some("invalid or unsupported reserved Excel name".to_string())
        } else if n.local {
            Some("worksheet-scoped names are not supported".to_string())
        } else if counts[&n.name.to_lowercase()] != 1 {
            Some(
                "duplicate or shadowed name; importing it would change name resolution".to_string(),
            )
        } else {
            None
        };
        let result = match reason {
            Some(reason) => Err(reason),
            None => target(wb, &n.name, &n.formula).and_then(|mut name| {
                name.description = n.description;
                wb.named_ranges_mut().set(name)
            }),
        };
        match result {
            Err(e) => warnings.push(format!("Defined name '{}' was not imported: {e}", n.name)),
            Ok(()) if n.hidden => warnings.push(format!(
                "Defined name '{}' was imported as a visible name; hidden names are not supported",
                n.name
            )),
            _ => {}
        }
    }
}
pub(crate) fn export_warnings(wb: &Workbook) -> Vec<String> {
    wb.named_ranges().list().into_iter().filter_map(|name| {
        visigrid_engine::named_range::is_valid_name(&name.name).err().map(|reason|
            format!("Defined name '{}' is omitted from Excel export: {reason} Formulas using it may show an error in Excel; the native definition is unchanged.", name.name))
    }).collect()
}

pub(crate) fn export(wb: &Workbook, out: &mut rust_xlsxwriter::Workbook) -> Result<(), String> {
    for name in wb.named_ranges().list() {
        // Legacy native files can retain identifiers that new creation rejects
        // (notably R1C1 names). Report the omission in review and ExportResult.
        if visigrid_engine::named_range::is_valid_name(&name.name).is_err() { continue; }
        if name.name.chars().count() > 255 || name.name.to_ascii_lowercase().starts_with("_xlnm.") {
            return Err(format!(
                "Defined name '{}' is not a valid Excel user name",
                name.name
            ));
        }
        let (sheet, r0, c0, r1, c1, range) = match name.target {
            NamedRangeTarget::RefError => {
                out.define_name(&name.name, "=#REF!").map_err(|e| e.to_string())?;
                continue;
            },
            NamedRangeTarget::Cell { sheet, row, col } => (sheet, row, col, row, col, false),
            NamedRangeTarget::Range {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
            } => (sheet, start_row, start_col, end_row, end_col, true),
        };
        let Some(sheet) = wb.sheet(sheet) else {
            // A name left pointing at a sheet that is already gone, for
            // example after a delete that predates name invalidation.
            out.define_name(&name.name, "=#REF!")
                .map_err(|e| e.to_string())?;
            continue;
        };
        if r0 > r1 || c0 > c1 || r1 >= NUM_ROWS || c1 >= NUM_COLS {
            return Err(format!(
                "Defined name '{}' has an invalid Excel target",
                name.name
            ));
        }
        let address = |r, c| format!("${}${}", super::xlsx::col_to_letter(c), r + 1);
        let mut reference = format!("='{}'!{}", sheet.name.replace('\'', "''"), address(r0, c0));
        if range {
            reference.push_str(&format!(":{}", address(r1, c1)));
        }
        out.define_name(&name.name, &reference)
            .map_err(|e| format!("Cannot export defined name '{}': {e}", name.name))?;
    }
    Ok(())
}
// rust_xlsxwriter's public name API has no description field. Patch only the
// workbook part of our own generated package, copying all other ZIP parts raw.
pub(crate) fn finish(bytes: Vec<u8>, wb: &Workbook) -> Result<Vec<u8>, String> {
    let descriptions: HashMap<_, _> = wb
        .named_ranges()
        .list()
        .into_iter()
        .filter_map(|n| {
            n.description
                .as_ref()
                .map(|d| (n.name.as_str(), d.as_str()))
        })
        .collect();
    if descriptions.is_empty() {
        return Ok(bytes);
    }
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| e.to_string())?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..input.len() {
        let mut entry = input.by_index(i).map_err(|e| e.to_string())?;
        if entry.name() != "xl/workbook.xml" {
            output.raw_copy_file(entry).map_err(|e| e.to_string())?;
            continue;
        }
        let mut xml = String::new();
        entry.read_to_string(&mut xml).map_err(|e| e.to_string())?;
        let mut reader = Reader::from_str(&xml);
        let mut writer = Writer::new(Vec::new());
        loop {
            let mut event = reader.read_event().map_err(|e| e.to_string())?;
            if let Event::Start(e) | Event::Empty(e) = &mut event {
                if e.local_name().as_ref() == b"definedName" {
                    if let Some(name) = super::xlsx_tables::attr(e, b"name")? {
                        if let Some(comment) = descriptions.get(name.as_str()) {
                            e.push_attribute(("comment", *comment));
                        }
                    }
                }
            }
            if matches!(event, Event::Eof) {
                break;
            }
            writer.write_event(event).map_err(|e| e.to_string())?;
        }
        output
            .start_file(
                "xl/workbook.xml",
                zip::write::SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Deflated),
            )
            .map_err(|e| e.to_string())?;
        output
            .write_all(&writer.into_inner())
            .map_err(|e| e.to_string())?;
    }
    Ok(output.finish().map_err(|e| e.to_string())?.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn package(xml: &str) -> tempfile::NamedTempFile {
        let file = tempfile::NamedTempFile::new().unwrap();
        let mut zip = zip::ZipWriter::new(File::create(file.path()).unwrap());
        zip.start_file("xl/workbook.xml", zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
        zip.finish().unwrap();
        file
    }
    #[test]
    fn name_metadata_is_bounded_without_partial_publication() {
        let xml = format!(
            "<workbook><definedNames>{}</definedNames></workbook>",
            "<definedName name=\"Name\">Sheet1!$A$1</definedName>".repeat(MAX_NAMES + 1)
        );
        let file = package(&xml);
        let mut wb = Workbook::new();
        let mut warnings = Vec::new();
        import(file.path(), &mut wb, &mut warnings);
        assert!(wb.named_ranges().is_empty());
        assert!(warnings[0].contains("8,192"));
        let file = package(&" ".repeat(MAX_BYTES as usize + 1));
        assert!(read(file.path()).err().unwrap().contains("8 MiB"));
    }
    #[test]
    fn hidden_names_warn_and_malformed_xml_does_not_install_earlier_definitions() {
        let file = package("<workbook><definedNames><definedName name=\"Private\" hidden=\"1\">Sheet1!$A$1</definedName></definedNames></workbook>");
        let mut wb = Workbook::new();
        let mut warnings = Vec::new();
        import(file.path(), &mut wb, &mut warnings);
        assert!(wb.named_ranges().get("Private").is_some());
        assert!(warnings[0].contains("visible name"));
        let file = package("<workbook><definedNames><definedName name=\"Good\">Sheet1!$A$1</definedName><definedName name=\"Broken\">Sheet1!$B$1");
        import(file.path(), &mut wb, &mut warnings);
        assert!(wb.named_ranges().get("Good").is_none());
    }
}
