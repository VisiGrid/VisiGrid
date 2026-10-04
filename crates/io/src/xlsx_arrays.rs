//! Cached array members are derived values, not spill obstructions.
use quick_xml::{events::Event, Reader};
use std::{collections::HashSet, fs::File, io::BufReader, path::Path};
use visigrid_engine::{
    cell::{Cell, CellValue},
    formula::eval::Value,
    table::TableRange,
    workbook::Workbook,
};

#[derive(Clone, Copy)]
struct Array {
    sheet: usize,
    range: TableRange,
}
pub(crate) struct Pending {
    array: Array,
    cells: Vec<(usize, usize, Cell)>,
}
const MAX_ARRAYS: usize = 4096;
const MAX_ARRAY_CELLS: usize = 5_000_000;

// Excel namespaces function identifiers, not text, sheet names or Table fields.
// Scan before parsing: combined _xlfn._xlws identifiers exceed the native
// parser's grammar. Preserve all other bytes, including formula grouping.
pub(crate) fn normalize_formula(source: &str) -> String {
    if !source
        .as_bytes()
        .windows(6)
        .any(|s| s.eq_ignore_ascii_case(b"_xlfn.") || s.eq_ignore_ascii_case(b"_xlws."))
    {
        return source.to_owned();
    }
    let mut output = String::with_capacity(source.len());
    let mut i = 0;
    let mut brackets = 0usize;
    while i < source.len() {
        let start = i;
        let ch = source[i..].chars().next().unwrap();
        i += ch.len_utf8();
        if brackets > 0 {
            match ch {
                // Apostrophes escape special characters in structured headers.
                '\'' => {
                    if let Some(next) = source[i..].chars().next() {
                        i += next.len_utf8();
                    }
                }
                '[' => brackets += 1,
                ']' => brackets -= 1,
                _ => {}
            }
        } else if ch == '[' {
            brackets = 1;
        } else if ch == '"' || ch == '\'' {
            // Strings and quoted worksheet names use doubled delimiter escapes.
            while i < source.len() {
                let next = source[i..].chars().next().unwrap();
                i += next.len_utf8();
                if next == ch {
                    if source[i..].starts_with(ch) {
                        i += ch.len_utf8();
                    } else {
                        break;
                    }
                }
            }
        } else if ch.is_alphanumeric() || ch == '_' || ch == '\\' || ch == '.' {
            while let Some(next) = source[i..].chars().next() {
                if next.is_alphanumeric() || next == '_' || next == '\\' || next == '.' {
                    i += next.len_utf8();
                } else {
                    break;
                }
            }
            let mut name = &source[start..i];
            if source[i..].trim_start().starts_with('(') {
                for prefix in ["_xlfn.", "_xlws."] {
                    if name
                        .get(..prefix.len())
                        .is_some_and(|s| s.eq_ignore_ascii_case(prefix))
                    {
                        name = &name[prefix.len()..];
                    }
                }
            }
            output.push_str(name);
            continue;
        }
        output.push_str(&source[start..i]);
    }
    output
}

#[cfg(test)]
mod tests {
    use super::normalize_formula;
    #[test]
    fn function_namespaces_leave_literals_and_reference_names_intact() {
        assert_eq!(
            normalize_formula("=_xlfn._xlws.FILTER(A1:A2,A1:A2>0)"),
            "=FILTER(A1:A2,A1:A2>0)"
        );
        assert_eq!(
            normalize_formula("=_xlfn.SEQUENCE(2)+LEN(\"_xlfn.SEQUENCE(2)\")"),
            "=SEQUENCE(2)+LEN(\"_xlfn.SEQUENCE(2)\")"
        );
        let source = "='_xlfn.SEQUENCE'!A1+Sales[_xlfn.SEQUENCE]+_xlfn.Name";
        assert_eq!(normalize_formula(source), source);
        assert_eq!(
            normalize_formula("=(_xlfn.SEQUENCE(2)+1)*2"),
            "=(SEQUENCE(2)+1)*2"
        );
        assert_eq!(
            normalize_formula(
                "=_xlfn.SEQUENCE(2)+LEN(\"é\"\"_xlfn.SEQUENCE(2)\")+Sales[a']_xlfn.SEQUENCE(2)]"
            ),
            "=SEQUENCE(2)+LEN(\"é\"\"_xlfn.SEQUENCE(2)\")+Sales[a']_xlfn.SEQUENCE(2)]"
        );
    }
}

fn read(path: &Path) -> Result<Vec<Array>, String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let Ok(mut zip) = zip::ZipArchive::new(file) else {
        return Ok(Vec::new());
    };
    let Some(workbook) = super::xlsx::read_zip_file_for_shared(&mut zip, "xl/workbook.xml") else {
        return Ok(Vec::new());
    };
    let Some(rels) = super::xlsx::read_zip_file_for_shared(&mut zip, "xl/_rels/workbook.xml.rels")
    else {
        return Ok(Vec::new());
    };
    let mut arrays = Vec::new();
    let mut cells = 0usize;
    for (sheet, path) in super::xlsx::resolve_worksheet_paths(&workbook, &rels)
        .iter()
        .enumerate()
    {
        let entry = zip.by_name(path).map_err(|e| e.to_string())?;
        let mut reader = Reader::from_reader(BufReader::new(entry));
        let mut buffer = Vec::new();
        let mut anchor = None;
        loop {
            buffer.clear();
            match reader
                .read_event_into(&mut buffer)
                .map_err(|e| e.to_string())?
            {
                Event::Start(e) if e.local_name().as_ref() == b"c" => {
                    anchor = super::xlsx_tables::attr(&e, b"r")?
                        .and_then(|s| super::xlsx_tables::range(&s).ok());
                }
                Event::End(e) if e.local_name().as_ref() == b"c" => anchor = None,
                Event::Start(e) | Event::Empty(e)
                    if e.local_name().as_ref() == b"f"
                        && super::xlsx_tables::attr(&e, b"t")?.as_deref() == Some("array") =>
                {
                    let target = super::xlsx_tables::attr(&e, b"ref")?
                        .ok_or("Array formula has no result range")?;
                    let range = super::xlsx_tables::range(&target)?;
                    let anchor = anchor.ok_or("Array formula has no anchor cell")?;
                    if (anchor.start_row, anchor.start_col) != (range.start_row, range.start_col) {
                        return Err("Array master is not at the start of its range".into());
                    }
                    cells = cells
                        .checked_add(
                            (range.end_row - range.start_row + 1)
                                .checked_mul(range.end_col - range.start_col + 1)
                                .ok_or("Array range is too large")?,
                        )
                        .ok_or("Array ranges are too large")?;
                    if arrays.len() >= MAX_ARRAYS || cells > MAX_ARRAY_CELLS {
                        return Err(
                            "Array metadata exceeds 4,096 arrays or 5,000,000 result cells".into(),
                        );
                    }
                    arrays.push(Array { sheet, range });
                }
                Event::Eof => break,
                _ => {}
            }
        }
    }
    Ok(arrays)
}

pub(crate) fn prepare(
    path: &Path,
    wb: &mut Workbook,
    warnings: &mut Vec<String>,
) -> Result<Vec<Pending>, String> {
    let arrays = match read(path) {
        Ok(a) => a,
        Err(e) => {
            warnings.push(format!("Array caches were kept as cells: {e}"));
            return Ok(Vec::new());
        }
    };
    let mut overlaps = HashSet::new();
    for (i, a) in arrays.iter().enumerate() {
        for (j, b) in arrays.iter().enumerate().skip(i + 1) {
            if a.sheet == b.sheet && a.range.intersects(b.range) {
                overlaps.insert(i);
                overlaps.insert(j);
            }
        }
    }
    let mut pending = Vec::new();
    // Validate every affected cell before changing the fresh import workbook.
    for (i, array) in arrays.into_iter().enumerate() {
        let r = array.range;
        let Some(sheet) = wb.sheet(array.sheet) else {
            continue;
        };
        let label = format!(
            "{}!{}{}",
            sheet.name,
            super::xlsx::col_to_letter(r.start_col),
            r.start_row + 1
        );
        let check = (|| -> Result<Option<Pending>, String> {
            if overlaps.contains(&i) {
                return Err("overlapping array ranges".into());
            }
            if r.end_row >= sheet.rows || r.end_col >= sheet.cols {
                return Err("range exceeds the imported sheet".into());
            }
            let anchor = sheet.get_cell(r.start_row, r.start_col);
            if !matches!(anchor.value, CellValue::Formula { ast: Some(_), .. }) {
                return Err("anchor is not a supported formula".into());
            }
            if sheet.get_spill_info(r.start_row, r.start_col).is_some() {
                return Ok(None);
            }
            if sheet.tables().iter().any(|t| t.full_range().intersects(r))
                || sheet.merged_regions.iter().any(|m| {
                    r.intersects(TableRange {
                        start_row: m.start.0,
                        start_col: m.start.1,
                        end_row: m.end.0,
                        end_col: m.end.1,
                    })
                })
            {
                return Err("range intersects a Table or merged cells".into());
            }
            let mut cells = Vec::new();
            for (row, col) in sheet.cells_in_range(r.start_row, r.end_row, r.start_col, r.end_col) {
                if (row, col) == (r.start_row, r.start_col) {
                    continue;
                }
                let cell = sheet.get_cell(row, col);
                if matches!(cell.value, CellValue::Formula { .. })
                    || cell.frozen_formula().is_some()
                    || cell.is_spill_receiver()
                    || cell.is_spill_parent()
                    || sheet.is_pivot_owned(row, col)
                {
                    return Err("range includes independent formulas or owned cells".into());
                }
                cells.push((row, col, cell));
            }
            Ok(Some(Pending { array, cells }))
        })();
        match check {
            Ok(Some(p)) => pending.push(p),
            Ok(None) => {}
            Err(e) => warnings.push(format!(
                "Array {label}: cached cells were kept ({e}); use values-only import if needed."
            )),
        }
    }
    let auto = wb.auto_recalc();
    wb.set_auto_recalc(false);
    let cleared = (|| -> Result<(), String> {
        for p in &pending {
            for (row, col, cell) in &p.cells {
                let mut blank = cell.clone();
                blank.value = CellValue::Empty;
                wb.restore_cell_tracked(p.array.sheet, *row, *col, Some(blank))?;
            }
        }
        Ok(())
    })();
    wb.set_auto_recalc(auto);
    cleared?;
    Ok(pending)
}

// Unsupported or failed array formulas must not erase their cached members.
// Supported results may shrink/grow through the ordinary collision checks.
pub(crate) fn finish(
    pending: Vec<Pending>,
    wb: &mut Workbook,
    warnings: &mut Vec<String>,
) -> Result<bool, String> {
    let mut failed = Vec::new();
    for p in pending {
        let r = p.array.range;
        let sheet = wb.sheet(p.array.sheet).unwrap();
        if sheet.get_spill_info(r.start_row, r.start_col).is_none()
            && (matches!(
                sheet.get_computed_value(r.start_row, r.start_col),
                Value::Error(_)
            ) || !matches!(
                sheet.get_cell(r.start_row, r.start_col).value,
                CellValue::Formula { .. }
            ) || r.end_row > r.start_row
                || r.end_col > r.start_col)
        {
            for (row, col, _) in &p.cells {
                if sheet.is_spill_receiver(*row, *col) || sheet.is_spill_parent(*row, *col) {
                    return Err("An array's cached results conflict with another recalculated spill. Import values only to preserve the file's results.".into());
                }
            }
            warnings.push(format!("Array {}!{}{} could not be recalculated; its cached result cells were kept. Values-only import also retains the cached anchor result.", sheet.name, super::xlsx::col_to_letter(r.start_col), r.start_row + 1));
            failed.push(p);
        }
    }
    let restored = !failed.is_empty();
    let auto = wb.auto_recalc();
    wb.set_auto_recalc(false);
    let result = (|| -> Result<(), String> {
        for p in failed {
            for (row, col, cell) in p.cells {
                wb.restore_cell_tracked(p.array.sheet, row, col, Some(cell))?;
            }
        }
        Ok(())
    })();
    wb.set_auto_recalc(auto);
    result?;
    Ok(restored)
}
