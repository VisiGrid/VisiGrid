//! Formula conditional formatting. The writer owns worksheet placement, formula
//! namespace encoding and global priorities. Patch its generated rules/styles
//! because its Format API drops explicit false font flags from differential
//! formats. No-rule workbooks take the byte-preserving fast path.
use quick_xml::{
    events::{BytesEnd, BytesStart, Event},
    Reader, Writer,
};
use rust_xlsxwriter::{ConditionalFormatFormula, Worksheet};
use std::{
    collections::{BTreeMap, BTreeSet},
    io::{Cursor, Read, Write},
};
use visigrid_engine::{
    cell::{BorderStyle, CellBorder, CellFormatOverride, CellStyle},
    formula::parser,
    sheet::{Sheet, NUM_COLS, NUM_ROWS},
    validation::CellRange,
    workbook::Workbook,
};

const MAX_RANGES: usize = 100_000;
const MAX_WORK: usize = 10_000_000;
const MAX_TEXT: usize = 8 * 1024 * 1024;
struct Rule {
    range: CellRange,
    formula: String,
    dxf: String,
}
struct Plan {
    rules: Vec<Rule>,
    warnings: Vec<String>,
}
fn error() -> String {
    "Conditional formatting exceeds the supported export complexity limit; no file was written."
        .into()
}
fn esc(s: &str) -> String {
    quick_xml::escape::escape(s)
        .replace('\n', "&#xA;")
        .replace('\r', "&#xD;")
        .replace('\t', "&#x9;")
}
fn rgb(c: [u8; 4]) -> String {
    format!("FF{:02X}{:02X}{:02X}", c[0], c[1], c[2])
}
fn color(n: u32) -> [u8; 4] {
    [(n >> 16) as u8, (n >> 8) as u8, n as u8, 255]
}
fn named(style: CellStyle) -> CellFormatOverride {
    let mut result = CellFormatOverride::default();
    let (bg, fg, border) = match style {
        CellStyle::None => return result,
        CellStyle::Error => (0xFEE4E2, 0xB42318, Some(0xD92D20)),
        CellStyle::Warning => (0xFEF0C7, 0xB54708, Some(0xF79009)),
        CellStyle::Success => (0xECFDF3, 0x067647, Some(0x12B76A)),
        CellStyle::Input => (0xEFF8FF, 0x2E90FA, Some(0x2E90FA)),
        CellStyle::Note => {
            result.italic = Some(true);
            (0xF9FAFB, 0x475467, None)
        }
        CellStyle::Total => {
            result.bold = Some(true);
            result.background_color = Some(Some(color(0xE2E3E5)));
            result.border_top = Some(CellBorder {
                style: BorderStyle::Thin,
                color: Some(color(0x101828)),
            });
            return result;
        }
    };
    result.background_color = Some(Some(color(bg)));
    result.font_color = Some(Some(color(fg)));
    if let Some(c) = border {
        let b = Some(CellBorder {
            style: BorderStyle::Thin,
            color: Some(color(c)),
        });
        result.border_left = b;
        result.border_right = b;
        result.border_top = b;
        result.border_bottom = b;
    }
    result
}
fn style_xml(mut style: CellFormatOverride, notes: &mut BTreeSet<String>) -> String {
    if let Some(semantic) = style.cell_style.take() {
        if semantic != CellStyle::None {
            notes.insert("semantic styles flatten to fixed Ledger Light colors; theme appearance and interaction with explicit cell formatting can differ".into());
            let mut base = named(semantic);
            base.merge_from(&style);
            style = base;
        } else {
            notes.insert(
                "resetting a semantic style to its underlying cell style is not represented".into(),
            );
        }
    }
    if style.font_family.is_some()
        || style.font_size.is_some()
        || style.alignment.is_some()
        || style.vertical_alignment.is_some()
        || style.text_overflow.is_some()
    {
        notes.insert("font family/size, alignment and text overflow cannot be changed by Excel conditional formatting and are omitted".into());
    }
    let mut font = String::new();
    for (tag, value) in [
        ("b", style.bold),
        ("i", style.italic),
        ("strike", style.strikethrough),
    ] {
        if let Some(on) = value {
            font.push_str(&format!("<{tag} val=\"{}\"/>", u8::from(on)));
        }
    }
    if let Some(on) = style.underline {
        font.push_str(if on {
            "<u val=\"single\"/>"
        } else {
            "<u val=\"none\"/>"
        });
    }
    let mut color_tag = |value: Option<Option<[u8; 4]>>, tag: &str| -> String {
        match value {
            Some(Some(c)) => {
                if c[3] != 255 {
                    notes.insert("transparent colors are exported as opaque RGB".into());
                }
                format!("<{tag} rgb=\"{}\"/>", rgb(c))
            }
            Some(None) => {
                notes.insert("explicit color inheritance resets are omitted".into());
                String::new()
            }
            None => String::new(),
        }
    };
    font.push_str(&color_tag(style.font_color, "color"));
    let fill = color_tag(style.background_color, "bgColor");
    let mut xml = String::from("<dxf>");
    if !font.is_empty() {
        xml.push_str(&format!("<font>{font}</font>"));
    }
    if let Some(number) = style.number_format {
        let code = super::xlsx::excel_number_format(&number).unwrap_or_else(|| "General".into());
        // A DXF carries its own code; the numeric id is replaced with a fresh
        // workbook-wide id when appended to styles.xml.
        xml.push_str(&format!(
            "<numFmt numFmtId=\"0\" formatCode=\"{}\"/>",
            esc(&code)
        ));
    }
    if !fill.is_empty() {
        xml.push_str(&format!("<fill><patternFill>{fill}</patternFill></fill>"));
    }
    let mut borders = String::new();
    for (tag, border) in [
        ("left", style.border_left),
        ("right", style.border_right),
        ("top", style.border_top),
        ("bottom", style.border_bottom),
    ] {
        if let Some(b) = border {
            let name = match b.style {
                BorderStyle::None => "none",
                BorderStyle::Thin => "thin",
                BorderStyle::Medium => "medium",
                BorderStyle::Thick => "thick",
            };
            let c = b.color.unwrap_or([0, 0, 0, 255]);
            if c[3] != 255 {
                notes.insert("transparent colors are exported as opaque RGB".into());
            }
            borders.push_str(&format!(
                "<{tag} style=\"{name}\"><color rgb=\"{}\"/></{tag}>",
                rgb(c)
            ));
        }
    }
    if !borders.is_empty() {
        xml.push_str(&format!("<border>{borders}</border>"));
    }
    xml.push_str("</dxf>");
    xml
}
fn subtract(r: CellRange, cut: CellRange) -> Vec<CellRange> {
    if !r.overlaps(&cut) {
        return vec![r];
    }
    let (top, bottom, left, right) = (
        r.start_row.max(cut.start_row),
        r.end_row.min(cut.end_row),
        r.start_col.max(cut.start_col),
        r.end_col.min(cut.end_col),
    );
    let mut result = Vec::new();
    if r.start_row < top {
        result.push(CellRange {
            end_row: top - 1,
            ..r
        });
    }
    if bottom < r.end_row {
        result.push(CellRange {
            start_row: bottom + 1,
            ..r
        });
    }
    if r.start_col < left {
        result.push(CellRange::new(top, r.start_col, bottom, left - 1));
    }
    if right < r.end_col {
        result.push(CellRange::new(top, right + 1, bottom, r.end_col));
    }
    result
}
fn plan(sheet: &Sheet) -> Result<Plan, String> {
    let store: Vec<_> = sheet.cond_formats.iter().collect();
    if store.len() > MAX_RANGES
        || store
            .iter()
            .fold(0usize, |n, rule| n.saturating_add(rule.ranges.len()))
            > MAX_RANGES
        || store
            .iter()
            .fold(0usize, |n, rule| n.saturating_add(rule.predicate.len()))
            > MAX_TEXT
    {
        return Err(error());
    }
    let mut result = Plan {
        rules: Vec::new(),
        warnings: Vec::new(),
    };
    let mut work = MAX_WORK;
    let mut bytes = 0usize;
    let mut notes = BTreeSet::new();
    let mut skipped_disabled = 0;
    let mut skipped_bad = 0;
    // Lower Excel priorities win; VisiGrid applies later rules last.
    for rule in store.into_iter().rev() {
        if rule.ranges.len() > MAX_RANGES {
            return Err(error());
        }
        for r in &rule.ranges {
            if r.start_row > r.end_row
                || r.start_col > r.end_col
                || r.end_row >= sheet.rows.min(NUM_ROWS)
                || r.end_col >= sheet.cols.min(NUM_COLS)
            {
                return Err(format!(
                    "Conditional formatting on '{}' has an invalid range; no file was written.",
                    sheet.name
                ));
            }
        }
        if rule.ranges.is_empty() {
            continue;
        }
        if !rule.enabled {
            skipped_disabled += 1;
            continue;
        }
        if parser::parse(&rule.predicate).is_err() {
            skipped_bad += 1;
            continue;
        }
        let dxf = style_xml(rule.style.as_override(), &mut notes);
        if rule.predicate.chars().chain(dxf.chars()).any(|c| {
            !matches!(c, '\t' | '\n' | '\r') && (c < ' ' || c == '\u{fffe}' || c == '\u{ffff}')
        }) {
            return Err(format!("Conditional formatting on '{}' contains text XML cannot represent; no file was written.", sheet.name));
        }
        let mut earlier = Vec::new();
        for range in &rule.ranges {
            let mut pieces = vec![*range];
            for cut in &earlier {
                if pieces.is_empty() {
                    break;
                }
                work = work.checked_sub(pieces.len()).ok_or_else(error)?;
                pieces = pieces.into_iter().flat_map(|r| subtract(r, *cut)).collect();
                if pieces.len() > MAX_RANGES {
                    return Err(error());
                }
            }
            for piece in pieces {
                let formula = parser::adjust_formula_refs(
                    &rule.predicate,
                    (piece.start_row - range.start_row) as i32,
                    (piece.start_col - range.start_col) as i32,
                );
                let formula = super::xlsx_tables::excel_formula(&formula);
                bytes = bytes
                    .saturating_add(formula.len())
                    .saturating_add(dxf.len());
                if bytes > MAX_TEXT || result.rules.len() == MAX_RANGES {
                    return Err(error());
                }
                result.rules.push(Rule {
                    range: piece,
                    formula,
                    dxf: dxf.clone(),
                });
            }
            earlier.push(*range);
        }
    }
    if skipped_disabled > 0 {
        notes.insert(format!("{skipped_disabled} disabled rule(s) are omitted; Excel has no enabled/disabled rule state"));
    }
    if skipped_bad > 0 {
        notes.insert(format!(
            "{skipped_bad} unparseable, inert rule(s) are omitted"
        ));
    }
    result.warnings = notes
        .into_iter()
        .map(|n| format!("Conditional formatting on '{}': {n}.", sheet.name))
        .collect();
    Ok(result)
}
pub(crate) fn warnings(wb: &Workbook) -> Result<Vec<String>, String> {
    wb.sheets().iter().try_fold(Vec::new(), |mut notes, sheet| {
        notes.extend(plan(sheet)?.warnings);
        Ok(notes)
    })
}
pub(crate) fn write(sheet: &Sheet, worksheet: &mut Worksheet) -> Result<(), String> {
    for rule in plan(sheet)?.rules {
        let r = rule.range;
        worksheet
            .add_conditional_format(
                r.start_row as u32,
                r.start_col as u16,
                r.end_row as u32,
                r.end_col as u16,
                &ConditionalFormatFormula::new().set_rule(rule.formula.as_str()),
            )
            .map_err(|e| format!("Could not export conditional formatting: {e}"))?;
    }
    Ok(())
}
pub(crate) fn finish(bytes: Vec<u8>, wb: &Workbook) -> Result<Vec<u8>, String> {
    if wb.sheets().iter().all(|s| s.cond_formats.is_empty()) {
        return Ok(bytes);
    }
    let plans: Vec<_> = wb.sheets().iter().map(plan).collect::<Result<_, _>>()?;
    if plans.iter().all(|p| p.rules.is_empty()) {
        return Ok(bytes);
    }
    let mut styles = Vec::new();
    let mut ids = BTreeMap::new();
    let mut sheet_ids = Vec::new();
    for p in plans {
        let mut refs = Vec::new();
        for rule in p.rules {
            let id = *ids.entry(rule.dxf.clone()).or_insert_with(|| {
                let id = styles.len();
                styles.push(rule.dxf);
                id
            });
            refs.push(id);
        }
        sheet_ids.push(refs);
    }
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| e.to_string())?;
    let mut xml = String::new();
    input
        .by_name("xl/styles.xml")
        .map_err(|e| e.to_string())?
        .read_to_string(&mut xml)
        .map_err(|e| e.to_string())?;
    let (style_xml, base) = append_styles(&xml, &styles)?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..input.len() {
        let mut entry = input.by_index(i).map_err(|e| e.to_string())?;
        let name = entry.name().to_string();
        let part = if name == "xl/styles.xml" {
            Some(style_xml.clone())
        } else if let Some(sheet) = name
            .strip_prefix("xl/worksheets/sheet")
            .and_then(|n| n.strip_suffix(".xml"))
            .and_then(|n| n.parse::<usize>().ok())
            .and_then(|n| n.checked_sub(1))
            .filter(|&n| n < sheet_ids.len() && !sheet_ids[n].is_empty())
        {
            let mut xml = String::new();
            entry.read_to_string(&mut xml).map_err(|e| e.to_string())?;
            Some(patch_rules(&xml, &sheet_ids[sheet], base)?)
        } else {
            None
        };
        if let Some(part) = part {
            output
                .start_file(
                    name,
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(|e| e.to_string())?;
            output.write_all(&part).map_err(|e| e.to_string())?;
        } else {
            output.raw_copy_file(entry).map_err(|e| e.to_string())?;
        }
    }
    Ok(output.finish().map_err(|e| e.to_string())?.into_inner())
}
fn append_styles(xml: &str, styles: &[String]) -> Result<(Vec<u8>, usize), String> {
    let mut reader = Reader::from_str(xml);
    let mut base = 0;
    let mut next_number = 164u32;
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) | Event::Empty(e) => {
                if e.local_name().as_ref() == b"dxf" {
                    base += 1;
                }
                if e.local_name().as_ref() == b"numFmt" {
                    if let Some(id) = super::xlsx_tables::attr(&e, b"numFmtId")?
                        .and_then(|n| n.parse::<u32>().ok())
                    {
                        next_number = next_number.max(id.checked_add(1).ok_or_else(error)?);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let mut extra = String::new();
    for style in styles {
        if style.contains("numFmtId=\"0\"") {
            extra
                .push_str(&style.replace("numFmtId=\"0\"", &format!("numFmtId=\"{next_number}\"")));
            next_number = next_number.checked_add(1).ok_or_else(error)?;
        } else {
            extra.push_str(style);
        }
    }
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::new());
    let mut found = false;
    loop {
        let event = reader.read_event().map_err(|e| e.to_string())?;
        match event {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"dxfs" => {
                found = true;
                let mut start = BytesStart::new("dxfs");
                start.push_attribute(("count", (base + styles.len()).to_string().as_str()));
                writer
                    .write_event(Event::Start(start))
                    .map_err(|e| e.to_string())?;
                // Generated styles use an empty dxfs element when base is zero.
                if base == 0 {
                    writer.get_mut().extend_from_slice(extra.as_bytes());
                    writer
                        .write_event(Event::End(BytesEnd::new("dxfs")))
                        .map_err(|e| e.to_string())?;
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"dxfs" => {
                if base > 0 {
                    writer.get_mut().extend_from_slice(extra.as_bytes());
                    writer
                        .write_event(Event::End(e))
                        .map_err(|e| e.to_string())?;
                }
            }
            Event::Eof => break,
            other => writer.write_event(other).map_err(|e| e.to_string())?,
        }
    }
    if !found {
        return Err("Generated styles have no differential-format collection.".into());
    }
    Ok((writer.into_inner(), base))
}
fn patch_rules(xml: &str, ids: &[usize], base: usize) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::new());
    let mut seen = BTreeSet::new();
    loop {
        let mut event = reader.read_event().map_err(|e| e.to_string())?;
        if let Event::Start(e) | Event::Empty(e) = &mut event {
            if e.local_name().as_ref() == b"cfRule" {
                let index = super::xlsx_tables::attr(e, b"priority")?
                    .and_then(|n| n.parse::<usize>().ok())
                    .and_then(|n| n.checked_sub(1))
                    .ok_or("Generated conditional-format priority is invalid.")?;
                let id = ids
                    .get(index)
                    .ok_or("Generated conditional-format priority is outside the plan.")?;
                if !seen.insert(index) {
                    return Err("Generated conditional-format priority is duplicated.".into());
                }
                e.push_attribute(("dxfId", (base + id).to_string().as_str()));
            }
        }
        if matches!(event, Event::Eof) {
            break;
        }
        writer.write_event(event).map_err(|e| e.to_string())?;
    }
    if seen.len() != ids.len() {
        return Err("Generated conditional formatting does not match the export plan.".into());
    }
    Ok(writer.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn no_rules_leave_the_generated_package_byte_identical() {
        let bytes = vec![1, 2, 3, 4];
        assert_eq!(finish(bytes.clone(), &Workbook::new()).unwrap(), bytes);
    }
    #[test]
    fn appended_styles_keep_existing_dxf_ids_and_allocate_number_ids_above_existing() {
        let xml = r#"<styleSheet><numFmts count="1"><numFmt numFmtId="208" formatCode="0.00"/></numFmts><dxfs count="1"><dxf><font><b/></font></dxf></dxfs><tableStyles/></styleSheet>"#;
        let extra = vec![r#"<dxf><numFmt numFmtId="0" formatCode="0%"/></dxf>"#.into()];
        let (out, base) = append_styles(xml, &extra).unwrap();
        let text = String::from_utf8(out).unwrap();
        assert_eq!(base, 1);
        assert!(text.contains(r#"<dxfs count="2"><dxf><font><b/></font></dxf>"#));
        assert!(text.contains(r#"numFmtId="209" formatCode="0%""#));
        let patched = patch_rules(r#"<worksheet><conditionalFormatting sqref="A1"><cfRule type="expression" priority="2"/><cfRule type="expression" priority="1"/></conditionalFormatting></worksheet>"#, &[0, 1], base).unwrap();
        let patched = String::from_utf8(patched).unwrap();
        assert!(patched.contains(r#"priority="2" dxfId="2""#));
        assert!(patched.contains(r#"priority="1" dxfId="1""#));
    }
    #[test]
    fn duplicate_ranges_keep_one_anchor_without_quadratic_empty_intersections() {
        let mut wb = Workbook::new();
        wb.active_sheet_mut().cond_formats.add(
            vec![CellRange::single(0, 0); 50_000],
            "=TRUE",
            visigrid_engine::cond_format::CondStyle::Inline(Default::default()),
        );
        assert_eq!(plan(wb.active_sheet()).unwrap().rules.len(), 1);
        wb.active_sheet_mut().cond_formats.add(
            vec![CellRange::single(1, 0); 50_001],
            "=TRUE",
            visigrid_engine::cond_format::CondStyle::Inline(Default::default()),
        );
        assert!(warnings(&wb).unwrap_err().contains("complexity"));
    }

    #[test]
    fn excessive_metadata_refuses_before_building_a_package() {
        let mut wb = Workbook::new();
        wb.active_sheet_mut().cond_formats.add(
            vec![CellRange::single(0, 0); MAX_RANGES + 1],
            "=TRUE",
            visigrid_engine::cond_format::CondStyle::Inline(Default::default()),
        );
        assert!(warnings(&wb).unwrap_err().contains("complexity"));
    }
}
