//! Traditional XLSX comments (called Notes in current Excel).
//! Follow package relationships; part numbers need not match sheet numbers.
use std::{
    collections::{HashMap, HashSet},
    fs::File,
    io::{Cursor, Read, Write},
    path::Path,
};

use quick_xml::{
    events::{BytesStart, Event},
    Reader,
};
use visigrid_engine::{cell::CellComment, sheet::Sheet, workbook::Workbook};

type Comments = Vec<(String, Vec<((usize, usize), CellComment)>)>;
const MAX_PART_BYTES: u64 = 32 * 1024 * 1024;
const THREADED_ERROR: &str = "This workbook contains threaded Excel comments. VisiGrid currently supports Excel Notes only. Import was stopped to avoid losing discussions. Open the original in Excel; use a separate copy with Notes to work in VisiGrid.";

fn attr(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<String>, String> {
    for a in e.attributes() {
        let a = a.map_err(|e| e.to_string())?;
        if a.key.local_name().as_ref() == key {
            let raw = std::str::from_utf8(a.value.as_ref()).map_err(|e| e.to_string())?;
            return quick_xml::escape::unescape(raw)
                .map(|s| Some(s.into_owned()))
                .map_err(|e| e.to_string());
        }
    }
    Ok(None)
}

fn part(zip: &mut zip::ZipArchive<File>, name: &str) -> Result<String, String> {
    let file = zip
        .by_name(name)
        .map_err(|e| format!("Missing comment package part {name}: {e}"))?;
    if file.size() > MAX_PART_BYTES {
        return Err(format!("Comment package part {name} is too large"));
    }
    let mut xml = String::new();
    file.take(MAX_PART_BYTES + 1)
        .read_to_string(&mut xml)
        .map_err(|e| e.to_string())?;
    if xml.len() as u64 > MAX_PART_BYTES {
        return Err(format!("Comment package part {name} is too large"));
    }
    Ok(xml)
}

fn target_path(source: &str, target: &str) -> Result<String, String> {
    let base = source.rsplit_once('/').map_or("", |(base, _)| base);
    let joined = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        format!("{base}/{target}")
    };
    let mut parts = Vec::new();
    for item in joined.split('/') {
        match item {
            "" | "." => {}
            ".." => {
                if parts.pop().is_none() {
                    return Err("Invalid comment relationship path".into());
                }
            }
            _ => parts.push(item),
        }
    }
    Ok(parts.join("/"))
}

struct Relationship {
    id: String,
    target: String,
    kind: String,
}
fn relationships(xml: &str, source: &str) -> Result<Vec<Relationship>, String> {
    let mut reader = Reader::from_str(xml);
    let mut result = Vec::new();
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                let kind = attr(&e, b"Type")?.unwrap_or_default();
                if kind.to_ascii_lowercase().contains("threadedcomment") {
                    return Err(THREADED_ERROR.into());
                }
                if attr(&e, b"TargetMode")?.as_deref() == Some("External") {
                    if kind.ends_with("/comments") {
                        return Err("External comment relationships are unsupported".into());
                    }
                    continue;
                }
                result.push(Relationship {
                    id: attr(&e, b"Id")?.unwrap_or_default(),
                    target: target_path(source, &attr(&e, b"Target")?.unwrap_or_default())?,
                    kind,
                });
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}

pub(crate) fn read(path: &Path) -> Result<Comments, String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let Ok(mut zip) = zip::ZipArchive::new(file) else {
        return Ok(Vec::new());
    }; // Legacy XLS.
    if !zip.file_names().any(|n| n == "xl/workbook.xml") {
        return Ok(Vec::new());
    } // ODS/XLSB.
      // Includes nonstandard part names via content types, not just Excel's usual directory.
    let types = part(&mut zip, "[Content_Types].xml")?;
    if types.to_ascii_lowercase().contains("threadedcomment") {
        return Err(THREADED_ERROR.into());
    }
    let workbook_xml = part(&mut zip, "xl/workbook.xml")?;
    let rel_xml = part(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let rels: HashMap<_, _> = relationships(&rel_xml, "xl/workbook.xml")?
        .into_iter()
        .map(|r| (r.id.clone(), r))
        .collect();
    let mut result = Vec::new();
    let mut reader = Reader::from_str(&workbook_xml);
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheet" => {
                let name = attr(&e, b"name")?.ok_or("Missing sheet name")?;
                let id = attr(&e, b"id")?.ok_or("Missing sheet relationship")?;
                let rel = rels.get(&id).ok_or("Missing sheet relationship target")?;
                if !rel.kind.ends_with("/worksheet") {
                    continue;
                }
                let (base, file) = rel
                    .target
                    .rsplit_once('/')
                    .ok_or("Invalid worksheet path")?;
                let rel_path = format!("{base}/_rels/{file}.rels");
                if !zip.file_names().any(|n| n == rel_path) {
                    continue;
                }
                let xml = part(&mut zip, &rel_path)?;
                for rel in relationships(&xml, &rel.target)? {
                    if rel.kind.ends_with("/comments") {
                        let xml = part(&mut zip, &rel.target)?;
                        result.push((
                            name.clone(),
                            parse(&xml)
                                .map_err(|e| format!("Cannot import Notes on '{name}': {e}"))?,
                        ));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}

// SpreadsheetML's escaped UTF-16 code units. Decode once so _x005F_x0041_
// stays the literal string _x0041_, rather than becoming A.
fn decode_excel(text: &str) -> String {
    let mut units = Vec::new();
    let mut rest = text;
    while !rest.is_empty() {
        let b = rest.as_bytes();
        if b.len() >= 7
            && b[0] == b'_'
            && b[1] == b'x'
            && b[6] == b'_'
            && b[2..6].iter().all(u8::is_ascii_hexdigit)
        {
            units.push(u16::from_str_radix(&rest[2..6], 16).unwrap());
            rest = &rest[7..];
        } else {
            let c = rest.chars().next().unwrap();
            units.extend_from_slice(c.encode_utf16(&mut [0; 2]));
            rest = &rest[c.len_utf8()..];
        }
    }
    String::from_utf16_lossy(&units)
}

// Unlike shared strings, this writer version's note XML does not protect
// literal SpreadsheetML escape sequences. Protect them before serialization.
fn encode_literals(text: &str) -> String {
    let mut result = String::new();
    for (i, c) in text.char_indices() {
        let b = &text.as_bytes()[i..];
        if b.len() >= 7
            && b[0] == b'_'
            && b[1] == b'x'
            && b[6] == b'_'
            && b[2..6].iter().all(u8::is_ascii_hexdigit)
        {
            result.push_str("_x005F");
        }
        result.push(c);
    }
    result
}

fn address(reference: &str) -> Result<(usize, usize), String> {
    let split = reference
        .find(|c: char| c.is_ascii_digit())
        .ok_or("Invalid comment cell reference")?;
    let (letters, digits) = reference.split_at(split);
    if letters.is_empty()
        || !letters.bytes().all(|b| b.is_ascii_uppercase())
        || !digits.bytes().all(|b| b.is_ascii_digit())
    {
        return Err("Invalid comment cell reference".into());
    }
    let col = letters
        .bytes()
        .try_fold(0usize, |n, b| {
            n.checked_mul(26)?.checked_add((b - b'A' + 1) as usize)
        })
        .ok_or("Invalid comment column")?;
    let row = digits.parse::<usize>().map_err(|e| e.to_string())?;
    if row == 0 || row > visigrid_engine::sheet::NUM_ROWS || col > visigrid_engine::sheet::NUM_COLS
    {
        return Err(format!("Comment at {reference} is outside VisiGrid's grid"));
    }
    Ok((row - 1, col - 1))
}

fn element_text(reader: &mut Reader<&[u8]>) -> Result<String, String> {
    let mut result = String::new();
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Text(t) => result.push_str(&t.xml_content().map_err(|e| e.to_string())?),
            Event::CData(t) => result.push_str(&t.xml_content().map_err(|e| e.to_string())?),
            Event::GeneralRef(r) => {
                let entity = format!(
                    "&{};",
                    std::str::from_utf8(r.as_ref()).map_err(|e| e.to_string())?
                );
                result.push_str(&quick_xml::escape::unescape(&entity).map_err(|e| e.to_string())?);
            }
            Event::End(_) => return Ok(decode_excel(&result)),
            _ => return Err("Invalid note text element".into()),
        }
    }
}

fn parse(xml: &str) -> Result<Vec<((usize, usize), CellComment)>, String> {
    let mut reader = Reader::from_str(xml);
    let mut authors = Vec::new();
    let mut comments = Vec::new();
    let mut positions = HashSet::new();
    let mut current = None;
    let mut text = String::new();
    let mut in_text = false;
    let mut found_comments = false;
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"comments" => {
                found_comments = true;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"comment" => {
                return Err("Comment is missing its text element".into());
            }
            Event::Start(e) if e.local_name().as_ref() == b"author" => {
                authors.push(element_text(&mut reader)?);
            }
            Event::Empty(e) if e.local_name().as_ref() == b"author" => authors.push(String::new()),
            Event::Start(e) if e.local_name().as_ref() == b"comment" => {
                if current.is_some() {
                    return Err("Nested comment".into());
                }
                let pos = address(&attr(&e, b"ref")?.ok_or("Missing comment reference")?)?;
                let author = attr(&e, b"authorId")?
                    .ok_or("Missing comment author")?
                    .parse::<usize>()
                    .map_err(|e| e.to_string())?;
                current = Some((pos, author));
                text.clear();
            }
            Event::Start(e) if e.local_name().as_ref() == b"text" => in_text = true,
            Event::Start(e) if e.local_name().as_ref() == b"rPh" => {
                reader.read_to_end(e.name()).map_err(|e| e.to_string())?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"t" && in_text && current.is_some() => {
                text.push_str(&element_text(&mut reader)?);
            }
            Event::End(e) if e.local_name().as_ref() == b"text" => in_text = false,
            Event::End(e) if e.local_name().as_ref() == b"comment" => {
                let (pos, author) = current.take().ok_or("Invalid comment end")?;
                let author = authors
                    .get(author)
                    .ok_or("Invalid comment author index")?
                    .clone();
                if !positions.insert(pos) {
                    return Err("Duplicate comment cell".into());
                }
                comments.push((
                    pos,
                    CellComment {
                        text: std::mem::take(&mut text),
                        author,
                    },
                ));
            }
            Event::Eof => {
                if current.is_some() {
                    return Err("Unterminated comment".into());
                }
                break;
            }
            _ => {}
        }
    }
    if !found_comments {
        return Err("Missing comments document root".into());
    }
    Ok(comments)
}

pub(crate) fn apply(comments: Comments, workbook: &mut Workbook) -> Result<usize, String> {
    let mut count = 0;
    for (name, comments) in comments {
        let idx = workbook
            .sheets()
            .iter()
            .position(|s| s.name == name)
            .ok_or_else(|| format!("Notes belong to missing sheet '{name}'"))?;
        let sheet = workbook.sheet_mut(idx).unwrap();
        for ((row, col), comment) in comments {
            if sheet.comment(row, col).is_some() {
                return Err(format!("Duplicate Notes on '{name}'"));
            }
            sheet.set_comment(row, col, Some(comment));
            count += 1;
        }
    }
    Ok(count)
}

pub(crate) fn write(
    sheet: &Sheet,
    worksheet: &mut rust_xlsxwriter::Worksheet,
) -> Result<usize, String> {
    let mut count = 0;
    for ((row, col), comment) in sheet.comments() {
        if row >= 1_048_576 || col >= 16_384 {
            return Err(format!(
                "A comment on '{}' is outside Excel's grid",
                sheet.name
            ));
        }
        if comment.text.encode_utf16().count() > 32767 {
            return Err(format!("A comment on '{}' exceeds Excel's 32,767-character limit. Shorten it before exporting.", sheet.name));
        }
        // Use the library for VML geometry and package relationships. Its 0.79
        // serializer sorts authors but assigns IDs in cell order, mismatching
        // authors on multi-author sheets. finish() writes the authoritative XML.
        let note = rust_xlsxwriter::Note::new("VisiGrid note").add_author_prefix(false);
        worksheet
            .insert_note(row as u32, col as u16, &note)
            .map_err(|e| {
                format!(
                    "Cannot export comment on '{}' at row {}, column {}: {e}",
                    sheet.name,
                    row + 1,
                    col + 1
                )
            })?;
        count += 1;
    }
    Ok(count)
}

fn xml_text(text: &str) -> String {
    let literal = encode_literals(text);
    let mut result = String::new();
    for c in literal.chars() {
        match c {
            '&' => result.push_str("&amp;"),
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            '\r'
            | '\u{0}'..='\u{8}'
            | '\u{b}'..='\u{c}'
            | '\u{e}'..='\u{1f}'
            | '\u{fffe}'
            | '\u{ffff}' => result.push_str(&format!("_x{:04X}_", c as u32)),
            _ => result.push(c),
        }
    }
    result
}

fn comments_xml(sheet: &Sheet) -> String {
    let mut comments: Vec<_> = sheet.comments().collect();
    comments.sort_by_key(|(pos, _)| *pos);
    let mut authors = Vec::new();
    let mut author_ids = HashMap::new();
    for (_, comment) in &comments {
        if !author_ids.contains_key(&comment.author) {
            author_ids.insert(comment.author.clone(), authors.len());
            authors.push(&comment.author);
        }
    }
    let mut xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors>");
    for author in authors {
        xml.push_str(&format!("<author>{}</author>", xml_text(author)));
    }
    xml.push_str("</authors><commentList>");
    for ((row, col), comment) in comments {
        let mut n = col + 1;
        let mut letters = Vec::new();
        while n > 0 {
            letters.push((b'A' + ((n - 1) % 26) as u8) as char);
            n = (n - 1) / 26;
        }
        let reference: String = letters.into_iter().rev().collect();
        xml.push_str(&format!("<comment ref=\"{reference}{}\" authorId=\"{}\"><text><t xml:space=\"preserve\">{}</t></text></comment>", row + 1, author_ids[&comment.author], xml_text(&comment.text)));
    }
    xml.push_str("</commentList></comments>");
    xml
}

/// Replace only the generated note XML; copy all other ZIP entries compressed.
/// The writer emits commentsN.xml in workbook order, omitting uncommented sheets.
pub(crate) fn finish(bytes: Vec<u8>, workbook: &Workbook) -> Result<Vec<u8>, String> {
    let parts: HashMap<_, _> = workbook
        .sheets()
        .iter()
        .filter(|s| s.comments().next().is_some())
        .enumerate()
        .map(|(i, s)| (format!("xl/comments{}.xml", i + 1), comments_xml(s)))
        .collect();
    if parts.is_empty() {
        return Ok(bytes);
    }
    let mut input = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| e.to_string())?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let mut replaced = 0;
    for i in 0..input.len() {
        let entry = input.by_index(i).map_err(|e| e.to_string())?;
        if let Some(xml) = parts.get(entry.name()) {
            output
                .start_file(
                    entry.name(),
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .map_err(|e| e.to_string())?;
            output
                .write_all(xml.as_bytes())
                .map_err(|e| e.to_string())?;
            replaced += 1;
        } else {
            output.raw_copy_file(entry).map_err(|e| e.to_string())?;
        }
    }
    if replaced != parts.len() {
        return Err("XLSX writer did not produce all comment parts".into());
    }
    output
        .finish()
        .map(|c| c.into_inner())
        .map_err(|e| e.to_string())
}
