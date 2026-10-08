//! Read modern Excel errors from rich values plus a conventional #VALUE!
//! fallback. MS-XLSX 2.3.6.1.3, 2.2.4.4 and 2.6.175. Export does not write
//! these parts: a guessed rich value makes Excel repair dynamic arrays.
//! Never emit #SPILL! as a raw worksheet t="e" value.
use quick_xml::{
    events::{BytesEnd, BytesStart, BytesText, Event},
    Reader, Writer,
};
use std::{
    collections::{BTreeMap, HashMap},
    fs::File,
    io::{BufReader, Cursor, Read, Write},
    path::Path,
};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RICH: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
const REL: &str = "http://schemas.microsoft.com/office/2017/06/relationships/";
const META_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
const DATA: &str = "xl/richData/rdrichvalue.xml";
const STRUCT: &str = "xl/richData/rdrichvaluestructure.xml";
const MAX_BYTES: u64 = 16 * 1024 * 1024;
const MAX_NODES: usize = 200_000;

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) struct ModernError {
    code: u32,
    rows: usize,
    cols: usize,
}
impl ModernError {
    pub(crate) fn from_error(
        error: &str,
        dimensions: Option<&visigrid_engine::cell::SpillInfo>,
    ) -> Option<Self> {
        let (code, _) = [
            (8, "#SPILL!"),
            (9, "#CONNECT!"),
            (10, "#BLOCKED!"),
            (11, "#UNKNOWN!"),
            (13, "#CALC!"),
            (14, "#BUSY!"),
            (19, "#TIMEOUT!"),
        ]
        .into_iter()
        .find(|(_, label)| {
            error.strip_prefix(label).is_some_and(|rest| {
                rest.is_empty() || rest.starts_with(char::is_whitespace) || rest.starts_with(':')
            })
        })?;
        let (rows, cols) = if code == 8 {
            dimensions.map_or((0, 0), |d| {
                (d.rows.saturating_sub(1), d.cols.saturating_sub(1))
            })
        } else {
            (0, 0)
        };
        Some(Self { code, rows, cols })
    }
}

// A bounded tree is used only for small metadata parts, never worksheet XML.
#[derive(Default)]
struct Node {
    name: String,
    attrs: BTreeMap<String, String>,
    children: Vec<Node>,
    text: String,
}
impl Node {
    fn local(&self) -> &str {
        self.name.rsplit(':').next().unwrap_or(&self.name)
    }
    fn attr(&self, key: &str) -> Option<&str> {
        self.attrs.get(key).map(String::as_str)
    }
    fn index(&self, key: &str) -> Result<usize, String> {
        self.attr(key)
            .ok_or_else(|| format!("Missing {key} in {}", self.name))?
            .parse()
            .map_err(|_| format!("Invalid {key} in {}", self.name))
    }
    fn child(&self, name: &str) -> Option<&Self> {
        self.children.iter().find(|n| n.local() == name)
    }
    fn all<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a Self> {
        self.children.iter().filter(move |n| n.local() == name)
    }
    fn xml(&self) -> Result<String, String> {
        fn write(node: &Node, out: &mut Writer<Vec<u8>>) -> Result<(), String> {
            let mut start = BytesStart::new(&node.name);
            for (key, value) in &node.attrs {
                start.push_attribute((key.as_str(), value.as_str()));
            }
            out.write_event(Event::Start(start))
                .map_err(|e| e.to_string())?;
            if !node.text.is_empty() {
                out.write_event(Event::Text(BytesText::new(&node.text)))
                    .map_err(|e| e.to_string())?;
            }
            for child in &node.children {
                write(child, out)?;
            }
            out.write_event(Event::End(BytesEnd::new(&node.name)))
                .map_err(|e| e.to_string())?;
            Ok(())
        }
        let mut out = Writer::new(Vec::new());
        write(self, &mut out)?;
        String::from_utf8(out.into_inner()).map_err(|e| e.to_string())
    }
}
fn parse(xml: &str) -> Result<Node, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().expand_empty_elements = true;
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut count = 0;
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) => {
                count += 1;
                if count > MAX_NODES || stack.len() >= 32 {
                    return Err("Rich metadata exceeds its node/depth limit".into());
                }
                let mut node = Node {
                    name: String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                    ..Default::default()
                };
                for attr in e.attributes() {
                    let attr = attr.map_err(|e| e.to_string())?;
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    let value = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(|e| e.to_string())?
                        .into_owned();
                    if node.attrs.insert(key, value).is_some() {
                        return Err("Duplicate rich metadata attribute".into());
                    }
                }
                stack.push(node);
            }
            Event::Text(e) => {
                if let Some(n) = stack.last_mut() {
                    n.text.push_str(
                        &quick_xml::escape::unescape(&e.decode().map_err(|e| e.to_string())?)
                            .map_err(|e| e.to_string())?,
                    );
                }
            }
            Event::CData(e) => {
                if let Some(n) = stack.last_mut() {
                    n.text.push_str(&e.decode().map_err(|e| e.to_string())?);
                }
            }
            Event::GeneralRef(e) => {
                let reference = e.decode().map_err(|e| e.to_string())?;
                let encoded = format!("&{reference};");
                let decoded = quick_xml::escape::unescape(&encoded).map_err(|e| e.to_string())?;
                if let Some(n) = stack.last_mut() {
                    n.text.push_str(&decoded);
                }
            }
            Event::End(_) => {
                let node = stack.pop().ok_or("Unbalanced rich metadata")?;
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else if root.replace(node).is_some() {
                    return Err("Multiple rich metadata roots".into());
                }
            }
            Event::DocType(_) => return Err("Rich metadata cannot contain a DTD".into()),
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() {
        return Err("Unclosed rich metadata".into());
    }
    root.ok_or_else(|| "Empty rich metadata".into())
}
fn part<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    name: &str,
) -> Result<String, String> {
    let entry = zip.by_name(name).map_err(|e| e.to_string())?;
    if entry.size() > MAX_BYTES {
        return Err("Rich metadata exceeds 16 MiB per part".into());
    }
    let mut text = String::new();
    entry
        .take(MAX_BYTES + 1)
        .read_to_string(&mut text)
        .map_err(|e| e.to_string())?;
    if text.len() as u64 > MAX_BYTES {
        return Err("Rich metadata exceeds 16 MiB per part".into());
    }
    Ok(text)
}
fn relation(root: &Node, kind: &str) -> Result<Option<String>, String> {
    let mut found = None;
    for n in root
        .all("Relationship")
        .filter(|n| n.attr("Type") == Some(kind))
    {
        if n.attr("TargetMode") == Some("External") {
            return Err("External rich metadata relationship".into());
        }
        let target = n.attr("Target").ok_or("Rich relationship has no target")?;
        let path = super::xlsx_tables::target("xl/workbook.xml", target)?;
        if found.replace(path).is_some() {
            return Err("Duplicate rich metadata relationship".into());
        }
    }
    Ok(found)
}
fn add_relation(root: &mut Node, kind: &str, target: &str) {
    let mut id = root.children.len() + 1;
    while root
        .children
        .iter()
        .any(|n| n.attr("Id") == Some(format!("rId{id}").as_str()))
    {
        id += 1;
    }
    let mut node = Node {
        name: "Relationship".into(),
        ..Default::default()
    };
    for (k, v) in [
        ("Id", format!("rId{id}")),
        ("Type", kind.into()),
        ("Target", target.into()),
    ] {
        node.attrs.insert(k.into(), v);
    }
    root.children.push(node);
}

pub(crate) fn finish(
    bytes: Vec<u8>,
    errors: &BTreeMap<ModernError, usize>,
) -> Result<Vec<u8>, String> {
    if errors.is_empty() {
        return Ok(bytes);
    }
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| e.to_string())?;
    let mut rels = parse(&part(&mut zip, "xl/_rels/workbook.xml.rels")?)?;
    let mut types = parse(&part(&mut zip, "[Content_Types].xml")?)?;
    if relation(&rels, &format!("{REL}rdRichValue"))?.is_some()
        || zip.file_names().any(|n| n == DATA || n == STRUCT)
    {
        return Err(
            "XLSX writer produced rich values that cannot be combined with error caches yet".into(),
        );
    }
    let has_metadata = zip.file_names().any(|n| n == "xl/metadata.xml");
    let mut metadata = if has_metadata {
        parse(&part(&mut zip, "xl/metadata.xml")?)?
    } else {
        parse(&format!(
            "<metadata xmlns=\"{MAIN}\"><metadataTypes count=\"0\"/></metadata>"
        ))?
    };
    if metadata.child("valueMetadata").is_some() {
        return Err("Unexpected value metadata in XLSX writer output".into());
    }
    metadata.attrs.insert("xmlns:xlrd".into(), RICH.into());
    let mt = metadata
        .children
        .iter_mut()
        .find(|n| n.local() == "metadataTypes")
        .ok_or("Missing metadata types")?;
    let type_id = mt.children.len() + 1;
    mt.children.push(parse("<metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\"/>")?);
    mt.attrs.insert("count".into(), type_id.to_string());
    let count = errors.len();
    let mut future = format!("<futureMetadata name=\"XLRICHVALUE\" count=\"{count}\">");
    let mut values = format!("<valueMetadata count=\"{count}\">");
    let mut data = format!("<rvData xmlns=\"{RICH}\" count=\"{count}\">");
    for (error, index) in errors {
        future.push_str(&format!("<bk><extLst><ext uri=\"{{3E2802C4-A4D2-4D8B-9148-E3BE6C30E623}}\"><xlrd:rvb i=\"{index}\"/></ext></extLst></bk>"));
        values.push_str(&format!("<bk><rc t=\"{type_id}\" v=\"{index}\"/></bk>"));
        data.push_str(&format!(
            "<rv s=\"{}\"><fb t=\"e\">#VALUE!</fb><v>{}</v>",
            if error.code == 8 { 1 } else { 0 },
            error.code
        ));
        if error.code == 8 {
            data.push_str(&format!("<v>{}</v><v>{}</v>", error.rows, error.cols));
        }
        data.push_str("</rv>");
    }
    future.push_str("</futureMetadata>");
    values.push_str("</valueMetadata>");
    data.push_str("</rvData>");
    let at = metadata
        .children
        .iter()
        .position(|n| n.local() == "cellMetadata")
        .unwrap_or(metadata.children.len());
    metadata.children.insert(at, parse(&future)?);
    metadata.children.push(parse(&values)?);
    let structure = format!("<rvStructures xmlns=\"{RICH}\" count=\"2\"><s t=\"_error\"><k n=\"errorType\" t=\"i\"/></s><s t=\"_error\"><k n=\"errorType\" t=\"i\"/><k n=\"rwOffset\" t=\"i\"/><k n=\"colOffset\" t=\"i\"/></s></rvStructures>");
    for (path, kind, mime) in [
        (
            "xl/metadata.xml",
            META_REL.to_owned(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
        ),
        (
            DATA,
            format!("{REL}rdRichValue"),
            "application/vnd.ms-excel.rdRichValue+xml",
        ),
        (
            STRUCT,
            format!("{REL}rdRichValueStructure"),
            "application/vnd.ms-excel.rdRichValueStructure+xml",
        ),
    ] {
        if relation(&rels, &kind)?.is_none() {
            add_relation(&mut rels, &kind, path.strip_prefix("xl/").unwrap());
        }
        let name = format!("/{path}");
        if !types
            .children
            .iter()
            .any(|n| n.attr("PartName") == Some(name.as_str()))
        {
            let mut node = Node {
                name: "Override".into(),
                ..Default::default()
            };
            node.attrs.insert("PartName".into(), name);
            node.attrs.insert("ContentType".into(), mime.into());
            types.children.push(node);
        }
    }
    let mut parts: BTreeMap<String, String> = [
        ("xl/metadata.xml".into(), metadata.xml()?),
        ("xl/_rels/workbook.xml.rels".into(), rels.xml()?),
        ("[Content_Types].xml".into(), types.xml()?),
        (DATA.into(), data),
        (STRUCT.into(), structure),
    ]
    .into();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..zip.len() {
        let entry = zip.by_index(i).map_err(|e| e.to_string())?;
        if parts.contains_key(entry.name()) {
            continue;
        }
        out.raw_copy_file(entry).map_err(|e| e.to_string())?;
    }
    for (name, xml) in &mut parts {
        out.start_file(
            name,
            zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated),
        )
        .map_err(|e| e.to_string())?;
        out.write_all(xml.as_bytes()).map_err(|e| e.to_string())?;
    }
    Ok(out.finish().map_err(|e| e.to_string())?.into_inner())
}

pub(crate) type Imported = HashMap<(usize, usize, usize), &'static str>;

fn error_label(code: usize) -> Option<&'static str> {
    Some(match code {
        4 => "#NAME?",
        8 => "#SPILL!",
        9 => "#CONNECT!",
        10 => "#BLOCKED!",
        11 => "#UNKNOWN!",
        12 => "#FIELD!",
        13 => "#CALC!",
        14 | 17 => "#BUSY!",
        19 => "#TIMEOUT!",
        _ => return None,
    })
}

pub(crate) fn read(path: &Path, warnings: &mut Vec<String>) -> Imported {
    match read_inner(path) {
        Ok(cells) => cells,
        Err(error) => {
            warnings.push(format!("Modern Excel error metadata was not restored: {error}. Conventional cached errors were kept."));
            HashMap::new()
        }
    }
}

fn read_inner(path: &Path) -> Result<Imported, String> {
    let Ok(mut zip) = zip::ZipArchive::new(File::open(path).map_err(|e| e.to_string())?) else {
        return Ok(HashMap::new());
    };
    if !zip.file_names().any(|n| n == "xl/_rels/workbook.xml.rels") {
        return Ok(HashMap::new());
    }
    let rels = parse(&part(&mut zip, "xl/_rels/workbook.xml.rels")?)?;
    let Some(data_path) = relation(&rels, &format!("{REL}rdRichValue"))? else {
        return Ok(HashMap::new());
    };
    let structure_path = relation(&rels, &format!("{REL}rdRichValueStructure"))?
        .ok_or("Missing rich value structure relationship")?;
    let metadata_path = relation(&rels, META_REL)?.ok_or("Missing metadata relationship")?;
    let structures = parse(&part(&mut zip, &structure_path)?)?;
    let data = parse(&part(&mut zip, &data_path)?)?;
    let metadata = parse(&part(&mut zip, &metadata_path)?)?;
    let schemas: Vec<_> = structures.all("s").collect();
    let mut codes = Vec::new();
    for value in data.all("rv") {
        let schema = schemas
            .get(value.index("s")?)
            .ok_or("Invalid rich structure index")?;
        let keys: Vec<_> = schema.all("k").collect();
        let vals: Vec<_> = value.all("v").collect();
        if keys.len() != vals.len() {
            return Err("Rich value key/value count mismatch".into());
        }
        let code = if schema.attr("t") == Some("_error") {
            let index = keys
                .iter()
                .position(|k| k.attr("n") == Some("errorType") && k.attr("t") == Some("i"))
                .ok_or("Missing rich error type")?;
            let number = vals[index]
                .text
                .trim()
                .parse::<usize>()
                .map_err(|_| "Invalid rich error number")?;
            if number == 8 {
                for key in ["rwOffset", "colOffset"] {
                    let at = keys
                        .iter()
                        .position(|k| k.attr("n") == Some(key) && k.attr("t") == Some("i"))
                        .ok_or("Missing spill error dimensions")?;
                    vals[at]
                        .text
                        .trim()
                        .parse::<usize>()
                        .map_err(|_| "Invalid spill error dimensions")?;
                }
            }
            Some(
                error_label(number)
                    .ok_or_else(|| format!("Unsupported modern error type {number}"))?,
            )
        } else {
            None
        };
        codes.push(code);
    }
    if codes.iter().all(Option::is_none) {
        return Ok(HashMap::new());
    }
    let types = metadata
        .child("metadataTypes")
        .ok_or("Missing metadata types")?;
    let type_id = types
        .all("metadataType")
        .position(|n| n.attr("name") == Some("XLRICHVALUE"))
        .map(|i| i + 1)
        .ok_or("Missing rich metadata type")?;
    let future = metadata
        .all("futureMetadata")
        .find(|n| n.attr("name") == Some("XLRICHVALUE"))
        .ok_or("Missing rich future metadata")?;
    let mut future_codes = Vec::new();
    for block in future.all("bk") {
        let rvb = block
            .child("extLst")
            .and_then(|e| {
                e.all("ext")
                    .filter(|e| {
                        e.attr("uri").is_some_and(|u| {
                            u.eq_ignore_ascii_case("{3E2802C4-A4D2-4D8B-9148-E3BE6C30E623}")
                        })
                    })
                    .find_map(|e| e.child("rvb"))
            })
            .ok_or("Missing rich value binding")?;
        future_codes.push(
            *codes
                .get(rvb.index("i")?)
                .ok_or("Invalid rich value index")?,
        );
    }
    let mut bindings = Vec::new();
    for block in metadata
        .child("valueMetadata")
        .ok_or("Missing value metadata")?
        .all("bk")
    {
        let mut code = None;
        let mut found = false;
        for rc in block.all("rc") {
            if rc.index("t")? == type_id {
                if found {
                    return Err("Duplicate rich value binding".into());
                }
                found = true;
                code = *future_codes
                    .get(rc.index("v")?)
                    .ok_or("Invalid future metadata index")?;
            }
        }
        bindings.push(code);
    }
    let workbook = part(&mut zip, "xl/workbook.xml")?;
    let rels_xml = part(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let mut cells = HashMap::new();
    for (sheet, path) in super::xlsx::resolve_worksheet_paths(&workbook, &rels_xml)
        .iter()
        .enumerate()
    {
        let mut reader = Reader::from_reader(BufReader::new(
            zip.by_name(path).map_err(|e| e.to_string())?,
        ));
        let mut buffer = Vec::new();
        loop {
            buffer.clear();
            match reader
                .read_event_into(&mut buffer)
                .map_err(|e| e.to_string())?
            {
                Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"c" => {
                    if super::xlsx_tables::attr(&e, b"t")?.as_deref() != Some("e") {
                        continue;
                    }
                    let Some(vm) = super::xlsx_tables::attr(&e, b"vm")? else {
                        continue;
                    };
                    let index = vm
                        .parse::<usize>()
                        .map_err(|_| "Invalid cell value metadata index")?;
                    // Zero is the OOXML sentinel for no value metadata.
                    let Some(index) = index.checked_sub(1) else {
                        continue;
                    };
                    let code = bindings
                        .get(index)
                        .ok_or("Invalid cell value metadata index")?;
                    if let Some(code) = code {
                        let address = super::xlsx_tables::attr(&e, b"r")?
                            .ok_or("Error cache has no cell address")?;
                        let r = super::xlsx_tables::range(&address)?;
                        if cells.len() >= 1_000_000 {
                            return Err("Too many modern error cells".into());
                        }
                        cells.insert((sheet, r.start_row, r.start_col), *code);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
    }
    Ok(cells)
}
