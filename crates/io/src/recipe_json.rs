//! JSON files as a recipe source: an array of records in a `.json` file
//! (at a dotted path, or found by itself), or one record per line in
//! `.jsonl` / `.ndjson`.
//!
//! Each record is one row. Nested objects become dotted columns
//! (`customer.address.city`), in the order they first appear; arrays inside
//! a record stay JSON text in their cell. A column holding only numbers is a
//! number column, except integers too large to hold exactly, which stay text
//! so IDs are never rounded; strings stay text even when they look like
//! numbers, since the file said so.

use serde::de::{Deserializer, MapAccess, Visitor};
use serde::{Deserialize, Serialize};

use super::{Frame, OutColumn, Snapshot, ValueKind};
use crate::csv_import::ColumnRule;

/// The largest integer a number column holds exactly (2^53).
const MAX_EXACT_INTEGER: u64 = 1 << 53;
/// Nesting deeper than this is kept as JSON text rather than more columns.
const MAX_FLATTEN_DEPTH: usize = 16;

/// A JSON value that keeps object keys in file order (serde_json's own
/// Value sorts them unless a crate-wide feature is on) and each number's
/// text exactly as written, so a 30-digit ID is never rounded.
#[derive(Debug, Clone, PartialEq)]
enum Value {
    Null,
    Bool(bool),
    /// The number as the file wrote it.
    Number(String),
    String(String),
    Array(Vec<Value>),
    Object(Vec<(String, Value)>),
}

impl Value {
    fn is_object(&self) -> bool {
        matches!(self, Value::Object(_))
    }

    fn get(&self, key: &str) -> Option<&Value> {
        match self {
            Value::Object(m) => m.iter().find(|(k, _)| k == key).map(|(_, v)| v),
            _ => None,
        }
    }

    /// Compact JSON, keys in file order.
    fn to_json(&self) -> String {
        match self {
            Value::Null => "null".into(),
            Value::Bool(b) => b.to_string(),
            Value::Number(n) => n.clone(),
            Value::String(s) => serde_json::to_string(s).unwrap_or_default(),
            Value::Array(items) => format!("[{}]", items.iter().map(Value::to_json).collect::<Vec<_>>().join(",")),
            Value::Object(m) => format!(
                "{{{}}}",
                m.iter().map(|(k, v)| format!("{}:{}", serde_json::to_string(k).unwrap_or_default(), v.to_json())).collect::<Vec<_>>().join(",")
            ),
        }
    }
}

/// Parse JSON text. Containers are read as raw pieces (serde_json keeps
/// their text), so object order and number text survive.
fn parse(text: &str) -> Result<Value, serde_json::Error> {
    use serde_json::value::RawValue;
    let raw: Box<RawValue> = serde_json::from_str(text)?;
    from_raw(raw.get(), 0)
}

fn from_raw(text: &str, depth: usize) -> Result<Value, serde_json::Error> {
    use serde_json::value::RawValue;
    let t = text.trim();
    // Past this depth containers stay text; flatten keeps them as JSON
    if depth > 64 {
        return Ok(Value::String(t.to_string()));
    }
    Ok(match t.as_bytes().first() {
        Some(b'{') => {
            let entries: OrderedMap = serde_json::from_str(t)?;
            let mut out: Vec<(String, Value)> = Vec::with_capacity(entries.0.len());
            for (k, v) in entries.0 {
                let v = from_raw(v.get(), depth + 1)?;
                // A repeated key: the last one wins, as in most readers
                match out.iter_mut().find(|(e, _)| *e == k) {
                    Some(e) => e.1 = v,
                    None => out.push((k, v)),
                }
            }
            Value::Object(out)
        }
        Some(b'[') => {
            let items: Vec<Box<RawValue>> = serde_json::from_str(t)?;
            Value::Array(items.iter().map(|i| from_raw(i.get(), depth + 1)).collect::<Result<_, _>>()?)
        }
        Some(b'"') => Value::String(serde_json::from_str(t)?),
        Some(b't') | Some(b'f') => Value::Bool(serde_json::from_str(t)?),
        Some(b'n') => Value::Null,
        _ => Value::Number(t.to_string()),
    })
}

/// An object's members in file order, values left as raw text.
struct OrderedMap(Vec<(String, Box<serde_json::value::RawValue>)>);

impl<'de> Deserialize<'de> for OrderedMap {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        struct V;
        impl<'de> Visitor<'de> for V {
            type Value = OrderedMap;
            fn expecting(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
                f.write_str("a JSON object")
            }
            fn visit_map<A: MapAccess<'de>>(self, mut map: A) -> Result<OrderedMap, A::Error> {
                let mut out = Vec::new();
                while let Some(entry) = map.next_entry()? {
                    out.push(entry);
                }
                Ok(OrderedMap(out))
            }
        }
        d.deserialize_map(V)
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct JsonSource {
    /// Relative paths resolve against the recipe file's folder.
    pub path: String,
    /// Where the records are in a `.json` file: a dotted path such as
    /// `data.items`. Empty: found by itself (the file itself if it is an
    /// array, else its largest array of objects). Not used for `.jsonl` /
    /// `.ndjson`.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub records: String,
    /// The column names when the recipe was saved.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
    /// With a pattern path: append every matching file.
    #[serde(default, skip_serializing_if = "super::is_false")]
    pub combine: bool,
}

impl JsonSource {
    pub fn new(path: String) -> Self {
        JsonSource { path, records: String::new(), columns: Vec::new(), combine: false }
    }
}

/// Whether a file name says one record per line.
pub fn is_lines(name: &str) -> bool {
    let lower = name.to_lowercase();
    lower.ends_with(".jsonl") || lower.ends_with(".ndjson")
}

/// Whether a file name is JSON this source reads.
pub fn is_json(name: &str) -> bool {
    let lower = name.to_lowercase();
    lower.ends_with(".json") || is_lines(&lower)
}

fn text(bytes: &[u8]) -> Result<&str, String> {
    let s = std::str::from_utf8(bytes).map_err(|_| "the file isn't UTF-8 text, so it isn't JSON".to_string())?;
    Ok(s.strip_prefix('\u{feff}').unwrap_or(s))
}

/// Every place in a `.json` file that holds an array of objects, as dotted
/// paths ("" is the file itself), largest first: what Records at offers.
pub fn record_paths(bytes: &[u8]) -> Result<Vec<(String, usize)>, String> {
    let value = parse(text(bytes)?).map_err(|e| format!("not JSON: {e}"))?;
    let mut found = Vec::new();
    find_arrays(&value, String::new(), 0, &mut found);
    found.sort_by(|a, b| b.1.cmp(&a.1));
    Ok(found)
}

fn find_arrays(v: &Value, path: String, depth: usize, out: &mut Vec<(String, usize)>) {
    if depth > 6 {
        return;
    }
    match v {
        Value::Array(items) if items.iter().any(Value::is_object) => out.push((path, items.len())),
        Value::Object(map) => {
            for (k, child) in map {
                let p = if path.is_empty() { k.clone() } else { format!("{path}.{k}") };
                find_arrays(child, p, depth + 1, out);
            }
        }
        _ => {}
    }
}

/// The records of one file, each with the line (or position) it came from.
fn records(src: &JsonSource, snapshot: &Snapshot) -> Result<(Vec<(usize, Value)>, Option<String>), String> {
    let body = text(&snapshot.bytes)?;
    let name = snapshot.path.file_name().and_then(|n| n.to_str()).unwrap_or("");
    if is_lines(name) {
        let mut out = Vec::new();
        for (i, line) in body.lines().enumerate() {
            let line = line.trim();
            if line.is_empty() {
                continue;
            }
            let v = parse(line).map_err(|e| format!("line {}: not JSON: {e}", i + 1))?;
            out.push((i + 1, v));
        }
        return Ok((out, None));
    }
    let root = parse(body).map_err(|e| format!("not JSON: {e}"))?;
    let (at, found_path) = if !src.records.trim().is_empty() {
        let mut v = &root;
        for part in src.records.trim().split('.') {
            v = match v {
                Value::Object(_) => v.get(part),
                Value::Array(a) => part.parse::<usize>().ok().and_then(|i| a.get(i)),
                _ => None,
            }
            .ok_or_else(|| format!("the file has nothing at {}", src.records.trim()))?;
        }
        (v, None)
    } else {
        // Found by itself: the file is the array, or its largest array of
        // objects; a single object is one record
        let mut found = Vec::new();
        find_arrays(&root, String::new(), 0, &mut found);
        found.sort_by(|a, b| b.1.cmp(&a.1));
        match (&root, found.first()) {
            (Value::Array(_), _) => (&root, None),
            (_, Some((path, _))) => {
                let mut v = &root;
                for part in path.split('.') {
                    v = v.get(part).expect("a path find_arrays reported");
                }
                (v, Some(path.clone()))
            }
            _ => (&root, None),
        }
    };
    let list = match at {
        Value::Array(items) => items.iter().cloned().enumerate().map(|(i, v)| (i + 1, v)).collect(),
        other => vec![(1, other.clone())],
    };
    Ok((list, found_path))
}

/// One cell: (text, how the column should treat it).
#[derive(Clone, Copy, PartialEq)]
enum Kind {
    Number,
    Text,
}

fn flatten(prefix: &str, v: &Value, depth: usize, out: &mut Vec<(String, String, Kind)>) {
    match v {
        Value::Object(map) if depth < MAX_FLATTEN_DEPTH && !map.is_empty() => {
            for (k, child) in map {
                let key = if prefix.is_empty() { k.clone() } else { format!("{prefix}.{k}") };
                flatten(&key, child, depth + 1, out);
            }
        }
        Value::Null => out.push((prefix.to_string(), String::new(), Kind::Number)),
        Value::Bool(b) => out.push((prefix.to_string(), if *b { "TRUE" } else { "FALSE" }.to_string(), Kind::Text)),
        Value::Number(n) => {
            // An integer past 2^53 can't be held exactly: keep its digits
            let integer = !n.contains(['.', 'e', 'E']);
            let digits = n.trim_start_matches('-');
            let big = integer && digits.parse::<u64>().map_or(true, |u| u > MAX_EXACT_INTEGER);
            out.push((prefix.to_string(), n.clone(), if big { Kind::Text } else { Kind::Number }));
        }
        Value::String(s) => out.push((prefix.to_string(), s.clone(), Kind::Text)),
        // Arrays (and over-deep or empty objects) stay as JSON in one cell
        other => out.push((prefix.to_string(), other.to_json(), Kind::Text)),
    }
}

/// The snapshot's records as a frame. Also notes where the records were
/// found when that was guessed.
pub(super) fn read_frame(src: &JsonSource, snapshot: &Snapshot) -> Result<(Frame, Option<String>), String> {
    let (recs, found) = records(src, snapshot)?;
    let mut names: Vec<String> = Vec::new();
    let mut index: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
    let mut numeric: Vec<bool> = Vec::new();
    let mut seen: Vec<bool> = Vec::new();
    let mut rows: Vec<Vec<String>> = Vec::with_capacity(recs.len());
    let mut lines = Vec::with_capacity(recs.len());
    for (line, rec) in &recs {
        let mut cells = Vec::new();
        match rec {
            Value::Object(_) => flatten("", rec, 0, &mut cells),
            // A record that isn't an object: one "value" column
            other => flatten("value", other, 0, &mut cells),
        }
        let mut row = vec![String::new(); names.len()];
        for (name, value, kind) in cells {
            let name = if name.is_empty() { "value".to_string() } else { name };
            let i = *index.entry(name.clone()).or_insert_with(|| {
                names.push(name);
                numeric.push(true);
                seen.push(false);
                names.len() - 1
            });
            if row.len() < names.len() {
                row.resize(names.len(), String::new());
            }
            if !value.is_empty() {
                seen[i] = true;
                if kind == Kind::Text {
                    numeric[i] = false;
                }
            }
            row[i] = value;
        }
        rows.push(row);
        lines.push(*line);
    }
    if names.len() > visigrid_engine::sheet::NUM_COLS {
        return Err(format!("the records have {} fields; a sheet holds {} columns", names.len(), visigrid_engine::sheet::NUM_COLS));
    }
    for row in &mut rows {
        row.resize(names.len(), String::new());
    }
    let columns = names
        .into_iter()
        .enumerate()
        .map(|(i, name)| {
            let number = numeric[i] && seen[i];
            OutColumn {
                name,
                rule: if number { ColumnRule::Number } else { ColumnRule::Text },
                kind: if number { ValueKind::Number } else { ValueKind::Plain },
            }
        })
        .collect();
    Ok((Frame { columns, rows, lines, files: Vec::new(), file_names: Vec::new(), decimal_comma: false }, found))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::Path;

    fn frame(name: &str, body: &str, records: &str) -> (Frame, Option<String>) {
        let mut src = JsonSource::new(name.into());
        src.records = records.into();
        read_frame(&src, &Snapshot::from_bytes(Path::new(name), body.as_bytes().to_vec())).unwrap()
    }

    fn names(f: &Frame) -> Vec<&str> {
        f.columns.iter().map(|c| c.name.as_str()).collect()
    }

    #[test]
    fn nested_objects_flatten_in_first_seen_order() {
        let body = r#"[
            {"id": 1, "customer": {"name": "Acme", "address": {"city": "Franklin"}}, "total": 10.5},
            {"id": 2, "total": 3, "customer": {"name": "Beta"}, "paid": true, "tags": ["a", "b"]}
        ]"#;
        let (f, found) = frame("x.json", body, "");
        assert_eq!(found, None);
        assert_eq!(names(&f), ["id", "customer.name", "customer.address.city", "total", "paid", "tags"]);
        assert_eq!(f.rows[0], ["1", "Acme", "Franklin", "10.5", "", ""]);
        assert_eq!(f.rows[1], ["2", "Beta", "", "3", "TRUE", r#"["a","b"]"#]);
        assert_eq!(f.columns[0].rule, ColumnRule::Number);
        assert_eq!(f.columns[1].rule, ColumnRule::Text);
        assert_eq!(f.columns[5].rule, ColumnRule::Text, "arrays stay JSON text");
    }

    #[test]
    fn records_are_found_or_named_by_a_dotted_path() {
        let body = r#"{"meta": {"count": 2}, "data": {"items": [{"a": 1}, {"a": 2}], "other": [{"b": 1}]}}"#;
        let (f, found) = frame("x.json", body, "");
        assert_eq!(found.as_deref(), Some("data.items"), "the largest array of objects");
        assert_eq!(f.rows.len(), 2);
        let (f, found) = frame("x.json", body, "data.other");
        assert_eq!((f.rows.len(), found), (1, None));
        let src = JsonSource { records: "data.nope".into(), ..JsonSource::new("x.json".into()) };
        let snap = Snapshot::from_bytes(Path::new("x.json"), body.as_bytes().to_vec());
        assert!(read_frame(&src, &snap).err().unwrap().contains("nothing at data.nope"));
        assert_eq!(record_paths(body.as_bytes()).unwrap(), vec![("data.items".to_string(), 2), ("data.other".to_string(), 1)]);
        // A single object is one record
        let (f, _) = frame("x.json", r#"{"a": 1, "b": "x"}"#, "");
        assert_eq!(f.rows, vec![vec!["1", "x"]]);
    }

    #[test]
    fn json_lines_keep_their_line_numbers() {
        let body = "{\"id\": 1}\n\n{\"id\": 2, \"note\": \"x\"}\n";
        let (f, _) = frame("x.ndjson", body, "");
        assert_eq!(f.rows, vec![vec!["1", ""], vec!["2", "x"]]);
        assert_eq!(f.lines, vec![1, 3]);
        let src = JsonSource::new("x.jsonl".into());
        let bad = Snapshot::from_bytes(Path::new("x.jsonl"), b"{\"id\": 1}\n{oops\n".to_vec());
        assert!(read_frame(&src, &bad).err().unwrap().starts_with("line 2:"));
    }

    #[test]
    fn large_integers_and_numeric_strings_stay_text() {
        let body = r#"[{"id": 9007199254740993, "n": 5, "zip": "02134", "maybe": null}, {"id": 1, "n": 6, "zip": "10001", "maybe": 7}]"#;
        let (f, _) = frame("x.json", body, "");
        assert_eq!(f.rows[0][0], "9007199254740993", "digits kept exactly");
        assert_eq!(f.columns[0].rule, ColumnRule::Text, "a column with an inexact integer is text");
        assert_eq!(f.columns[1].rule, ColumnRule::Number);
        assert_eq!(f.columns[2].rule, ColumnRule::Text, "the file said string");
        assert_eq!(f.columns[3].rule, ColumnRule::Number, "null is empty, not text");
        assert_eq!(f.rows[0][3], "");
        let (f, _) = frame("x.json", r#"[{"id": 123456789012345678901234567890, "amt": 10.50}]"#, "");
        assert_eq!(f.rows[0], vec!["123456789012345678901234567890", "10.50"], "the file's own text");
        assert_eq!(f.columns[0].rule, ColumnRule::Text);
    }
}
