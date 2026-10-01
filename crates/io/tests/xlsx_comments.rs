use std::io::{Read, Write};
use visigrid_engine::{cell::CellComment, sheet::MergedRegion, workbook::Workbook};
use visigrid_io::{native, xlsx};

fn note(text: &str, author: &str) -> CellComment {
    CellComment {
        text: text.into(),
        author: author.into(),
    }
}

fn rewrite(
    source: &std::path::Path,
    target: &std::path::Path,
    change: impl Fn(&str, Vec<u8>) -> (String, Vec<u8>),
) {
    let mut input = zip::ZipArchive::new(std::fs::File::open(source).unwrap()).unwrap();
    let mut output = zip::ZipWriter::new(std::fs::File::create(target).unwrap());
    for i in 0..input.len() {
        let mut entry = input.by_index(i).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        let (name, bytes) = change(entry.name(), bytes);
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        output.write_all(&bytes).unwrap();
    }
    output.finish().unwrap();
}

#[test]
fn notes_survive_repeated_xlsx_native_roundtrips_and_edits() {
    let dir = tempfile::tempdir().unwrap();
    let file = dir.path().join("notes.xlsx");
    let mut wb = Workbook::new();
    wb.add_sheet();
    let text = "  Leading space\n日本語 café 🙂 & <tag>\nLiteral _x0041_\tend  ";
    let n = note(text, "A & B");
    wb.sheet_mut(0).unwrap().set_value(1, 1, "=1+2");
    wb.sheet_mut(0).unwrap().set_comment(1, 1, Some(n.clone()));
    wb.sheet_mut(0)
        .unwrap()
        .add_merge(MergedRegion::new(4, 0, 4, 2))
        .unwrap();
    wb.sheet_mut(0)
        .unwrap()
        .set_comment(4, 0, Some(note("Merged", "")));
    wb.sheet_mut(1)
        .unwrap()
        .set_comment(80, 12, Some(note("Blank distant cell", "QA")));
    for _ in 0..3 {
        assert_eq!(xlsx::export(&wb, &file, None).unwrap().comments_exported, 3);
        let (loaded, report) = xlsx::import(&file).unwrap();
        assert_eq!(report.comments_imported, 3);
        assert_eq!(loaded.sheets()[0].comment(1, 1), Some(&n));
        assert_eq!(loaded.sheets()[0].comment(4, 0), Some(&note("Merged", "")));
        assert_eq!(
            loaded.sheets()[1].comment(80, 12),
            Some(&note("Blank distant cell", "QA"))
        );
        assert_eq!(loaded.sheets()[0].get_raw(1, 1), "=1+2");
        assert_eq!(loaded.sheets()[1].get_raw(80, 12), "");
        let native_path = dir.path().join("notes.sheet");
        native::save_workbook(&loaded, &native_path).unwrap();
        wb = native::load_workbook(&native_path).unwrap();
    }
    wb.sheet_mut(0)
        .unwrap()
        .set_comment(1, 1, Some(note("Edited", "New author")));
    wb.sheet_mut(0).unwrap().set_comment(4, 0, None);
    wb.sheet_mut(1)
        .unwrap()
        .set_comment(3, 3, Some(note("New", "")));
    let (bytes, report) = xlsx::export_to_buffer(&wb, None).unwrap();
    assert_eq!(report.comments_exported, 3);
    std::fs::write(&file, bytes).unwrap();
    let (loaded, _) = xlsx::import(&file).unwrap();
    assert_eq!(
        loaded.sheets()[0].comment(1, 1),
        Some(&note("Edited", "New author"))
    );
    assert!(loaded.sheets()[0].comment(4, 0).is_none());
    assert_eq!(loaded.sheets()[1].comment(3, 3), Some(&note("New", "")));
    let mut empty = loaded;
    for i in 0..empty.sheet_count() {
        let positions: Vec<_> = empty.sheets()[i].comments().map(|(p, _)| p).collect();
        for (r, c) in positions {
            empty.sheet_mut(i).unwrap().set_comment(r, c, None);
        }
    }
    xlsx::export(&empty, &file, None).unwrap();
    let archive = zip::ZipArchive::new(std::fs::File::open(&file).unwrap()).unwrap();
    assert!(!archive
        .file_names()
        .any(|n| n.contains("comments") || n.ends_with(".vml")));
}

#[test]
fn excel_rich_text_and_nonstandard_relationship_targets_import() {
    let dir = tempfile::tempdir().unwrap();
    let source = dir.path().join("source.xlsx");
    let target = dir.path().join("renamed.xlsx");
    let mut excel = rust_xlsxwriter::Workbook::new();
    excel
        .add_worksheet()
        .set_name("Other & sheet")
        .unwrap()
        .insert_note(
            2,
            1,
            &rust_xlsxwriter::Note::new("Body").set_author("Excel Author"),
        )
        .unwrap();
    excel.save(&source).unwrap();
    rewrite(&source, &target, |name, bytes| {
        if name == "xl/comments1.xml" {
            return ("xl/notes/custom.xml".into(), br#"<?xml version="1.0"?><s:comments xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:authors><s:author>A &amp; B</s:author></s:authors><s:commentList><s:comment ref="B3" authorId="0"><s:text><s:r><s:rPr><s:b/></s:rPr><s:t xml:space="preserve"> First &lt;part&gt; </s:t></s:r><s:r><s:t>second&#10;_x0041_ _x005F_x0041_</s:t></s:r><s:rPh sb="0" eb="1"><s:t>phonetic</s:t></s:rPh></s:text></s:comment></s:commentList></s:comments>"#.to_vec());
        }
        if name.ends_with(".rels") || name == "[Content_Types].xml" {
            let xml = String::from_utf8(bytes)
                .unwrap()
                .replace("../comments1.xml", "/xl/notes/./custom.xml")
                .replace("/xl/comments1.xml", "/xl/notes/custom.xml");
            return (name.into(), xml.into_bytes());
        }
        (name.into(), bytes)
    });
    let (wb, report) = xlsx::import(&target).unwrap();
    assert_eq!(report.comments_imported, 1);
    assert_eq!(wb.sheets()[0].name, "Other & sheet");
    assert_eq!(
        wb.sheets()[0].comment(2, 1),
        Some(&note(" First <part> second\nA _x0041_", "A & B"))
    );
    // Real writer author prefixes are part of the body: retain them exactly,
    // without adding another author prefix on each export.
    let (original, _) = xlsx::import(&source).unwrap();
    let comment = original.sheets()[0].comment(2, 1).unwrap().clone();
    assert!(comment.text.contains("Excel Author"));
    xlsx::export(&original, &target, None).unwrap();
    assert_eq!(
        xlsx::import(&target).unwrap().0.sheets()[0].comment(2, 1),
        Some(&comment)
    );
}

#[test]
fn unsupported_or_broken_notes_never_stop_an_import() {
    // A workbook must open whatever state its notes are in: unreadable
    // notes are skipped and named in a warning, never the whole file.
    // Threaded comments (Excel 365) keep the Note Excel saves beside them.
    let dir = tempfile::tempdir().unwrap();
    let source = dir.path().join("source.xlsx");
    let target = dir.path().join("broken.xlsx");
    let mut wb = Workbook::new();
    wb.active_sheet_mut()
        .set_comment(0, 0, Some(note("Keep me", "QA")));
    wb.active_sheet_mut().set_value(0, 1, "data");
    xlsx::export(&wb, &source, None).unwrap();
    for replacement in [
        "threaded",
        "threaded-rels",
        "author",
        "reference",
        "missing",
        "document",
        "empty",
    ] {
        rewrite(&source, &target, |name, bytes| {
            let xml = String::from_utf8(bytes.clone()).unwrap_or_default();
            let changed = match replacement {
                "document" if name == "xl/comments1.xml" => "<unrelated/>".into(),
                "empty" if name == "xl/comments1.xml" => "<comments><authors><author>QA</author></authors><commentList><comment ref=\"A1\" authorId=\"0\"/></commentList></comments>".into(),
                "threaded" if name == "[Content_Types].xml" => xml.replace("</Types>", "<Override PartName=\"/xl/custom.xml\" ContentType=\"application/vnd.ms-excel.threadedcomments+xml\"/></Types>"),
                // As Excel 365 writes them: a threadedComment part beside the
                // legacy comments part, and a person list on the workbook.
                "threaded-rels" if name.ends_with("sheet1.xml.rels") => xml.replace("</Relationships>", "<Relationship Id=\"rIdT\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/threadedComment\" Target=\"../threadedComments/threadedComment1.xml\"/></Relationships>"),
                "threaded-rels" if name == "xl/_rels/workbook.xml.rels" => xml.replace("</Relationships>", "<Relationship Id=\"rIdP\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/person\" Target=\"persons/person.xml\"/></Relationships>"),
                "author" if name == "xl/comments1.xml" => xml.replace("authorId=\"0\"", "authorId=\"900\"").replace("authorId=\"1\"", "authorId=\"900\""),
                "reference" if name == "xl/comments1.xml" => xml.replace("ref=\"A1\"", "ref=\"A0\""),
                "missing" if name.ends_with("sheet1.xml.rels") => xml.replace("comments1.xml", "missing.xml"),
                _ => return (name.into(), bytes),
            };
            (name.into(), changed.into_bytes())
        });
        let (loaded, report) = xlsx::import(&target).unwrap_or_else(|e| panic!("{replacement}: {e}"));
        assert_eq!(loaded.sheet(0).unwrap().get_raw(0, 1), "data", "{replacement}: cells still import");
        if replacement.starts_with("threaded") {
            assert_eq!(loaded.sheet(0).unwrap().comment(0, 0).map(|c| c.text.as_str()), Some("Keep me"), "{replacement}");
            if replacement == "threaded" {
                assert!(report.warnings.iter().any(|w| w.contains("threaded")), "{:?}", report.warnings);
            }
        } else {
            assert_eq!(report.comments_imported, 0, "{replacement}");
            assert!(loaded.sheet(0).unwrap().comment(0, 0).is_none(), "{replacement}");
            assert!(report.warnings.iter().any(|w| w.contains("Notes on 'Sheet1'")), "{replacement}: {:?}", report.warnings);
        }
    }
    wb.active_sheet_mut()
        .set_comment(0, 0, Some(note(&"a".repeat(32768), "QA")));
    assert!(xlsx::export_to_buffer(&wb, None)
        .unwrap_err()
        .contains("32,767-character"));
}

#[test]
fn author_order_empty_authors_and_excel_string_limits_are_preserved() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("authors.xlsx");
    let mut wb = Workbook::new();
    // An uncommented first sheet must not offset the comment-part mapping.
    wb.add_sheet();
    let notes = [
        note("Z first", "Zelda"),
        note("A second", "Alice"),
        note("No author", ""),
        note("\r\n\u{1} _x005F_ _x0041_ 🙂", "_x0041_"),
        note(&"a".repeat(32767), &"Long author ".repeat(8)),
        note("", "Empty text still has metadata"),
    ];
    for (r, comment) in notes.iter().enumerate() {
        wb.sheet_mut(1)
            .unwrap()
            .set_comment(r, 1, Some(comment.clone()));
    }
    xlsx::export(&wb, &path, None).unwrap();
    let (loaded, report) = xlsx::import(&path).unwrap();
    assert_eq!(report.comments_imported, notes.len());
    assert_eq!(loaded.sheets()[0].comments().count(), 0);
    for (r, comment) in notes.iter().enumerate() {
        assert_eq!(loaded.sheets()[1].comment(r, 1), Some(comment));
    }
}
