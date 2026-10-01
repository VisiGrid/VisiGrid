//! Reproducible Excel Notes QA: cargo run -p visigrid-io --example
//! xlsx_comments_probe -- OUTPUT_DIRECTORY
use std::{env, fs, path::PathBuf};
use visigrid_engine::cell::CellComment;
use visigrid_io::{native, xlsx};

fn main() -> Result<(), String> {
    let output = PathBuf::from(env::args_os().nth(1).ok_or("Pass an output directory")?);
    fs::create_dir_all(&output).map_err(|e| e.to_string())?;
    let source = output.join("notes-source.xlsx");
    let mut excel = rust_xlsxwriter::Workbook::new();
    let title = rust_xlsxwriter::Format::new()
        .set_bold()
        .set_font_color("#FFFFFF")
        .set_background_color("#275D7E");
    let sheet = excel.add_worksheet();
    sheet.set_name("Notes QA").map_err(|e| e.to_string())?;
    sheet.set_column_width(0, 36).map_err(|e| e.to_string())?;
    sheet.set_column_width(1, 24).map_err(|e| e.to_string())?;
    sheet
        .write_string_with_format(0, 0, "Excel Notes round-trip", &title)
        .map_err(|e| e.to_string())?;
    for (row, label) in [
        (2, "Amount — edit this note"),
        (4, "Blank B5 — keep this note"),
        (6, "Delete B7 note in edited copy"),
    ] {
        sheet
            .write_string(row, 0, label)
            .map_err(|e| e.to_string())?;
    }
    sheet
        .write_number(2, 1, 1234.5)
        .map_err(|e| e.to_string())?;
    for (row, text) in [
        (
            2,
            "Check this amount.\nSecond line — café 日本語 🙂 & <text>",
        ),
        (4, "A note on a blank cell."),
        (6, "This note will be removed from the edited copy."),
    ] {
        sheet
            .insert_note(
                row,
                1,
                &rust_xlsxwriter::Note::new(text).set_author("QA Author"),
            )
            .map_err(|e| e.to_string())?;
    }
    excel
        .add_worksheet()
        .set_name("Other sheet")
        .map_err(|e| e.to_string())?
        .insert_note(
            0,
            0,
            &rust_xlsxwriter::Note::new("Cross-sheet note").set_author("QA Author"),
        )
        .map_err(|e| e.to_string())?;
    excel.save(&source).map_err(|e| e.to_string())?;
    let (mut wb, stats) = xlsx::import(&source)?;
    assert_eq!(stats.comments_imported, 4);
    let layouts: Vec<_> = stats
        .imported_layouts
        .iter()
        .map(|l| l.to_export_layout())
        .collect();
    xlsx::export(&wb, &output.join("notes-unchanged.xlsx"), Some(&layouts))?;
    native::save_workbook(&wb, &output.join("notes.sheet"))?;
    let restored = native::load_workbook(&output.join("notes.sheet"))?;
    xlsx::export(
        &restored,
        &output.join("notes-native-roundtrip.xlsx"),
        Some(&layouts),
    )?;
    wb.sheet_mut(0).unwrap().set_comment(
        2,
        1,
        Some(CellComment {
            text: "Edited in VisiGrid.\nAmount checked.".into(),
            author: "Zelda".into(),
        }),
    );
    wb.sheet_mut(0).unwrap().set_comment(6, 1, None);
    wb.sheet_mut(0).unwrap().set_comment(
        8,
        1,
        Some(CellComment {
            text: "Created in VisiGrid — no author.".into(),
            author: String::new(),
        }),
    );
    wb.sheet_mut(0).unwrap().set_text(8, 0, "New note at B9");
    xlsx::export(&wb, &output.join("notes-edited.xlsx"), Some(&layouts))?;
    for name in [
        "notes-unchanged.xlsx",
        "notes-native-roundtrip.xlsx",
        "notes-edited.xlsx",
    ] {
        let (_, stats) = xlsx::import(&output.join(name))?;
        assert_eq!(stats.comments_imported, 4);
        println!("{name}: {}", stats.summary());
    }
    Ok(())
}
