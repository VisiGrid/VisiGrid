use visigrid_engine::{
    cell::CellComment,
    sheet::{Sheet, SheetId},
};
fn note() -> CellComment {
    CellComment {
        text: "First line\n日本語 café 🙂".into(),
        author: "Robert".into(),
    }
}
#[test]
fn comments_follow_structural_edits_and_do_not_change_values() {
    let mut sheet = Sheet::new(SheetId(1), 100, 20);
    sheet.set_value(4, 3, "=1+2");
    sheet.set_comment(4, 3, Some(note()));
    sheet.set_comment(7, 3, Some(note())); // A blank cell must still retain its comment.
    sheet.insert_rows(2, 2);
    sheet.insert_cols(1, 1);
    assert_eq!(sheet.comment(6, 4), Some(&note()));
    assert_eq!(sheet.comment(9, 4), Some(&note()));
    sheet.delete_rows(0, 2);
    sheet.delete_cols(0, 1);
    assert_eq!(sheet.comment(4, 3), Some(&note()));
    assert_eq!(sheet.get_raw(4, 3), "=1+2");
    sheet.set_value(4, 3, "");
    assert_eq!(
        sheet.comment(4, 3),
        Some(&note()),
        "clear contents preserves comments"
    );
    let clone = sheet.clone();
    sheet.set_comment(4, 3, None);
    assert_eq!(clone.comment(4, 3), Some(&note()));
    assert!(sheet.comment(4, 3).is_none());
    sheet.delete_rows(7, 1);
    assert_eq!(sheet.comments().count(), 0);
}
#[test]
fn comments_serialize_on_blank_cells() {
    let mut sheet = Sheet::new(SheetId(1), 20, 10);
    sheet.set_comment(8, 6, Some(note()));
    let json = serde_json::to_string(&sheet.get_cell(8, 6)).unwrap();
    let decoded: visigrid_engine::cell::Cell = serde_json::from_str(&json).unwrap();
    assert_eq!(decoded.comment(), Some(&note()));
}
