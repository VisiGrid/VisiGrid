use visigrid_engine::{cell::ValueRef, sheet::MergedRegion, workbook::Workbook};
use visigrid_io::native;

#[test]
fn native_save_modes_preserve_literal_text_and_per_sheet_merges() {
    let mut wb = Workbook::new();
    wb.add_sheet();
    wb.add_sheet();
    let text = ["00123", "=SUM(A1:A2)", "1e6", "007", ""];
    for (r, value) in text.iter().enumerate() {
        wb.sheet_mut(0).unwrap().set_text(r, 3, value);
    }
    wb.sheet_mut(0)
        .unwrap()
        .add_merge(MergedRegion::new(8, 0, 9, 1))
        .unwrap();
    wb.sheet_mut(2)
        .unwrap()
        .add_merge(MergedRegion::new(1, 1, 3, 3))
        .unwrap();
    wb.set_active_sheet(2);
    let dir = tempfile::tempdir().unwrap();
    for mode in 0..3 {
        let path = dir.path().join(format!("mode-{mode}.sheet"));
        match mode {
            0 => native::save_workbook(&wb, &path),
            1 => native::save_workbook_with_metadata(&wb, &Default::default(), &path),
            _ => native::save_workbook_full(&wb, &Default::default(), &[], &[], &path),
        }
        .unwrap();
        let loaded = native::load_workbook(&path).unwrap();
        assert_eq!(loaded.active_sheet_index(), 2);
        for i in 0..3 {
            assert_eq!(
                loaded.sheets()[i].merged_regions,
                wb.sheets()[i].merged_regions,
                "mode {mode}, sheet {i}"
            );
        }
        for (r, value) in text.iter().enumerate() {
            assert_eq!(
                loaded.sheets()[0].get_raw(r, 3),
                *value,
                "mode {mode}, row {r}"
            );
            if !value.is_empty() {
                assert!(matches!(
                    loaded.sheets()[0].get_cell(r, 3).value(),
                    ValueRef::Text(_)
                ));
            }
        }
    }
}

#[test]
fn legacy_single_sheet_loader_preserves_literal_text() {
    let mut wb = Workbook::new();
    wb.active_sheet_mut().set_text(0, 0, "00123");
    wb.active_sheet_mut().set_text(1, 0, "=1+2");
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("legacy.sheet");
    native::save(wb.active_sheet(), &path).unwrap();
    let loaded = native::load(&path).unwrap();
    assert_eq!(loaded.get_raw(0, 0), "00123");
    assert!(matches!(
        loaded.get_cell(1, 0).value(),
        ValueRef::Text("=1+2")
    ));
}

#[test]
fn native_comments_survive_all_save_modes_and_deletion() {
    use visigrid_engine::cell::CellComment;
    let note = CellComment {
        text: "Line 1\n日本語 café 🙂 <text>".into(),
        author: "Robert".into(),
    };
    let dir = tempfile::tempdir().unwrap();
    for mode in 0..3 {
        let mut wb = Workbook::new();
        wb.add_sheet();
        wb.sheet_mut(0)
            .unwrap()
            .set_comment(0, 0, Some(note.clone()));
        wb.sheet_mut(1)
            .unwrap()
            .set_comment(80, 9, Some(note.clone()));
        wb.sheet_mut(1).unwrap().set_value(80, 9, "42");
        let path = dir.path().join(format!("comments-{mode}.sheet"));
        let save = |wb: &Workbook| match mode {
            0 => native::save_workbook(wb, &path),
            1 => native::save_workbook_with_metadata(wb, &Default::default(), &path),
            _ => native::save_workbook_full(wb, &Default::default(), &[], &[], &path),
        };
        save(&wb).unwrap();
        let mut loaded = native::load_workbook(&path).unwrap();
        assert_eq!(loaded.sheet(0).unwrap().comment(0, 0), Some(&note));
        assert_eq!(loaded.sheet(1).unwrap().comment(80, 9), Some(&note));
        assert_eq!(loaded.sheet(1).unwrap().get_raw(80, 9), "42");
        loaded.sheet_mut(1).unwrap().set_comment(80, 9, None);
        save(&loaded).unwrap();
        let loaded = native::load_workbook(&path).unwrap();
        assert!(loaded.sheet(1).unwrap().comment(80, 9).is_none());
        assert_eq!(loaded.sheet(0).unwrap().comment(0, 0), Some(&note));
    }
}

#[test]
fn legacy_comments_roundtrip() {
    use visigrid_engine::cell::CellComment;
    let mut wb = Workbook::new();
    let note = CellComment {
        text: "Blank cell note".into(),
        author: "".into(),
    };
    wb.active_sheet_mut().set_comment(50, 4, Some(note.clone()));
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("legacy-comments.sheet");
    native::save(wb.active_sheet(), &path).unwrap();
    assert_eq!(native::load(&path).unwrap().comment(50, 4), Some(&note));
}
