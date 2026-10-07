//! Large sheets as row bands: a banded export, loaded band by band in any
//! order, is the same workbook as the inline one.
use visigrid_engine::cell::CellFormat;
use visigrid_engine::sheet::{Sheet, SheetId, NUM_COLS, NUM_ROWS};
use visigrid_engine::workbook::Workbook;
use visigrid_io::json::{bands, export_workbook, import_any, SheetLayout};

fn big() -> Workbook {
    let mut s = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
    s.set_name("Data");
    let bold = CellFormat { bold: true, ..Default::default() };
    // 150,000 rows x 2 = 300,000 cells: over the threshold, three bands.
    for r in 0..150_000 {
        s.set_value_deferred(r, 0, &format!("{}", r % 997));
        s.set_value_deferred(r, 1, &format!("t{r}"));
    }
    s.set_format(70_000, 1, bold);
    // Formulas reaching across bands, and one that spills.
    s.set_value_deferred(0, 3, "=SUM(A1:A150000)");
    s.set_value_deferred(140_000, 3, "=A1+A70001");
    s.set_value_deferred(5, 5, "=SEQUENCE(3)");
    let small = {
        let mut t = Sheet::new(SheetId(2), NUM_ROWS, NUM_COLS);
        t.set_name("Small");
        t.set_value_deferred(0, 0, "=Data!D1");
        t
    };
    let mut wb = Workbook::from_sheets(vec![s, small], 0);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    wb
}

fn shown(wb: &Workbook, sheet: usize, r: usize, c: usize) -> (String, String, bool) {
    let s = &wb.sheets()[sheet];
    (s.get_raw(r, c), s.get_formatted_display(r, c), s.get_format(r, c).bold)
}

#[test]
fn banded_export_round_trips_in_any_band_order() {
    let wb = big();
    let layouts = vec![SheetLayout::default(), SheetLayout::default()];
    let (manifest, out) = bands::export_banded(&wb, &layouts, 0).unwrap();
    assert_eq!(out.len(), 10, "150,000 rows make ten 16,384-row bands");
    assert!(out.iter().all(|b| b.sheet == 0), "the small sheet stays inline");
    assert!(manifest.len() < 4096, "the manifest carries no cells: {} bytes", manifest.len());
    let refs = bands::manifest_bands(&manifest).unwrap();
    assert_eq!(refs.len(), 10);

    let (mut loaded, _, _) = import_any(&manifest).unwrap();
    assert_eq!(loaded.sheets()[0].cells_iter().count(), 0);
    for b in out.iter().rev() {
        bands::apply(&mut loaded, &b.data, Some(&b.reference.key)).unwrap();
    }
    bands::finish(&mut loaded).unwrap();

    for &(sheet, r, c) in &[(0, 0, 0), (0, 70_000, 1), (0, 149_999, 1), (0, 0, 3), (0, 140_000, 3), (0, 5, 5), (0, 7, 5), (1, 0, 0)] {
        assert_eq!(shown(&loaded, sheet, r, c), shown(&wb, sheet, r, c), "cell {sheet}:{r}:{c}");
    }
    assert_eq!(loaded.sheets()[0].cells_iter().count(), wb.sheets()[0].cells_iter().count());
    // Inline and banded agree as documents too.
    assert_eq!(export_workbook(&loaded, &layouts, 0).unwrap(), export_workbook(&wb, &layouts, 0).unwrap());
}

#[test]
fn a_band_is_checked_against_its_key() {
    let wb = big();
    let (_, out) = bands::export_banded(&wb, &[SheetLayout::default(), SheetLayout::default()], 0).unwrap();
    let mut other = Workbook::new();
    let err = bands::apply(&mut other, &out[0].data, Some(&out[1].reference.key)).unwrap_err();
    assert!(err.contains("does not match"), "{err}");
}

#[test]
fn small_workbooks_are_not_banded() {
    let mut wb = Workbook::new();
    wb.sheet_mut(0).unwrap().set_value(0, 0, "1");
    let (manifest, out) = bands::export_banded(&wb, &[SheetLayout::default()], 0).unwrap();
    assert!(out.is_empty());
    let v = |s: &str| serde_json::from_str::<serde_json::Value>(s).unwrap();
    assert_eq!(v(&manifest), v(&export_workbook(&wb, &[SheetLayout::default()], 0).unwrap()));
}

fn tiny_band(cell: serde_json::Value) -> (String, Vec<u8>, String) {
    use sha2::{Digest, Sha256};
    let doc = format!(r#"{{"format":"visigrid-band","version":1,"sheet":0,"r0":0,"r1":2,"cells":[{cell}]}}"#);
    let data = miniz_oxide::deflate::compress_to_vec(doc.as_bytes(), 6);
    let key: String = Sha256::digest(&data).iter().map(|b| format!("{b:02x}")).collect();
    let manifest = serde_json::json!({"format":"visigrid-json","version":2,"sheets":[{"name":"Data","cells":[],"bands":[{"r0":0,"r1":2,"cells":1,"key":key,"bytes":data.len()}]}]}).to_string();
    (manifest, data, key)
}

#[test]
fn partial_workbook_cannot_edit_save_or_finish_early() {
    let (manifest, data, key) = tiny_band(serde_json::json!({"row":0,"col":0,"value":42}));
    let (mut wb, layouts, active) = import_any(&manifest).unwrap();
    assert!(wb.ensure_writable().is_err());
    assert!(bands::finish(&mut wb).is_err());
    assert!(export_workbook(&wb, &layouts, active).is_err());
    wb.sheet_mut(0).unwrap().read_only_reason = None;
    assert!(wb.ensure_writable().is_err());
    assert!(export_workbook(&wb, &layouts, active).is_err());
    bands::apply(&mut wb, &data, Some(&key)).unwrap();
    assert!(bands::apply(&mut wb, &data, Some(&key)).is_err());
    bands::finish(&mut wb).unwrap();
    wb.ensure_writable().unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(0,0), "42");
    assert!(export_workbook(&wb, &layouts, active).is_ok());
}

#[test]
fn unknown_or_unrepresentable_band_content_never_changes_the_preview() {
    for cell in [
        serde_json::json!({"row":0,"col":0,"value":42,"future":true}),
        serde_json::json!({"row":0,"col":0,"value":42,"fmt":{"future":true}}),
        serde_json::json!({"row":0,"col":0,"value":42,"fmt":{"align":"future"}}),
        serde_json::json!({"row":0,"col":0,"value":9007199254740993_u64}),
        serde_json::json!({"row":2,"col":0,"value":42}),
        serde_json::json!({"row":0,"col":0,"value":{"future":true}}),
        serde_json::json!({"row":0,"col":0,"value":42,"spill_from":[1,0],"fmt":{"bold":true}}),
        serde_json::json!({"row":0,"col":0,"value":42,"stale_custom_fn":true}),
    ] {
        let (manifest, data, key) = tiny_band(cell);
        let (mut wb, _, _) = import_any(&manifest).unwrap();
        assert!(bands::apply(&mut wb, &data, Some(&key)).is_err());
        assert!(wb.sheet(0).unwrap().cells_iter().next().is_none());
        assert_eq!(wb.pending_bands.as_ref().unwrap().remaining.len(),1);
        assert!(bands::finish(&mut wb).is_err());
    }
}

#[test]
fn unsupported_manifest_stays_locked_and_retains_original_source() {
    let (manifest, data, key) = tiny_band(serde_json::json!({"row":0,"col":0,"value":42}));
    let mut source: serde_json::Value = serde_json::from_str(&manifest).unwrap();
    source["future"] = serde_json::json!({"setting":"keep"});
    let original = source.to_string();
    let (mut wb, _, _) = import_any(&original).unwrap();
    assert!(bands::apply(&mut wb, &data, Some(&key)).is_err());
    assert_eq!(wb.sheet(0).unwrap().canonical_content_protection.as_ref().unwrap().source.as_str(), original);
    assert!(wb.ensure_writable().is_err());
}

#[test]
fn stable_collaboration_identities_survive_editable_json_round_trip() {
    let source = r#"{"format":"visigrid-json","version":2,"collab_sheet_ids":[17,4],"sheets":[{"name":"A","cells":[]},{"name":"B","cells":[]}]}"#;
    let (mut wb, layouts, active) = import_any(source).unwrap();
    wb.ensure_writable().unwrap();
    assert_eq!(wb.sheet(0).unwrap().id,SheetId(17));
    assert_eq!(wb.sheet(1).unwrap().id,SheetId(4));
    wb.sheet_mut(0).unwrap().set_value(0,0,"edited");
    let result: serde_json::Value = serde_json::from_str(&export_workbook(&wb,&layouts,active).unwrap()).unwrap();
    assert_eq!(result["collab_sheet_ids"],serde_json::json!([17,4]));
    assert!(import_any(&source.replace("[17,4]","[17,17]")).is_err());
    assert!(import_any(&source.replace("[17,4]","[17]")).is_err());
}

#[test]
fn band_loading_retains_cached_custom_formula_results() {
    let (manifest,data,key) = tiny_band(serde_json::json!({"row":0,"col":0,"formula":"=FUTURE_CUSTOM(1)","value":123}));
    let (mut wb,_,_) = import_any(&manifest).unwrap();
    bands::apply(&mut wb,&data,Some(&key)).unwrap();
    bands::finish(&mut wb).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(0,0),"123");
}

#[test]
fn a_late_band_error_is_atomic() {
    use sha2::{Digest,Sha256};
    let doc = r#"{"format":"visigrid-band","version":1,"sheet":0,"r0":0,"r1":2,"cells":[{"row":0,"col":0,"value":42},{"row":1,"col":0,"value":43,"future":true}]}"#;
    let data = miniz_oxide::deflate::compress_to_vec(doc.as_bytes(),6);
    let key: String = Sha256::digest(&data).iter().map(|b| format!("{b:02x}")).collect();
    let manifest = serde_json::json!({"format":"visigrid-json","version":2,"sheets":[{"name":"Data","bands":[{"r0":0,"r1":2,"cells":2,"key":key,"bytes":data.len()}]}]}).to_string();
    let (mut wb,_,_) = import_any(&manifest).unwrap();
    assert!(bands::apply(&mut wb,&data,Some(&key)).is_err());
    assert!(wb.sheet(0).unwrap().cells_iter().next().is_none());
    assert_eq!(wb.pending_bands.as_ref().unwrap().remaining.len(),1);
}

#[test]
fn bands_cannot_silently_move_cells_into_a_merge_anchor() {
    let (manifest,data,key) = tiny_band(serde_json::json!({"row":1,"col":0,"value":42}));
    let mut source: serde_json::Value = serde_json::from_str(&manifest).unwrap();
    source["sheets"][0]["merges"] = serde_json::json!([{"start_row":0,"start_col":0,"end_row":1,"end_col":0}]);
    let (mut wb,_,_) = import_any(&source.to_string()).unwrap();
    assert!(bands::apply(&mut wb,&data,Some(&key)).is_err());
    assert!(wb.sheet(0).unwrap().cells_iter().next().is_none());
    assert!(bands::finish(&mut wb).is_err());
}

#[test]
fn numeric_literals_beyond_u64_or_double_precision_keep_the_original() {
    for number in ["90071992547409930000124", "0.10000000000000000001"] {
        let source = format!(r#"{{"format":"visigrid-json","version":2,"sheets":[{{"name":"Precision","cells":[{{"row":0,"col":0,"value":{number}}}]}}]}}"#);
        let (wb,layouts,active) = import_any(&source).unwrap();
        assert!(wb.read_only_reason().is_some(), "unsafe numeric literal {number} remained editable");
        assert_eq!(export_workbook(&wb,&layouts,active).unwrap(),source);
    }
}

#[test]
fn numeric_precision_is_checked_in_opaque_metadata_and_formula_caches() {
    let source = r#"{"format":"visigrid-json","version":2,"sheets":[{"name":"Precision","cells":[{"row":0,"col":0,"formula":"=CUSTOM(1)","value":9007199254740993}]}]}"#;
    let (wb,layouts,active) = import_any(source).unwrap();
    assert!(wb.read_only_reason().is_some());
    assert_eq!(export_workbook(&wb,&layouts,active).unwrap(),source);
    let source = r#"{"format":"visigrid-json","version":2,"sheets":[{"name":"Precision","cells":[],"charts":[{"future_id":90071992547409930000124}]}]}"#;
    let (wb,layouts,active) = import_any(source).unwrap();
    assert!(wb.read_only_reason().is_some());
    assert_eq!(export_workbook(&wb,&layouts,active).unwrap(),source);
    let source = r#"{"format":"visigrid-json","version":2,"sheets":[{"name":"Text","cells":[{"row":0,"col":0,"value":"90071992547409930000124 \"quoted\" \\ 0.10000000000000000001"}]}]}"#;
    let (wb,_,_) = import_any(source).unwrap();
    wb.ensure_writable().unwrap();
}

#[test]
fn raw_band_numbers_are_checked_before_json_value_rounding() {
    use sha2::{Digest,Sha256};
    for number in ["90071992547409930000124", "0.10000000000000000001"] {
        let doc = format!(r#"{{"format":"visigrid-band","version":1,"sheet":0,"r0":0,"r1":2,"cells":[{{"row":0,"col":0,"value":{number}}}]}}"#);
        let data = miniz_oxide::deflate::compress_to_vec(doc.as_bytes(),6);
        let key: String = Sha256::digest(&data).iter().map(|b| format!("{b:02x}")).collect();
        let manifest = serde_json::json!({"format":"visigrid-json","version":2,"sheets":[{"name":"Data","bands":[{"r0":0,"r1":2,"cells":1,"key":key,"bytes":data.len()}]}]}).to_string();
        let (mut wb,_,_) = import_any(&manifest).unwrap();
        assert!(bands::apply(&mut wb,&data,Some(&key)).is_err());
        assert!(wb.sheet(0).unwrap().cells_iter().next().is_none());
        assert!(bands::finish(&mut wb).is_err());
    }
}

#[test]
fn ordinary_native_float_values_remain_editable_after_json_round_trip() {
    let mut seed = 0x5eed_u64;
    let cells: Vec<_> = (0..1000).map(|row| {
        seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1);
        let value = f64::from_bits(0x3ff0000000000000 | (seed & 0x000fffffffffffff));
        serde_json::json!({"row":row,"col":0,"value":value})
    }).collect();
    let source = serde_json::json!({"format":"visigrid-json","version":2,"sheets":[{"name":"Floats","cells":cells}]}).to_string();
    let (wb,_,_) = import_any(&source).unwrap();
    wb.ensure_writable().unwrap();
}

fn tiny_band_raw(cell: &str) -> (String, Vec<u8>, String) {
    use sha2::{Digest, Sha256};
    let doc = format!(r#"{{"format":"visigrid-band","version":1,"sheet":0,"r0":0,"r1":2,"cells":[{cell}]}}"#);
    let data = miniz_oxide::deflate::compress_to_vec(doc.as_bytes(), 6);
    let key: String = Sha256::digest(&data).iter().map(|b| format!("{b:02x}")).collect();
    let manifest = serde_json::json!({"format":"visigrid-json","version":2,"sheets":[{"name":"Data","cells":[],"bands":[{"r0":0,"r1":2,"cells":1,"key":key,"bytes":data.len()}]}]}).to_string();
    (manifest, data, key)
}

/// Band cells as literal text, accepted or refused exactly as the full
/// content check decides. The fast paths for plain cells and f64-exact
/// numbers must not change any of these.
const BAND_CELL_DECISIONS: &[(&str, bool)] = &[
    (r#"{"row":0,"col":0,"value":42}"#, true),
    (r#"{"row":1,"col":3,"value":7}"#, true),
    (r#"{"row":0,"col":0}"#, true),
    (r#"{"row":0,"col":0,"value":1.5}"#, true),
    (r#"{"row":0,"col":0,"value":0.1}"#, true),
    (r#"{"row":0,"col":0,"value":1.10}"#, true),
    (r#"{"row":0,"col":0,"value":-0.000123}"#, true),
    (r#"{"row":0,"col":0,"value":1e5}"#, true),
    (r#"{"row":0,"col":0,"value":1.0e-5}"#, true),
    (r#"{"row":0,"col":0,"value":-0}"#, true),
    (r#"{"row":0,"col":0,"value":123456789012345}"#, true),
    (r#"{"row":0,"col":0,"value":1234567890123456}"#, true),
    (r#"{"row":0,"col":0,"value":12345678901234567}"#, false),
    (r#"{"row":0,"col":0,"value":9007199254740993}"#, false),
    (r#"{"row":0,"col":0,"value":0.1234567890123456789}"#, false),
    (r#"{"row":0,"col":0,"value":0.30000000000000004}"#, true),
    (r#"{"row":0,"col":0,"value":1e400}"#, false),
    (r#"{"row":0,"col":0,"value":2.5e-320}"#, true),
    (r#"{"row":0,"col":0,"value":"text"}"#, true),
    (r#"{"row":0,"col":0,"value":"007"}"#, true),
    (r#"{"row":0,"col":0,"value":true}"#, true),
    (r#"{"row":0,"col":0,"value":null}"#, false),
    (r#"{"row":0,"col":0,"formula":null,"value":1}"#, false),
    (r#"{"row":0,"col":0,"value":1,"value":2}"#, false),
    (r#"{"row":0,"col":0,"value":[1]}"#, false),
    (r#"{"row":0,"col":0,"formula":"=1+1","value":2}"#, true),
    (r#"{"row":0,"col":0,"formula":"=1+1","value":0.1234567890123456789}"#, false),
    (r#"{"row":0,"col":0,"formula":"=\"value\""}"#, true),
    (r#"{"row":0,"col":0,"formula":"=1+1"}"#, true),
    (r#"{"row":0,"col":0,"value":42,"fmt":{"bold":true}}"#, true),
    (r#"{"row":0,"col":0,"value":42,"fmt":{"future":true}}"#, false),
    (r#"{"row":0,"col":0,"value":42,"future":true}"#, false),
    (r#"{"row":0,"col":0.0,"value":42}"#, false),
    (r#"{"row":2,"col":0,"value":42}"#, false),
];

#[test]
fn band_cell_decisions_match_the_full_check() {
    let mut report = Vec::new();
    for (cell, expected) in BAND_CELL_DECISIONS {
        let (manifest, data, key) = tiny_band_raw(cell);
        let (mut wb, _, _) = import_any(&manifest).unwrap();
        let accepted = bands::apply(&mut wb, &data, Some(&key)).is_ok();
        report.push(format!("    (r#\"{cell}\"#, {accepted}),"));
        if !accepted {
            assert!(wb.sheet(0).unwrap().cells_iter().next().is_none(), "{cell}: a refused band wrote cells");
        }
        if std::env::var_os("BAND_DECISIONS_PRINT").is_none() {
            assert_eq!(accepted, *expected, "{cell}");
        }
    }
    if std::env::var_os("BAND_DECISIONS_PRINT").is_some() {
        eprintln!("{}", report.join("\n"));
    }
}

/// A formula in the first band over a column that spans every band must
/// follow an edit made after the load, as on any other sheet.
#[test]
fn edits_after_a_banded_load_recalculate_formulas_over_other_bands() {
    let mut s = Sheet::new(SheetId(1), NUM_ROWS, NUM_COLS);
    s.set_name("Data");
    let rows = 210_000;
    for r in 0..rows {
        s.set_value_deferred(r, 1, &format!("{}", (r * 7 + 1) % 1000));
    }
    s.set_value_deferred(0, 20, &format!("=SUM(B1:B{rows})"));
    let mut wb = Workbook::from_sheets(vec![s], 0);
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let before = wb.sheets()[0].get_display(0, 20);

    let (manifest, out) = bands::export_banded(&wb, &[SheetLayout::default()], 0).unwrap();
    assert!(out.len() >= 3, "the column spans several bands");
    let (mut loaded, _, _) = import_any(&manifest).unwrap();
    for b in &out {
        bands::apply(&mut loaded, &b.data, Some(&b.reference.key)).unwrap();
    }
    bands::finish(&mut loaded).unwrap();
    assert_eq!(loaded.sheets()[0].get_display(0, 20), before, "SUM after the load");

    // B3 was (2*7+1)%1000 = 15; B40001 is in a later band.
    let b3: i64 = loaded.sheets()[0].get_display(2, 1).parse().unwrap();
    let late: i64 = loaded.sheets()[0].get_display(200_000, 1).parse().unwrap();
    loaded.set_cell_value_tracked(0, 2, 1, "5000");
    loaded.set_cell_value_tracked(0, 200_000, 1, "7000");
    let expected = before.parse::<i64>().unwrap() + (5000 - b3) + (7000 - late);
    assert_eq!(loaded.sheets()[0].get_display(0, 20), expected.to_string(), "SUM after edits in the first and a later band");
}

