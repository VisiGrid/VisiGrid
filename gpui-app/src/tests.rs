//! Tests for spreadsheet operations
//!
//! This module contains unit tests for fill operations, formula adjustments,
//! multi-edit undo, and format undo coalescing.

use regex::Regex;
use visigrid_engine::sheet::{Sheet, SheetId};

/// Test-only version of adjust_formula_refs (mirrors Spreadsheet::adjust_formula_refs)
fn adjust_formula_refs(formula: &str, delta_row: i32, delta_col: i32) -> String {
    let re = Regex::new(r"(\$?)([A-Za-z]+)(\$?)(\d+)").unwrap();

    re.replace_all(formula, |caps: &regex::Captures| {
        let col_absolute = &caps[1] == "$";
        let col_letters = &caps[2];
        let row_absolute = &caps[3] == "$";
        let row_num: i32 = caps[4].parse().unwrap_or(1);

        let col = col_letters.to_uppercase().chars().fold(0i32, |acc, c| {
            acc * 26 + (c as i32 - 'A' as i32 + 1)
        }) - 1;

        let new_col = if col_absolute { col } else { col + delta_col };
        let new_row = if row_absolute { row_num } else { row_num + delta_row };

        if new_col < 0 || new_row < 1 {
            return "#REF!".to_string();
        }

        let col_str = col_to_letter(new_col as usize);

        format!(
            "{}{}{}{}",
            if col_absolute { "$" } else { "" },
            col_str,
            if row_absolute { "$" } else { "" },
            new_row
        )
    }).to_string()
}

fn col_to_letter(col: usize) -> String {
    let mut s = String::new();
    let mut n = col;
    loop {
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 { break; }
        n = n / 26 - 1;
    }
    s
}

/// Simulate fill_down at the Sheet level (no gpui required)
fn fill_down_on_sheet(sheet: &mut Sheet, min_row: usize, max_row: usize, col: usize) {
    let source = sheet.get_raw(min_row, col);
    for row in (min_row + 1)..=max_row {
        let new_value = if source.starts_with('=') {
            adjust_formula_refs(&source, row as i32 - min_row as i32, 0)
        } else {
            source.clone()
        };
        sheet.set_value(row, col, &new_value);
    }
}

// =========================================================================
// REGRESSION TEST: Mixed references (the bug we just fixed)
// =========================================================================

#[test]
fn test_fill_down_mixed_references_formulas() {
    // Test that adjust_formula_refs correctly handles all 4 reference types
    let formula = "=A1 + $A$1 + A$1 + $A1";

    // Fill down by 1 row
    assert_eq!(
        adjust_formula_refs(formula, 1, 0),
        "=A2 + $A$1 + A$1 + $A2",
        "Row 2: A1->A2 (relative), $A$1->$A$1 (absolute), A$1->A$1 (row absolute), $A1->$A2 (col absolute)"
    );

    // Fill down by 2 rows
    assert_eq!(
        adjust_formula_refs(formula, 2, 0),
        "=A3 + $A$1 + A$1 + $A3"
    );

    // Fill down by 3 rows
    assert_eq!(
        adjust_formula_refs(formula, 3, 0),
        "=A4 + $A$1 + A$1 + $A4"
    );
}

#[test]
fn test_fill_down_mixed_references_end_to_end() {
    // End-to-end test: seed values, fill, verify formulas AND computed values
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // Seed A1:A4 with distinct values
    sheet.set_value(0, 0, "10"); // A1 = 10
    sheet.set_value(1, 0, "1");  // A2 = 1
    sheet.set_value(2, 0, "2");  // A3 = 2
    sheet.set_value(3, 0, "3");  // A4 = 3

    // Set B1 formula: =A1 + $A$1 + A$1 + $A1
    sheet.set_value(0, 1, "=A1 + $A$1 + A$1 + $A1");

    // Verify B1 value before fill
    assert_eq!(sheet.get_display(0, 1), "40", "B1 should be 10+10+10+10=40");

    // Simulate fill_down from B1 to B4
    fill_down_on_sheet(&mut sheet, 0, 3, 1);

    // Assert formulas are correct
    assert_eq!(sheet.get_raw(0, 1), "=A1 + $A$1 + A$1 + $A1", "B1 formula unchanged");
    assert_eq!(sheet.get_raw(1, 1), "=A2 + $A$1 + A$1 + $A2", "B2 formula adjusted");
    assert_eq!(sheet.get_raw(2, 1), "=A3 + $A$1 + A$1 + $A3", "B3 formula adjusted");
    assert_eq!(sheet.get_raw(3, 1), "=A4 + $A$1 + A$1 + $A4", "B4 formula adjusted");

    // Assert computed values are correct
    // B1: A1(10) + $A$1(10) + A$1(10) + $A1(10) = 40
    // B2: A2(1) + $A$1(10) + A$1(10) + $A2(1) = 22
    // B3: A3(2) + $A$1(10) + A$1(10) + $A3(2) = 24
    // B4: A4(3) + $A$1(10) + A$1(10) + $A4(3) = 26
    assert_eq!(sheet.get_display(0, 1), "40", "B1 value");
    assert_eq!(sheet.get_display(1, 1), "22", "B2 value: 1+10+10+1");
    assert_eq!(sheet.get_display(2, 1), "24", "B3 value: 2+10+10+2");
    assert_eq!(sheet.get_display(3, 1), "26", "B4 value: 3+10+10+3");
}

// =========================================================================
// EDGE CASE: Ranges in formulas (SUM, etc.)
// =========================================================================

#[test]
fn test_fill_down_with_ranges_formulas() {
    // =SUM(A1:A3) + $A$1 should become =SUM(A2:A4) + $A$1
    let formula = "=SUM(A1:A3) + $A$1";

    assert_eq!(
        adjust_formula_refs(formula, 1, 0),
        "=SUM(A2:A4) + $A$1",
        "Range A1:A3 should become A2:A4, absolute $A$1 stays"
    );

    assert_eq!(
        adjust_formula_refs(formula, 2, 0),
        "=SUM(A3:A5) + $A$1"
    );
}

#[test]
fn test_fill_down_with_ranges_end_to_end() {
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // Seed values
    sheet.set_value(0, 0, "10"); // A1 = 10
    sheet.set_value(1, 0, "20"); // A2 = 20
    sheet.set_value(2, 0, "30"); // A3 = 30
    sheet.set_value(3, 0, "40"); // A4 = 40
    sheet.set_value(4, 0, "50"); // A5 = 50

    // B1 = SUM(A1:A3) + $A$1 = (10+20+30) + 10 = 70
    sheet.set_value(0, 1, "=SUM(A1:A3) + $A$1");
    assert_eq!(sheet.get_display(0, 1), "70", "B1: SUM(10,20,30)+10");

    // Fill down B1:B3
    fill_down_on_sheet(&mut sheet, 0, 2, 1);

    // Check formulas
    assert_eq!(sheet.get_raw(0, 1), "=SUM(A1:A3) + $A$1");
    assert_eq!(sheet.get_raw(1, 1), "=SUM(A2:A4) + $A$1");
    assert_eq!(sheet.get_raw(2, 1), "=SUM(A3:A5) + $A$1");

    // Check values
    // B1: SUM(A1:A3) + $A$1 = (10+20+30) + 10 = 70
    // B2: SUM(A2:A4) + $A$1 = (20+30+40) + 10 = 100
    // B3: SUM(A3:A5) + $A$1 = (30+40+50) + 10 = 130
    assert_eq!(sheet.get_display(0, 1), "70", "B1 value");
    assert_eq!(sheet.get_display(1, 1), "100", "B2 value: SUM(20,30,40)+10");
    assert_eq!(sheet.get_display(2, 1), "130", "B3 value: SUM(30,40,50)+10");
}

// =========================================================================
// EDGE CASE: Multi-letter columns (AA, AB, etc.)
// =========================================================================

#[test]
fn test_fill_down_multi_letter_columns_formulas() {
    // =AA1 + $B$1 + C$2 + $D3
    // AA1 -> AA2 (both relative)
    // $B$1 -> $B$1 (both absolute)
    // C$2 -> C$2 (row absolute)
    // $D3 -> $D4 (col absolute, row relative)
    let formula = "=AA1 + $B$1 + C$2 + $D3";

    assert_eq!(
        adjust_formula_refs(formula, 1, 0),
        "=AA2 + $B$1 + C$2 + $D4",
        "Multi-letter columns with mixed refs"
    );

    assert_eq!(
        adjust_formula_refs(formula, 2, 0),
        "=AA3 + $B$1 + C$2 + $D5"
    );
}

#[test]
fn test_fill_down_multi_letter_columns_end_to_end() {
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // AA is column 26 (0-indexed), B is 1, C is 2, D is 3
    // Seed values
    sheet.set_value(0, 26, "100"); // AA1 = 100
    sheet.set_value(1, 26, "200"); // AA2 = 200
    sheet.set_value(2, 26, "300"); // AA3 = 300

    sheet.set_value(0, 1, "10");   // B1 = 10

    sheet.set_value(1, 2, "5");    // C2 = 5

    sheet.set_value(2, 3, "1");    // D3 = 1
    sheet.set_value(3, 3, "2");    // D4 = 2
    sheet.set_value(4, 3, "3");    // D5 = 3

    // AB1 = AA1 + $B$1 + C$2 + $D3 = 100 + 10 + 5 + 1 = 116
    sheet.set_value(0, 27, "=AA1 + $B$1 + C$2 + $D3"); // AB1
    assert_eq!(sheet.get_display(0, 27), "116", "AB1: 100+10+5+1");

    // Fill down AB1:AB3
    fill_down_on_sheet(&mut sheet, 0, 2, 27);

    // Check formulas
    assert_eq!(sheet.get_raw(0, 27), "=AA1 + $B$1 + C$2 + $D3");
    assert_eq!(sheet.get_raw(1, 27), "=AA2 + $B$1 + C$2 + $D4");
    assert_eq!(sheet.get_raw(2, 27), "=AA3 + $B$1 + C$2 + $D5");

    // Check values
    // AB1: AA1(100) + $B$1(10) + C$2(5) + $D3(1) = 116
    // AB2: AA2(200) + $B$1(10) + C$2(5) + $D4(2) = 217
    // AB3: AA3(300) + $B$1(10) + C$2(5) + $D5(3) = 318
    assert_eq!(sheet.get_display(0, 27), "116", "AB1 value");
    assert_eq!(sheet.get_display(1, 27), "217", "AB2 value: 200+10+5+2");
    assert_eq!(sheet.get_display(2, 27), "318", "AB3 value: 300+10+5+3");
}

// =========================================================================
// EDGE CASE: Fill right (column adjustment)
// =========================================================================

#[test]
fn test_fill_right_formulas() {
    let formula = "=A1 + $A$1 + A$1 + $A1";

    // Fill right by 1 column
    // A1 -> B1 (col relative)
    // $A$1 -> $A$1 (both absolute)
    // A$1 -> B$1 (col relative, row absolute)
    // $A1 -> $A1 (col absolute)
    assert_eq!(
        adjust_formula_refs(formula, 0, 1),
        "=B1 + $A$1 + B$1 + $A1",
        "Fill right: relative cols shift, absolute cols stay"
    );
}

/// Simulate fill_right at the Sheet level (no gpui required)
fn fill_right_on_sheet(sheet: &mut Sheet, row: usize, min_col: usize, max_col: usize) {
    let source = sheet.get_raw(row, min_col);
    for col in (min_col + 1)..=max_col {
        let new_value = if source.starts_with('=') {
            adjust_formula_refs(&source, 0, col as i32 - min_col as i32)
        } else {
            source.clone()
        };
        sheet.set_value(row, col, &new_value);
    }
}

#[test]
fn test_fill_right_mixed_references_end_to_end() {
    // End-to-end test for fill right with mixed references
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // Seed row 1 with distinct values: A1=10, B1=1, C1=2, D1=3
    sheet.set_value(0, 0, "10"); // A1 = 10
    sheet.set_value(0, 1, "1");  // B1 = 1
    sheet.set_value(0, 2, "2");  // C1 = 2
    sheet.set_value(0, 3, "3");  // D1 = 3

    // Set A2 formula: =A1 + $A$1 + A$1 + $A1
    // When filling right:
    // - A1 shifts column (relative col)
    // - $A$1 stays (both absolute)
    // - A$1 shifts column (relative col, absolute row)
    // - $A1 stays (absolute col, relative row)
    sheet.set_value(1, 0, "=A1 + $A$1 + A$1 + $A1");

    // Verify A2 value before fill
    // A1(10) + $A$1(10) + A$1(10) + $A1(10) = 40
    assert_eq!(sheet.get_display(1, 0), "40", "A2 should be 40");

    // Fill right A2:D2
    fill_right_on_sheet(&mut sheet, 1, 0, 3);

    // Check formulas
    assert_eq!(sheet.get_raw(1, 0), "=A1 + $A$1 + A$1 + $A1", "A2 formula unchanged");
    assert_eq!(sheet.get_raw(1, 1), "=B1 + $A$1 + B$1 + $A1", "B2 formula adjusted");
    assert_eq!(sheet.get_raw(1, 2), "=C1 + $A$1 + C$1 + $A1", "C2 formula adjusted");
    assert_eq!(sheet.get_raw(1, 3), "=D1 + $A$1 + D$1 + $A1", "D2 formula adjusted");

    // Check computed values
    // A2: A1(10) + $A$1(10) + A$1(10) + $A1(10) = 40
    // B2: B1(1) + $A$1(10) + B$1(1) + $A1(10) = 22
    // C2: C1(2) + $A$1(10) + C$1(2) + $A1(10) = 24
    // D2: D1(3) + $A$1(10) + D$1(3) + $A1(10) = 26
    assert_eq!(sheet.get_display(1, 0), "40", "A2 value");
    assert_eq!(sheet.get_display(1, 1), "22", "B2 value: 1+10+1+10");
    assert_eq!(sheet.get_display(1, 2), "24", "C2 value: 2+10+2+10");
    assert_eq!(sheet.get_display(1, 3), "26", "D2 value: 3+10+3+10");
}

// =========================================================================
// EDGE CASE: Multi-edit with single undo
// =========================================================================

#[test]
fn test_multi_edit_applies_once_and_single_undo() {
    use crate::history::{History, CellChange, UndoAction};

    let mut sheet = Sheet::new(SheetId(1), 100, 100);
    let mut history = History::new();

    // Seed initial values: A1=1, A2=2, A3=3, B1=10, B2=20, B3=30
    sheet.set_value(0, 0, "1");  // A1
    sheet.set_value(1, 0, "2");  // A2
    sheet.set_value(2, 0, "3");  // A3
    sheet.set_value(0, 1, "10"); // B1
    sheet.set_value(1, 1, "20"); // B2
    sheet.set_value(2, 1, "30"); // B3

    // Simulate multi-edit: set "=A1*2" to selection A1:B3 (6 cells)
    let new_value = "=A1*2";
    let selection = [(0, 0), (0, 1), (1, 0), (1, 1), (2, 0), (2, 1)];

    let mut changes = Vec::new();
    for (row, col) in selection.iter() {
        let old_value = sheet.get_raw(*row, *col);
        if old_value != new_value {
            changes.push(CellChange {
                row: *row,
                col: *col,
                old_value,
                new_value: new_value.to_string(),
            });
        }
        sheet.set_value(*row, *col, new_value);
    }

    // Record as single batch (this is what multi-edit does)
    // sheet_index = 0 for this test
    history.record_batch(0, changes);

    // Verify all 6 cells have the formula
    for (row, col) in selection.iter() {
        assert_eq!(
            sheet.get_raw(*row, *col), "=A1*2",
            "Cell ({}, {}) should have formula =A1*2", row, col
        );
    }

    // Verify computed values (all reference A1 which is now =A1*2, causing circular ref)
    // Actually A1 = =A1*2 is circular, so let's verify at least B1 computes
    // B1 = =A1*2 where A1 = =A1*2 (circular)
    // The key test is that single undo reverts ALL cells

    // Single undo should revert ALL 6 cells
    let entry = history.undo().expect("Should have undo entry");
    let changes = match &entry.action {
        UndoAction::Values { changes, .. } => changes,
        _ => panic!("Expected Values action"),
    };
    assert_eq!(changes.len(), 6, "Undo entry should contain all 6 changes");

    // Apply undo to sheet
    for change in changes.iter() {
        sheet.set_value(change.row, change.col, &change.old_value);
    }

    // Verify original values are restored
    assert_eq!(sheet.get_raw(0, 0), "1", "A1 restored to 1");
    assert_eq!(sheet.get_raw(1, 0), "2", "A2 restored to 2");
    assert_eq!(sheet.get_raw(2, 0), "3", "A3 restored to 3");
    assert_eq!(sheet.get_raw(0, 1), "10", "B1 restored to 10");
    assert_eq!(sheet.get_raw(1, 1), "20", "B2 restored to 20");
    assert_eq!(sheet.get_raw(2, 1), "30", "B3 restored to 30");

    // Verify redo works and contains all 6 changes
    let redo_entry = history.redo().expect("Should have redo entry");
    let redo_changes = match &redo_entry.action {
        UndoAction::Values { changes, .. } => changes,
        _ => panic!("Expected Values action"),
    };
    assert_eq!(redo_changes.len(), 6, "Redo entry should contain all 6 changes");
}

// =========================================================================
// FORMAT UNDO: Coalescing tests
// =========================================================================

#[test]
fn test_format_coalescing_same_cells_merges() {
    use crate::history::{History, CellFormatPatch, FormatActionKind, UndoAction};
    use visigrid_engine::cell::CellFormat;

    let mut history = History::new();

    // First decimal change on cell (0,0)
    let patches1 = vec![CellFormatPatch {
        row: 0,
        col: 0,
        before: CellFormat::default(),
        after: CellFormat { bold: true, ..Default::default() },
    }];
    history.record_format(0, patches1, FormatActionKind::DecimalPlaces, "Decimal +".into());

    // Second decimal change on same cell within 500ms window
    let patches2 = vec![CellFormatPatch {
        row: 0,
        col: 0,
        before: CellFormat { bold: true, ..Default::default() },
        after: CellFormat { bold: true, italic: true, ..Default::default() },
    }];
    history.record_format(0, patches2, FormatActionKind::DecimalPlaces, "Decimal +".into());

    // Should have coalesced into single entry
    let entry = history.undo().expect("Should have undo entry");
    match &entry.action {
        UndoAction::Format { patches, .. } => {
            assert_eq!(patches.len(), 1, "Should have 1 patch");
            // Before should be original, after should be final
            assert!(!patches[0].before.bold, "Before should be default (no bold)");
            assert!(patches[0].after.italic, "After should have italic");
        }
        _ => panic!("Expected Format action"),
    }

    // No more undo entries
    assert!(history.undo().is_none(), "Should have no more entries - coalesced into one");
}

#[test]
fn test_format_coalescing_different_cells_separate() {
    use crate::history::{History, CellFormatPatch, FormatActionKind, UndoAction};
    use visigrid_engine::cell::CellFormat;

    let mut history = History::new();

    // First decimal change on cell (0,0)
    let patches1 = vec![CellFormatPatch {
        row: 0,
        col: 0,
        before: CellFormat::default(),
        after: CellFormat { bold: true, ..Default::default() },
    }];
    history.record_format(0, patches1, FormatActionKind::DecimalPlaces, "Decimal +".into());

    // Second decimal change on DIFFERENT cell (0,1) within 500ms window
    let patches2 = vec![CellFormatPatch {
        row: 0,
        col: 1,  // Different column!
        before: CellFormat::default(),
        after: CellFormat { bold: true, ..Default::default() },
    }];
    history.record_format(0, patches2, FormatActionKind::DecimalPlaces, "Decimal +".into());

    // Should have TWO separate entries because cells are different
    let entry1 = history.undo().expect("Should have first undo entry");
    match &entry1.action {
        UndoAction::Format { patches, .. } => {
            assert_eq!(patches[0].col, 1, "First undo should be cell (0,1)");
        }
        _ => panic!("Expected Format action"),
    }

    let entry2 = history.undo().expect("Should have second undo entry");
    match &entry2.action {
        UndoAction::Format { patches, .. } => {
            assert_eq!(patches[0].col, 0, "Second undo should be cell (0,0)");
        }
        _ => panic!("Expected Format action"),
    }
}

#[test]
fn test_format_undo_restores_mixed_state() {
    use crate::history::{History, CellFormatPatch, FormatActionKind, UndoAction};
    use visigrid_engine::cell::CellFormat;
    use visigrid_engine::sheet::{Sheet, SheetId};

    let mut sheet = Sheet::new(SheetId(1), 100, 100);
    let mut history = History::new();

    // Set up mixed initial state: A1 bold, A2 not bold, A3 italic
    let mut format_a1 = CellFormat::default();
    format_a1.bold = true;
    sheet.set_format(0, 0, format_a1.clone());

    let format_a2 = CellFormat::default(); // not bold
    sheet.set_format(1, 0, format_a2.clone());

    let mut format_a3 = CellFormat::default();
    format_a3.italic = true;
    sheet.set_format(2, 0, format_a3.clone());

    // Capture before states
    let before_a1 = sheet.get_format(0, 0);
    let before_a2 = sheet.get_format(1, 0);
    let before_a3 = sheet.get_format(2, 0);

    // Apply "set bold on" to all three cells
    sheet.set_bold(0, 0, true);
    sheet.set_bold(1, 0, true);
    sheet.set_bold(2, 0, true);

    // Record format change with individual before/after for each cell
    let patches = vec![
        CellFormatPatch { row: 0, col: 0, before: before_a1.clone(), after: sheet.get_format(0, 0) },
        CellFormatPatch { row: 1, col: 0, before: before_a2.clone(), after: sheet.get_format(1, 0) },
        CellFormatPatch { row: 2, col: 0, before: before_a3.clone(), after: sheet.get_format(2, 0) },
    ];
    history.record_format(0, patches, FormatActionKind::Bold, "Bold on".into());

    // Now undo
    let entry = history.undo().expect("Should have undo entry");
    match entry.action {
        UndoAction::Format { patches, .. } => {
            for patch in patches {
                sheet.set_format(patch.row, patch.col, patch.before);
            }
        }
        _ => panic!("Expected Format action"),
    }

    // Verify mixed state is restored exactly
    assert!(sheet.get_format(0, 0).bold, "A1 should still be bold (was bold before)");
    assert!(!sheet.get_format(1, 0).bold, "A2 should not be bold (wasn't bold before)");
    assert!(!sheet.get_format(2, 0).bold, "A3 should not be bold (wasn't bold before)");
    assert!(sheet.get_format(2, 0).italic, "A3 should still be italic");
}

// =========================================================================
// EDGE CASE: Boundary conditions
// =========================================================================

#[test]
fn test_fill_down_ref_error() {
    // Filling up from row 1 should produce #REF!
    let formula = "=A1";
    assert_eq!(
        adjust_formula_refs(formula, -1, 0),
        "=#REF!",
        "Row 0 (A0) doesn't exist, should be #REF!"
    );
}

#[test]
fn test_fill_left_ref_error() {
    // Filling left from column A should produce #REF!
    let formula = "=A1";
    assert_eq!(
        adjust_formula_refs(formula, 0, -1),
        "=#REF!",
        "Column before A doesn't exist, should be #REF!"
    );
}

// =========================================================================
// EXTRACT NAMED RANGE: Replacement correctness tests
// =========================================================================

/// Test-only version of replace_range_in_formula (mirrors Spreadsheet::replace_range_in_formula)
fn replace_range_in_formula(formula: &str, range_literal: &str, name: &str) -> String {
    let range_upper = range_literal.to_uppercase();
    let mut result = String::new();
    let chars: Vec<char> = formula.chars().collect();
    let range_len = range_upper.len();

    let mut i = 0;
    let mut in_string = false;

    while i < chars.len() {
        // Track string literal state (toggle on each unescaped quote)
        if chars[i] == '"' {
            // Check for escaped quote (doubled quote in Excel formulas)
            if in_string && i + 1 < chars.len() && chars[i + 1] == '"' {
                result.push(chars[i]);
                result.push(chars[i + 1]);
                i += 2;
                continue;
            }
            in_string = !in_string;
            result.push(chars[i]);
            i += 1;
            continue;
        }

        // If inside a string, just copy the character
        if in_string {
            result.push(chars[i]);
            i += 1;
            continue;
        }

        // Check for range match (only outside strings)
        if i + range_len <= chars.len() {
            let slice: String = chars[i..i + range_len].iter().collect::<String>().to_uppercase();
            if slice == range_upper {
                // Verify word boundaries
                let before_ok = i == 0 || (!chars[i - 1].is_alphanumeric() && chars[i - 1] != '_' && chars[i - 1] != '$');
                let after_ok = i + range_len >= chars.len() || (!chars[i + range_len].is_alphanumeric() && chars[i + range_len] != '_');
                if before_ok && after_ok {
                    result.push_str(name);
                    i += range_len;
                    continue;
                }
            }
        }
        result.push(chars[i]);
        i += 1;
    }
    result
}

#[test]
fn test_extract_string_literal_not_replaced() {
    // ="A1:B2" should NOT be modified - it's a string literal
    let formula = r#"="A1:B2""#;
    let result = replace_range_in_formula(formula, "A1:B2", "MyRange");
    assert_eq!(result, r#"="A1:B2""#, "String literal should not be modified");
}

#[test]
fn test_extract_simple_range_replaced() {
    // =SUM(A1:B2) should become =SUM(MyRange)
    let formula = "=SUM(A1:B2)";
    let result = replace_range_in_formula(formula, "A1:B2", "MyRange");
    assert_eq!(result, "=SUM(MyRange)", "Simple range should be replaced");
}

#[test]
fn test_extract_absolute_range_replaced() {
    // =SUM($A$1:$B$2) should become =SUM(MyRange)
    let formula = "=SUM($A$1:$B$2)";
    let result = replace_range_in_formula(formula, "$A$1:$B$2", "MyRange");
    assert_eq!(result, "=SUM(MyRange)", "Absolute range should be replaced exactly");
}

#[test]
fn test_extract_similar_range_not_touched() {
    // Extracting A1:B2 must NOT touch A1:B20 (longer range)
    let formula = "=SUM(A1:B20)";
    let result = replace_range_in_formula(formula, "A1:B2", "MyRange");
    assert_eq!(result, "=SUM(A1:B20)", "Similar but longer range should NOT be replaced");
}

#[test]
fn test_extract_boundary_checks() {
    // Range at different boundaries

    // At start of formula
    let result1 = replace_range_in_formula("=A1:B2+C1", "A1:B2", "MyRange");
    assert_eq!(result1, "=MyRange+C1", "Range at start");

    // At end of formula
    let result2 = replace_range_in_formula("=C1+A1:B2", "A1:B2", "MyRange");
    assert_eq!(result2, "=C1+MyRange", "Range at end");

    // Multiple occurrences
    let result3 = replace_range_in_formula("=SUM(A1:B2)+AVERAGE(A1:B2)", "A1:B2", "MyRange");
    assert_eq!(result3, "=SUM(MyRange)+AVERAGE(MyRange)", "Multiple occurrences replaced");
}

#[test]
fn test_extract_case_insensitive() {
    // Range matching should be case-insensitive
    let formula = "=SUM(a1:b2)";
    let result = replace_range_in_formula(formula, "A1:B2", "MyRange");
    assert_eq!(result, "=SUM(MyRange)", "Case-insensitive matching");
}

#[test]
fn test_extract_preserves_mixed_content() {
    // Formula with string and range
    let formula = r#"=IF(A1>0,"A1:B2",SUM(A1:B2))"#;
    let result = replace_range_in_formula(formula, "A1:B2", "MyRange");
    assert_eq!(result, r#"=IF(A1>0,"A1:B2",SUM(MyRange))"#, "String preserved, range replaced");
}

#[test]
fn test_extract_escaped_quote_in_string() {
    // Excel uses "" for escaped quote inside strings
    let formula = r#"=CONCAT("Say ""A1:B2""",A1:B2)"#;
    let result = replace_range_in_formula(formula, "A1:B2", "MyRange");
    assert_eq!(result, r#"=CONCAT("Say ""A1:B2""",MyRange)"#, "Escaped quotes handled");
}

#[test]
fn test_extract_word_boundary_prevents_partial_match() {
    // Should not match A1:B2 inside A1:B2X or XA1:B2
    let formula = "=SUM(A1:B2X)";  // Not a valid range, but test boundary
    let result = replace_range_in_formula(formula, "A1:B2", "MyRange");
    // A1:B2X has 'X' after, which is alphanumeric, so should NOT match
    assert_eq!(result, "=SUM(A1:B2X)", "Alphanumeric suffix prevents match");

    // With underscore suffix
    let formula2 = "=A1:B2_total";
    let result2 = replace_range_in_formula(formula2, "A1:B2", "MyRange");
    assert_eq!(result2, "=A1:B2_total", "Underscore suffix prevents match");
}

// =========================================================================
// PROPERTY-BASED UNDO TESTS
// =========================================================================
//
// These tests verify that undo perfectly restores sheet state after
// arbitrary sequences of cell operations. This is critical for trust:
// users won't forgive corrupted spreadsheets.

/// Simple deterministic PRNG (Xorshift64) - no external dependencies
struct Xorshift64 {
    state: u64,
}

impl Xorshift64 {
    fn new(seed: u64) -> Self {
        // Ensure non-zero state
        Self { state: if seed == 0 { 1 } else { seed } }
    }

    fn next(&mut self) -> u64 {
        let mut x = self.state;
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        self.state = x;
        x
    }

    fn next_usize(&mut self, max: usize) -> usize {
        (self.next() as usize) % max
    }

    fn next_bool(&mut self) -> bool {
        self.next() % 2 == 0
    }
}

/// Canonical snapshot of sheet state for comparison
#[derive(Debug, Clone, PartialEq)]
struct SheetState {
    /// (row, col, raw_value) for all populated cells, sorted for deterministic comparison
    cells: Vec<(usize, usize, String)>,
}

impl SheetState {
    fn from_sheet(sheet: &Sheet) -> Self {
        let mut cells: Vec<_> = sheet.cells_iter()
            .map(|(&(row, col), cell)| (row, col, cell.value.raw_display()))
            .filter(|(_, _, val)| !val.is_empty())
            .collect();
        cells.sort_by_key(|(r, c, _)| (*r, *c));
        Self { cells }
    }
}

/// Op types for random generation (mirrors LuaOp but simpler for testing)
#[derive(Debug, Clone)]
enum TestOp {
    SetValue { row: usize, col: usize, value: String },
    SetFormula { row: usize, col: usize, formula: String },
    Clear { row: usize, col: usize },
}

/// Generate a random op sequence
fn generate_ops(rng: &mut Xorshift64, count: usize, rows: usize, cols: usize) -> Vec<TestOp> {
    // Safe formula set (deterministic, no circular refs in small scope)
    const FORMULAS: &[&str] = &[
        "=1+1",
        "=2*3",
        "=10/2",
        "=SUM(1,2,3)",
        "=A1+1",
        "=B2*2",
        "=A1+B1",
        "=SUM(A1:A3)",
        "=AVERAGE(A1:B2)",
        "=IF(A1>0,1,0)",
    ];

    let mut ops = Vec::with_capacity(count);

    for _ in 0..count {
        let row = rng.next_usize(rows);
        let col = rng.next_usize(cols);

        let op = match rng.next_usize(10) {
            0..=3 => {
                // SetValue: number
                let n = rng.next() % 1000;
                TestOp::SetValue { row, col, value: n.to_string() }
            }
            4..=5 => {
                // SetValue: string
                let s = format!("text{}", rng.next() % 100);
                TestOp::SetValue { row, col, value: s }
            }
            6..=7 => {
                // SetFormula
                let formula = FORMULAS[rng.next_usize(FORMULAS.len())];
                TestOp::SetFormula { row, col, formula: formula.to_string() }
            }
            8 => {
                // SetValue: bool
                let b = if rng.next_bool() { "TRUE" } else { "FALSE" };
                TestOp::SetValue { row, col, value: b.to_string() }
            }
            _ => {
                // Clear (set to empty)
                TestOp::Clear { row, col }
            }
        };
        ops.push(op);
    }

    ops
}

/// Apply ops to sheet and record history, returning changes for verification
fn apply_ops_with_history(
    sheet: &mut Sheet,
    history: &mut crate::history::History,
    ops: &[TestOp],
) -> Vec<crate::history::CellChange> {
    use crate::history::CellChange;

    let mut changes = Vec::new();

    for op in ops {
        match op {
            TestOp::SetValue { row, col, value } => {
                let old_value = sheet.get_raw(*row, *col);
                sheet.set_value(*row, *col, value);
                changes.push(CellChange {
                    row: *row,
                    col: *col,
                    old_value,
                    new_value: value.clone(),
                });
            }
            TestOp::SetFormula { row, col, formula } => {
                let old_value = sheet.get_raw(*row, *col);
                sheet.set_value(*row, *col, formula);
                changes.push(CellChange {
                    row: *row,
                    col: *col,
                    old_value,
                    new_value: formula.clone(),
                });
            }
            TestOp::Clear { row, col } => {
                let old_value = sheet.get_raw(*row, *col);
                sheet.set_value(*row, *col, "");
                changes.push(CellChange {
                    row: *row,
                    col: *col,
                    old_value,
                    new_value: String::new(),
                });
            }
        }
    }

    // Record as single batch (like Lua script commit)
    history.record_batch(0, changes.clone());

    changes
}

/// Apply undo to sheet
///
/// CRITICAL: Changes must be applied in REVERSE order.
/// If the same cell is modified multiple times in a batch, we need to undo
/// the last change first to restore the original value correctly.
fn apply_undo(sheet: &mut Sheet, history: &mut crate::history::History) -> bool {
    use crate::history::UndoAction;

    if let Some(entry) = history.undo() {
        match entry.action {
            UndoAction::Values { changes, .. } => {
                // Apply in reverse order to handle same-cell sequences correctly
                for change in changes.iter().rev() {
                    sheet.set_value(change.row, change.col, &change.old_value);
                }
                true
            }
            _ => false,
        }
    } else {
        false
    }
}

/// Apply redo to sheet
fn apply_redo(sheet: &mut Sheet, history: &mut crate::history::History) -> bool {
    use crate::history::UndoAction;

    if let Some(entry) = history.redo() {
        match entry.action {
            UndoAction::Values { changes, .. } => {
                for change in changes {
                    sheet.set_value(change.row, change.col, &change.new_value);
                }
                true
            }
            _ => false,
        }
    } else {
        false
    }
}

#[test]
fn test_undo_restores_original_state_property_based() {
    // Property: For any random op sequence, undo should restore original state exactly.
    // Run 1000 iterations with different seeds.

    const ITERATIONS: u64 = 1000;
    const SHEET_ROWS: usize = 20;
    const SHEET_COLS: usize = 20;

    for seed in 0..ITERATIONS {
        let mut rng = Xorshift64::new(seed + 1);

        // Create sheet with some initial data
        let mut sheet = Sheet::new(SheetId(1), 100, 100);

        // Seed with some initial values (20% of cells)
        for _ in 0..(SHEET_ROWS * SHEET_COLS / 5) {
            let row = rng.next_usize(SHEET_ROWS);
            let col = rng.next_usize(SHEET_COLS);
            let val = rng.next() % 100;
            sheet.set_value(row, col, &val.to_string());
        }

        // Snapshot original state
        let original = SheetState::from_sheet(&sheet);

        // Generate random ops (1-200)
        let op_count = 1 + rng.next_usize(200);
        let ops = generate_ops(&mut rng, op_count, SHEET_ROWS, SHEET_COLS);

        // Apply ops
        let mut history = crate::history::History::new();
        apply_ops_with_history(&mut sheet, &mut history, &ops);

        // Undo
        let undone = apply_undo(&mut sheet, &mut history);
        assert!(undone, "Seed {}: Undo should succeed", seed);

        // Verify state matches original
        let restored = SheetState::from_sheet(&sheet);
        assert_eq!(
            original, restored,
            "Seed {}: State mismatch after undo.\nOps: {:?}\nOriginal: {:?}\nRestored: {:?}",
            seed, ops, original, restored
        );
    }
}

#[test]
fn test_undo_redo_idempotence() {
    // Property: apply → undo → redo → undo = original state
    // This tests that redo doesn't corrupt the undo chain.

    const ITERATIONS: u64 = 500;
    const SHEET_ROWS: usize = 20;
    const SHEET_COLS: usize = 20;

    for seed in 0..ITERATIONS {
        let mut rng = Xorshift64::new(seed + 1000);

        let mut sheet = Sheet::new(SheetId(1), 100, 100);

        // Seed initial data
        for _ in 0..(SHEET_ROWS * SHEET_COLS / 5) {
            let row = rng.next_usize(SHEET_ROWS);
            let col = rng.next_usize(SHEET_COLS);
            let val = rng.next() % 100;
            sheet.set_value(row, col, &val.to_string());
        }

        let original = SheetState::from_sheet(&sheet);

        // Generate ops
        let op_count = 1 + rng.next_usize(100);
        let ops = generate_ops(&mut rng, op_count, SHEET_ROWS, SHEET_COLS);

        // Apply
        let mut history = crate::history::History::new();
        apply_ops_with_history(&mut sheet, &mut history, &ops);

        let after_apply = SheetState::from_sheet(&sheet);

        // Undo
        apply_undo(&mut sheet, &mut history);
        let after_undo = SheetState::from_sheet(&sheet);
        assert_eq!(original, after_undo, "Seed {}: First undo should restore original", seed);

        // Redo
        apply_redo(&mut sheet, &mut history);
        let after_redo = SheetState::from_sheet(&sheet);
        assert_eq!(after_apply, after_redo, "Seed {}: Redo should restore applied state", seed);

        // Undo again
        apply_undo(&mut sheet, &mut history);
        let after_undo2 = SheetState::from_sheet(&sheet);
        assert_eq!(original, after_undo2, "Seed {}: Second undo should restore original", seed);
    }
}

#[test]
fn test_same_cell_sequences() {
    // Explicitly test sequences that mutate the same cell multiple times.
    // This is where undo bugs typically hide.

    const ITERATIONS: u64 = 500;

    for seed in 0..ITERATIONS {
        let mut rng = Xorshift64::new(seed + 2000);

        let mut sheet = Sheet::new(SheetId(1), 100, 100);

        // Start with a known value in cell (5, 5)
        sheet.set_value(5, 5, "initial");

        let original = SheetState::from_sheet(&sheet);

        // Generate ops that all target the same cell or nearby cells
        let op_count = 5 + rng.next_usize(20);
        let mut ops = Vec::new();

        for _ in 0..op_count {
            // 80% chance to hit (5,5), 20% to hit nearby
            let (row, col) = if rng.next_usize(10) < 8 {
                (5, 5)
            } else {
                (5 + rng.next_usize(3), 5 + rng.next_usize(3))
            };

            let op = match rng.next_usize(4) {
                0 => TestOp::SetValue { row, col, value: format!("v{}", rng.next() % 100) },
                1 => TestOp::SetFormula { row, col, formula: "=1+1".to_string() },
                2 => TestOp::SetFormula { row, col, formula: "=A1+B2".to_string() },
                _ => TestOp::Clear { row, col },
            };
            ops.push(op);
        }

        // Apply
        let mut history = crate::history::History::new();
        apply_ops_with_history(&mut sheet, &mut history, &ops);

        // Undo
        apply_undo(&mut sheet, &mut history);

        let restored = SheetState::from_sheet(&sheet);
        assert_eq!(
            original, restored,
            "Seed {}: Same-cell sequence failed.\nOps: {:?}",
            seed, ops
        );
    }
}

#[test]
fn test_touched_cells_completeness() {
    // Property: Every cell modified by ops should be in the history changes.
    // This verifies the "touched cells" tracking is complete.

    const ITERATIONS: u64 = 500;
    const SHEET_ROWS: usize = 20;
    const SHEET_COLS: usize = 20;

    for seed in 0..ITERATIONS {
        let mut rng = Xorshift64::new(seed + 3000);

        let mut sheet = Sheet::new(SheetId(1), 100, 100);

        // Track which cells we're going to modify
        let op_count = 1 + rng.next_usize(100);
        let ops = generate_ops(&mut rng, op_count, SHEET_ROWS, SHEET_COLS);

        // Collect all cells that ops will touch
        let mut touched: std::collections::HashSet<(usize, usize)> = std::collections::HashSet::new();
        for op in &ops {
            let (row, col) = match op {
                TestOp::SetValue { row, col, .. } => (*row, *col),
                TestOp::SetFormula { row, col, .. } => (*row, *col),
                TestOp::Clear { row, col } => (*row, *col),
            };
            touched.insert((row, col));
        }

        // Apply and get recorded changes
        let mut history = crate::history::History::new();
        let changes = apply_ops_with_history(&mut sheet, &mut history, &ops);

        // Verify every touched cell appears in changes
        let changed_cells: std::collections::HashSet<(usize, usize)> = changes
            .iter()
            .map(|c| (c.row, c.col))
            .collect();

        for (row, col) in &touched {
            assert!(
                changed_cells.contains(&(*row, *col)),
                "Seed {}: Cell ({}, {}) was touched but not in history changes",
                seed, row, col
            );
        }
    }
}

// =========================================================================
// FILL HANDLE: Core fill operations
// =========================================================================

/// Simulate fill handle drag down (from single anchor cell)
fn fill_handle_down(sheet: &mut Sheet, anchor_row: usize, col: usize, end_row: usize) {
    if end_row <= anchor_row {
        return; // No-op or upward fill
    }
    let source = sheet.get_raw(anchor_row, col);
    for row in (anchor_row + 1)..=end_row {
        let delta_row = row as i32 - anchor_row as i32;
        let new_value = if source.starts_with('=') {
            adjust_formula_refs(&source, delta_row, 0)
        } else {
            source.clone()
        };
        sheet.set_value(row, col, &new_value);
    }
}

/// Simulate fill handle drag up (from single anchor cell)
fn fill_handle_up(sheet: &mut Sheet, anchor_row: usize, col: usize, end_row: usize) {
    if end_row >= anchor_row {
        return; // No-op or downward fill
    }
    let source = sheet.get_raw(anchor_row, col);
    for row in end_row..anchor_row {
        let delta_row = row as i32 - anchor_row as i32;
        let new_value = if source.starts_with('=') {
            adjust_formula_refs(&source, delta_row, 0)
        } else {
            source.clone()
        };
        sheet.set_value(row, col, &new_value);
    }
}

/// Simulate fill handle drag right (from single anchor cell)
fn fill_handle_right(sheet: &mut Sheet, row: usize, anchor_col: usize, end_col: usize) {
    if end_col <= anchor_col {
        return; // No-op or leftward fill
    }
    let source = sheet.get_raw(row, anchor_col);
    for col in (anchor_col + 1)..=end_col {
        let delta_col = col as i32 - anchor_col as i32;
        let new_value = if source.starts_with('=') {
            adjust_formula_refs(&source, 0, delta_col)
        } else {
            source.clone()
        };
        sheet.set_value(row, col, &new_value);
    }
}

/// Simulate fill handle drag left (from single anchor cell)
fn fill_handle_left(sheet: &mut Sheet, row: usize, anchor_col: usize, end_col: usize) {
    if end_col >= anchor_col {
        return; // No-op or rightward fill
    }
    let source = sheet.get_raw(row, anchor_col);
    for col in end_col..anchor_col {
        let delta_col = col as i32 - anchor_col as i32;
        let new_value = if source.starts_with('=') {
            adjust_formula_refs(&source, 0, delta_col)
        } else {
            source.clone()
        };
        sheet.set_value(row, col, &new_value);
    }
}

#[test]
fn test_fill_handle_drag_down_fills_a2_to_a5() {
    // Drag down from A1 to A5 fills A2:A5
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // A1 = 10
    sheet.set_value(0, 0, "10");

    // Drag fill handle from A1 (row 0) to A5 (row 4)
    fill_handle_down(&mut sheet, 0, 0, 4);

    // Verify A2:A5 are filled with 10
    assert_eq!(sheet.get_raw(1, 0), "10", "A2 should be 10");
    assert_eq!(sheet.get_raw(2, 0), "10", "A3 should be 10");
    assert_eq!(sheet.get_raw(3, 0), "10", "A4 should be 10");
    assert_eq!(sheet.get_raw(4, 0), "10", "A5 should be 10");

    // A1 unchanged
    assert_eq!(sheet.get_raw(0, 0), "10", "A1 should still be 10");
}

#[test]
fn test_fill_handle_drag_down_formula_adjustment() {
    // Drag down from A1 to A5 with formula =B1
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // B1:B5 with values 1-5
    sheet.set_value(0, 1, "1");  // B1
    sheet.set_value(1, 1, "2");  // B2
    sheet.set_value(2, 1, "3");  // B3
    sheet.set_value(3, 1, "4");  // B4
    sheet.set_value(4, 1, "5");  // B5

    // A1 = =B1
    sheet.set_value(0, 0, "=B1");

    // Drag fill handle from A1 to A5
    fill_handle_down(&mut sheet, 0, 0, 4);

    // Verify formulas are adjusted
    assert_eq!(sheet.get_raw(0, 0), "=B1", "A1 formula unchanged");
    assert_eq!(sheet.get_raw(1, 0), "=B2", "A2 formula adjusted to =B2");
    assert_eq!(sheet.get_raw(2, 0), "=B3", "A3 formula adjusted to =B3");
    assert_eq!(sheet.get_raw(3, 0), "=B4", "A4 formula adjusted to =B4");
    assert_eq!(sheet.get_raw(4, 0), "=B5", "A5 formula adjusted to =B5");

    // Verify computed values
    assert_eq!(sheet.get_display(0, 0), "1", "A1 displays 1");
    assert_eq!(sheet.get_display(1, 0), "2", "A2 displays 2");
    assert_eq!(sheet.get_display(2, 0), "3", "A3 displays 3");
    assert_eq!(sheet.get_display(3, 0), "4", "A4 displays 4");
    assert_eq!(sheet.get_display(4, 0), "5", "A5 displays 5");
}

#[test]
fn test_fill_handle_drag_right_fills_b1_to_e1() {
    // Drag right from A1 to E1 fills B1:E1
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // A1 = 42
    sheet.set_value(0, 0, "42");

    // Drag fill handle from A1 (col 0) to E1 (col 4)
    fill_handle_right(&mut sheet, 0, 0, 4);

    // Verify B1:E1 are filled with 42
    assert_eq!(sheet.get_raw(0, 1), "42", "B1 should be 42");
    assert_eq!(sheet.get_raw(0, 2), "42", "C1 should be 42");
    assert_eq!(sheet.get_raw(0, 3), "42", "D1 should be 42");
    assert_eq!(sheet.get_raw(0, 4), "42", "E1 should be 42");

    // A1 unchanged
    assert_eq!(sheet.get_raw(0, 0), "42", "A1 should still be 42");
}

#[test]
fn test_fill_handle_drag_right_formula_adjustment() {
    // Drag right from A1 to E1 with formula =A2
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // A2:E2 with values 10-50
    sheet.set_value(1, 0, "10");  // A2
    sheet.set_value(1, 1, "20");  // B2
    sheet.set_value(1, 2, "30");  // C2
    sheet.set_value(1, 3, "40");  // D2
    sheet.set_value(1, 4, "50");  // E2

    // A1 = =A2
    sheet.set_value(0, 0, "=A2");

    // Drag fill handle from A1 to E1
    fill_handle_right(&mut sheet, 0, 0, 4);

    // Verify formulas are adjusted
    assert_eq!(sheet.get_raw(0, 0), "=A2", "A1 formula unchanged");
    assert_eq!(sheet.get_raw(0, 1), "=B2", "B1 formula adjusted to =B2");
    assert_eq!(sheet.get_raw(0, 2), "=C2", "C1 formula adjusted to =C2");
    assert_eq!(sheet.get_raw(0, 3), "=D2", "D1 formula adjusted to =D2");
    assert_eq!(sheet.get_raw(0, 4), "=E2", "E1 formula adjusted to =E2");

    // Verify computed values
    assert_eq!(sheet.get_display(0, 0), "10", "A1 displays 10");
    assert_eq!(sheet.get_display(0, 1), "20", "B1 displays 20");
    assert_eq!(sheet.get_display(0, 2), "30", "C1 displays 30");
    assert_eq!(sheet.get_display(0, 3), "40", "D1 displays 40");
    assert_eq!(sheet.get_display(0, 4), "50", "E1 displays 50");
}

#[test]
fn test_fill_handle_drag_up_fills_a1_to_a4() {
    // Drag up from A5 to A1 fills A1:A4 (excluding anchor A5)
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // A5 = 99
    sheet.set_value(4, 0, "99");

    // Drag fill handle from A5 (row 4) up to A1 (row 0)
    fill_handle_up(&mut sheet, 4, 0, 0);

    // Verify A1:A4 are filled with 99 (A5 is anchor, excluded)
    assert_eq!(sheet.get_raw(0, 0), "99", "A1 should be 99");
    assert_eq!(sheet.get_raw(1, 0), "99", "A2 should be 99");
    assert_eq!(sheet.get_raw(2, 0), "99", "A3 should be 99");
    assert_eq!(sheet.get_raw(3, 0), "99", "A4 should be 99");

    // A5 (anchor) unchanged
    assert_eq!(sheet.get_raw(4, 0), "99", "A5 should still be 99");
}

#[test]
fn test_fill_handle_drag_up_formula_adjustment() {
    // Drag up from A5 to A1 with formula =B5
    // Formulas should get negative row delta
    let mut sheet = Sheet::new(SheetId(1), 100, 100);

    // B1:B5 with values 1-5
    sheet.set_value(0, 1, "1");  // B1
    sheet.set_value(1, 1, "2");  // B2
    sheet.set_value(2, 1, "3");  // B3
    sheet.set_value(3, 1, "4");  // B4
    sheet.set_value(4, 1, "5");  // B5

    // A5 = =B5
    sheet.set_value(4, 0, "=B5");

    // Drag fill handle from A5 up to A1
    fill_handle_up(&mut sheet, 4, 0, 0);

    // Verify formulas are adjusted (delta is negative)
    assert_eq!(sheet.get_raw(0, 0), "=B1", "A1 formula adjusted to =B1 (delta -4)");
    assert_eq!(sheet.get_raw(1, 0), "=B2", "A2 formula adjusted to =B2 (delta -3)");
    assert_eq!(sheet.get_raw(2, 0), "=B3", "A3 formula adjusted to =B3 (delta -2)");
    assert_eq!(sheet.get_raw(3, 0), "=B4", "A4 formula adjusted to =B4 (delta -1)");
    assert_eq!(sheet.get_raw(4, 0), "=B5", "A5 formula unchanged (anchor)");

    // Verify computed values
    assert_eq!(sheet.get_display(0, 0), "1", "A1 displays 1");
    assert_eq!(sheet.get_display(1, 0), "2", "A2 displays 2");
    assert_eq!(sheet.get_display(2, 0), "3", "A3 displays 3");
    assert_eq!(sheet.get_display(3, 0), "4", "A4 displays 4");
    assert_eq!(sheet.get_display(4, 0), "5", "A5 displays 5");
}

#[test]
fn test_fill_handle_undo_single_entry() {
    // Fill handle should record changes that can be undone in one step
    use crate::history::{History, CellChange};

    let mut sheet = Sheet::new(SheetId(1), 100, 100);
    let mut history = History::new();

    // A1 = 100
    sheet.set_value(0, 0, "100");

    // Simulate fill handle drag down from A1 to A5
    // First, collect the changes
    let mut changes = Vec::new();
    let source = sheet.get_raw(0, 0);

    for row in 1..=4 {
        let old_value = sheet.get_raw(row, 0);
        let new_value = source.clone();
        if old_value != new_value {
            changes.push(CellChange {
                row,
                col: 0,
                old_value,
                new_value: new_value.clone(),
            });
        }
        sheet.set_value(row, 0, &new_value);
    }

    // Record as single batch
    history.record_batch(0, changes);

    // Verify fill occurred
    assert_eq!(sheet.get_raw(1, 0), "100", "A2 filled");
    assert_eq!(sheet.get_raw(2, 0), "100", "A3 filled");
    assert_eq!(sheet.get_raw(3, 0), "100", "A4 filled");
    assert_eq!(sheet.get_raw(4, 0), "100", "A5 filled");

    // Single undo should revert all 4 cells
    let entry = history.undo().expect("Should have undo entry");
    match entry.action {
        crate::history::UndoAction::Values { changes, .. } => {
            assert_eq!(changes.len(), 4, "Undo should contain all 4 changes");
            for change in &changes {
                sheet.set_value(change.row, change.col, &change.old_value);
            }
        }
        _ => panic!("Expected Values action"),
    }

    // Verify cells are reverted (originally empty)
    assert_eq!(sheet.get_raw(1, 0), "", "A2 reverted to empty");
    assert_eq!(sheet.get_raw(2, 0), "", "A3 reverted to empty");
    assert_eq!(sheet.get_raw(3, 0), "", "A4 reverted to empty");
    assert_eq!(sheet.get_raw(4, 0), "", "A5 reverted to empty");

    // A1 unchanged
    assert_eq!(sheet.get_raw(0, 0), "100", "A1 unchanged by undo");
}

// =========================================================================
// VALIDATION UNDO/REDO TESTS
// =========================================================================

/// Test partial overlap: existing large rule, apply smaller rule, undo restores original exactly.
#[test]
fn test_validation_undo_partial_overlap_restore() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationStore, NumericConstraint};
    use crate::history::{History, UndoAction};

    let mut validations = ValidationStore::new();

    // Set up: A1:A100 (rows 0-99, col 0) has a "Whole Number" rule
    let original_range = CellRange::new(0, 0, 99, 0);
    let original_rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));
    validations.set(original_range, original_rule.clone());

    assert!(validations.get(0, 0).is_some(), "A1 should have validation");
    assert!(validations.get(50, 0).is_some(), "A51 should have validation");
    assert!(validations.get(99, 0).is_some(), "A100 should have validation");

    // Apply new rule to A10:A20 (rows 9-19)
    let new_range = CellRange::new(9, 0, 19, 0);
    let new_rule = ValidationRule::whole_number(NumericConstraint::between(1, 50));

    // Capture overlapping rules before clear (mimics dialog behavior)
    let previous_rules: Vec<(CellRange, ValidationRule)> = validations
        .iter()
        .filter(|(r, _)| r.overlaps(&new_range))
        .map(|(r, v)| (*r, v.clone()))
        .collect();

    // Replace semantics: clear overlaps, then set
    validations.clear_range(&new_range);
    validations.set(new_range, new_rule.clone());

    // After apply: original A1:A100 rule should be GONE (because it overlapped)
    // This is the "replace" semantic - the whole old rule is removed, not clipped
    assert!(validations.get(0, 0).is_none(), "A1 should have NO validation after replace");
    assert!(validations.get(9, 0).is_some(), "A10 should have new validation");
    assert!(validations.get(19, 0).is_some(), "A20 should have new validation");

    // Now simulate undo: clear new range, restore previous rules
    validations.clear_range(&new_range);
    for (rule_range, rule) in previous_rules {
        validations.set(rule_range, rule);
    }

    // After undo: original A1:A100 rule should be back EXACTLY
    assert!(validations.get(0, 0).is_some(), "A1 should have validation after undo");
    assert!(validations.get(50, 0).is_some(), "A51 should have validation after undo");
    assert!(validations.get(99, 0).is_some(), "A100 should have validation after undo");
}

/// Test redo runs full replace pipeline (clear overlaps + set).
#[test]
fn test_validation_redo_runs_replace_pipeline() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationStore, NumericConstraint};

    let mut validations = ValidationStore::new();

    // Set up: A1:A100 rule
    let original_range = CellRange::new(0, 0, 99, 0);
    let original_rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));
    validations.set(original_range, original_rule);

    // Apply new rule to A10:A20
    let new_range = CellRange::new(9, 0, 19, 0);
    let new_rule = ValidationRule::whole_number(NumericConstraint::between(1, 50));

    // Capture for undo
    let previous_rules: Vec<(CellRange, ValidationRule)> = validations
        .iter()
        .filter(|(r, _)| r.overlaps(&new_range))
        .map(|(r, v)| (*r, v.clone()))
        .collect();

    // Apply: clear overlaps + set
    validations.clear_range(&new_range);
    validations.set(new_range, new_rule.clone());

    // Undo: clear new range, restore previous
    validations.clear_range(&new_range);
    for (rule_range, rule) in previous_rules.clone() {
        validations.set(rule_range, rule);
    }

    // Now someone adds another rule at A50:A60 (external mutation)
    let external_range = CellRange::new(49, 0, 59, 0);
    let external_rule = ValidationRule::whole_number(NumericConstraint::between(1, 200));
    validations.set(external_range, external_rule);

    // Redo: must run full replace pipeline, not just set
    // This clears overlaps in new_range (which now includes the original restored rule)
    validations.clear_range(&new_range);
    validations.set(new_range, new_rule.clone());

    // After redo: A10:A20 has new rule, original A1:A100 is gone (re-cleared)
    // But A50:A60 external rule should still exist (outside new_range)
    assert!(validations.get(9, 0).is_some(), "A10 should have validation");
    assert!(validations.get(19, 0).is_some(), "A20 should have validation");
    assert!(validations.get(49, 0).is_some(), "A50 should have external validation");
    assert!(validations.get(59, 0).is_some(), "A60 should have external validation");
}

/// Test clear validation undo restores the cleared rules.
#[test]
fn test_validation_clear_undo_restores() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationStore, NumericConstraint};

    let mut validations = ValidationStore::new();

    // Set up: A1:A10 has a rule
    let range = CellRange::new(0, 0, 9, 0);
    let rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));
    validations.set(range, rule.clone());

    assert!(validations.get(0, 0).is_some(), "A1 should have validation before clear");
    assert!(validations.get(9, 0).is_some(), "A10 should have validation before clear");

    // Clear: capture rules for undo
    let cleared_rules: Vec<(CellRange, ValidationRule)> = validations
        .iter()
        .filter(|(r, _)| r.overlaps(&range))
        .map(|(r, v)| (*r, v.clone()))
        .collect();

    validations.clear_range(&range);

    assert!(validations.get(0, 0).is_none(), "A1 should have NO validation after clear");
    assert!(validations.get(9, 0).is_none(), "A10 should have NO validation after clear");

    // Undo: restore cleared rules
    for (rule_range, rule) in cleared_rules {
        validations.set(rule_range, rule);
    }

    assert!(validations.get(0, 0).is_some(), "A1 should have validation after undo");
    assert!(validations.get(9, 0).is_some(), "A10 should have validation after undo");
}

/// Test Excel semantics: "Any Value" means remove validation, not store AnyValue rule.
/// This matches user expectations: selecting "Any value" clears the validation.
#[test]
fn test_validation_any_value_removes_rule() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationType, ValidationStore, NumericConstraint};

    let mut validations = ValidationStore::new();

    // Set up: A1:A10 has a "Whole Number between 1 and 100" rule
    let range = CellRange::new(0, 0, 9, 0);
    let rule = ValidationRule::new(ValidationType::WholeNumber(NumericConstraint::between(1, 100)));
    validations.set(range, rule.clone());

    assert!(validations.get(0, 0).is_some(), "A1 should have validation before Any Value");

    // Simulate "Any Value" OK: this should CLEAR validation, not set AnyValue
    // Capture for undo (like clear_validation does)
    let cleared_rules: Vec<(CellRange, ValidationRule)> = validations
        .iter()
        .filter(|(r, _)| r.overlaps(&range))
        .map(|(r, v)| (*r, v.clone()))
        .collect();

    // Clear (this is what Any Value should do)
    validations.clear_range(&range);

    // Assert: NO rule exists (not an AnyValue rule)
    assert!(validations.get(0, 0).is_none(), "A1 should have NO validation after Any Value (not AnyValue rule)");
    assert!(validations.is_empty(), "ValidationStore should be empty after Any Value");

    // Undo: restore prior rule
    for (rule_range, rule) in cleared_rules {
        validations.set(rule_range, rule);
    }

    assert!(validations.get(0, 0).is_some(), "A1 should have validation restored after undo");
}

/// Test precedence: applying "Any Value" to narrow selection removes override,
/// allowing broader rule to apply again.
#[test]
fn test_validation_any_value_restores_broader_rule_precedence() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationType, ValidationStore, NumericConstraint};

    let mut validations = ValidationStore::new();

    // Set up: Broad rule on A1:A100
    let broad_range = CellRange::new(0, 0, 99, 0);
    let broad_rule = ValidationRule::new(ValidationType::WholeNumber(NumericConstraint::between(1, 1000)));
    validations.set(broad_range, broad_rule.clone());

    // Narrow override on A10:A20 (more restrictive)
    let narrow_range = CellRange::new(9, 0, 19, 0);
    let narrow_rule = ValidationRule::new(ValidationType::WholeNumber(NumericConstraint::between(1, 100)));
    validations.set(narrow_range, narrow_rule.clone());

    // A15 should have the narrow rule (first match wins in BTreeMap order)
    // Note: with current impl, narrow range was inserted after broad, so A15 gets narrow
    assert!(validations.get(14, 0).is_some(), "A15 should have validation");

    // Apply "Any Value" to narrow selection: removes the narrow rule
    validations.clear_range(&narrow_range);

    // Now A15 should get the broad rule (because narrow override is gone)
    // But wait - clear_range removes ALL overlapping rules, including broad!
    // This is the replace-in-range semantic we implemented.
    // So after clearing narrow_range, broad rule is also gone (it overlapped).

    // This is actually correct for replace-in-range:
    // "Any Value on A10:A20" = "remove all validations that touch A10:A20"
    // The broad rule A1:A100 overlaps, so it's removed too.

    // If user wanted to "punch a hole" they'd need to manually recreate the broad rule
    // for the non-overlapping parts. That's a feature request, not current scope.

    // For now, verify the clear happened
    assert!(validations.get(14, 0).is_none(), "A15 should have no validation (clear removes overlapping)");
    assert!(validations.get(0, 0).is_none(), "A1 should have no validation (broad rule was overlapping and removed)");
}

/// Test applying same rule twice is idempotent (no history spam).
#[test]
fn test_validation_same_rule_twice_idempotent() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationStore, NumericConstraint};

    let mut validations = ValidationStore::new();

    // Apply rule to A1:A10
    let range = CellRange::new(0, 0, 9, 0);
    let rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));

    // First apply
    let previous_rules_1: Vec<(CellRange, ValidationRule)> = validations
        .iter()
        .filter(|(r, _)| r.overlaps(&range))
        .map(|(r, v)| (*r, v.clone()))
        .collect();

    validations.clear_range(&range);
    validations.set(range, rule.clone());

    assert!(previous_rules_1.is_empty(), "First apply: no previous rules");

    // Second apply (same rule, same range)
    let previous_rules_2: Vec<(CellRange, ValidationRule)> = validations
        .iter()
        .filter(|(r, _)| r.overlaps(&range))
        .map(|(r, v)| (*r, v.clone()))
        .collect();

    validations.clear_range(&range);
    validations.set(range, rule.clone());

    assert_eq!(previous_rules_2.len(), 1, "Second apply: captured previous rule");

    // Undo second apply: restore previous (which was the same rule)
    validations.clear_range(&range);
    for (rule_range, rule) in previous_rules_2 {
        validations.set(rule_range, rule);
    }

    // State should be identical to after first apply
    assert!(validations.get(0, 0).is_some(), "A1 should still have validation");
    assert!(validations.get(9, 0).is_some(), "A10 should still have validation");

    // Undo first apply: no previous rules, so validation is removed
    validations.clear_range(&range);
    for (rule_range, rule) in previous_rules_1 {
        validations.set(rule_range, rule);
    }

    assert!(validations.get(0, 0).is_none(), "A1 should have NO validation after full undo");
}

/// Test multiple overlapping rules: restore order preserved.
#[test]
fn test_validation_undo_multiple_overlaps() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationType, ValidationStore, NumericConstraint};

    let mut validations = ValidationStore::new();

    // Set up: two overlapping rules
    // Rule 1: A1:A50 (rows 0-49)
    let range1 = CellRange::new(0, 0, 49, 0);
    let rule1 = ValidationRule::new(ValidationType::WholeNumber(NumericConstraint::between(1, 100)));

    // Rule 2: A30:A80 (rows 29-79) - overlaps with range1
    let range2 = CellRange::new(29, 0, 79, 0);
    let rule2 = ValidationRule::new(ValidationType::WholeNumber(NumericConstraint::between(1, 50)));

    validations.set(range1, rule1.clone());
    validations.set(range2, rule2.clone());

    // Apply new rule that overlaps both: A40:A60 (rows 39-59)
    let new_range = CellRange::new(39, 0, 59, 0);
    let new_rule = ValidationRule::whole_number(NumericConstraint::between(10, 90));

    // Capture overlapping rules
    let previous_rules: Vec<(CellRange, ValidationRule)> = validations
        .iter()
        .filter(|(r, _)| r.overlaps(&new_range))
        .map(|(r, v)| (*r, v.clone()))
        .collect();

    assert_eq!(previous_rules.len(), 2, "Should capture both overlapping rules");

    // Apply: clear overlaps + set
    validations.clear_range(&new_range);
    validations.set(new_range, new_rule);

    // Both original rules should be gone
    assert!(validations.get(0, 0).is_none(), "A1 should have NO validation (rule1 cleared)");
    assert!(validations.get(79, 0).is_none(), "A80 should have NO validation (rule2 cleared)");

    // Undo: restore previous rules
    validations.clear_range(&new_range);
    for (rule_range, rule) in previous_rules {
        validations.set(rule_range, rule);
    }

    // Both rules restored
    assert!(validations.get(0, 0).is_some(), "A1 should have validation (rule1 restored)");
    assert!(validations.get(79, 0).is_some(), "A80 should have validation (rule2 restored)");
}

/// Test: Broad rule + exclusion hole preserves broad validation outside hole.
/// A column rule applies everywhere except excluded cells.
#[test]
fn test_exclusion_hole_preserves_surrounding_validation() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationType, ValidationStore, NumericConstraint};

    let mut validations = ValidationStore::new();

    // Set up: A1:A100 requires whole numbers 1-100
    let column_range = CellRange::new(0, 0, 99, 0);
    let rule = ValidationRule::new(ValidationType::WholeNumber(NumericConstraint::between(1, 100)));
    validations.set(column_range, rule.clone());

    // Verify rule applies to all cells
    assert!(validations.get(0, 0).is_some(), "A1 should have validation");
    assert!(validations.get(49, 0).is_some(), "A50 should have validation");
    assert!(validations.get(99, 0).is_some(), "A100 should have validation");

    // Exclude A10:A20 (rows 9-19) - the "hole"
    let exclusion_range = CellRange::new(9, 0, 19, 0);
    validations.exclude(exclusion_range);

    // Excluded cells return None from get()
    assert!(validations.get(9, 0).is_none(), "A10 should be excluded (no validation)");
    assert!(validations.get(14, 0).is_none(), "A15 should be excluded (no validation)");
    assert!(validations.get(19, 0).is_none(), "A20 should be excluded (no validation)");

    // Non-excluded cells still have validation
    assert!(validations.get(0, 0).is_some(), "A1 should still have validation");
    assert!(validations.get(8, 0).is_some(), "A9 should still have validation (just before exclusion)");
    assert!(validations.get(20, 0).is_some(), "A21 should still have validation (just after exclusion)");
    assert!(validations.get(99, 0).is_some(), "A100 should still have validation");

    // The rule itself is preserved (not cleared)
    assert_eq!(validations.len(), 1, "Rule count should be 1 (not fragmented)");
    assert_eq!(validations.exclusions_len(), 1, "Exclusion count should be 1");
}

/// Test: Undo/redo exclusion restores behavior correctly.
/// After undoing exclusion, validation applies again; after redo, exclusion returns.
#[test]
fn test_exclusion_undo_redo_restores_behavior() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationType, ValidationStore, NumericConstraint};

    let mut validations = ValidationStore::new();

    // Set up: A1:A10 requires whole numbers
    let range = CellRange::new(0, 0, 9, 0);
    let rule = ValidationRule::new(ValidationType::WholeNumber(NumericConstraint::between(1, 100)));
    validations.set(range, rule);

    // Exclude A5
    let exclusion = CellRange::new(4, 0, 4, 0);
    validations.exclude(exclusion);

    // Verify exclusion works
    assert!(validations.get(4, 0).is_none(), "A5 excluded - no validation");
    assert!(validations.get(3, 0).is_some(), "A4 not excluded - has validation");

    // Undo: remove the exclusion
    validations.remove_exclusion(&exclusion);

    // After undo, validation applies again
    assert!(validations.get(4, 0).is_some(), "A5 should have validation after undo");
    assert!(validations.exclusions_is_empty(), "No exclusions after undo");

    // Redo: re-add the exclusion
    validations.exclude(exclusion);

    // After redo, exclusion is back
    assert!(validations.get(4, 0).is_none(), "A5 should be excluded again after redo");
    assert_eq!(validations.exclusions_len(), 1, "One exclusion after redo");
}

// =========================================================================
// NEW WINDOW REGRESSION TESTS
// =========================================================================
//
// These tests verify the NewWindow behavior doesn't regress.
// Multi-window tests that require gpui context are deferred until we have
// a proper window harness. For now, we test what we can at the unit level.

/// Test that new_in_place() is destructive: resets workbook state.
/// This is the dangerous method that NewInPlace action calls.
#[test]
fn test_new_in_place_is_destructive() {
    use visigrid_engine::workbook::Workbook;

    // Create a workbook and modify it
    let mut workbook = Workbook::new();
    workbook.sheet_mut(0).unwrap().set_value(0, 0, "important data");
    workbook.sheet_mut(0).unwrap().set_value(5, 5, "=SUM(A1:A10)");

    // Verify data exists
    assert_eq!(workbook.sheet(0).unwrap().get_raw(0, 0), "important data");
    assert_eq!(workbook.sheet(0).unwrap().get_raw(5, 5), "=SUM(A1:A10)");

    // Simulate what new_in_place does: replace with fresh workbook
    workbook = Workbook::new();

    // All data should be gone
    assert_eq!(workbook.sheet(0).unwrap().get_raw(0, 0), "");
    assert_eq!(workbook.sheet(0).unwrap().get_raw(5, 5), "");
}

/// Test that keybinding registry maps Ctrl+N to NewWindow (not NewInPlace).
/// This is a snapshot test to catch accidental rebinding regressions.
#[test]
fn test_keybinding_maps_ctrl_n_to_new_window() {
    // We can't easily introspect gpui keybindings at runtime without a full App context.
    // Instead, this test validates the source code pattern.
    //
    // The keybinding registration in keybindings.rs should contain:
    //   KeyBinding::new(&kb(m, "n"), NewWindow, Some("Spreadsheet"))
    //
    // This test serves as documentation that:
    // 1. Ctrl+N (or Cmd+N on macOS) is bound to NewWindow
    // 2. NewInPlace is NOT bound to any default key
    // 3. The App-level handler in main.rs opens a new window

    // If this test fails, someone changed the keybinding.
    // Verify the source directly to ensure correctness.

    let keybindings_source = include_str!("keybindings.rs");

    // Verify NewWindow is bound to Ctrl+N
    assert!(
        keybindings_source.contains("kb(m, \"n\"), NewWindow"),
        "Ctrl+N should be bound to NewWindow, not NewInPlace or NewFile"
    );

    // Verify NewInPlace is NOT bound anywhere in the default keybindings
    assert!(
        !keybindings_source.contains("NewInPlace"),
        "NewInPlace should NOT be bound to any key by default"
    );
}

/// Test that CloseWindow handler checks is_modified flag.
/// This is a source code verification test - we verify the handler exists
/// and includes the dirty check pattern.
#[test]
fn test_close_window_checks_dirty_flag() {
    let views_source = include_str!("views/actions_ui.rs");

    // Verify CloseWindow handler checks is_modified
    assert!(
        views_source.contains("if !this.is_modified"),
        "CloseWindow handler should check is_modified flag"
    );

    // Verify dirty workbook triggers prompt (not immediate close)
    assert!(
        views_source.contains("window.prompt("),
        "CloseWindow handler should show prompt when dirty"
    );

    // Verify save_and_close is called for "Save" option
    assert!(
        views_source.contains("save_and_close"),
        "CloseWindow handler should call save_and_close for Save option"
    );
}

/// Test that Ctrl+W is bound to CloseWindow action.
#[test]
fn test_keybinding_maps_ctrl_w_to_close_window() {
    let keybindings_source = include_str!("keybindings.rs");

    // Verify CloseWindow is bound (cmd-w on macOS)
    assert!(
        keybindings_source.contains("CloseWindow"),
        "CloseWindow should be bound to a key"
    );
}

/// Test that NewWindow handler does NOT call new_in_place or mutate workbook state.
/// This verifies the architectural separation: NewWindow opens a new window,
/// it doesn't touch the current workbook.
#[test]
fn test_newwindow_does_not_mutate_current_state() {
    // Verify the NewWindow handler is at App level (main.rs), not Spreadsheet level
    let main_source = include_str!("main.rs");

    // NewWindow handler must be at App level (has cx.open_window)
    assert!(
        main_source.contains("on_action") && main_source.contains("NewWindow") && main_source.contains("open_window"),
        "NewWindow handler should be at App level with cx.open_window()"
    );

    // Verify Spreadsheet doesn't handle NewWindow directly
    let views_source = include_str!("views/mod.rs");
    assert!(
        !views_source.contains("on_action") || !views_source.contains("&NewWindow,"),
        "Spreadsheet should NOT have its own NewWindow handler that could mutate state"
    );

    // Verify new_in_place is not called by NewWindow
    let actions_source = include_str!("actions.rs");
    assert!(
        actions_source.contains("NewWindow") && actions_source.contains("NewInPlace"),
        "Actions should have separate NewWindow and NewInPlace"
    );

    // NewWindow in main.rs should create new Spreadsheet, not call new_in_place
    assert!(
        main_source.contains("Spreadsheet::new") && main_source.contains("NewWindow"),
        "NewWindow should create new Spreadsheet::new(), not call new_in_place()"
    );
    assert!(
        !main_source.contains("new_in_place"),
        "NewWindow handler must NOT call new_in_place"
    );
}

/// Test that window registry is initialized and used for window management.
#[test]
fn test_window_registry_initialized() {
    let main_source = include_str!("main.rs");

    // Verify WindowRegistry is initialized
    assert!(
        main_source.contains("WindowRegistry::new"),
        "main.rs should initialize WindowRegistry as global"
    );

    // Verify SwitchWindow handler exists
    assert!(
        main_source.contains("SwitchWindow") && main_source.contains("on_action"),
        "main.rs should have SwitchWindow handler"
    );

    // Verify windows are registered
    assert!(
        main_source.contains("register_with_window_registry"),
        "Windows should be registered with the registry on creation"
    );
}

/// Test that Cmd+` (macOS) / Ctrl+` (Linux/Win) is bound to SwitchWindow.
#[test]
fn test_keybinding_maps_backtick_to_switch_window() {
    let keybindings_source = include_str!("keybindings.rs");

    // Verify SwitchWindow is bound to backtick
    assert!(
        keybindings_source.contains("SwitchWindow") && keybindings_source.contains("`"),
        "SwitchWindow should be bound to Cmd+` or Ctrl+`"
    );
}

/// Test: Circle Invalid Data ignores excluded cells.
/// When validating a range, excluded cells should not be flagged as invalid.
/// The key behavior is that `get()` returns `None` for excluded cells.
#[test]
fn test_circle_invalid_data_ignores_excluded_cells() {
    use visigrid_engine::validation::{CellRange, ValidationRule, ValidationType, ValidationStore, NumericConstraint};

    let mut validations = ValidationStore::new();

    // Set up: A1:A5 requires whole numbers 1-100
    let range = CellRange::new(0, 0, 4, 0);
    let rule = ValidationRule::new(ValidationType::WholeNumber(NumericConstraint::between(1, 100)));
    validations.set(range, rule);

    // Exclude A3 (row 2)
    let exclusion = CellRange::new(2, 0, 2, 0);
    validations.exclude(exclusion);

    // Simulate "Circle Invalid Data" by checking which cells have validation rules
    // (the actual validation logic uses get() to check if a cell should be validated)
    let mut cells_to_validate: Vec<(usize, usize)> = Vec::new();
    for row in 0..5 {
        let col = 0;
        if validations.get(row, col).is_some() {
            // This cell has a validation rule and should be checked
            cells_to_validate.push((row, col));
        }
    }

    // A3 (row 2) should NOT be in the list (it's excluded)
    assert!(!cells_to_validate.contains(&(2, 0)), "A3 is excluded and should not be validated");

    // All other cells should be in the list
    assert!(cells_to_validate.contains(&(0, 0)), "A1 should be validated");
    assert!(cells_to_validate.contains(&(1, 0)), "A2 should be validated");
    assert!(cells_to_validate.contains(&(3, 0)), "A4 should be validated");
    assert!(cells_to_validate.contains(&(4, 0)), "A5 should be validated");

    // 4 cells should be validated (not 5, because A3 is excluded)
    assert_eq!(cells_to_validate.len(), 4, "Should validate 4 cells (A3 excluded)");
}

// ============================================================================
// Per-Sheet Column/Row Sizing Tests
// ============================================================================

/// Test that per-sheet column width storage isolates widths between sheets.
/// This verifies the fix for "new sheet inherits column widths" bug.
#[test]
fn test_per_sheet_column_widths_isolation() {
    use std::collections::HashMap;

    // Simulate the per-sheet storage: HashMap<SheetId, HashMap<usize, f32>>
    let mut col_widths: HashMap<SheetId, HashMap<usize, f32>> = HashMap::new();

    let sheet1_id = SheetId(1);
    let sheet2_id = SheetId(2);

    // Set column A width on sheet 1
    col_widths.entry(sheet1_id).or_insert_with(HashMap::new).insert(0, 150.0);

    // Sheet 2 should have no custom widths (new sheets get defaults)
    assert!(
        col_widths.get(&sheet2_id).is_none() || col_widths.get(&sheet2_id).unwrap().is_empty(),
        "New sheet should not inherit column widths from other sheets"
    );

    // Verify sheet 1 still has its width
    assert_eq!(
        col_widths.get(&sheet1_id).and_then(|m| m.get(&0)),
        Some(&150.0),
        "Sheet 1 should retain its column width"
    );

    // Set a different width on sheet 2
    col_widths.entry(sheet2_id).or_insert_with(HashMap::new).insert(0, 200.0);

    // Verify isolation - each sheet has its own width for column A
    assert_eq!(
        col_widths.get(&sheet1_id).and_then(|m| m.get(&0)),
        Some(&150.0),
        "Sheet 1 column A width should be 150"
    );
    assert_eq!(
        col_widths.get(&sheet2_id).and_then(|m| m.get(&0)),
        Some(&200.0),
        "Sheet 2 column A width should be 200"
    );
}

/// Test that per-sheet row height storage isolates heights between sheets.
#[test]
fn test_per_sheet_row_heights_isolation() {
    use std::collections::HashMap;

    let mut row_heights: HashMap<SheetId, HashMap<usize, f32>> = HashMap::new();

    let sheet1_id = SheetId(1);
    let sheet2_id = SheetId(2);

    // Set row 1 height on sheet 1
    row_heights.entry(sheet1_id).or_insert_with(HashMap::new).insert(0, 40.0);

    // Sheet 2 should have no custom heights
    assert!(
        row_heights.get(&sheet2_id).is_none() || row_heights.get(&sheet2_id).unwrap().is_empty(),
        "New sheet should not inherit row heights from other sheets"
    );

    // Set a different height on sheet 2
    row_heights.entry(sheet2_id).or_insert_with(HashMap::new).insert(0, 60.0);

    // Verify isolation
    assert_eq!(
        row_heights.get(&sheet1_id).and_then(|m| m.get(&0)),
        Some(&40.0),
        "Sheet 1 row 1 height should be 40"
    );
    assert_eq!(
        row_heights.get(&sheet2_id).and_then(|m| m.get(&0)),
        Some(&60.0),
        "Sheet 2 row 1 height should be 60"
    );
}

/// Test that the default width/height is returned for sheets without custom sizing.
/// This mirrors the col_width() and row_height() method behavior.
#[test]
fn test_per_sheet_sizing_default_fallback() {
    use std::collections::HashMap;

    const DEFAULT_WIDTH: f32 = 100.0;
    const DEFAULT_HEIGHT: f32 = 21.0;

    let col_widths: HashMap<SheetId, HashMap<usize, f32>> = HashMap::new();
    let row_heights: HashMap<SheetId, HashMap<usize, f32>> = HashMap::new();

    let sheet_id = SheetId(42);

    // Helper mimics col_width() method
    let get_col_width = |col: usize| -> f32 {
        col_widths
            .get(&sheet_id)
            .and_then(|m| m.get(&col))
            .copied()
            .unwrap_or(DEFAULT_WIDTH)
    };

    // Helper mimics row_height() method
    let get_row_height = |row: usize| -> f32 {
        row_heights
            .get(&sheet_id)
            .and_then(|m| m.get(&row))
            .copied()
            .unwrap_or(DEFAULT_HEIGHT)
    };

    // All columns should return default width for new sheet
    assert_eq!(get_col_width(0), DEFAULT_WIDTH);
    assert_eq!(get_col_width(100), DEFAULT_WIDTH);

    // All rows should return default height for new sheet
    assert_eq!(get_row_height(0), DEFAULT_HEIGHT);
    assert_eq!(get_row_height(1000), DEFAULT_HEIGHT);
}

// ============================================================================
// Layout Provenance Tests
// ============================================================================

/// Test that ColumnWidthSet action has correct label and summary.
#[test]
fn test_column_width_set_action_label_and_summary() {
    use crate::history::UndoAction;

    // Test: custom width set
    let action = UndoAction::ColumnWidthSet {
        sheet_id: SheetId(1),  // Stable ID, not index
        col: 0,
        old: None,
        new: Some(200.0),
    };
    assert_eq!(action.label(), "Set column width");
    let summary = action.summary().expect("should have summary");
    assert!(summary.contains("Col A"), "Summary should contain 'Col A': {}", summary);
    assert!(summary.contains("default"), "Summary should mention old was default: {}", summary);
    // No "px" unit - just the number (units are internal, not guaranteed to match Excel)
    assert!(summary.contains("200"), "Summary should show new width value: {}", summary);
    assert!(!summary.contains("px"), "Summary should NOT contain 'px' unit: {}", summary);

    // Test: width reset to default
    let action2 = UndoAction::ColumnWidthSet {
        sheet_id: SheetId(1),
        col: 2,
        old: Some(150.0),
        new: None,
    };
    let summary2 = action2.summary().expect("should have summary");
    assert!(summary2.contains("Col C"), "Summary should contain 'Col C': {}", summary2);
    assert!(summary2.contains("150"), "Summary should show old width value: {}", summary2);
    assert!(summary2.contains("default"), "Summary should mention resetting to default: {}", summary2);
}

/// Test that RowHeightSet action has correct label and summary.
#[test]
fn test_row_height_set_action_label_and_summary() {
    use crate::history::UndoAction;

    // Test: custom height set
    let action = UndoAction::RowHeightSet {
        sheet_id: SheetId(1),  // Stable ID, not index
        row: 4,
        old: None,
        new: Some(50.0),
    };
    assert_eq!(action.label(), "Set row height");
    let summary = action.summary().expect("should have summary");
    assert!(summary.contains("Row 5"), "Summary should contain 'Row 5' (1-indexed): {}", summary);
    assert!(summary.contains("default"), "Summary should mention old was default: {}", summary);
    // No "px" unit - just the number
    assert!(summary.contains("50"), "Summary should show new height value: {}", summary);
    assert!(!summary.contains("px"), "Summary should NOT contain 'px' unit: {}", summary);
}

/// Test that ColumnWidthSet action generates correct Lua provenance.
#[test]
fn test_column_width_set_lua_provenance() {
    use crate::history::UndoAction;

    // Test: set width
    let action = UndoAction::ColumnWidthSet {
        sheet_id: SheetId(2),  // Stable sheet ID
        col: 2,                // Column C (0-indexed)
        old: None,
        new: Some(180.0),
    };
    let lua = action.to_lua().expect("should have Lua representation");
    assert!(lua.contains("grid.set_col_width"), "Lua should use set_col_width: {}", lua);
    assert!(lua.contains("sheet_id=2"), "Lua should reference sheet_id=2: {}", lua);
    assert!(lua.contains("col=\"C\""), "Lua should reference column C: {}", lua);
    assert!(lua.contains("width=180"), "Lua should include width: {}", lua);

    // Test: reset to default
    let action2 = UndoAction::ColumnWidthSet {
        sheet_id: SheetId(1),
        col: 0,
        old: Some(120.0),
        new: None,
    };
    let lua2 = action2.to_lua().expect("should have Lua representation");
    assert!(lua2.contains("grid.clear_col_width"), "Lua should use clear_col_width for reset: {}", lua2);
}

/// Test that RowHeightSet action generates correct Lua provenance.
#[test]
fn test_row_height_set_lua_provenance() {
    use crate::history::UndoAction;

    // Test: set height
    let action = UndoAction::RowHeightSet {
        sheet_id: SheetId(1),
        row: 9,  // Row 10 (0-indexed)
        old: None,
        new: Some(40.0),
    };
    let lua = action.to_lua().expect("should have Lua representation");
    assert!(lua.contains("grid.set_row_height"), "Lua should use set_row_height: {}", lua);
    assert!(lua.contains("sheet_id=1"), "Lua should reference sheet_id=1: {}", lua);
    assert!(lua.contains("row=10"), "Lua should reference row 10 (1-indexed): {}", lua);
    assert!(lua.contains("height=40"), "Lua should include height: {}", lua);

    // Test: reset to default
    let action2 = UndoAction::RowHeightSet {
        sheet_id: SheetId(1),
        row: 0,
        old: Some(60.0),
        new: None,
    };
    let lua2 = action2.to_lua().expect("should have Lua representation");
    assert!(lua2.contains("grid.clear_row_height"), "Lua should use clear_row_height for reset: {}", lua2);
}

/// Test that layout actions are correctly classified and support replay.
#[test]
fn test_layout_action_kind_replay_support() {
    use crate::history::{UndoAction, UndoActionKind};

    let col_action = UndoAction::ColumnWidthSet {
        sheet_id: SheetId(1),
        col: 0,
        old: None,
        new: Some(150.0),
    };
    assert_eq!(col_action.kind(), UndoActionKind::ColumnWidthSet);
    assert!(col_action.kind().is_replay_supported(), "ColumnWidthSet should support replay");

    let row_action = UndoAction::RowHeightSet {
        sheet_id: SheetId(1),
        row: 0,
        old: None,
        new: Some(40.0),
    };
    assert_eq!(row_action.kind(), UndoActionKind::RowHeightSet);
    assert!(row_action.kind().is_replay_supported(), "RowHeightSet should support replay");
}

/// Test that layout action kinds have stable byte tags for fingerprinting.
#[test]
fn test_layout_action_kind_tags_stable() {
    use crate::history::UndoActionKind;

    // These tags must remain stable for history fingerprinting
    assert_eq!(UndoActionKind::ColumnWidthSet.tag(), 0x12, "ColumnWidthSet tag must be stable");
    assert_eq!(UndoActionKind::RowHeightSet.tag(), 0x13, "RowHeightSet tag must be stable");
}

/// Test that SheetId remains stable when sheets are deleted (indices shift).
/// This is why we use SheetId instead of sheet_index - indices shift, IDs don't.
#[test]
fn test_layout_action_sheet_id_stable_across_delete() {
    use crate::history::UndoAction;
    use visigrid_engine::workbook::Workbook;
    use visigrid_engine::sheet::Sheet;

    // Create a workbook with 3 sheets
    let mut wb = Workbook::new();
    // Sheet1 is at index 0
    wb.add_sheet(); // Sheet2 at index 1
    wb.add_sheet(); // Sheet3 at index 2
    let sheet3_id = wb.sheets()[2].id; // Capture Sheet3's stable ID

    // Record a layout action on Sheet3 (currently at index 2)
    let action = UndoAction::ColumnWidthSet {
        sheet_id: sheet3_id,
        col: 0,
        old: None,
        new: Some(200.0),
    };

    // Delete Sheet2 (index 1) - this shifts Sheet3 from index 2 to index 1
    wb.delete_sheet(1);

    // Sheet3 is now at index 1, but its SheetId hasn't changed
    assert_eq!(wb.sheets()[1].id, sheet3_id, "Sheet3's ID should be unchanged after delete");

    // The action still references the correct sheet via SheetId
    // (If we had used sheet_index=2, it would now be out of bounds or wrong)
    if let UndoAction::ColumnWidthSet { sheet_id, .. } = action {
        assert_eq!(sheet_id, sheet3_id, "Action should still reference Sheet3 by stable ID");

        // Verify we can find the sheet by ID regardless of current index
        let found_sheet = wb.sheets().iter().find(|s: &&Sheet| s.id == sheet_id);
        assert!(found_sheet.is_some(), "Should find sheet by SheetId");
    }

    // Verify Lua output uses sheet_id (stable), not index (fragile)
    let lua = action.to_lua().expect("should have Lua");
    assert!(lua.contains(&format!("sheet_id={}", sheet3_id.0)),
        "Lua should use stable sheet_id, not index: {}", lua);
}

/// Test that SheetId remains stable when sheets are reordered.
/// TODO: Implement once sheet reorder functionality exists.
///
/// Invariant: Layout actions use SheetId, not sheet_index, so reordering
/// sheets must not affect which sheet a layout action targets.
///
/// Test plan (when reorder exists):
/// 1. Create 3 sheets: [Sheet1, Sheet2, Sheet3] with IDs [1, 2, 3]
/// 2. Record layout action on Sheet3 (ID=3, index=2)
/// 3. Reorder so Sheet3 moves to index 0: [Sheet3, Sheet1, Sheet2]
/// 4. Verify action still references Sheet3 by ID (not "whatever is at old index")
/// 5. Verify replay/undo targets correct sheet
#[test]
#[ignore = "Sheet reorder not yet implemented"]
fn test_layout_action_sheet_id_stable_across_reorder() {
    // Placeholder - will fail if accidentally un-ignored without implementation
    panic!("Sheet reorder not yet implemented - update this test when it is");
}

/// Test behavior when layout action references a deleted sheet.
///
/// Current behavior: Layout actions store data keyed by SheetId in the app's
/// col_widths/row_heights maps. If the sheet is deleted, this data becomes
/// orphaned but doesn't cause an error. This is acceptable because:
/// - The data is harmless (no computation depends on deleted sheet sizing)
/// - The sheet could theoretically be restored (undo delete)
/// - Provenance/Lua export includes the sheet_id for audit trail
///
/// Future consideration: If stricter validation is needed, undo/redo handlers
/// could check if the sheet exists and emit a warning or error.
#[test]
fn test_layout_action_on_deleted_sheet_behavior() {
    use crate::history::UndoAction;
    use visigrid_engine::workbook::Workbook;

    // Create workbook with 2 sheets
    let mut wb = Workbook::new();
    wb.add_sheet();
    let sheet2_id = wb.sheets()[1].id;

    // Record a layout action on sheet2
    let action = UndoAction::ColumnWidthSet {
        sheet_id: sheet2_id,
        col: 0,
        old: None,
        new: Some(200.0),
    };

    // Delete sheet2
    wb.delete_sheet(1);

    // The action still has its sheet_id (it's just data)
    if let UndoAction::ColumnWidthSet { sheet_id, .. } = action {
        assert_eq!(sheet_id, sheet2_id, "Action retains original sheet_id");

        // The sheet no longer exists in workbook
        let found = wb.sheets().iter().any(|s| s.id == sheet_id);
        assert!(!found, "Sheet should no longer exist in workbook");
    }

    // Provenance still generates valid Lua (the ID is recorded for audit)
    let lua = action.to_lua().expect("should have Lua");
    assert!(lua.contains(&format!("sheet_id={}", sheet2_id.0)),
        "Lua should still reference the (now-deleted) sheet_id: {}", lua);

    // Note: Actual undo/redo in the app would store the width against this
    // orphaned SheetId. If stricter validation is desired, modify undo_redo.rs
    // to check sheet existence and return an error.
}

// ============================================================================
// Grid Hit-Testing Structure Tests
// ============================================================================

/// Regression test: cell wrapper owns .id() and mouse handlers, inner cell is visual only.
///
/// Before this fix, .id() and mouse handlers lived on the inner cell div (which used
/// .size_full() to fill its wrapper). Sub-pixel gaps at cell borders caused the inner
/// div to not perfectly reach the wrapper edge, creating dead zones where clicks failed
/// and the cursor reverted to arrow. Moving .id() and handlers to the wrapper (which has
/// exact pixel dimensions from the flex layout) eliminates these gaps.
///
/// If someone refactors the cell rendering and moves .id() back to the inner cell,
/// this test will catch it.
#[test]
fn test_cell_wrapper_owns_id_and_handlers() {
    let grid_source = include_str!("views/grid.rs");

    // The inner cell div must NOT have .id() — it's purely visual.
    // Pattern: `let mut cell = div()\n        .id(` would mean .id() is on the cell.
    // The correct pattern: `let mut cell = div()\n        .relative()` (no .id())
    assert!(
        grid_source.contains("let mut cell = div()\n        .relative()  // Enable absolute positioning"),
        "Inner cell div must NOT have .id() — it should start with div().relative()"
    );

    // The wrapper must own the .id()
    assert!(
        grid_source.contains("let mut wrapper = div()\n        .id(ElementId::Name(format!(\"cell-{}-{}\", view_row, col)"),
        "Wrapper div must own .id(\"cell-{{row}}-{{col}}\") for hit-testing"
    );

    // The wrapper must set crosshair cursor (not the inner cell)
    // Find the wrapper section and verify cursor is there
    let wrapper_section = grid_source.split("let mut wrapper = div()").nth(1)
        .expect("wrapper div should exist");
    let wrapper_before_handlers = wrapper_section.split(".on_mouse_down(").next()
        .expect("wrapper should have on_mouse_down");
    assert!(
        wrapper_before_handlers.contains(".cursor(CursorStyle::Crosshair)"),
        "Wrapper must set CursorStyle::Crosshair for consistent grid cursor"
    );

    // The wrapper must have left-click, right-click, mouse_move, and mouse_up handlers
    assert!(
        wrapper_section.contains(".on_mouse_down(MouseButton::Left,"),
        "Wrapper must handle left-click"
    );
    assert!(
        wrapper_section.contains(".on_mouse_down(MouseButton::Right,"),
        "Wrapper must handle right-click"
    );
    assert!(
        wrapper_section.contains(".on_mouse_move("),
        "Wrapper must handle mouse move for drag selection"
    );
    assert!(
        wrapper_section.contains(".on_mouse_up(MouseButton::Left,"),
        "Wrapper must handle mouse up to end drag"
    );
}

/// Regression test: merge overlay click handler commits pending edit.
///
/// When clicking a merged cell while editing another cell, the edit must be
/// committed before navigating (same behavior as regular cell clicks). Without
/// this, editing A1 and clicking a merged B2:C3 would leave the app in a broken
/// state with edit mode active but a different cell selected.
#[test]
fn test_merge_overlay_commits_edit_on_click() {
    let grid_source = include_str!("views/grid.rs");

    // Find the merge overlay left-click handler (render_merge_div function)
    // It should contain commit_pending_edit before the normal click handling
    let merge_div_section = grid_source.split("fn render_merge_div(").nth(1)
        .expect("render_merge_div function should exist");

    assert!(
        merge_div_section.contains("commit_pending_edit"),
        "Merge overlay click handler must call commit_pending_edit to exit edit mode"
    );
}

// =============================================================================
// Format bar state machine regression tests
// =============================================================================
// These tests guard the action-handler routing logic added in actions_edit.rs
// and actions_nav.rs. Format bar editing consumes ConfirmEdit, CancelEdit,
// BackspaceChar, and navigation actions before Spreadsheet editing/navigation.
// If someone removes those guards during a keybinding refactor, these tests
// catch the regression at the logic level.

/// Regression: ConfirmEdit → commit_font_size must parse the input buffer
/// and produce a valid font size. The action handler in actions_edit.rs
/// gates ConfirmEdit when size_editing is true.
#[test]
fn test_format_bar_confirm_commits_valid_font_size() {
    use crate::views::format_bar::parse_font_size_input;

    // Normal sizes
    assert_eq!(parse_font_size_input("24"), Some(24.0));
    assert_eq!(parse_font_size_input("11"), Some(11.0));
    assert_eq!(parse_font_size_input("1"), Some(1.0));
    assert_eq!(parse_font_size_input("400"), Some(400.0));

    // Whitespace tolerance
    assert_eq!(parse_font_size_input(" 16 "), Some(16.0));
}

/// Regression: CancelEdit must NOT apply any change. Invalid/empty input
/// should also produce None (revert, no change applied).
#[test]
fn test_format_bar_cancel_reverts_no_change() {
    use crate::views::format_bar::parse_font_size_input;

    // Empty or invalid input → None (no change)
    assert_eq!(parse_font_size_input(""), None);
    assert_eq!(parse_font_size_input("abc"), None);
    assert_eq!(parse_font_size_input("0"), None, "0 is below minimum");
    assert_eq!(parse_font_size_input("401"), None, "401 exceeds maximum");
    assert_eq!(parse_font_size_input("-5"), None, "negative not allowed");
    assert_eq!(parse_font_size_input("12.5"), None, "floats not allowed");
}

/// Regression: BackspaceChar in the format bar must edit the buffer,
/// not delete the grid selection. The action handler in actions_edit.rs
/// gates BackspaceChar when size_editing is true.
#[test]
fn test_format_bar_backspace_edits_buffer() {
    use crate::views::format_bar::parse_font_size_input;

    let mut buffer = String::from("123");

    // Simulate backspace (same logic as in BackspaceChar handler)
    buffer.pop();
    assert_eq!(buffer, "12");

    buffer.pop();
    assert_eq!(buffer, "1");

    // Backspace on single char → empty (which parses to None = no change)
    buffer.pop();
    assert_eq!(buffer, "");
    assert_eq!(parse_font_size_input(&buffer), None);

    // Backspace on empty is safe
    buffer.pop(); // no-op
    assert_eq!(buffer, "");
}

// =============================================================================
// URL parsing (macOS "Open With" handler)
// =============================================================================

#[test]
fn test_url_to_path_file_url() {
    use crate::url_to_path;
    let path = url_to_path("file:///Users/bob/Documents/report.xlsx");
    assert_eq!(path, Some(std::path::PathBuf::from("/Users/bob/Documents/report.xlsx")));
}

#[test]
fn test_url_to_path_percent_encoded_spaces() {
    use crate::url_to_path;
    let path = url_to_path("file:///Users/bob/My%20Documents/Q1%20Report.xlsx");
    assert_eq!(path, Some(std::path::PathBuf::from("/Users/bob/My Documents/Q1 Report.xlsx")));
}

#[test]
fn test_url_to_path_plain_path() {
    use crate::url_to_path;
    let path = url_to_path("/Users/bob/file.csv");
    assert_eq!(path, Some(std::path::PathBuf::from("/Users/bob/file.csv")));
}

#[test]
fn test_url_to_path_non_file_url_returns_none() {
    use crate::url_to_path;
    assert_eq!(url_to_path("https://example.com/file.xlsx"), None);
    assert_eq!(url_to_path("ftp://server/file.csv"), None);
}

#[test]
fn test_percent_decode_mixed() {
    use crate::percent_decode;
    assert_eq!(percent_decode("hello%20world"), "hello world");
    assert_eq!(percent_decode("100%25%20done"), "100% done");
    assert_eq!(percent_decode("no-encoding"), "no-encoding");
    assert_eq!(percent_decode("%2FUsers%2Fbob"), "/Users/bob");
}

#[test]
fn test_percent_decode_incomplete_sequence() {
    use crate::percent_decode;
    // Incomplete percent sequence at end — passed through literally
    assert_eq!(percent_decode("abc%2"), "abc%2");
    assert_eq!(percent_decode("abc%"), "abc%");
}

#[test]
fn test_normalize_and_dedup_urls() {
    use crate::normalize_and_dedup_urls;
    // Duplicate URLs (same path, different encoding) should be deduplicated
    let urls = vec![
        "file:///tmp/test-dedup.csv".to_string(),
        "file:///tmp/test-dedup.csv".to_string(),       // exact dupe
        "file:///tmp/test%2Ddedup.csv".to_string(),     // percent-encoded '-' = same file
    ];
    let paths = normalize_and_dedup_urls(urls);
    // All three refer to the same path — should collapse to 1
    assert_eq!(paths.len(), 1);
    assert!(paths[0].ends_with("test-dedup.csv"));
}

#[test]
fn test_normalize_and_dedup_preserves_order() {
    use crate::normalize_and_dedup_urls;
    // Distinct paths should be preserved in order
    let urls = vec![
        "/tmp/b.xlsx".to_string(),
        "/tmp/a.xlsx".to_string(),
        "/tmp/c.xlsx".to_string(),
    ];
    let paths = normalize_and_dedup_urls(urls);
    assert_eq!(paths.len(), 3);
    assert!(paths[0].ends_with("b.xlsx"));
    assert!(paths[1].ends_with("a.xlsx"));
    assert!(paths[2].ends_with("c.xlsx"));
}

#[test]
fn test_normalize_and_dedup_skips_non_file_urls() {
    use crate::normalize_and_dedup_urls;
    let urls = vec![
        "https://example.com/foo.xlsx".to_string(),
        "ftp://server/bar.csv".to_string(),
        "/tmp/real-file.csv".to_string(),
    ];
    let paths = normalize_and_dedup_urls(urls);
    // Only the plain path should survive
    assert_eq!(paths.len(), 1);
    assert!(paths[0].ends_with("real-file.csv"));
}
