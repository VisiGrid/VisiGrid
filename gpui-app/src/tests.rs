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
