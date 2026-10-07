use visigrid_engine::{custom_fns, formula::eval::{EvalArg, EvalResult}, table::TableRange, workbook::Workbook};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Mutex;
static CALLS: AtomicUsize = AtomicUsize::new(0);
static HANDLER: Mutex<()> = Mutex::new(());
fn alternating(name: &str, _: &[EvalArg]) -> Option<EvalResult> {
    (name == "RESIZEADDRESS").then(|| EvalResult::Text(if CALLS.fetch_add(1, Ordering::SeqCst) % 2 == 0 { "F1" } else { "F2" }.into()))
}
#[test]
fn resize_refuses_unsettled_indirect_without_publishing_candidate() {
    let _handler = HANDLER.lock().unwrap();
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0,0,0,"Amount"); wb.set_cell_value_tracked(0,1,0,"10");
    let id = wb.create_table(wb.active_sheet_id(), TableRange {start_row:0,start_col:0,end_row:1,end_col:0}, "Sales").unwrap().table_id();
    wb.set_cell_value_tracked(0,0,2,"=A3");
    wb.set_cell_value_tracked(0,0,5,"1"); wb.set_cell_value_tracked(0,1,5,"2");
    struct Reset(Option<custom_fns::CustomFnHandler>);
    impl Drop for Reset { fn drop(&mut self) { custom_fns::set_default_custom_fn_handler(self.0); } }
    let _reset = Reset(custom_fns::default_custom_fn_handler());
    custom_fns::set_default_custom_fn_handler(Some(alternating));
    wb.set_cell_value_tracked(0,0,4,"=INDIRECT(RESIZEADDRESS())");
    assert!(wb.recompute_full_ordered().errors.iter().any(|e|e.error.contains("not settled")));
    wb.take_incremental_errors();
    let before = format!("{wb:?}");
    let error = wb.resize_table(id,TableRange {start_row:0,start_col:0,end_row:2,end_col:0}).unwrap_err();
    assert!(error.contains("not settled"), "{error}");
    assert_eq!(format!("{wb:?}"),before);
}

#[test]
fn stale_not_settled_error_does_not_block_a_later_resize_after_indirect_is_fixed() {
    let _handler = HANDLER.lock().unwrap();
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0,0,0,"Amount"); wb.set_cell_value_tracked(0,1,0,"10");
    let id = wb.create_table(wb.active_sheet_id(), TableRange {start_row:0,start_col:0,end_row:1,end_col:0}, "Sales").unwrap().table_id();
    wb.set_cell_value_tracked(0,0,5,"1"); wb.set_cell_value_tracked(0,1,5,"2");
    struct Reset(Option<custom_fns::CustomFnHandler>);
    impl Drop for Reset { fn drop(&mut self) { custom_fns::set_default_custom_fn_handler(self.0); } }
    let _reset = Reset(custom_fns::default_custom_fn_handler());
    custom_fns::set_default_custom_fn_handler(Some(alternating));
    wb.set_cell_value_tracked(0,0,4,"=INDIRECT(RESIZEADDRESS())");
    assert!(wb.recompute_full_ordered().errors.iter().any(|e| e.error.contains("not settled")));
    assert!(!wb.take_incremental_errors().is_empty(), "the unsettled pass must leave an error behind");
    // Put the error back: take_incremental_errors emptied it, and the bug is
    // a leftover that the next edit does not clear.
    wb.set_cell_value_tracked(0,0,4,"=INDIRECT(RESIZEADDRESS())");
    wb.set_cell_value_tracked(0,0,4,"1");
    wb.resize_table(id, TableRange {start_row:0,start_col:0,end_row:2,end_col:0}).expect("a fixed sheet must resize");
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 2);
}
