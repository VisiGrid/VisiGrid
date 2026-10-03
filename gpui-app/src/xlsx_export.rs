//! Review state shared by the Excel order chooser and its final write check.
use visigrid_engine::workbook::Workbook;
use visigrid_io::xlsx::{self, ExportLayout, ExportOrder};

pub(crate) struct ExportReview {
    pub has_sort: bool,
    pub sorted: Result<Vec<String>, String>,
    stored: Vec<String>,
    order: ExportOrder,
}
impl ExportReview {
    pub fn new(wb: &Workbook, layouts: &[ExportLayout]) -> Result<Self, String> {
        // Stored order must still pass recovery, schema and writer checks.
        let stored =
            xlsx::table_export_warnings_with_order(wb, Some(layouts), ExportOrder::Stored)?;
        let sorted = xlsx::table_export_warnings_with_order(wb, Some(layouts), ExportOrder::Sorted);
        let order = if sorted.is_ok() {
            ExportOrder::Sorted
        } else {
            ExportOrder::Stored
        };
        Ok(Self {
            has_sort: wb
                .sheets()
                .iter()
                .any(|s| s.table_view_spec().is_some_and(|v| v.sort.is_some())),
            sorted,
            stored,
            order,
        })
    }
    pub fn order(&self) -> ExportOrder {
        self.order
    }
    pub fn select(&mut self, order: ExportOrder) {
        if order != ExportOrder::Sorted || self.sorted.is_ok() {
            self.order = order;
        }
    }
    pub fn warnings(&self) -> &[String] {
        match self.order {
            ExportOrder::Stored => &self.stored,
            ExportOrder::Sorted => self
                .sorted
                .as_deref()
                .expect("sorted choice passed preflight"),
        }
    }
    pub fn needs_review(&self) -> bool {
        self.has_sort || !self.warnings().is_empty()
    }
    pub fn check_current(&self, wb: &Workbook, layouts: &[ExportLayout]) -> Result<(), String> {
        let warnings = xlsx::table_export_warnings_with_order(wb, Some(layouts), self.order)?;
        if warnings != self.warnings() {
            return Err(
                "The workbook's export details changed. Export again to review them.".into(),
            );
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use visigrid_engine::{
        filter::SortDirection,
        sheet::SheetId,
        table::TableRange,
        table_view::{TableSort, TableViewSpec},
    };
    fn book() -> Workbook {
        let mut wb = Workbook::new();
        for (r, value) in ["Key", "1", "2"].iter().enumerate() {
            wb.set_cell_value_tracked(0, r, 0, value);
        }
        let id = wb
            .create_table(
                SheetId(1),
                TableRange {
                    start_row: 0,
                    end_row: 2,
                    start_col: 0,
                    end_col: 0,
                },
                "Items",
            )
            .unwrap()
            .table_id();
        let mut spec = TableViewSpec::new(id);
        spec.sort = Some(TableSort {
            column: wb.table(id).unwrap().1.columns[0].id,
            direction: SortDirection::Descending,
        });
        wb.set_table_view_spec(SheetId(1), Some(spec)).unwrap();
        wb
    }
    #[test]
    fn choices_review_the_selected_mode_and_retain_the_choice() {
        let wb = book();
        let mut review = ExportReview::new(&wb, &[]).unwrap();
        assert!(review.needs_review());
        assert_eq!(review.order(), ExportOrder::Sorted);
        assert!(review.warnings()[0].contains("saved sort order"));
        review.select(ExportOrder::Stored);
        assert!(review.warnings()[0].contains("use Reapply"));
        review.check_current(&wb, &[]).unwrap();
        assert_eq!(review.order(), ExportOrder::Stored);
        review.select(ExportOrder::Sorted);
        review.check_current(&wb, &[]).unwrap();
    }
    #[test]
    fn unsupported_sorted_export_offers_stored_and_cannot_select_disabled_choice() {
        let mut wb = book();
        wb.set_cell_value_tracked(0, 4, 0, "=ROW(A2)");
        let mut review = ExportReview::new(&wb, &[]).unwrap();
        assert!(review.sorted.as_ref().unwrap_err().contains("Function ROW"));
        assert_eq!(review.order(), ExportOrder::Stored);
        assert!(review.needs_review());
        review.select(ExportOrder::Sorted);
        assert_eq!(review.order(), ExportOrder::Stored);
        review.check_current(&wb, &[]).unwrap();
    }
    #[test]
    fn host_layout_is_checked_before_choosing_an_order() {
        let wb = book();
        let mut layout = ExportLayout::default();
        layout.row_heights.insert(1, 30.0);
        let review = ExportReview::new(&wb, &[layout]).unwrap();
        assert!(review.sorted.is_err());
        assert_eq!(review.order(), ExportOrder::Stored);
    }
    #[test]
    fn changed_warning_details_require_review_again() {
        let mut wb = book();
        let mut review = ExportReview::new(&wb, &[]).unwrap();
        review.select(ExportOrder::Stored);
        wb.define_name_for_cell("FirstKey", 0, 1, 0).unwrap();
        assert!(review
            .check_current(&wb, &[])
            .unwrap_err()
            .contains("details changed"));
    }
    #[test]
    fn sorted_eligibility_is_rechecked_without_silently_switching_modes() {
        let mut wb = book();
        let review = ExportReview::new(&wb, &[]).unwrap();
        wb.set_cell_value_tracked(0, 4, 0, "=ROW(A2)");
        assert!(review
            .check_current(&wb, &[])
            .unwrap_err()
            .contains("Function ROW"));
        assert_eq!(review.order(), ExportOrder::Sorted);
    }
    #[test]
    fn workbooks_without_sort_or_losses_skip_the_order_dialog() {
        let review = ExportReview::new(&Workbook::new(), &[]).unwrap();
        assert!(!review.has_sort);
        assert!(!review.needs_review());
    }
}
