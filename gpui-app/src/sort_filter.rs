//! Sort and filter operations
//!
//! Contains:
//! - Row View Layer (view space <-> data space conversion)
//! - Sort operations (sort by column, clear sort)
//! - AutoFilter (toggle, dropdown, apply filters)
//! - Filter helpers

use gpui::*;
use visigrid_engine::provenance::{MutationOp, SortKey};
use crate::app::Spreadsheet;

impl Spreadsheet {
    // ========================================================================
    // Row View Layer (view space <-> data space conversion)
    // ========================================================================
    // UI uses VIEW space (what user sees after sort/filter)
    // Storage uses DATA space (canonical row numbers)
    // Convert at boundaries only

    /// Convert view row to data row
    #[inline]
    pub fn view_to_data(&self, view_row: usize) -> usize {
        self.row_view.view_to_data(view_row)
    }

    /// Convert data row to view row (None if hidden by filter)
    #[inline]
    pub fn data_to_view(&self, data_row: usize) -> Option<usize> {
        self.row_view.data_to_view(data_row)
    }

    /// Get visible row count
    #[inline]
    pub fn visible_row_count(&self) -> usize {
        self.row_view.visible_count()
    }

    /// Get the nth visible row (view_row, data_row) for rendering
    /// Returns None if index is out of bounds
    #[inline]
    pub fn nth_visible_row(&self, visible_index: usize) -> Option<(usize, usize)> {
        let view_row = self.row_view.nth_visible(visible_index)?;
        let data_row = self.row_view.view_to_data(view_row);
        Some((view_row, data_row))
    }

    /// View row indices that are visible after filtering
    /// (Not to be confused with visible_rows() which returns screen row count)
    #[inline]
    pub fn filtered_row_indices(&self) -> &[usize] {
        self.row_view.visible_rows()
    }

    /// Ensure row_view has enough capacity for current sheet
    pub fn ensure_row_view_capacity(&mut self) {
        // For now, use a large default. Later this can track actual data extent.
        let needed = 100000;
        if self.row_view.row_count() < needed {
            self.row_view.resize(needed);
        }
    }

    // ========================================================================
    // Sort Operations
    // ========================================================================

    /// Sort by the column containing the active cell
    ///
    /// This is the ONE command UI calls. Engine owns the permutation.
    pub fn sort_by_current_column(
        &mut self,
        direction: visigrid_engine::filter::SortDirection,
        cx: &mut Context<Self>,
    ) {
        use visigrid_engine::filter::{sort_by_column, SortState};

        // Ensure filter range is set (use current selection if not)
        if self.filter_state.filter_range.is_none() {
            // Auto-detect range: from row 0 to last non-empty row in current column
            let col = self.view_state.selected.1;
            let max_row = self.find_last_data_row(col);
            if max_row == 0 {
                self.status_message = Some("No data to sort".to_string());
                cx.notify();
                return;
            }
            // Set filter range covering the data extent
            self.filter_state.filter_range = Some((0, col, max_row, col));
        }

        let col = self.view_state.selected.1;

        // Create value_at closure (captures sheet for computed values)
        let sheet = self.sheet();
        let value_at = |data_row: usize, c: usize| -> visigrid_engine::formula::eval::Value {
            sheet.get_computed_value(data_row, c)
        };

        // Call engine's sort function
        let (new_order, undo_item) = sort_by_column(
            &self.row_view,
            &self.filter_state,
            value_at,
            col,
            direction,
        );

        // Record undo for sort operation
        let previous_sort_state = undo_item.previous_sort_state.map(|s| {
            (s.column, s.direction == visigrid_engine::filter::SortDirection::Ascending)
        });
        let is_ascending = direction == visigrid_engine::filter::SortDirection::Ascending;

        // Build provenance
        let provenance = if let Some((start_row, start_col, end_row, end_col)) = self.filter_state.filter_range {
            Some(MutationOp::Sort {
                sheet: self.sheet().id,
                range_start_row: start_row,
                range_start_col: start_col,
                range_end_row: end_row,
                range_end_col: end_col,
                keys: vec![SortKey { col, ascending: is_ascending }],
                has_header: false,  // TODO: detect header row
            }.to_provenance(&self.sheet().name))
        } else {
            None
        };

        self.history.record_action_with_provenance(crate::history::UndoAction::SortApplied {
            previous_row_order: undo_item.previous_row_order,
            previous_sort_state,
            new_row_order: new_order.clone(),
            new_sort_state: (col, is_ascending),
        }, provenance);

        // Apply the sort
        self.row_view.apply_sort(new_order);

        // Update filter_state.sort
        self.filter_state.sort = Some(SortState { column: col, direction });

        // Invalidate caches
        self.filter_state.invalidate_all_caches();

        self.is_modified = true;
        self.status_message = Some(format!(
            "Sorted by column {} {}",
            Self::col_to_letter(col),
            if direction == visigrid_engine::filter::SortDirection::Ascending { "A→Z" } else { "Z→A" }
        ));
        cx.notify();
    }

    /// Find the last non-empty row in a column (for auto-detecting sort range)
    fn find_last_data_row(&self, col: usize) -> usize {
        let sheet = self.sheet();
        let mut last_row = 0;
        // Scan up to a reasonable limit (or could use sheet's actual data extent)
        for row in 0..10000 {
            let cell = sheet.get_cell(row, col);
            if !cell.value.raw_display().is_empty() {
                last_row = row;
            }
        }
        last_row
    }

    /// Toggle AutoFilter on/off for current selection
    pub fn toggle_auto_filter(&mut self, cx: &mut Context<Self>) {
        if self.filter_state.is_enabled() {
            // Disable: restore original order, clear filters
            self.row_view.clear_sort();
            self.row_view.clear_filter();
            self.filter_state.disable();
            self.status_message = Some("AutoFilter disabled".to_string());
        } else {
            // Enable: set filter range based on selection or data region
            let (row, col) = self.view_state.selected;
            let max_row = self.find_last_data_row(col);
            let max_col = self.find_last_data_col(row);

            if max_row == 0 && max_col == 0 {
                self.status_message = Some("No data for AutoFilter".to_string());
                cx.notify();
                return;
            }

            // Filter range: row 0 is header, data starts at row 1
            self.filter_state.filter_range = Some((0, 0, max_row, max_col));
            self.status_message = Some(format!(
                "AutoFilter enabled: A1:{}{}",
                Self::col_to_letter(max_col),
                max_row + 1
            ));
        }
        cx.notify();
    }

    /// Open the filter dropdown for a column
    pub fn open_filter_dropdown(&mut self, col: usize, cx: &mut Context<Self>) {
        if !self.filter_state.is_enabled() {
            return;
        }

        // Collect values for the column first (to avoid borrow conflicts)
        let values: Vec<(usize, visigrid_engine::formula::eval::Value)> = if let Some((data_start, _, data_end, _)) = self.filter_state.data_range() {
            let sheet = self.sheet();
            (data_start..=data_end)
                .map(|data_row| (data_row, sheet.get_computed_value(data_row, col)))
                .collect()
        } else {
            Vec::new()
        };

        // Build unique values cache (max 500 unique values, sorted by frequency)
        self.filter_state.build_unique_values_from_vec(col, &values, 500);

        // Initialize checked items: all checked if no filter, or match current filter
        self.filter_checked_items.clear();
        if let Some(unique_vals) = self.filter_state.get_unique_values(col) {
            let col_filter = self.filter_state.column_filters.get(&col);
            for (idx, entry) in unique_vals.iter().enumerate() {
                // If no filter for this column, all items are checked
                // If filter exists, check if this value is in selected set
                let should_check = match col_filter {
                    None => true,
                    Some(cf) => match &cf.selected {
                        None => true, // No selection = all pass
                        Some(selected) => selected.contains(&entry.key),
                    },
                };
                if should_check {
                    self.filter_checked_items.insert(idx);
                }
            }
        }

        self.filter_dropdown_col = Some(col);
        self.filter_search_text.clear();
        cx.notify();
    }

    /// Close the filter dropdown without applying
    pub fn close_filter_dropdown(&mut self, cx: &mut Context<Self>) {
        self.filter_dropdown_col = None;
        self.filter_search_text.clear();
        self.filter_checked_items.clear();
        cx.notify();
    }

    /// Toggle a value in the filter dropdown
    pub fn toggle_filter_item(&mut self, idx: usize, cx: &mut Context<Self>) {
        if self.filter_checked_items.contains(&idx) {
            self.filter_checked_items.remove(&idx);
        } else {
            self.filter_checked_items.insert(idx);
        }
        cx.notify();
    }

    /// Select all items in filter dropdown
    pub fn filter_select_all(&mut self, cx: &mut Context<Self>) {
        let Some(col) = self.filter_dropdown_col else { return };
        if let Some(unique_vals) = self.filter_state.get_unique_values(col) {
            self.filter_checked_items.clear();
            for idx in 0..unique_vals.len() {
                self.filter_checked_items.insert(idx);
            }
        }
        cx.notify();
    }

    /// Clear all items in filter dropdown
    pub fn filter_clear_all(&mut self, cx: &mut Context<Self>) {
        self.filter_checked_items.clear();
        cx.notify();
    }

    /// Apply the current filter dropdown selection
    pub fn apply_filter_dropdown(&mut self, cx: &mut Context<Self>) {
        let Some(col) = self.filter_dropdown_col else { return };
        let Some(unique_vals) = self.filter_state.get_unique_values(col) else {
            self.close_filter_dropdown(cx);
            return;
        };

        // Build selected set from checked items
        let all_checked = self.filter_checked_items.len() == unique_vals.len();

        if all_checked {
            // All checked = no filter (remove filter for this column)
            self.filter_state.clear_column_filter(col);
        } else {
            // Build HashSet of selected normalized keys
            let selected: std::collections::HashSet<_> = self
                .filter_checked_items
                .iter()
                .filter_map(|&idx| unique_vals.get(idx).map(|e| e.key.clone()))
                .collect();

            self.filter_state.column_filters.insert(
                col,
                visigrid_engine::filter::ColumnFilter {
                    selected: Some(selected),
                    text_filter: None,
                },
            );
        }

        // Apply filters to row_view
        self.apply_all_filters();

        self.filter_dropdown_col = None;
        self.filter_search_text.clear();
        self.filter_checked_items.clear();
        self.is_modified = true;
        cx.notify();
    }

    /// Apply all column filters to update visible_mask
    fn apply_all_filters(&mut self) {
        let Some((data_start, min_col, data_end, max_col)) = self.filter_state.data_range() else {
            // No filter range - all visible
            self.row_view.clear_filter();
            return;
        };

        // Build visible_mask for all data rows
        let row_count = self.row_view.row_count();
        let mut visible_mask = vec![true; row_count];

        // Header row always visible
        if let Some(header) = self.filter_state.header_row() {
            visible_mask[header] = true;
        }

        // Check each data row against all column filters
        for data_row in data_start..=data_end {
            if data_row >= row_count {
                break;
            }

            let mut passes = true;
            for col in min_col..=max_col {
                if let Some(col_filter) = self.filter_state.column_filters.get(&col) {
                    if col_filter.is_active() {
                        let value = self.sheet().get_computed_value(data_row, col);
                        let filter_key = visigrid_engine::filter::FilterKey::from_value(&value);
                        if !col_filter.passes(&filter_key) {
                            passes = false;
                            break;
                        }
                    }
                }
            }
            visible_mask[data_row] = passes;
        }

        self.row_view.apply_filter(visible_mask);
    }

    /// Check if a column has an active filter
    pub fn column_has_filter(&self, col: usize) -> bool {
        self.filter_state
            .column_filters
            .get(&col)
            .map_or(false, |f| f.is_active())
    }

    /// Find the last non-empty column in a row
    fn find_last_data_col(&self, row: usize) -> usize {
        let sheet = self.sheet();
        let mut last_col = 0;
        for col in 0..256 {
            let cell = sheet.get_cell(row, col);
            if !cell.value.raw_display().is_empty() {
                last_col = col;
            }
        }
        last_col
    }

    /// Clear sort (restore original data order)
    pub fn clear_sort(&mut self, cx: &mut Context<Self>) {
        self.row_view.clear_sort();
        self.filter_state.sort = None;
        self.is_modified = true;
        self.status_message = Some("Sort cleared".to_string());
        cx.notify();
    }
}
