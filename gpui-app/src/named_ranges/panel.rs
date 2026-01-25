//! Named Ranges Panel Actions - delete, jump, filter, and usage tracking

use gpui::*;
use crate::app::Spreadsheet;
use crate::history::UndoAction;

impl Spreadsheet {
    // =========================================================================
    // Named Ranges Panel Actions
    // =========================================================================

    /// Delete a named range by name (shows impact preview first)
    pub fn delete_named_range(&mut self, name: &str, cx: &mut Context<Self>) {
        // Check if named range exists
        if self.workbook.get_named_range(name).is_none() {
            self.status_message = Some(format!("Named range '{}' not found", name));
            cx.notify();
            return;
        }

        // Show impact preview instead of deleting directly
        self.show_impact_preview_for_delete(name, cx);
    }

    /// Internal method to delete a named range (called from impact preview)
    pub(crate) fn delete_named_range_internal(&mut self, name: &str, usage_count: usize, cx: &mut Context<Self>) {
        // Get the named range first (need to clone for undo)
        let named_range = self.workbook.get_named_range(name).cloned();

        if let Some(nr) = named_range {
            // Record undo action BEFORE deleting
            self.history.record_named_range_action(UndoAction::NamedRangeDeleted {
                named_range: nr.clone(),
            });

            // Now delete
            self.workbook.delete_named_range(name);
            self.is_modified = true;
            self.bump_cells_rev();

            // Log the deletion
            let impact = if usage_count > 0 {
                Some(format!("{} formula{} will show #NAME? error", usage_count, if usage_count == 1 { "" } else { "s" }))
            } else {
                None
            };
            self.log_refactor(
                "Deleted named range",
                name,
                impact.as_deref(),
            );

            cx.notify();
        }
    }

    /// Count how many formula cells reference a named range
    fn count_named_range_references(&self, name: &str) -> usize {
        let name_upper = name.to_uppercase();
        let mut count = 0;

        for ((_, _), cell) in self.sheet().cells_iter() {
            let raw = cell.value.raw_display();
            if raw.starts_with('=') {
                // Simple check: does the formula contain this name as a word?
                // More sophisticated: parse the formula and check identifiers
                // For now, do case-insensitive word boundary check
                let formula_upper = raw.to_uppercase();
                // Check if name appears as a standalone identifier
                // This is a simple heuristic - a proper check would parse the formula
                for word in formula_upper.split(|c: char| !c.is_alphanumeric() && c != '_' && c != '.') {
                    if word == name_upper {
                        count += 1;
                        break; // Count each cell only once
                    }
                }
            }
        }
        count
    }

    /// Get usage count for a named range (with caching)
    pub fn get_named_range_usage_count(&mut self, name: &str) -> usize {
        // Check if cache is stale
        if self.named_range_usage_cache.cached_rev != self.cells_rev {
            self.rebuild_named_range_usage_cache();
        }

        // Return cached count (or 0 if not found)
        self.named_range_usage_cache.counts
            .get(&name.to_lowercase())
            .copied()
            .unwrap_or(0)
    }

    /// Rebuild the usage count cache for all named ranges
    fn rebuild_named_range_usage_cache(&mut self) {
        self.named_range_usage_cache.counts.clear();

        // Get all named range names (lowercase for lookup)
        let names: Vec<String> = self.workbook.list_named_ranges()
            .iter()
            .map(|nr| nr.name.to_lowercase())
            .collect();

        // Also store uppercase versions for matching
        let names_upper: Vec<String> = names.iter()
            .map(|n| n.to_uppercase())
            .collect();

        // Initialize all counts to 0
        for name in &names {
            self.named_range_usage_cache.counts.insert(name.clone(), 0);
        }

        // Collect all formulas first (to avoid borrow issues)
        let formulas: Vec<String> = self.sheet().cells_iter()
            .filter_map(|((_, _), cell)| {
                let raw = cell.value.raw_display();
                if raw.starts_with('=') {
                    Some(raw.to_uppercase())
                } else {
                    None
                }
            })
            .collect();

        // Now process formulas and update counts
        for formula_upper in formulas {
            // Check each named range
            for (i, name_upper) in names_upper.iter().enumerate() {
                // Check if name appears as a standalone identifier
                for word in formula_upper.split(|c: char| !c.is_alphanumeric() && c != '_' && c != '.') {
                    if word == name_upper {
                        if let Some(count) = self.named_range_usage_cache.counts.get_mut(&names[i]) {
                            *count += 1;
                        }
                        break; // Count each cell only once per name
                    }
                }
            }
        }

        // Mark cache as fresh
        self.named_range_usage_cache.cached_rev = self.cells_rev;
    }

    /// Jump to a named range definition and select the whole range
    pub fn jump_to_named_range(&mut self, name: &str, cx: &mut Context<Self>) {
        use visigrid_engine::named_range::NamedRangeTarget;

        let target_info = self.workbook.get_named_range(name).map(|nr| {
            match &nr.target {
                NamedRangeTarget::Cell { row, col, .. } => {
                    (*row, *col, *row, *col, nr.reference_string())
                }
                NamedRangeTarget::Range { start_row, start_col, end_row, end_col, .. } => {
                    (*start_row, *start_col, *end_row, *end_col, nr.reference_string())
                }
            }
        });

        if let Some((start_row, start_col, end_row, end_col, ref_str)) = target_info {
            // Select the whole range
            self.view_state.selected = (start_row, start_col);
            if start_row == end_row && start_col == end_col {
                self.view_state.selection_end = None;
            } else {
                self.view_state.selection_end = Some((end_row, end_col));
            }

            // Center the view on the selection
            self.ensure_cell_visible(start_row, start_col);

            self.status_message = Some(format!("'{}' = {}", name, ref_str));
            cx.notify();
        } else {
            self.status_message = Some(format!("Named range '{}' not found", name));
            cx.notify();
        }
    }

    /// Filter named ranges by query (for Names panel search)
    pub fn set_names_filter(&mut self, query: String, cx: &mut Context<Self>) {
        self.names_filter_query = query;
        cx.notify();
    }

    /// Get filtered named ranges for the Names panel
    pub fn filtered_named_ranges(&self) -> Vec<&visigrid_engine::named_range::NamedRange> {
        let query = self.names_filter_query.to_lowercase();
        let mut ranges: Vec<_> = self.workbook.list_named_ranges()
            .into_iter()
            .filter(|nr| {
                if query.is_empty() {
                    return true;
                }
                // Match against name or description
                nr.name.to_lowercase().contains(&query)
                    || nr.description.as_ref()
                        .map(|d| d.to_lowercase().contains(&query))
                        .unwrap_or(false)
            })
            .collect();

        // Sort alphabetically by name
        ranges.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()));
        ranges
    }

    /// Trace a named range's dependencies (highlight cells and their precedents in grid)
    pub fn trace_named_range(&mut self, name: &str, cx: &mut Context<Self>) {
        use visigrid_engine::named_range::NamedRangeTarget;
        use visigrid_engine::cell_id::CellId;

        let range_info = self.workbook.get_named_range(name).map(|nr| {
            let sheet_index = match &nr.target {
                NamedRangeTarget::Cell { sheet, .. } => *sheet,
                NamedRangeTarget::Range { sheet, .. } => *sheet,
            };
            let cells: Vec<(usize, usize)> = match &nr.target {
                NamedRangeTarget::Cell { row, col, .. } => vec![(*row, *col)],
                NamedRangeTarget::Range { start_row, start_col, end_row, end_col, .. } => {
                    let mut cells = Vec::new();
                    for r in *start_row..=*end_row {
                        for c in *start_col..=*end_col {
                            cells.push((r, c));
                        }
                    }
                    cells
                }
            };
            (sheet_index, cells)
        });

        if let Some((sheet_index, cells)) = range_info {
            // Get the sheet ID for CellId construction
            let sheet_id = self.workbook.sheets().get(sheet_index)
                .map(|s| s.id)
                .unwrap_or_else(|| self.sheet().id);

            // Build trace path: cells in the range + their precedents
            let mut trace_cells: Vec<CellId> = cells.iter()
                .map(|(r, c)| CellId::new(sheet_id, *r, *c))
                .collect();

            // Add precedents of each cell (limited to avoid huge traces)
            let max_precedents = 50;
            let mut precedent_count = 0;

            for (row, col) in &cells {
                if precedent_count >= max_precedents {
                    break;
                }

                let raw = self.sheet().get_raw(*row, *col);
                if raw.starts_with('=') {
                    // Get precedents from dependency graph
                    let precedents = self.workbook.get_precedents(sheet_id, *row, *col);
                    for prec in precedents {
                        if !trace_cells.contains(&prec) {
                            trace_cells.push(prec);
                            precedent_count += 1;
                            if precedent_count >= max_precedents {
                                break;
                            }
                        }
                    }
                }
            }

            self.inspector_trace_path = Some(trace_cells);
            self.inspector_trace_incomplete = precedent_count >= max_precedents;
            cx.notify();
        }
    }

    /// Clear trace when named range is deselected
    pub fn clear_named_range_trace(&mut self, cx: &mut Context<Self>) {
        if self.selected_named_range.is_some() {
            self.inspector_trace_path = None;
            self.inspector_trace_incomplete = false;
        }
        cx.notify();
    }
}
