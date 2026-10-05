//! Named Ranges Panel Actions - delete, jump, filter, and usage tracking

use gpui::{*};
use crate::app::Spreadsheet;
use visigrid_engine::workbook::NamedRangeEdit;

impl Spreadsheet {
    // =========================================================================
    // Named Ranges Panel Actions
    // =========================================================================

    /// Delete a named range by name (shows impact preview first)
    pub fn delete_named_range(&mut self, name: &str, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) { return; }
        // Check if named range exists
        if self.wb(cx).get_named_range(name).is_none() {
            self.status_message = Some(format!("Named range '{}' not found", name));
            cx.notify();
            return;
        }

        // Show impact preview instead of deleting directly
        self.show_impact_preview_for_delete(name, cx);
    }

    /// Internal method to delete a named range (called from impact preview)
    pub(crate) fn delete_named_range_internal(&mut self, name: &str, _usage_count: usize, cx: &mut Context<Self>) -> bool {
        let before = match self.named_range_draft(cx) {
            Ok(before) if before.name.eq_ignore_ascii_case(name) => before,
            _ => { let error = "The named range changed. Reopen the preview and try again.".to_string(); self.name_draft_error = Some(error.clone()); self.status_message = Some(error); cx.notify(); return false; }
        };
        if !self.apply_named_range_edit(NamedRangeEdit::Delete(before), format!("Delete named range: {name}"), cx) { return false; }
        self.log_refactor("Deleted named range", name, Some("Workbook references recalculated"));
        true
    }

    /// Count parsed references across the workbook, including Table rules.
    pub fn get_named_range_usage_count(&mut self, name: &str, cx: &App) -> usize {
        if self.named_range_usage_cache.cached_rev != self.cells_rev {
            self.rebuild_named_range_usage_cache(cx);
        }
        self.named_range_usage_cache.counts.get(&name.to_lowercase()).copied().unwrap_or(0)
    }

    fn rebuild_named_range_usage_cache(&mut self, cx: &App) {
        self.named_range_usage_cache.counts = self.wb(cx).list_named_ranges().into_iter()
            .map(|name| (name.name.to_lowercase(), self.wb(cx).named_range_usages(&name.name).len())).collect();
        self.named_range_usage_cache.cached_rev = self.cells_rev;
    }

    /// Jump to a named range definition and select the whole range
    pub fn jump_to_named_range(&mut self, name: &str, cx: &mut Context<Self>) {
        use visigrid_engine::named_range::NamedRangeTarget;

        let target_info = self.wb(cx).get_named_range(name).map(|nr| {
            match &nr.target {
                NamedRangeTarget::Cell { sheet, row, col } => {
                    (*sheet, *row, *col, *row, *col, nr.reference_string())
                }
                NamedRangeTarget::Range { sheet, start_row, start_col, end_row, end_col } => {
                    (*sheet, *start_row, *start_col, *end_row, *end_col, nr.reference_string())
                }
            }
        });

        if let Some((sheet_idx, start_row, start_col, end_row, end_col, ref_str)) = target_info {
            // Switch to target sheet if different
            let current_sheet = self.sheet_index(cx);
            if sheet_idx != current_sheet {
                if !self.activate_sheet(sheet_idx, cx) {
                    return;
                }
            }

            self.sync_table_view(cx);
            let ranges = super::plan::project_named_range(&self.row_view, (start_row, start_col), (end_row, end_col));
            let Some(&(start, end)) = ranges.first() else {
                self.status_message = Some(format!("'{name}' = {ref_str} is hidden by the current view."));
                cx.notify(); return;
            };
            self.view_state.selected = start;
            self.view_state.selection_end = (start != end).then_some(end);
            self.view_state.additional_selections = ranges.into_iter().skip(1).map(|(a,b)| (a, (a != b).then_some(b))).collect();
            self.ensure_cell_visible(start.0, start.1);
            self.status_message = Some(format!("'{name}' = {ref_str} (visible cells selected)"));
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
    pub fn filtered_named_ranges(&self, cx: &App) -> Vec<visigrid_engine::named_range::NamedRange> {
        let query = self.names_filter_query.to_lowercase();
        let mut ranges: Vec<_> = self.wb(cx).list_named_ranges()
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
            .cloned()
            .collect();

        // Sort alphabetically by name
        ranges.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()));
        ranges
    }

    /// Trace a named range's dependencies (highlight cells and their precedents in grid)
    pub fn trace_named_range(&mut self, name: &str, cx: &mut Context<Self>) {
        use visigrid_engine::named_range::NamedRangeTarget;
        use visigrid_engine::cell_id::CellId;

        let range_info = self.wb(cx).get_named_range(name).map(|nr| {
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
            let sheet_id = self.wb(cx).sheets().get(sheet_index)
                .map(|s| s.id)
                .unwrap_or_else(|| self.sheet(cx).id);

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

                let raw = self.wb(cx).sheet(sheet_index).map(|s| s.get_raw(*row, *col)).unwrap_or_default();
                if raw.starts_with('=') {
                    // Get precedents from dependency graph
                    let precedents = self.wb(cx).get_precedents(sheet_id, *row, *col);
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
