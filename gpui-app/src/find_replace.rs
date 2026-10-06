//! Find and Replace functionality for Spreadsheet.
//!
//! This module contains:
//! - MatchKind and MatchHit types for search results
//! - Find/Replace dialog control methods
//! - Search and replace operations with formula-aware matching
use crate::app::Spreadsheet;
use gpui::*;
#[path = "find_replace_plan.rs"]
mod plan;
pub use plan::MatchHit;

impl Spreadsheet {
    // =========================================================================
    // Find and Replace
    // =========================================================================

    /// Show Find dialog (Ctrl+F)
    /// If already in Find mode, collapses to Find-only (hides Replace row)
    pub fn show_find(&mut self, cx: &mut Context<Self>) {
        use crate::mode::Mode;

        if self.mode == Mode::Find {
            // Already open: collapse to Find-only mode, preserve inputs
            self.find_replace_mode = false;
            self.find_focus_replace = false;
            cx.notify();
            return;
        }

        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );

        // Fresh open: clear state
        self.status_message = None;
        self.lua_console.visible = false;
        self.tab_chain_origin_col = None; // Dialog breaks tab chain
        self.mode = Mode::Find;
        self.find_input.clear();
        self.replace_input.clear();
        self.find_results.clear();
        self.find_index = 0;
        self.find_replace_mode = false;
        self.find_focus_replace = false;
        cx.notify();
    }

    /// Show Find and Replace dialog (Ctrl+H)
    /// If already in Find mode, expands to show Replace row
    pub fn show_find_replace(&mut self, cx: &mut Context<Self>) {
        use crate::mode::Mode;

        if self.mode == Mode::Find {
            // Already open: expand to Replace mode, preserve inputs
            self.find_replace_mode = true;
            // Focus Replace field if Find field has content, else stay on Find
            if !self.find_input.is_empty() {
                self.find_focus_replace = true;
            }
            cx.notify();
            return;
        }

        // Fresh open: clear state
        self.status_message = None;
        self.lua_console.visible = false;
        self.tab_chain_origin_col = None; // Dialog breaks tab chain
        self.mode = Mode::Find;
        self.find_input.clear();
        self.replace_input.clear();
        self.find_results.clear();
        self.find_index = 0;
        self.find_replace_mode = true;
        self.find_focus_replace = false;
        cx.notify();
    }

    pub fn hide_find(&mut self, cx: &mut Context<Self>) {
        use crate::mode::Mode;
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Toggle focus between find and replace input fields
    pub fn find_toggle_focus(&mut self, cx: &mut Context<Self>) {
        if self.find_replace_mode {
            self.find_focus_replace = !self.find_focus_replace;
            cx.notify();
        }
    }

    pub fn find_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        use crate::mode::Mode;

        if self.mode == Mode::Find {
            if self.find_focus_replace {
                self.replace_input.push(c);
            } else {
                self.find_input.push(c);
                self.perform_find(cx);
            }
            cx.notify();
        }
    }

    pub fn find_backspace(&mut self, cx: &mut Context<Self>) {
        use crate::mode::Mode;

        if self.mode == Mode::Find {
            if self.find_focus_replace {
                self.replace_input.pop();
            } else {
                self.find_input.pop();
                self.perform_find(cx);
            }
            cx.notify();
        }
    }

    /// Search only displayed cells; hits retain canonical cell identities.
    pub(crate) fn perform_find(&mut self, cx: &mut Context<Self>) {
        self.sync_table_view(cx);
        self.find_index = 0;
        match plan::search(
            self.sheet(cx),
            self.sheet_index(cx),
            &self.row_view,
            self.display_hidden_rows(),
            self.display_hidden_cols(),
            &self.find_input,
        ) {
            Ok(hits) => {
                self.find_results = hits;
                self.status_message = if self.find_input.is_empty() {
                    None
                } else if self.find_results.is_empty() {
                    Some("No matches in visible cells".into())
                } else {
                    Some(format!(
                        "Found {} matches in visible cells",
                        self.find_results.len()
                    ))
                };
                self.jump_to_find_result(cx);
            }
            Err(error) => {
                self.find_results.clear();
                self.status_message = Some(error);
            }
        }
        cx.notify();
    }

    pub fn find_next(&mut self, cx: &mut Context<Self>) {
        if self.find_results.is_empty() {
            return;
        }
        self.find_index = (self.find_index + 1) % self.find_results.len();
        self.jump_to_find_result(cx);
    }

    pub fn find_prev(&mut self, cx: &mut Context<Self>) {
        if self.find_results.is_empty() {
            return;
        }
        self.find_index = (self.find_index + self.find_results.len() - 1) % self.find_results.len();
        self.jump_to_find_result(cx);
    }

    fn jump_to_find_result(&mut self, cx: &mut Context<Self>) {
        self.sync_table_view(cx);
        let Some(hit) = self.find_results.get(self.find_index) else {
            return;
        };
        if hit.sheet_id != self.sheet(cx).id {
            self.status_message = Some("The search sheet changed. Search again.".into());
            cx.notify();
            return;
        }
        let Some(row) = plan::visible_row(&self.row_view, self.display_hidden_rows(), hit.row)
            .filter(|_| !self.is_col_hidden(hit.col))
        else {
            self.status_message =
                Some("This match is now hidden. Search again to update the results.".into());
            cx.notify();
            return;
        };
        self.view_state.select_cell(row, hit.col);
        self.ensure_visible(cx);
        self.status_message = Some(format!(
            "Match {} of {} in visible cells",
            self.find_index + 1,
            self.find_results.len()
        ));
        cx.notify();
    }

    /// Replace exactly one occurrence, with one guarded undo entry.
    pub fn replace_next(&mut self, cx: &mut Context<Self>) {
        if !self.find_replace_mode {
            self.find_next(cx);
            return;
        }
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        self.sync_table_view(cx);
        let Some(hit) = self.find_results.get(self.find_index).cloned() else {
            return;
        };
        if hit.kind.is_none() {
            if self.find_results.iter().any(|h| h.kind.is_some()) {
                self.find_next(cx);
            } else {
                self.status_message = Some("No replaceable matches".into());
                cx.notify();
            }
            return;
        }
        let old_slot = self
            .row_view
            .data_to_view(hit.row)
            .unwrap_or(self.view_state.selected.0);
        let Some(count) = self.apply_find_replacements(std::slice::from_ref(&hit), cx) else {
            return;
        };
        if count == 0 {
            self.find_next(cx);
            return;
        }
        // Resume after the inserted text, using the record's new displayed
        // position when the edit changed its sorting/filter membership.
        let visible = plan::visible_row(&self.row_view, self.display_hidden_rows(), hit.row);
        let anchor = (
            visible.unwrap_or(old_slot),
            hit.col,
            if visible.is_some() {
                hit.start + self.replace_input.len()
            } else {
                0
            },
        );
        self.perform_find(cx);
        self.find_index = self
            .find_results
            .iter()
            .position(|h| (self.row_view.data_to_view_unchecked(h.row), h.col, h.start) >= anchor)
            .unwrap_or(0);
        self.jump_to_find_result(cx);
    }

    /// Prepare every visible match before writing any; one failure refuses all.
    pub fn replace_all(&mut self, cx: &mut Context<Self>) {
        if (self.cloud_live_enabled() && self.block_if_previewing(cx)) || self.block_if_previewing_only(cx) {
            return;
        }
        if !self.find_replace_mode || self.find_results.is_empty() {
            return;
        }
        self.sync_table_view(cx);
        let hits = self.find_results.clone();
        if self
            .apply_find_replacements(&hits, cx)
            .is_some_and(|count| count > 0)
        {
            let message = self.status_message.clone();
            self.perform_find(cx);
            self.status_message = message;
            cx.notify();
        }
    }

    fn apply_find_replacements(
        &mut self,
        hits: &[MatchHit],
        cx: &mut Context<Self>,
    ) -> Option<usize> {
        let (writes, count) = match plan::replacements(
            self.wb(cx),
            self.sheet_index(cx),
            &self.row_view,
            self.display_hidden_rows(),
            self.display_hidden_cols(),
            hits,
            &self.replace_input,
        ) {
            Ok(plan) => plan,
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                return None;
            }
        };
        if writes.is_empty() {
            self.status_message = Some("No changes to make in the visible matches".into());
            cx.notify();
            return Some(0);
        }
        if !self.apply_table_cell_writes(writes, "Find and Replace", cx) {
            return None;
        }
        self.find_results.clear();
        self.find_index = 0;
        self.status_message = Some(format!(
            "Replaced {count} visible match{}",
            if count == 1 { "" } else { "es" }
        ));
        cx.notify();
        Some(count)
    }
}

#[cfg(test)]
#[path = "find_replace_tests.rs"]
mod tests;
