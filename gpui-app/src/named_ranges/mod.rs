//! Named Ranges module
//!
//! Provides functionality for creating, renaming, extracting, and managing named ranges.
//!
//! Submodules:
//! - `rename`: Rename Symbol (Ctrl+Shift+R) and Edit Description
//! - `extract`: Extract Named Range from formula literals
//! - `create`: Create Named Range (Ctrl+Shift+N)
//! - `panel`: Named Ranges panel actions (delete, jump, filter)

mod rename;
mod extract;
mod create;
mod panel;

use gpui::*;
use crate::app::Spreadsheet;
use crate::mode::Mode;
use crate::settings::{user_settings, update_user_settings, TipId};

impl Spreadsheet {
    // =========================================================================
    // Shared Helpers
    // =========================================================================

    /// Extract formula source from a CellValue if it's a formula
    pub(crate) fn get_formula_source(&self, value: &visigrid_engine::cell::CellValue) -> Option<String> {
        match value {
            visigrid_engine::cell::CellValue::Formula { source, .. } => Some(source.clone()),
            _ => None,
        }
    }

    // =========================================================================
    // Tour Methods
    // =========================================================================

    /// Show the named ranges tour
    pub fn show_tour(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.tour_step = 0;
        self.mode = Mode::Tour;
        cx.notify();
    }

    /// Hide the tour
    pub fn hide_tour(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Go to next tour step
    pub fn tour_next(&mut self, cx: &mut Context<Self>) {
        if self.tour_step < 3 {
            self.tour_step += 1;
            cx.notify();
        }
    }

    /// Go to previous tour step
    pub fn tour_back(&mut self, cx: &mut Context<Self>) {
        if self.tour_step > 0 {
            self.tour_step -= 1;
            cx.notify();
        }
    }

    /// Complete the tour
    pub fn tour_done(&mut self, cx: &mut Context<Self>) {
        self.tour_completed = true;
        self.mode = Mode::Navigation;
        self.status_message = Some("You just refactored a spreadsheet like code.".to_string());
        cx.notify();
    }

    /// Check if the name tooltip should be shown
    pub fn should_show_name_tooltip(&self, cx: &gpui::App) -> bool {
        // Show if: not dismissed, no named ranges exist, has a range selection
        !user_settings(cx).is_tip_dismissed(TipId::NamedRanges)
            && self.wb(cx).list_named_ranges().is_empty()
            && self.view_state.selection_end.is_some()
    }

    /// Dismiss the name tooltip permanently
    pub fn dismiss_name_tooltip(&mut self, cx: &mut Context<Self>) {
        update_user_settings(cx, |settings| {
            settings.dismiss_tip(TipId::NamedRanges);
        });
        cx.notify();
    }
}
