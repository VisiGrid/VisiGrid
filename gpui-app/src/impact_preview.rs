//! Impact Preview and Refactor Log
//!
//! Contains Spreadsheet methods for:
//! - Finding named range usages in formulas
//! - Showing impact preview for rename/delete operations
//! - Applying previewed actions
//! - Refactor log display

use gpui::*;
use crate::app::Spreadsheet;
use crate::mode::Mode;
use crate::settings::{user_settings, update_user_settings, TipId};

impl Spreadsheet {
    // =========================================================================
    // Impact Preview methods
    // =========================================================================

    /// Find all cells that reference a named range
    fn find_named_range_usages(&self, name: &str) -> Vec<crate::views::impact_preview::ImpactedFormula> {
        use crate::views::impact_preview::ImpactedFormula;

        let name_upper = name.to_uppercase();
        let mut usages = Vec::new();

        // Scan all cells for formulas containing the name
        for ((row, col), cell) in self.sheet().cells_iter() {
            let raw = cell.value.raw_display();
            if !raw.starts_with('=') {
                continue;
            }

            let formula_upper = raw.to_uppercase();

            // Check if name appears as a standalone identifier
            let contains_name = formula_upper
                .split(|c: char| !c.is_alphanumeric() && c != '_' && c != '.')
                .any(|word| word == name_upper);

            if contains_name {
                // Format cell reference
                let cell_ref = {
                    let mut col_name = String::new();
                    let mut c = *col;
                    loop {
                        col_name.insert(0, (b'A' + (c % 26) as u8) as char);
                        if c < 26 { break; }
                        c = c / 26 - 1;
                    }
                    format!("{}{}", col_name, *row + 1)
                };

                usages.push(ImpactedFormula {
                    cell_ref,
                    formula: raw.to_string(),
                });
            }
        }

        // Sort by cell reference for consistent display
        usages.sort_by(|a, b| a.cell_ref.cmp(&b.cell_ref));
        usages
    }

    /// Show impact preview for a rename operation
    pub fn show_impact_preview_for_rename(&mut self, old_name: &str, new_name: &str, cx: &mut Context<Self>) {
        use crate::views::impact_preview::ImpactAction;

        let usages = self.find_named_range_usages(old_name);
        self.impact_preview_action = Some(ImpactAction::Rename {
            old_name: old_name.to_string(),
            new_name: new_name.to_string(),
        });
        self.impact_preview_usages = usages;
        self.mode = Mode::ImpactPreview;
        cx.notify();
    }

    /// Show impact preview for a delete operation
    pub fn show_impact_preview_for_delete(&mut self, name: &str, cx: &mut Context<Self>) {
        use crate::views::impact_preview::ImpactAction;

        let usages = self.find_named_range_usages(name);
        self.impact_preview_action = Some(ImpactAction::Delete {
            name: name.to_string(),
        });
        self.impact_preview_usages = usages;
        self.mode = Mode::ImpactPreview;
        cx.notify();
    }

    /// Hide the impact preview modal
    pub fn hide_impact_preview(&mut self, cx: &mut Context<Self>) {
        self.impact_preview_action = None;
        self.impact_preview_usages.clear();
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Apply the previewed action (rename or delete)
    pub fn apply_impact_preview(&mut self, cx: &mut Context<Self>) {
        use crate::views::impact_preview::ImpactAction;

        let action = self.impact_preview_action.take();
        let usage_count = self.impact_preview_usages.len();
        self.impact_preview_usages.clear();

        match action {
            Some(ImpactAction::Rename { old_name, new_name }) => {
                // Perform the rename
                self.apply_rename_internal(&old_name, &new_name, cx);
                self.mode = Mode::Navigation;

                // Show one-time F12 hint after first rename
                if !user_settings(cx).is_tip_dismissed(TipId::RenameF12) {
                    update_user_settings(cx, |settings| {
                        settings.dismiss_tip(TipId::RenameF12);
                    });
                    self.status_message = Some(format!(
                        "Renamed \"{}\" → \"{}\". Tip: Press F12 to jump to this name's definition.",
                        old_name, new_name
                    ));
                } else {
                    self.status_message = Some(if usage_count > 0 {
                        format!("Renamed \"{}\" → \"{}\", updated {} formula{}",
                            old_name, new_name, usage_count, if usage_count == 1 { "" } else { "s" })
                    } else {
                        format!("Renamed \"{}\" → \"{}\"", old_name, new_name)
                    });
                }
            }
            Some(ImpactAction::Delete { name }) => {
                // Perform the delete
                self.delete_named_range_internal(&name, usage_count, cx);
                self.mode = Mode::Navigation;
                self.status_message = Some(if usage_count > 0 {
                    format!("Deleted \"{}\", {} formula{} affected",
                        name, usage_count, if usage_count == 1 { "" } else { "s" })
                } else {
                    format!("Deleted \"{}\"", name)
                });
            }
            None => {
                self.mode = Mode::Navigation;
            }
        }
        cx.notify();
    }

    // =========================================================================
    // Refactor Log methods
    // =========================================================================

    /// Show the refactor log modal
    pub fn show_refactor_log(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::RefactorLog;
        cx.notify();
    }

    /// Hide the refactor log modal
    pub fn hide_refactor_log(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Log a refactor action
    pub fn log_refactor(&mut self, action: &str, details: &str, impact: Option<&str>) {
        use crate::views::refactor_log::RefactorLogEntry;

        let mut entry = RefactorLogEntry::new(action, details);
        if let Some(imp) = impact {
            entry = entry.with_impact(imp);
        }
        self.refactor_log.push(entry);
    }
}
