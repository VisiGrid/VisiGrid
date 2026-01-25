//! Create Named Range (Ctrl+Shift+N) - create named ranges from selection

use gpui::*;
use visigrid_engine::named_range::is_valid_name;
use crate::app::{Spreadsheet, CreateNameFocus, col_to_letter};
use crate::mode::Mode;

impl Spreadsheet {
    // =========================================================================
    // Create Named Range (Ctrl+Shift+N)
    // =========================================================================

    /// Show the create named range dialog
    pub fn show_create_named_range(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        // Build target string from current selection
        let target = self.selection_to_reference_string();

        self.create_name_name = String::new();
        self.create_name_description = String::new();
        self.create_name_target = target;
        self.create_name_validation_error = None;
        self.create_name_focus = CreateNameFocus::Name;
        self.mode = Mode::CreateNamedRange;
        cx.notify();
    }

    /// Hide the create named range dialog
    pub fn hide_create_named_range(&mut self, cx: &mut Context<Self>) {
        self.create_name_name.clear();
        self.create_name_description.clear();
        self.create_name_target.clear();
        self.create_name_validation_error = None;
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Insert a character into the currently focused create name field
    pub fn create_name_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        match self.create_name_focus {
            CreateNameFocus::Name => self.create_name_name.push(c),
            CreateNameFocus::Description => self.create_name_description.push(c),
        }
        self.validate_create_name();
        cx.notify();
    }

    /// Backspace in the currently focused create name field
    pub fn create_name_backspace(&mut self, cx: &mut Context<Self>) {
        match self.create_name_focus {
            CreateNameFocus::Name => { self.create_name_name.pop(); }
            CreateNameFocus::Description => { self.create_name_description.pop(); }
        }
        self.validate_create_name();
        cx.notify();
    }

    /// Tab to next field in create named range dialog
    pub fn create_name_tab(&mut self, cx: &mut Context<Self>) {
        self.create_name_focus = match self.create_name_focus {
            CreateNameFocus::Name => CreateNameFocus::Description,
            CreateNameFocus::Description => CreateNameFocus::Name,
        };
        cx.notify();
    }

    /// Validate the name field
    fn validate_create_name(&mut self) {
        if self.create_name_name.is_empty() {
            self.create_name_validation_error = Some("Name is required".into());
            return;
        }

        if let Err(e) = is_valid_name(&self.create_name_name) {
            self.create_name_validation_error = Some(e);
            return;
        }

        // Check if name already exists
        if self.workbook.get_named_range(&self.create_name_name).is_some() {
            self.create_name_validation_error = Some(format!(
                "'{}' already exists",
                self.create_name_name
            ));
            return;
        }

        self.create_name_validation_error = None;
    }

    /// Confirm creation of the named range
    pub fn confirm_create_named_range(&mut self, cx: &mut Context<Self>) {
        // Validate first
        self.validate_create_name();
        if self.create_name_validation_error.is_some() {
            return;
        }

        let name = self.create_name_name.clone();
        let description = if self.create_name_description.is_empty() {
            None
        } else {
            Some(self.create_name_description.clone())
        };

        // Parse the selection and create the named range
        let (anchor_row, anchor_col) = self.view_state.selected;
        let (end_row, end_col) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
        let (start_row, start_col, end_row, end_col) = (
            anchor_row.min(end_row),
            anchor_col.min(end_col),
            anchor_row.max(end_row),
            anchor_col.max(end_col),
        );
        let sheet = self.workbook.active_sheet_index();

        let result = if start_row == end_row && start_col == end_col {
            // Single cell
            self.workbook.define_name_for_cell(&name, sheet, start_row, start_col)
        } else {
            // Range
            self.workbook.define_name_for_range(
                &name, sheet, start_row, start_col, end_row, end_col
            )
        };

        match result {
            Ok(()) => {
                // Add description if provided
                if let Some(desc) = description {
                    if let Some(nr) = self.workbook.named_ranges_mut().get(&name).cloned() {
                        let mut updated = nr;
                        updated.description = Some(desc);
                        let _ = self.workbook.named_ranges_mut().set(updated);
                    }
                }

                self.is_modified = true;

                // Log the creation
                let target = self.create_name_target.clone();
                self.log_refactor(
                    "Created named range",
                    &format!("{} → {}", name, target),
                    None,
                );

                self.status_message = Some(format!(
                    "Created named range '{}' → {}",
                    name,
                    self.create_name_target
                ));
                self.hide_create_named_range(cx);
            }
            Err(e) => {
                self.create_name_validation_error = Some(e);
                cx.notify();
            }
        }
    }

    /// Convert current selection to a reference string (e.g., "A1" or "A1:B10")
    pub(crate) fn selection_to_reference_string(&self) -> String {
        let (anchor_row, anchor_col) = self.view_state.selected;
        let (end_row, end_col) = self.view_state.selection_end.unwrap_or(self.view_state.selected);
        let (start_row, start_col, end_row, end_col) = (
            anchor_row.min(end_row),
            anchor_col.min(end_col),
            anchor_row.max(end_row),
            anchor_col.max(end_col),
        );

        let start_ref = format!("{}{}", col_to_letter(start_col), start_row + 1);

        if start_row == end_row && start_col == end_col {
            start_ref
        } else {
            format!("{}:{}{}", start_ref, col_to_letter(end_col), end_row + 1)
        }
    }
}
