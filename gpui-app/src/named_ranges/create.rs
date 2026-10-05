//! Create Named Range (Ctrl+Shift+N) - create named ranges from selection

use gpui::{*};
use visigrid_engine::named_range::is_valid_name;
use crate::app::{Spreadsheet, CreateNameFocus};
use crate::mode::Mode;

impl Spreadsheet {
    // =========================================================================
    // Create Named Range (Ctrl+Shift+N)
    // =========================================================================

    /// Show the create named range dialog
    pub fn show_create_named_range(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        if self.block_if_previewing_only(cx) { return; }
        self.sync_table_view(cx);
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Select one rectangular range to name.".into());
            cx.notify(); return;
        }
        let range = match super::plan::selection_range(self.wb(cx), &self.row_view,
            self.view_state.selected, self.view_state.selection_end.unwrap_or(self.view_state.selected)) {
            Ok(range) => range,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };
        let target = format!("'{}'!{}", self.sheet(cx).name.replace('\'', "''"), range.reference_string());
        self.name_draft_error = None;
        self.name_draft = Some(super::plan::NameDraft { revision: self.wb(cx).revision(), range });

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
        self.name_draft_error = None;
        self.name_draft = None;
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
        self.validate_create_name(cx);
        cx.notify();
    }

    /// Backspace in the currently focused create name field
    pub fn create_name_backspace(&mut self, cx: &mut Context<Self>) {
        match self.create_name_focus {
            CreateNameFocus::Name => { self.create_name_name.pop(); }
            CreateNameFocus::Description => { self.create_name_description.pop(); }
        }
        self.validate_create_name(cx);
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
    fn validate_create_name(&mut self, cx: &App) {
        if self.create_name_name.is_empty() {
            self.create_name_validation_error = Some("Name is required".into());
            return;
        }

        if let Err(e) = is_valid_name(&self.create_name_name) {
            self.create_name_validation_error = Some(e);
            return;
        }

        // Check if name already exists
        if self.wb(cx).get_named_range(&self.create_name_name).is_some() {
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
        if self.block_if_previewing_only(cx) { return; }
        self.validate_create_name(cx);
        if self.create_name_validation_error.is_some() { return; }
        let mut range = match self.named_range_draft(cx) {
            Ok(range) => range,
            Err(error) => { self.create_name_validation_error = Some(error); cx.notify(); return; }
        };
        range.name = self.create_name_name.trim().to_string();
        range.description = (!self.create_name_description.is_empty()).then(|| self.create_name_description.clone());
        let name = range.name.clone();
        if self.apply_named_range_edit(visigrid_engine::workbook::NamedRangeEdit::Create(range), format!("Create named range: {name}"), cx) {
            self.log_refactor("Created named range", &format!("{} → {}", name, self.create_name_target), None);
            self.status_message = Some(format!("Created named range '{}' → {}", name, self.create_name_target));
            self.hide_create_named_range(cx);
        } else {
            self.create_name_validation_error = self.status_message.clone();
        }
    }

}
