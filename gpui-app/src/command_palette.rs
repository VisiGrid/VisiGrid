//! Command Palette and picker functionality for Spreadsheet.
//!
//! This module contains:
//! - Command palette show/hide/navigation
//! - Menu scope filtering (Alt accelerators)
//! - Cell references and precedents navigation
//! - Named range navigation
//! - Font picker
//! - Theme picker

use gpui::{*};

use crate::app::{Spreadsheet, PaletteScope};
use crate::mode::Mode;
use crate::search::{
    MenuCategory, ReferenceEntry, ReferencesProvider, SearchProvider, SearchQuery,
    SearchAction, SearchItem, SearchKind, CommandId, PrecedentEntry, PrecedentsProvider,
    CellSearchProvider, RecentFilesProvider, NamedRangeSearchProvider, NamedRangeEntry,
};
use crate::user_keybindings;
use visigrid_engine::named_range::NamedRangeTarget;

impl Spreadsheet {
    // Command Palette
    pub fn toggle_palette(&mut self, cx: &mut Context<Self>) {
        if self.mode == Mode::Command {
            self.hide_palette(cx);
        } else {
            self.show_palette(cx);
        }
    }

    /// Open the command palette with an optional scope.
    /// Shared logic for show_palette(), show_quick_open(), and apply_menu_scope().
    pub fn show_palette_with_scope(&mut self, scope: Option<PaletteScope>, cx: &mut Context<Self>) {
        // Close validation dropdown when opening modal
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
        self.lua_console.visible = false;
        self.tab_chain_origin_col = None;  // Dialog breaks tab chain

        // Save pre-palette state for restore on Esc (only if not already in palette)
        if self.mode != Mode::Command {
            self.palette_pre_selection = self.view_state.selected;
            self.palette_pre_selection_end = self.view_state.selection_end;
            self.palette_pre_scroll = (self.view_state.scroll_row, self.view_state.scroll_col);
            self.palette_previewing = false;
        }

        self.mode = Mode::Command;
        self.palette_query.clear();
        self.palette_selected = 0;
        self.palette_scope = scope;
        self.update_palette_results(cx);
        cx.notify();
    }

    pub fn show_palette(&mut self, cx: &mut Context<Self>) {
        self.show_palette_with_scope(None, cx);

        // One-time KeyTips discovery hint (macOS only, once per session)
        #[cfg(target_os = "macos")]
        if !self.keytips_hint_shown {
            self.keytips_hint_shown = true;
            self.status_message = Some("Tip: ⌥Space shows KeyTips (F/E/V/O/D/T/H)".into());
        }
    }

    /// Open palette scoped to recent files (Ctrl+K / Cmd+K)
    pub fn show_quick_open(&mut self, cx: &mut Context<Self>) {
        self.show_palette_with_scope(Some(PaletteScope::QuickOpen), cx);
    }

    /// Apply a menu scope filter (for Alt accelerators).
    /// Works whether palette is already open or not.
    pub fn apply_menu_scope(&mut self, category: MenuCategory, cx: &mut Context<Self>) {
        self.show_palette_with_scope(Some(PaletteScope::Menu(category)), cx);
    }

    /// Clear palette scope (backspace with empty query).
    /// Returns true if scope was cleared, false if no scope was active.
    pub fn clear_palette_scope(&mut self, cx: &mut Context<Self>) -> bool {
        if self.palette_scope.is_some() {
            self.palette_scope = None;
            self.update_palette_results(cx);
            cx.notify();
            true
        } else {
            false
        }
    }

    /// Show cells that reference the given cell (Find References - Shift+F12)
    /// Opens the command palette populated with dependent cells
    pub fn show_references(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        use visigrid_engine::formula::parser::{parse, extract_cell_refs};
        use visigrid_engine::cell::CellValue;

        // Get the cell reference for display
        let source_cell_ref = self.cell_ref_at(row, col);

        // Find all cells that reference this cell (dependents)
        let mut references = Vec::new();
        for (&(cell_row, cell_col), cell) in self.sheet(cx).cells_iter() {
            if let CellValue::Formula { source, .. } = &cell.value {
                if let Ok(expr) = parse(source) {
                    let refs = extract_cell_refs(&expr);
                    if refs.contains(&(row, col)) {
                        let cell_ref = self.cell_ref_at(cell_row, cell_col);
                        references.push(ReferenceEntry::new(
                            cell_row,
                            cell_col,
                            cell_ref,
                            source.clone(),
                        ));
                    }
                }
            }
        }

        if references.is_empty() {
            self.status_message = Some(format!("No cells reference {}", source_cell_ref));
            cx.notify();
            return;
        }

        // Sort references by cell position for predictable order
        references.sort_by_key(|r| (r.row, r.col));

        // Save pre-palette state for restore on Esc
        self.palette_pre_selection = self.view_state.selected;
        self.palette_pre_selection_end = self.view_state.selection_end;
        self.palette_pre_scroll = (self.view_state.scroll_row, self.view_state.scroll_col);
        self.palette_previewing = false;

        // Build results using the ReferencesProvider
        let provider = ReferencesProvider::new(source_cell_ref.clone(), references);
        let query = SearchQuery::parse("");
        let results = provider.search(&query, 50);

        // Open palette with references
        self.mode = Mode::Command;
        self.palette_query = format!("References to {}", source_cell_ref);
        self.palette_selected = 0;
        self.palette_total_results = results.len();
        self.palette_results = results;
        cx.notify();
    }

    /// Show cells that the given cell references (Go to Precedents - F12)
    /// Opens the command palette populated with precedent cells
    pub fn show_precedents(&mut self, row: usize, col: usize, cx: &mut Context<Self>) {
        use visigrid_engine::formula::parser::{parse, extract_cell_refs};

        // Get the cell reference for display
        let source_cell_ref = self.cell_ref_at(row, col);

        // Get formula from cell
        let raw = self.sheet(cx).get_raw(row, col);
        if !raw.starts_with('=') {
            self.status_message = Some(format!("{} is not a formula", source_cell_ref));
            cx.notify();
            return;
        }

        // Parse formula and extract cell references
        let refs = match parse(&raw) {
            Ok(expr) => extract_cell_refs(&expr),
            Err(_) => {
                self.status_message = Some("Could not parse formula".to_string());
                cx.notify();
                return;
            }
        };

        if refs.is_empty() {
            self.status_message = Some(format!("{} has no cell references", source_cell_ref));
            cx.notify();
            return;
        }

        // Build precedent entries
        let mut precedents: Vec<PrecedentEntry> = refs.iter().map(|(r, c)| {
            let cell_ref = self.cell_ref_at(*r, *c);
            let display = self.sheet(cx).get_display(*r, *c);
            PrecedentEntry::new(*r, *c, cell_ref, display)
        }).collect();

        // Sort by cell position
        precedents.sort_by_key(|p| (p.row, p.col));

        // Save pre-palette state
        self.palette_pre_selection = self.view_state.selected;
        self.palette_pre_selection_end = self.view_state.selection_end;
        self.palette_pre_scroll = (self.view_state.scroll_row, self.view_state.scroll_col);
        self.palette_previewing = false;

        // Build results
        let provider = PrecedentsProvider::new(source_cell_ref.clone(), precedents);
        let query = SearchQuery::parse("");
        let results = provider.search(&query, 50);

        // Open palette
        self.mode = Mode::Command;
        self.palette_query = format!("Precedents of {}", source_cell_ref);
        self.palette_selected = 0;
        self.palette_total_results = results.len();
        self.palette_results = results;
        cx.notify();
    }

    /// Get named range at cursor (if in a formula referencing one)
    pub fn named_range_at_cursor(&self, cx: &App) -> Option<String> {
        // Only works in formula mode with edit_value containing a formula
        if !self.mode.is_formula() && !self.mode.is_editing() {
            return None;
        }

        if !self.edit_value.starts_with('=') {
            return None;
        }

        // Find named range token at cursor position
        // This is a simplified check - could be improved with proper tokenization
        let cursor = self.edit_cursor;
        let text = &self.edit_value;

        // Find word boundaries around cursor
        let start = text[..cursor].rfind(|c: char| !c.is_alphanumeric() && c != '_')
            .map(|i| i + 1)
            .unwrap_or(0);
        let end = text[cursor..].find(|c: char| !c.is_alphanumeric() && c != '_')
            .map(|i| cursor + i)
            .unwrap_or(text.len());

        if start >= end {
            return None;
        }

        let word = &text[start..end];

        // Check if this word is a named range
        if self.wb(cx).get_named_range(word).is_some() {
            Some(word.to_string())
        } else {
            None
        }
    }

    /// Go to the definition of a named range (F12 on named range in formula)
    pub fn go_to_named_range_definition(&mut self, name: &str, cx: &mut Context<Self>) {
        use visigrid_engine::named_range::NamedRangeTarget;

        // Extract data from named range before mutable borrows
        let target_info = self.wb(cx).get_named_range(name).map(|nr| {
            let (row, col) = match &nr.target {
                NamedRangeTarget::Cell { row, col, .. } => (*row, *col),
                NamedRangeTarget::Range { start_row, start_col, .. } => (*start_row, *start_col),
            };
            (row, col, nr.reference_string())
        });

        if let Some((row, col, ref_str)) = target_info {
            // Exit edit mode and jump to the named range's target
            self.mode = Mode::Navigation;
            self.edit_value.clear();
            self.edit_cursor = 0;
            self.view_state.selected = (row, col);
            self.view_state.selection_end = None;
            self.ensure_cell_visible(row, col);
            self.status_message = Some(format!("'{}' → {}", name, ref_str));
            cx.notify();
        } else {
            self.status_message = Some(format!("Named range '{}' not found", name));
            cx.notify();
        }
    }

    /// Show all formulas that use a named range (Shift+F12 on named range)
    pub fn show_named_range_references(&mut self, name: &str, cx: &mut Context<Self>) {
        use visigrid_engine::cell::CellValue;

        let name_upper = name.to_uppercase();

        // Find all cells that use this named range
        let mut references = Vec::new();
        for (&(cell_row, cell_col), cell) in self.sheet(cx).cells_iter() {
            if let CellValue::Formula { source, .. } = &cell.value {
                // Check if formula references this named range (word-boundary aware)
                if self.formula_references_name(source, &name_upper) {
                    let cell_ref = self.cell_ref_at(cell_row, cell_col);
                    references.push(ReferenceEntry::new(
                        cell_row,
                        cell_col,
                        cell_ref,
                        source.clone(),
                    ));
                }
            }
        }

        if references.is_empty() {
            self.status_message = Some(format!("No cells reference '{}'", name));
            cx.notify();
            return;
        }

        // Sort references by cell position
        references.sort_by_key(|r| (r.row, r.col));

        // Save pre-palette state
        self.palette_pre_selection = self.view_state.selected;
        self.palette_pre_selection_end = self.view_state.selection_end;
        self.palette_pre_scroll = (self.view_state.scroll_row, self.view_state.scroll_col);
        self.palette_previewing = false;

        // Build results
        let provider = ReferencesProvider::new(format!("${}", name), references);
        let query = SearchQuery::parse("");
        let results = provider.search(&query, 50);

        // Open palette with references
        self.mode = Mode::Command;
        self.palette_query = format!("References to ${}", name);
        self.palette_selected = 0;
        self.palette_total_results = results.len();
        self.palette_results = results;
        cx.notify();
    }

    pub fn hide_palette(&mut self, cx: &mut Context<Self>) {
        // Restore pre-palette state (Esc behavior)
        if self.palette_previewing {
            self.view_state.selected = self.palette_pre_selection;
            self.view_state.selection_end = self.palette_pre_selection_end;
            self.view_state.scroll_row = self.palette_pre_scroll.0;
            self.view_state.scroll_col = self.palette_pre_scroll.1;
        }

        self.mode = Mode::Navigation;
        self.palette_query.clear();
        self.palette_selected = 0;
        self.palette_scope = None;  // Clear scope on close
        self.palette_results.clear();
        self.palette_previewing = false;
        cx.notify();
    }

    /// Preview selected palette item (updates view but remembers pre-state)
    pub fn palette_preview(&mut self, cx: &mut Context<Self>) {
        if let Some(item) = self.palette_results.get(self.palette_selected) {
            match &item.action {
                SearchAction::JumpToCell { row, col } => {
                    self.palette_previewing = true;
                    self.view_state.selected = (*row, *col);
                    self.view_state.selection_end = None;
                    self.ensure_visible(cx);
                    cx.notify();
                }
                _ => {}
            }
        }
    }

    /// Get palette results for rendering (borrows immutably)
    pub fn palette_results(&self) -> &[SearchItem] {
        &self.palette_results
    }

    /// Update palette results based on current query and scope
    pub(crate) fn update_palette_results(&mut self, cx: &App) {
        // Clone query string first to avoid borrow conflicts with cache refresh
        let query_str = self.palette_query.clone();
        let query = SearchQuery::parse(&query_str);

        // When scoped (Alt accelerators), search a larger pool before filtering
        // This prevents false "no matches" when the top 12 happen to be non-scoped commands
        let search_limit = if self.palette_scope.is_some() { 200 } else { 12 };
        let mut results = self.search_engine.search(&query_str, search_limit);

        // Add recent files when there's no prefix (commands + recent files)
        if query.prefix.is_none() && !self.recent_files.is_empty() {
            let provider = RecentFilesProvider::new(self.recent_files.clone());
            let recent_results = provider.search(&query, 10);
            results.extend(recent_results);
        }

        // Add named ranges when no prefix (Quick Open behavior) or with $ prefix
        if query.prefix.is_none() || query.prefix == Some('$') {
            let entries: Vec<NamedRangeEntry> = self.wb(cx).list_named_ranges()
                .into_iter()
                .map(|nr| {
                    let (row, col) = match &nr.target {
                        NamedRangeTarget::Cell { row, col, .. } => (*row, *col),
                        NamedRangeTarget::Range { start_row, start_col, .. } => (*start_row, *start_col),
                    };
                    NamedRangeEntry::new(
                        nr.name.clone(),
                        nr.reference_string(),
                        nr.description.clone(),
                        row,
                        col,
                    )
                })
                .collect();

            if !entries.is_empty() {
                let provider = NamedRangeSearchProvider::new(entries);
                // Limit named ranges in default view to avoid overwhelming commands
                let limit = if query.prefix == Some('$') { 50 } else { 5 };
                let named_results = provider.search(&query, limit);
                results.extend(named_results);
            }
        }

        // Add cell search with @ prefix (uses generation-based cache for freshness)
        if query.prefix == Some('@') {
            // Ensure cache is fresh (rebuilds only if cells_rev changed)
            self.ensure_cell_search_cache_fresh(cx);

            // Search over cached entries
            let provider = CellSearchProvider::new(self.cell_search_cache.entries.clone());
            let cell_results = provider.search(&query, 50);
            results.extend(cell_results);
        }

        // Filter by palette scope if set, but prefix overrides scope.
        // Typing ":B5" in QuickOpen routes to GoToCell, not filtered out.
        if query.prefix.is_none() {
            if let Some(scope) = &self.palette_scope {
                match scope {
                    PaletteScope::Menu(category) => {
                        results.retain(|item| {
                            if let SearchAction::RunCommand(cmd) = &item.action {
                                cmd.menu_category() == Some(*category)
                            } else {
                                false // Non-command items filtered out in menu scope
                            }
                        });
                    }
                    PaletteScope::QuickOpen => {
                        results.retain(|item| item.kind == SearchKind::RecentFile);
                    }
                }
            }
        }

        // Apply recency boost to commands (makes the palette feel "adaptive")
        for result in &mut results {
            if let SearchAction::RunCommand(cmd) = &result.action {
                let boost = self.command_recency_score(cmd);
                result.score += boost;
            }
        }

        // Apply unified sorting: score (desc) → kind priority (asc) → title (asc)
        results.sort_by(|a, b| {
            match b.score.partial_cmp(&a.score) {
                Some(std::cmp::Ordering::Equal) | None => {}
                Some(ord) => return ord,
            }
            match a.kind.priority().cmp(&b.kind.priority()) {
                std::cmp::Ordering::Equal => {}
                ord => return ord,
            }
            a.title.cmp(&b.title)
        });

        // Track total before truncation
        self.palette_total_results = results.len();
        results.truncate(12);

        self.palette_results = results;
    }

    /// Track a command as recently used (for scoring boost)
    pub(crate) fn add_recent_command(&mut self, cmd: CommandId) {
        const MAX_RECENT_COMMANDS: usize = 20;

        // Remove if already present (we'll add to front)
        self.recent_commands.retain(|c| c != &cmd);

        // Add to front
        self.recent_commands.insert(0, cmd);

        // Limit size
        self.recent_commands.truncate(MAX_RECENT_COMMANDS);
    }

    /// Check if a command was recently used (returns recency score 0.0-1.0)
    pub fn command_recency_score(&self, cmd: &CommandId) -> f32 {
        if let Some(pos) = self.recent_commands.iter().position(|c| c == cmd) {
            // More recent = higher score, decays with position
            // Position 0 (most recent) = 0.15 boost, position 19 = ~0.0 boost
            0.15 * (1.0 - (pos as f32 / 20.0))
        } else {
            0.0
        }
    }

    pub fn palette_up(&mut self, cx: &mut Context<Self>) {
        if self.palette_selected > 0 {
            self.palette_selected -= 1;
            cx.notify();
        }
    }

    pub fn palette_down(&mut self, cx: &mut Context<Self>) {
        let count = self.palette_results.len();
        if self.palette_selected + 1 < count {
            self.palette_selected += 1;
            cx.notify();
        }
    }

    pub fn palette_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.palette_query.push(c);
        self.palette_selected = 0;  // Reset selection on filter change
        self.update_palette_results(cx);
        cx.notify();
    }

    pub fn palette_backspace(&mut self, cx: &mut Context<Self>) {
        // Scope-aware backspace behavior:
        // 1. If query non-empty → delete last char
        // 2. If query empty + scope active → clear scope
        // 3. If query empty + no scope → do nothing (Esc closes)

        if !self.palette_query.is_empty() {
            // Retain prefix character if it's the only thing left
            // Prefixes: >, =, @, :, #, $
            let query_len = self.palette_query.chars().count();
            if query_len == 1 {
                let first_char = self.palette_query.chars().next().unwrap();
                if matches!(first_char, '>' | '=' | '@' | ':' | '#' | '$') {
                    // Don't remove the prefix - user stays in that search mode
                    return;
                }
            }
            self.palette_query.pop();
            self.palette_selected = 0;  // Reset selection on filter change
            self.update_palette_results(cx);
            cx.notify();
        } else if self.palette_scope.is_some() {
            // Query empty but scoped - clear scope, return to full palette
            self.palette_scope = None;
            self.palette_selected = 0;
            self.update_palette_results(cx);
            cx.notify();
        }
        // Query empty and no scope - do nothing (Esc closes palette)
    }

    pub fn palette_execute(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if let Some(item) = self.palette_results.get(self.palette_selected).cloned() {
            // Clear palette state - don't restore since we're executing
            self.palette_query.clear();
            self.palette_selected = 0;
            self.palette_results.clear();
            self.palette_previewing = false;  // Clear previewing flag

            self.dispatch_action(item.action, window, cx);
            // Only return to Navigation if action didn't change mode
            if self.mode == Mode::Command {
                self.mode = Mode::Navigation;
            }
            cx.notify();
        } else {
            self.hide_palette(cx);
        }
    }

    /// Execute secondary action (Ctrl+Enter) for selected palette item
    pub fn palette_execute_secondary(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if let Some(item) = self.palette_results.get(self.palette_selected).cloned() {
            if let Some(secondary) = item.secondary_action {
                // Clear palette state
                self.palette_query.clear();
                self.palette_selected = 0;
                self.palette_results.clear();
                self.palette_previewing = false;

                self.dispatch_action(secondary, window, cx);
                if self.mode == Mode::Command {
                    self.mode = Mode::Navigation;
                }
                cx.notify();
            } else {
                // No secondary action - show hint
                self.status_message = Some("No secondary action available".to_string());
                cx.notify();
            }
        }
    }

    // Font Picker
    pub fn show_font_picker(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.mode = Mode::FontPicker;
        self.font_picker_query.clear();
        self.font_picker_selected = 0;
        self.font_picker_scroll_offset = 0;
        // Focus the picker so first click is an activation click, not a focus click
        window.focus(&self.font_picker_focus, cx);
        cx.notify();
    }

    pub fn hide_font_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.font_picker_query.clear();
        self.font_picker_selected = 0;
        self.font_picker_scroll_offset = 0;
        cx.notify();
    }

    /// Maximum visible items in the font list
    const FONT_PICKER_VISIBLE: usize = 12;

    pub fn font_picker_up(&mut self, cx: &mut Context<Self>) {
        if self.font_picker_selected > 0 {
            self.font_picker_selected -= 1;
            // Keep selected item visible
            if self.font_picker_selected < self.font_picker_scroll_offset {
                self.font_picker_scroll_offset = self.font_picker_selected;
            }
            cx.notify();
        }
    }

    pub fn font_picker_down(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_fonts();
        if self.font_picker_selected + 1 < filtered.len() {
            self.font_picker_selected += 1;
            // Keep selected item visible
            if self.font_picker_selected >= self.font_picker_scroll_offset + Self::FONT_PICKER_VISIBLE {
                self.font_picker_scroll_offset = self.font_picker_selected + 1 - Self::FONT_PICKER_VISIBLE;
            }
            cx.notify();
        }
    }

    pub fn font_picker_scroll(&mut self, delta: i32, cx: &mut Context<Self>) {
        let filtered_len = self.filter_fonts().len();
        let max_offset = filtered_len.saturating_sub(Self::FONT_PICKER_VISIBLE);
        if delta > 0 {
            // Scroll down
            self.font_picker_scroll_offset = (self.font_picker_scroll_offset + delta as usize).min(max_offset);
        } else {
            // Scroll up
            self.font_picker_scroll_offset = self.font_picker_scroll_offset.saturating_sub((-delta) as usize);
        }
        cx.notify();
    }

    pub fn font_picker_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.font_picker_query.push(c);
        self.font_picker_selected = 0;
        self.font_picker_scroll_offset = 0;
        cx.notify();
    }

    pub fn font_picker_backspace(&mut self, cx: &mut Context<Self>) {
        self.font_picker_query.pop();
        self.font_picker_selected = 0;
        self.font_picker_scroll_offset = 0;
        cx.notify();
    }

    pub fn font_picker_execute(&mut self, cx: &mut Context<Self>) {
        let filtered = self.filter_fonts();
        if let Some(font_name) = filtered.get(self.font_picker_selected) {
            let font = font_name.clone();
            self.apply_font_to_selection(&font, cx);
        }
        self.hide_font_picker(cx);
    }

    /// Filter available fonts by query
    pub fn filter_fonts(&self) -> Vec<String> {
        if self.font_picker_query.is_empty() {
            return self.available_fonts.clone();
        }
        let query_lower = self.font_picker_query.to_lowercase();
        self.available_fonts
            .iter()
            .filter(|f| f.to_lowercase().contains(&query_lower))
            .cloned()
            .collect()
    }

    /// Apply font to all cells in current selection (with history)
    pub fn apply_font_to_selection(&mut self, font_name: &str, cx: &mut Context<Self>) {
        let font = if font_name.is_empty() { None } else { Some(font_name.to_string()) };
        self.set_font_family_selection(font, cx);
    }

    /// Clear font from selection (reset to default, with history)
    pub fn clear_font_from_selection(&mut self, cx: &mut Context<Self>) {
        self.set_font_family_selection(None, cx);
    }

    // Color Picker
    pub fn show_color_picker(&mut self, target: crate::color_palette::ColorTarget, window: &mut Window, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.mode = Mode::ColorPicker;
        self.ui.color_picker.target = target;
        self.ui.color_picker.reset();
        // Pre-populate hex input with current cell's color
        let (row, col) = self.view_state.selected;
        let current = self.sheet(cx).get_background_color(row, col);
        if let Some(color) = current {
            self.ui.color_picker.hex_input = crate::color_palette::to_hex(color);
        }
        window.focus(&self.ui.color_picker.focus, cx);
        cx.notify();
    }

    pub fn hide_color_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.ui.color_picker.reset();
        cx.notify();
    }

    pub fn apply_color_from_picker(&mut self, color: Option<[u8; 4]>, window: &mut Window, cx: &mut Context<Self>) {
        match self.ui.color_picker.target {
            crate::color_palette::ColorTarget::Fill => {
                self.set_background_color(color, cx);
            }
        }
        if let Some(c) = color {
            self.ui.color_picker.push_recent(c);
        }
        window.focus(&self.ui.color_picker.focus, cx);
    }

    /// Handle a key-down event while the color picker is focused.
    ///
    /// Returns `true` if the event was consumed.
    pub fn color_picker_handle_key(
        &mut self,
        key: &str,
        key_char: Option<&str>,
        has_modifier: bool,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) -> bool {
        use crate::ui::text_input::{handle_input_key, InputAction};
        let cp = &mut self.ui.color_picker;
        match handle_input_key(&mut cp.hex_input, &mut cp.all_selected, key, key_char, has_modifier) {
            InputAction::Changed => { cx.notify(); true }
            InputAction::Submit => { self.color_picker_execute(window, cx); true }
            InputAction::Cancel => { self.hide_color_picker(cx); true }
            InputAction::Ignored => false,
        }
    }

    pub fn color_picker_paste(&mut self, cx: &mut Context<Self>) {
        if let Some(item) = cx.read_from_clipboard() {
            if let Some(text) = item.text() {
                // Smart extraction: if pasted text contains a color token, use just that
                let to_insert = crate::color_palette::extract_color_token(&text)
                    .unwrap_or_else(|| {
                        text.trim().chars().filter(|c| !c.is_control()).collect()
                    });
                crate::ui::text_input::handle_input_paste(
                    &mut self.ui.color_picker.hex_input,
                    &mut self.ui.color_picker.all_selected,
                    &to_insert,
                );
                cx.notify();
            }
        }
    }

    pub fn color_picker_select_all(&mut self, cx: &mut Context<Self>) {
        crate::ui::text_input::handle_input_select_all(&mut self.ui.color_picker.all_selected);
        cx.notify();
    }

    pub fn color_picker_execute(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if let Some(color) = crate::color_palette::parse_hex_color(&self.ui.color_picker.hex_input) {
            self.apply_color_from_picker(Some(color), window, cx);
        }
        self.hide_color_picker(cx);
    }

    // Theme Picker
    pub fn show_theme_picker(&mut self, cx: &mut Context<Self>) {
        self.lua_console.visible = false;
        self.mode = Mode::ThemePicker;
        self.theme_picker_query.clear();
        self.theme_picker_selected = 0;
        self.theme_preview = None;
        cx.notify();
    }

    pub fn hide_theme_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        self.theme_picker_query.clear();
        self.theme_picker_selected = 0;
        self.theme_preview = None;
        cx.notify();
    }

    // Open keybindings.json in user's editor
    pub fn open_keybindings(&mut self, cx: &mut Context<Self>) {
        match user_keybindings::open_keybindings_file() {
            Ok(_) => {
                self.status_message = Some("Opened keybindings.json - restart to apply changes".into());
            }
            Err(e) => {
                self.status_message = Some(format!("Failed to open keybindings: {}", e));
            }
        }
        cx.notify();
    }
}
