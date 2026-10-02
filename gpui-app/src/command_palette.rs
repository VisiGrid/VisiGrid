//! Command Palette and picker functionality for Spreadsheet.
//!
//! This module contains:
//! - Command palette show/hide/navigation
//! - Menu scope filtering (Alt accelerators)
//! - Cell references and precedents navigation
//! - Named range navigation
//! - Font picker
//! - Theme picker
use visigrid_engine::cell::ValueRef;

use gpui::{*};

use crate::app::{Spreadsheet, PaletteScope};
use crate::mode::Mode;
use crate::search::{
    MenuCategory, ReferenceEntry, ReferencesProvider, SearchProvider, SearchQuery,
    SearchAction, SearchItem, SearchKind, CommandId, PrecedentEntry, PrecedentsProvider,
    CellSearchProvider, RecentFilesProvider, NamedRangeSearchProvider, NamedRangeEntry,
    CommandSearchProvider,
};
use crate::user_keybindings;
use visigrid_engine::named_range::NamedRangeTarget;

/// A heading over a run of palette results.
#[derive(Clone, Debug)]
pub struct PaletteSection {
    /// Index of the first result under this heading.
    pub start: usize,
    pub len: usize,
    /// Empty = a divider with no heading (the "Open from disk" row).
    pub title: String,
}

/// A command as a palette row.
fn command_item(cmd: CommandId) -> SearchItem {
    let mut item = SearchItem::new(SearchKind::Command, cmd.name(), SearchAction::RunCommand(cmd));
    if let Some(shortcut) = cmd.shortcut() {
        item = item.with_subtitle(shortcut);
    }
    if let Some(category) = cmd.menu_category() {
        item = item.with_meta(category.name());
    }
    item
}

/// The last row of Ctrl+K: leave recent files for the file dialog.
fn open_from_disk_item() -> SearchItem {
    SearchItem::new(SearchKind::Command, "Open from disk…", SearchAction::RunCommand(CommandId::OpenFile))
        .with_subtitle("Ctrl+O")
}

/// Files, functions and settings carry a second line (folder, description);
/// a command's subtitle is its shortcut, drawn as key caps on the same line.
pub(crate) fn palette_row_has_second_line(item: &SearchItem) -> bool {
    item.subtitle.is_some() && item.kind != SearchKind::Command
}

pub(crate) fn is_open_from_disk(item: &SearchItem) -> bool {
    item.kind == SearchKind::Command
        && item.action == SearchAction::RunCommand(CommandId::OpenFile)
        && item.title == "Open from disk…"
}

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
            self.ui.palette_edit_mode = self.mode.is_editing().then_some(self.mode);
            self.palette_pre_selection = self.view_state.selected;
            self.palette_pre_selection_end = self.view_state.selection_end;
            self.palette_pre_scroll = (self.view_state.scroll_row, self.view_state.scroll_col);
            self.palette_previewing = false;
        }

        self.mode = Mode::Command;
        self.palette_query.clear();
        self.palette_selected = 0;
        self.palette_scroll_offset = 0;
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

    /// Open command palette with a pre-filled query string.
    /// Used by inspector to link to style commands, etc.
    pub fn open_command_palette_with_prefill(&mut self, query: &str, cx: &mut Context<Self>) {
        self.show_palette_with_scope(None, cx);
        self.palette_query = query.to_string();
        self.update_palette_results(cx);
        cx.notify();
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
        for ((cell_row, cell_col), cell) in self.sheet(cx).cells_iter() {
            if let ValueRef::Formula { source, .. } = cell.value() {
                if let Ok(expr) = parse(source) {
                    let refs = extract_cell_refs(&expr);
                    if refs.contains(&(row, col)) {
                        let cell_ref = self.cell_ref_at(cell_row, cell_col);
                        references.push(ReferenceEntry::new(
                            cell_row,
                            cell_col,
                            cell_ref,
                            source.to_string(),
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
        self.palette_scroll_offset = 0;
        self.set_palette_sections(vec![(String::new(), results)]);
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
        self.palette_scroll_offset = 0;
        self.set_palette_sections(vec![(String::new(), results)]);
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
        for ((cell_row, cell_col), cell) in self.sheet(cx).cells_iter() {
            if let ValueRef::Formula { source, .. } = cell.value() {
                // Check if formula references this named range (word-boundary aware)
                if self.formula_references_name(source, &name_upper) {
                    let cell_ref = self.cell_ref_at(cell_row, cell_col);
                    references.push(ReferenceEntry::new(
                        cell_row,
                        cell_col,
                        cell_ref,
                        source.to_string(),
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
        self.palette_scroll_offset = 0;
        self.set_palette_sections(vec![(String::new(), results)]);
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

        self.mode = self.ui.palette_edit_mode.take().unwrap_or(Mode::Navigation);
        self.palette_query.clear();
        self.palette_selected = 0;
        self.palette_scroll_offset = 0;
        self.palette_scope = None;  // Clear scope on close
        self.palette_results.clear();
        self.palette_sections.clear();
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

    /// Update palette results based on current query and scope.
    ///
    /// Results are kept flat (selection and execution index into one list) and
    /// grouped into sections for display: an empty palette shows Recent, For
    /// this selection, Recent files and All commands; a search shows Commands,
    /// Files, Named ranges and so on, each ranked within itself.
    pub(crate) fn update_palette_results(&mut self, cx: &App) {
        let query_str = self.palette_query.clone();
        let query = SearchQuery::parse(&query_str);

        let sections = match self.palette_scope {
            Some(PaletteScope::QuickOpen) if query.prefix.is_none() => self.quick_open_sections(&query),
            None if query.prefix.is_none() && query.needle.is_empty() => self.empty_palette_sections(cx),
            _ => self.search_sections(&query, cx),
        };
        self.set_palette_sections(sections);
    }

    /// Flatten titled groups into `palette_results` + `palette_sections`.
    /// Empty groups are dropped; an empty title draws a divider, not a heading.
    pub(crate) fn set_palette_sections(&mut self, groups: Vec<(String, Vec<SearchItem>)>) {
        self.palette_results.clear();
        self.palette_sections.clear();
        self.palette_total_results = 0;
        for (title, items) in groups {
            if items.is_empty() {
                continue;
            }
            self.palette_sections.push(PaletteSection {
                start: self.palette_results.len(),
                len: items.len(),
                title,
            });
            self.palette_total_results += items.iter().filter(|i| !is_open_from_disk(i)).count();
            self.palette_results.extend(items);
        }
        if self.palette_selected >= self.palette_results.len() {
            self.palette_selected = 0;
            self.palette_scroll_offset = 0;
        }
    }

    /// Ctrl+K: recent files, then a way out to the file dialog.
    fn quick_open_sections(&self, query: &SearchQuery) -> Vec<(String, Vec<SearchItem>)> {
        let files = RecentFilesProvider::new(self.recent_files.clone()).search(query, 50);
        let title = if query.needle.is_empty() { "Recent files" } else { "Files" };
        vec![
            (title.to_string(), files),
            (String::new(), vec![open_from_disk_item()]),
        ]
    }

    /// The palette before anything is typed.
    fn empty_palette_sections(&self, cx: &App) -> Vec<(String, Vec<SearchItem>)> {
        let recent: Vec<CommandId> = self.recent_commands.iter().copied().take(3).collect();
        let (selection_label, suggested) = self.selection_suggestions(cx);
        let suggested: Vec<CommandId> =
            suggested.into_iter().filter(|c| !recent.contains(c)).collect();

        let all: Vec<SearchItem> = {
            let mut items = CommandSearchProvider.search(&SearchQuery::parse(""), usize::MAX);
            items.retain(|i| match &i.action {
                SearchAction::RunCommand(c) => !recent.contains(c) && !suggested.contains(c),
                _ => true,
            });
            items.sort_by(|a, b| a.title.cmp(&b.title));
            items
        };

        let files = RecentFilesProvider::new(self.recent_files.iter().take(3).cloned().collect())
            .search(&SearchQuery::parse(""), 3);

        vec![
            ("Recent".to_string(), recent.into_iter().map(command_item).collect()),
            (format!("For this selection · {selection_label}"), suggested.into_iter().map(command_item).collect()),
            ("Recent files".to_string(), files),
            ("All commands".to_string(), all),
        ]
    }

    /// A typed query, a prefix, or a menu scope.
    fn search_sections(&mut self, query: &SearchQuery, cx: &App) -> Vec<(String, Vec<SearchItem>)> {
        const LIMIT: usize = 500;
        let mut results = self.search_engine.search(query.raw, LIMIT);

        // Recent files join unprefixed search
        if query.prefix.is_none() && !self.recent_files.is_empty() {
            let provider = RecentFilesProvider::new(self.recent_files.clone());
            results.extend(provider.search(query, 10));
        }

        // Named ranges join unprefixed search, or are the whole result with $
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
                let limit = if query.prefix == Some('$') { LIMIT } else { 10 };
                results.extend(NamedRangeSearchProvider::new(entries).search(query, limit));
            }
        }

        // Cell search with @ prefix (uses generation-based cache for freshness)
        if query.prefix == Some('@') {
            self.ensure_cell_search_cache_fresh(cx);
            let provider = CellSearchProvider::new(self.cell_search_cache.entries.clone());
            results.extend(provider.search(query, LIMIT));
        }

        // Filter by palette scope if set, but prefix overrides scope.
        // Typing ":B5" in QuickOpen routes to GoToCell, not filtered out.
        if query.prefix.is_none() {
            if let Some(PaletteScope::Menu(category)) = &self.palette_scope {
                results.retain(|item| match &item.action {
                    SearchAction::RunCommand(cmd) => cmd.menu_category() == Some(*category),
                    _ => false,
                });
            }
        }

        // Recency boost makes the palette feel adaptive
        for result in &mut results {
            if let SearchAction::RunCommand(cmd) = &result.action {
                result.score += self.command_recency_score(cmd);
            }
        }

        // score (desc) → kind priority (asc) → title (asc)
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

        // Group by kind, keeping rank order inside each group. Groups appear
        // in the order of their best result, so a query that only matches a
        // file shows Files first.
        let mut groups: Vec<(SearchKind, Vec<SearchItem>)> = Vec::new();
        for item in results {
            match groups.iter_mut().find(|(k, _)| *k == item.kind) {
                Some((_, items)) => items.push(item),
                None => groups.push((item.kind, vec![item])),
            }
        }
        groups
            .into_iter()
            .map(|(kind, items)| (kind.section_title().to_string(), items))
            .collect()
    }

    /// Commands that fit what is selected, and the label for their heading.
    ///
    /// Deliberately few rules: values → sort and filter; numbers → AutoSum and
    /// number formats; an empty cell under numbers → AutoSum.
    fn selection_suggestions(&self, cx: &App) -> (String, Vec<CommandId>) {
        use visigrid_engine::formula::eval::Value;

        let ((r0, c0), (r1, c1)) = self.selection_range();
        let label = if (r0, c0) == (r1, c1) {
            self.cell_ref_at(r0, c0)
        } else {
            format!("{}:{}", self.cell_ref_at(r0, c0), self.cell_ref_at(r1, c1))
        };

        let sheet = self.sheet(cx);
        if (r0, c0) == (r1, c1) {
            let empty = matches!(sheet.get_computed_value(r0, c0), Value::Empty);
            let above_is_number = r0 > 0
                && matches!(sheet.get_computed_value(r0 - 1, c0), Value::Number(_));
            let suggested = if empty && above_is_number {
                vec![CommandId::AutoSum]
            } else {
                Vec::new()
            };
            return (label, suggested);
        }

        let (mut numbers, mut texts) = (0usize, 0usize);
        for ((row, col), _) in sheet.cells_iter() {
            if row < r0 || row > r1 || col < c0 || col > c1 {
                continue;
            }
            match sheet.get_computed_value(row, col) {
                Value::Number(_) => numbers += 1,
                Value::Text(_) | Value::Boolean(_) => texts += 1,
                _ => {}
            }
        }

        let mut suggested = Vec::new();
        if numbers + texts > 1 {
            suggested.extend([CommandId::SortAscending, CommandId::SortDescending, CommandId::ToggleAutoFilter]);
        }
        if numbers > 0 && numbers >= texts {
            suggested.extend([CommandId::AutoSum, CommandId::FormatCurrency, CommandId::FormatPercent]);
        }
        (label, suggested)
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

    /// Rows a PageUp/PageDown moves.
    pub(crate) const PALETTE_VISIBLE: usize = 10;
    /// Pixel height of the result list; rows and headings fill it from the
    /// scroll offset down (see `palette_window_end`).
    pub(crate) const PALETTE_LIST_H: f32 = 444.0;
    pub(crate) const PALETTE_HEADER_H: f32 = 26.0;
    pub(crate) const PALETTE_DIVIDER_H: f32 = 9.0;
    pub(crate) const PALETTE_ROW_H: f32 = 30.0;
    pub(crate) const PALETTE_ROW2_H: f32 = 42.0;

    /// The heading drawn above row `idx`: the section starting there, or — for
    /// the first visible row — the section it belongs to, so a scrolled list
    /// still says what it is showing.
    pub(crate) fn palette_heading_at(&self, idx: usize, first_visible: bool) -> Option<&PaletteSection> {
        let section = self.palette_sections.iter().find(|s| idx >= s.start && idx < s.start + s.len)?;
        if section.start != idx && !first_visible {
            return None;
        }
        // A nameless section at the very top is just the list (References, Precedents)
        if section.title.is_empty() && (section.start == 0 || section.start != idx) {
            return None;
        }
        Some(section)
    }

    fn palette_row_height(&self, idx: usize, first_visible: bool) -> f32 {
        let heading = match self.palette_heading_at(idx, first_visible) {
            Some(s) if s.title.is_empty() => Self::PALETTE_DIVIDER_H,
            Some(_) => Self::PALETTE_HEADER_H,
            None => 0.0,
        };
        let row = if self.palette_results.get(idx).is_some_and(palette_row_has_second_line) {
            Self::PALETTE_ROW2_H
        } else {
            Self::PALETTE_ROW_H
        };
        heading + row
    }

    /// One past the last row that fits when the list starts at `offset`.
    pub(crate) fn palette_window_end(&self, offset: usize) -> usize {
        let mut used = 0.0;
        let mut idx = offset;
        while idx < self.palette_results.len() {
            used += self.palette_row_height(idx, idx == offset);
            if used > Self::PALETTE_LIST_H && idx > offset {
                break;
            }
            idx += 1;
        }
        idx
    }

    fn palette_max_offset(&self) -> usize {
        let len = self.palette_results.len();
        let mut offset = len.saturating_sub(1);
        while offset > 0 && self.palette_window_end(offset - 1) >= len {
            offset -= 1;
        }
        offset
    }

    fn palette_follow_selection(&mut self) {
        if self.palette_selected < self.palette_scroll_offset {
            self.palette_scroll_offset = self.palette_selected;
        }
        while self.palette_selected >= self.palette_window_end(self.palette_scroll_offset) {
            self.palette_scroll_offset += 1;
        }
    }

    pub fn palette_up(&mut self, cx: &mut Context<Self>) {
        if self.palette_selected > 0 {
            self.palette_selected -= 1;
            self.palette_follow_selection();
            cx.notify();
        }
    }

    pub fn palette_down(&mut self, cx: &mut Context<Self>) {
        let count = self.palette_results.len();
        if self.palette_selected + 1 < count {
            self.palette_selected += 1;
            self.palette_follow_selection();
            cx.notify();
        }
    }

    /// Move the selection by `delta` rows, clamped (PageUp/PageDown/Home/End).
    pub fn palette_move(&mut self, delta: isize, cx: &mut Context<Self>) {
        let count = self.palette_results.len();
        if count == 0 {
            return;
        }
        let target = (self.palette_selected as isize + delta).clamp(0, count as isize - 1) as usize;
        if target != self.palette_selected {
            self.palette_selected = target;
            self.palette_follow_selection();
            cx.notify();
        }
    }

    /// Mouse wheel: scroll the window without moving the selection.
    pub fn palette_scroll(&mut self, rows: isize, cx: &mut Context<Self>) {
        let max = self.palette_max_offset();
        let target = (self.palette_scroll_offset as isize + rows).clamp(0, max as isize) as usize;
        if target != self.palette_scroll_offset {
            self.palette_scroll_offset = target;
            cx.notify();
        }
    }

    pub fn palette_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        self.palette_query.push(c);
        self.palette_selected = 0;
        self.palette_scroll_offset = 0;  // Reset selection on filter change
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
            self.palette_selected = 0;
        self.palette_scroll_offset = 0;  // Reset selection on filter change
            self.update_palette_results(cx);
            cx.notify();
        } else if self.palette_scope.is_some() {
            // Query empty but scoped - clear scope, return to full palette
            self.palette_scope = None;
            self.palette_selected = 0;
        self.palette_scroll_offset = 0;
            self.update_palette_results(cx);
            cx.notify();
        }
        // Query empty and no scope - do nothing (Esc closes palette)
    }

    pub fn palette_execute(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if let Some(item) = self.palette_results.get(self.palette_selected).cloned() {
            // Layout actions change personal chrome, not the suspended editor.
            if matches!(item.action, SearchAction::RunCommand(CommandId::UseCompactToolbar
                | CommandId::UseRibbonToolbar | CommandId::ToggleRibbonCollapsed | CommandId::ToggleToolbar)) {
                self.view_state.selected = self.palette_pre_selection;
                self.view_state.selection_end = self.palette_pre_selection_end;
                self.view_state.scroll_row = self.palette_pre_scroll.0;
                self.view_state.scroll_col = self.palette_pre_scroll.1;
                self.mode = self.ui.palette_edit_mode.take().unwrap_or(Mode::Navigation);
            } else {
                self.ui.palette_edit_mode = None;
            }
            // Clear palette state - don't restore since we're executing
            self.palette_query.clear();
            self.palette_selected = 0;
        self.palette_scroll_offset = 0;
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
                self.ui.palette_edit_mode = None;
                // Clear palette state
                self.palette_query.clear();
                self.palette_selected = 0;
        self.palette_scroll_offset = 0;
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
        self.font_picker_for_default = false;
        self.font_picker_query.clear();
        self.font_picker_selected = 0;
        self.font_picker_scroll_offset = 0;
        // Focus the picker so first click is an activation click, not a focus click
        window.focus(&self.font_picker_focus, cx);
        cx.notify();
    }

    pub fn hide_font_picker(&mut self, cx: &mut Context<Self>) {
        self.mode = if self.font_picker_for_default { Mode::Preferences } else { Mode::Navigation };
        self.font_picker_for_default = false;
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
            self.apply_picked_font(&font, cx);
        }
        self.hide_font_picker(cx);
    }

    pub fn apply_picked_font(&mut self, font: &str, cx: &mut Context<Self>) {
        if self.font_picker_for_default {
            crate::settings::update_user_settings(cx, |settings| {
                settings.appearance.default_font_family = crate::settings::Setting::Value(font.to_string());
            });
        } else {
            self.apply_font_to_selection(font, cx);
        }
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
        // Pre-populate hex input with current color
        let current = match target {
            crate::color_palette::ColorTarget::Fill => {
                let (row, col) = self.view_state.selected;
                self.sheet(cx).get_background_color(row, col)
            }
            crate::color_palette::ColorTarget::Text => {
                let (row, col) = self.view_state.selected;
                self.sheet(cx).get_format(row, col).font_color
            }
            crate::color_palette::ColorTarget::Border => self.current_border_color,
        };
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
            crate::color_palette::ColorTarget::Text => {
                self.set_font_color_selection(color, cx);
            }
            crate::color_palette::ColorTarget::Border => {
                self.current_border_color = color;
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
