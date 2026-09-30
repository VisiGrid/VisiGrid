pub mod data;
mod ops;

use std::io::{self, stdout, Write};
use std::time::Duration;

use crossterm::{
    event::{self, Event, KeyCode, KeyEvent, KeyModifiers},
    terminal::{self, EnterAlternateScreen, LeaveAlternateScreen},
    ExecutableCommand,
};
use ratatui::{
    backend::CrosstermBackend,
    layout::{Constraint, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Clear, Paragraph},
    Frame, Terminal,
};

use crate::util;
use data::{PeekData, SheetData};

struct TuiApp {
    /// All sheets (for .sheet files) or a single sheet (for CSV)
    sheets: Vec<SheetData>,
    /// Index into `sheets` for the active sheet
    active_sheet: usize,
    cursor_row: usize,
    cursor_col: usize,
    scroll_row: usize,
    scroll_col: usize,
    file_name: String,
    should_quit: bool,
    show_help: bool,
    /// Width of the row-number gutter, computed from max file row number
    row_num_width: usize,
    /// Whether there is more than one tab (workbook sheets or derived tabs)
    multi_sheet: bool,
    /// Per tab: display order of data rows (indices into `data.rows`).
    /// The cursor and scroll positions index this order.
    order: Vec<Vec<usize>>,
    /// Per tab: the tab a derived tab (frequency, pivot) was made from.
    parent: Vec<Option<usize>>,
    /// Per tab: (cursor_row, cursor_col, scroll_row, scroll_col), restored
    /// when the tab is shown again.
    positions: Vec<(usize, usize, usize, usize)>,
    /// Text entry on the status line (search or pivot spec).
    prompt: Option<Prompt>,
    /// Last search query, for n / N.
    search: Option<String>,
    /// One-shot status message, cleared on the next key.
    message: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum PromptKind {
    Search,
    Pivot,
}

struct Prompt {
    kind: PromptKind,
    input: String,
}

impl TuiApp {
    fn new(data: PeekData, file_name: String) -> Self {
        let row_num_width = Self::compute_row_num_width(&data);
        Self {
            sheets: vec![SheetData {
                name: String::new(),
                data,
            }],
            active_sheet: 0,
            cursor_row: 0,
            cursor_col: 0,
            scroll_row: 0,
            scroll_col: 0,
            file_name,
            should_quit: false,
            show_help: false,
            row_num_width,
            multi_sheet: false,
            order: Vec::new(),
            parent: vec![None],
            positions: Vec::new(),
            prompt: None,
            search: None,
            message: None,
        }
        .with_identity_orders()
    }

    fn new_multi(sheets: Vec<SheetData>, file_name: String, initial_sheet: usize) -> Self {
        let active = initial_sheet.min(sheets.len().saturating_sub(1));
        let row_num_width = Self::compute_row_num_width(&sheets[active].data);
        let multi = sheets.len() > 1;
        Self {
            sheets,
            active_sheet: active,
            cursor_row: 0,
            cursor_col: 0,
            scroll_row: 0,
            scroll_col: 0,
            file_name,
            should_quit: false,
            show_help: false,
            row_num_width,
            multi_sheet: multi,
            order: Vec::new(),
            parent: Vec::new(),
            positions: Vec::new(),
            prompt: None,
            search: None,
            message: None,
        }
        .with_identity_orders()
    }

    fn with_identity_orders(mut self) -> Self {
        self.order = self.sheets.iter().map(|s| (0..s.data.num_rows).collect()).collect();
        self.parent = vec![None; self.sheets.len()];
        self.positions = vec![(0, 0, 0, 0); self.sheets.len()];
        self
    }

    /// Data row index behind a view position.
    fn row_at(&self, pos: usize) -> usize {
        self.order[self.active_sheet].get(pos).copied().unwrap_or(pos)
    }

    fn cell(&self, pos: usize, col: usize) -> &str {
        self.data().rows.get(self.row_at(pos)).and_then(|r| r.get(col)).map(|s| s.as_str()).unwrap_or("")
    }

    fn col_name(&self, col: usize) -> String {
        self.data().col_names.get(col).cloned().unwrap_or_else(|| util::col_to_letter(col))
    }

    /// Open a derived tab and switch to it.
    fn push_tab(&mut self, name: String, data: PeekData) {
        let from = self.active_sheet;
        if self.sheets[from].name.is_empty() {
            self.sheets[from].name = self.file_name.clone();
        }
        self.order.push((0..data.num_rows).collect());
        self.sheets.push(SheetData { name, data });
        self.parent.push(Some(from));
        self.positions.push((0, 0, 0, 0));
        self.multi_sheet = true;
        self.switch_sheet(self.sheets.len() - 1);
    }

    /// Close the active derived tab and return to the tab it came from.
    fn close_tab(&mut self) {
        let i = self.active_sheet;
        let Some(back) = self.parent[i] else { return };
        self.sheets.remove(i);
        self.order.remove(i);
        self.parent.remove(i);
        self.positions.remove(i);
        for p in self.parent.iter_mut().flatten() {
            if *p == i {
                *p = back;
            } else if *p > i {
                *p -= 1;
            }
        }
        self.multi_sheet = self.sheets.len() > 1;
        self.active_sheet = usize::MAX; // force switch_sheet to reset state
        self.switch_sheet(if back > i { back - 1 } else { back });
    }

    /// Suffix for derived tab names when the preview is not the whole file.
    fn partial_note(&self) -> &'static str {
        let d = self.data();
        if d.total_rows.is_some_and(|t| t > d.num_rows) { " (loaded rows)" } else { "" }
    }

    fn sort(&mut self, descending: bool) {
        let col = self.cursor_col;
        let keep = self.row_at(self.cursor_row);
        let order = ops::sorted_order(self.data(), &self.order[self.active_sheet], col, descending);
        self.cursor_row = order.iter().position(|&r| r == keep).unwrap_or(0);
        self.order[self.active_sheet] = order;
        self.message = Some(format!(
            "sorted by {} {}{}",
            self.col_name(col),
            if descending { "descending" } else { "ascending" },
            if self.partial_note().is_empty() { "" } else { " (loaded rows only)" }
        ));
    }

    fn search_next(&mut self, backward: bool) {
        let Some(query) = self.search.clone() else {
            self.message = Some("no search yet: press / to search".into());
            return;
        };
        match ops::find(self.data(), &self.order[self.active_sheet], (self.cursor_row, self.cursor_col), &query, backward) {
            Some((pos, col)) => {
                self.cursor_row = pos;
                self.cursor_col = col;
                let n = ops::count_matches(self.data(), &query);
                self.message = Some(format!("/{query}  {n} match{}", if n == 1 { "" } else { "es" }));
            }
            None => self.message = Some(format!("/{query}  not found")),
        }
    }

    fn submit_prompt(&mut self, prompt: Prompt) {
        match prompt.kind {
            PromptKind::Search => {
                if prompt.input.is_empty() {
                    return;
                }
                self.search = Some(prompt.input);
                // Start just before the cursor so a match in the current cell counts.
                let cols = self.data().num_cols.max(1);
                let rows = self.order[self.active_sheet].len().max(1);
                let flat = (self.cursor_row * cols + self.cursor_col + rows * cols - 1) % (rows * cols);
                let (r, c) = (flat / cols, flat % cols);
                let (save_r, save_c) = (self.cursor_row, self.cursor_col);
                self.cursor_row = r;
                self.cursor_col = c;
                self.search_next(false);
                if self.message.as_deref().is_some_and(|m| m.ends_with("not found")) {
                    self.cursor_row = save_r;
                    self.cursor_col = save_c;
                }
            }
            PromptKind::Pivot => match ops::pivot(self.data(), &prompt.input) {
                Ok((data, label)) => {
                    let name = format!("{label}{}", self.partial_note());
                    let headerless = !self.data().has_headers;
                    self.push_tab(name, data);
                    self.message = Some(if headerless {
                        "pivot: first row used as headers (open with --headers to show them) · q returns".into()
                    } else {
                        "pivot: q returns to the source".into()
                    });
                }
                Err(e) => self.message = Some(format!("pivot: {e}")),
            },
        }
    }

    fn handle_prompt_key(&mut self, key: KeyEvent) {
        let Some(prompt) = self.prompt.as_mut() else { return };
        match key.code {
            KeyCode::Esc => self.prompt = None,
            KeyCode::Enter => {
                let p = self.prompt.take().unwrap();
                self.submit_prompt(p);
            }
            KeyCode::Backspace => {
                if prompt.input.pop().is_none() {
                    self.prompt = None;
                }
            }
            KeyCode::Char('u') if key.modifiers.contains(KeyModifiers::CONTROL) => prompt.input.clear(),
            KeyCode::Char(c) if !key.modifiers.contains(KeyModifiers::CONTROL) => prompt.input.push(c),
            _ => {}
        }
    }

    fn compute_row_num_width(data: &PeekData) -> usize {
        let max_file_row = data.file_row(data.num_rows.saturating_sub(1));
        let digits = if max_file_row == 0 {
            1
        } else {
            (max_file_row as f64).log10().floor() as usize + 1
        };
        digits.max(3) + 1
    }

    fn data(&self) -> &PeekData {
        &self.sheets[self.active_sheet].data
    }

    fn switch_sheet(&mut self, idx: usize) {
        if idx >= self.sheets.len() || idx == self.active_sheet {
            return;
        }
        if let Some(p) = self.positions.get_mut(self.active_sheet) {
            *p = (self.cursor_row, self.cursor_col, self.scroll_row, self.scroll_col);
        }
        self.active_sheet = idx;
        (self.cursor_row, self.cursor_col, self.scroll_row, self.scroll_col) =
            self.positions.get(idx).copied().unwrap_or((0, 0, 0, 0));
        self.row_num_width = Self::compute_row_num_width(self.data());
    }

    fn next_sheet(&mut self) {
        if self.sheets.len() > 1 {
            let next = (self.active_sheet + 1) % self.sheets.len();
            self.switch_sheet(next);
        }
    }

    fn prev_sheet(&mut self) {
        if self.sheets.len() > 1 {
            let prev = if self.active_sheet == 0 {
                self.sheets.len() - 1
            } else {
                self.active_sheet - 1
            };
            self.switch_sheet(prev);
        }
    }

    fn handle_key(&mut self, key: KeyEvent) {
        if self.show_help {
            // Any key dismisses help
            self.show_help = false;
            return;
        }
        if self.prompt.is_some() {
            self.handle_prompt_key(key);
            return;
        }
        self.message = None;

        match key.code {
            KeyCode::Char('q') | KeyCode::Esc => {
                if self.parent[self.active_sheet].is_some() {
                    self.close_tab();
                } else {
                    self.should_quit = true;
                }
            }
            KeyCode::Char('/') => self.prompt = Some(Prompt { kind: PromptKind::Search, input: String::new() }),
            KeyCode::Char('n') => self.search_next(false),
            KeyCode::Char('N') => self.search_next(true),
            KeyCode::Char('[') => self.sort(false),
            KeyCode::Char(']') => self.sort(true),
            KeyCode::Char('R') => {
                let keep = self.row_at(self.cursor_row);
                self.order[self.active_sheet] = (0..self.data().num_rows).collect();
                self.cursor_row = keep;
                self.message = Some("file order".into());
            }
            KeyCode::Char('F') if self.data().num_cols > 0 => {
                let col = self.cursor_col;
                let data = ops::frequency(self.data(), col);
                let name = format!("freq {}{}", self.col_name(col), self.partial_note());
                self.push_tab(name, data);
            }
            KeyCode::Char('P') if self.data().num_cols > 0 => {
                // The header row is the first row in file order, not whatever
                // sorts first.
                let field = if self.data().has_headers {
                    self.col_name(self.cursor_col)
                } else {
                    self.data().rows.first().and_then(|r| r.get(self.cursor_col)).cloned().unwrap_or_default()
                };
                self.prompt = Some(Prompt { kind: PromptKind::Pivot, input: format!("rows={field} values=count:{field}") });
            }
            KeyCode::Char('?') => self.show_help = true,
            KeyCode::Up | KeyCode::Char('k') => self.move_cursor(-1, 0),
            KeyCode::Down | KeyCode::Char('j') => self.move_cursor(1, 0),
            KeyCode::Left | KeyCode::Char('h') => self.move_cursor(0, -1),
            KeyCode::Right | KeyCode::Char('l') => self.move_cursor(0, 1),
            KeyCode::PageUp => {
                if key.modifiers.contains(KeyModifiers::CONTROL) {
                    self.prev_sheet();
                } else {
                    self.page_up();
                }
            }
            KeyCode::PageDown => {
                if key.modifiers.contains(KeyModifiers::CONTROL) {
                    self.next_sheet();
                } else {
                    self.page_down();
                }
            }
            KeyCode::Home | KeyCode::Char('g') => self.cursor_row = 0,
            KeyCode::End | KeyCode::Char('G') => {
                if self.data().num_rows > 0 {
                    self.cursor_row = self.data().num_rows - 1;
                }
            }
            KeyCode::Char('0') => self.cursor_col = 0,
            KeyCode::Char('$') => {
                if self.data().num_cols > 0 {
                    self.cursor_col = self.data().num_cols - 1;
                }
            }
            // 1-9: jump to sheet by index
            KeyCode::Char(c @ '1'..='9') if self.multi_sheet => {
                let idx = (c as usize) - ('1' as usize);
                self.switch_sheet(idx);
            }
            KeyCode::Tab => {
                if self.multi_sheet {
                    if key.modifiers.contains(KeyModifiers::SHIFT) {
                        self.prev_sheet();
                    } else {
                        self.next_sheet();
                    }
                } else if key.modifiers.contains(KeyModifiers::SHIFT) {
                    self.move_cursor(0, -1);
                } else {
                    self.move_cursor(0, 1);
                }
            }
            KeyCode::BackTab => {
                if self.multi_sheet {
                    self.prev_sheet();
                } else {
                    self.move_cursor(0, -1);
                }
            }
            _ => {}
        }
    }

    fn move_cursor(&mut self, drow: i32, dcol: i32) {
        let data = self.data();
        if data.num_rows == 0 || data.num_cols == 0 {
            return;
        }
        let new_row = (self.cursor_row as i32 + drow)
            .max(0)
            .min(data.num_rows as i32 - 1) as usize;
        let new_col = (self.cursor_col as i32 + dcol)
            .max(0)
            .min(data.num_cols as i32 - 1) as usize;
        self.cursor_row = new_row;
        self.cursor_col = new_col;
    }

    fn page_up(&mut self) {
        let jump = 20;
        self.cursor_row = self.cursor_row.saturating_sub(jump);
    }

    fn page_down(&mut self) {
        let jump = 20;
        let num_rows = self.data().num_rows;
        if num_rows > 0 {
            self.cursor_row = (self.cursor_row + jump).min(num_rows - 1);
        }
    }

    fn ensure_visible(&mut self, visible_rows: usize, area_width: u16) {
        if self.cursor_row < self.scroll_row {
            self.scroll_row = self.cursor_row;
        }
        if visible_rows > 0 && self.cursor_row >= self.scroll_row + visible_rows {
            self.scroll_row = self.cursor_row - visible_rows + 1;
        }

        let available = (area_width as usize).saturating_sub(self.row_num_width + 1);
        let vis_cols = self.visible_columns(self.scroll_col, available);

        if self.cursor_col < self.scroll_col {
            self.scroll_col = self.cursor_col;
        }
        if !vis_cols.is_empty() {
            let last_vis = vis_cols[vis_cols.len() - 1];
            if self.cursor_col > last_vis {
                let mut sc = self.scroll_col;
                loop {
                    let cols = self.visible_columns(sc, available);
                    if cols.is_empty() || *cols.last().unwrap() >= self.cursor_col {
                        break;
                    }
                    sc += 1;
                    if sc >= self.data().num_cols {
                        break;
                    }
                }
                self.scroll_col = sc;
            }
        }
    }

    fn visible_columns(&self, start_col: usize, available: usize) -> Vec<usize> {
        let data = self.data();
        let mut cols = Vec::new();
        let mut used = 0usize;
        for c in start_col..data.num_cols {
            let w = data.col_widths.get(c).copied().unwrap_or(3) + 1;
            if used + w > available && !cols.is_empty() {
                break;
            }
            used += w;
            cols.push(c);
        }
        cols
    }

    /// Column letter for index, using col_names if they look like headers, else generated.
    fn col_letter(&self, c: usize) -> String {
        util::col_to_letter(c)
    }

    fn draw(&self, frame: &mut Frame) {
        let area = frame.area();
        if self.multi_sheet {
            let chunks = Layout::vertical([
                Constraint::Length(1),
                Constraint::Length(1),
                Constraint::Min(3),
                Constraint::Length(1),
            ])
            .split(area);

            self.draw_title(frame, chunks[0]);
            self.draw_tab_bar(frame, chunks[1]);
            self.draw_grid(frame, chunks[2]);
            self.draw_status(frame, chunks[3]);
        } else {
            let chunks = Layout::vertical([
                Constraint::Length(1),
                Constraint::Min(3),
                Constraint::Length(1),
            ])
            .split(area);

            self.draw_title(frame, chunks[0]);
            self.draw_grid(frame, chunks[1]);
            self.draw_status(frame, chunks[2]);
        }

        if self.show_help {
            self.draw_help(frame, area);
        }
    }

    fn draw_tab_bar(&self, frame: &mut Frame, area: Rect) {
        let mut spans = Vec::new();
        for (i, sheet) in self.sheets.iter().enumerate() {
            let label = if i < 9 {
                format!(" {}:{} ", i + 1, sheet.name)
            } else {
                format!(" {} ", sheet.name)
            };
            if i == self.active_sheet {
                spans.push(Span::styled(
                    label,
                    Style::default()
                        .fg(Color::Black)
                        .bg(Color::White)
                        .add_modifier(Modifier::BOLD),
                ));
            } else {
                spans.push(Span::styled(
                    label,
                    Style::default().fg(Color::Gray).bg(Color::DarkGray),
                ));
            }
            spans.push(Span::styled(" ", Style::default().bg(Color::Black)));
        }
        let line = Line::from(spans);
        let para = Paragraph::new(line).style(Style::default().bg(Color::Black));
        frame.render_widget(para, area);
    }

    fn draw_title(&self, frame: &mut Frame, area: Rect) {
        let data = self.data();
        let row_info = if let Some(total) = data.total_rows {
            format!(
                "{} rows x {} cols (showing {})",
                total, data.num_cols, data.num_rows
            )
        } else {
            format!("{} rows x {} cols", data.num_rows, data.num_cols)
        };

        let sheet_info = if self.multi_sheet {
            format!(" | {} sheets", self.sheets.len())
        } else {
            String::new()
        };

        let title = format!(" visigrid: {} | {}{} ", self.file_name, row_info, sheet_info);
        let para = Paragraph::new(Line::from(vec![Span::styled(
            title,
            Style::default()
                .fg(Color::Black)
                .bg(Color::Cyan)
                .add_modifier(Modifier::BOLD),
        )]))
        .style(Style::default().bg(Color::Cyan));
        frame.render_widget(para, area);
    }

    fn draw_grid(&self, frame: &mut Frame, area: Rect) {
        let data = self.data();
        if data.num_rows == 0 || data.num_cols == 0 {
            let msg =
                Paragraph::new("(empty)").style(Style::default().fg(Color::DarkGray));
            frame.render_widget(msg, area);
            return;
        }

        let grid_available =
            (area.width as usize).saturating_sub(self.row_num_width + 1);
        let vis_cols = self.visible_columns(self.scroll_col, grid_available);

        let header_height: u16 = 1;
        let data_height = area.height.saturating_sub(header_height);

        // Header line
        let gutter_blank = " ".repeat(self.row_num_width);
        let mut header_spans = vec![Span::styled(
            format!("{} ", gutter_blank),
            Style::default().fg(Color::DarkGray),
        )];
        for &c in &vis_cols {
            let name = data
                .col_names
                .get(c)
                .map(|s| s.as_str())
                .unwrap_or("?");
            let w = data.col_widths.get(c).copied().unwrap_or(3);
            let display = util::pad_right(&util::truncate_display(name, w), w);
            let style = if c == self.cursor_col {
                Style::default()
                    .fg(Color::Yellow)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default()
                    .fg(Color::Cyan)
                    .add_modifier(Modifier::BOLD)
            };
            header_spans.push(Span::styled(format!("{} ", display), style));
        }

        // Data lines
        let visible_rows = data_height as usize;
        let end_row = (self.scroll_row + visible_rows).min(data.num_rows);

        let mut lines: Vec<Line> = Vec::with_capacity(visible_rows + 1);
        lines.push(Line::from(header_spans));

        let query = self.search.as_deref().map(str::to_lowercase).unwrap_or_default();
        for r in self.scroll_row..end_row {
            let row_data = &data.rows[self.row_at(r)];
            let is_cursor_row = r == self.cursor_row;
            let file_row = data.file_row(self.row_at(r));

            let row_num_style = if is_cursor_row {
                Style::default()
                    .fg(Color::Yellow)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default().fg(Color::DarkGray)
            };

            let mut spans = vec![Span::styled(
                format!("{:>width$} ", file_row, width = self.row_num_width),
                row_num_style,
            )];

            for &c in &vis_cols {
                let value = row_data.get(c).map(|s| s.as_str()).unwrap_or("");
                let w = data.col_widths.get(c).copied().unwrap_or(3);
                let display = util::pad_right(&util::truncate_display(value, w), w);

                let style = if is_cursor_row && c == self.cursor_col {
                    Style::default()
                        .fg(Color::Black)
                        .bg(Color::White)
                        .add_modifier(Modifier::BOLD)
                } else if is_cursor_row {
                    Style::default().fg(Color::White)
                } else if ops::cell_matches(value, &query) {
                    Style::default().fg(Color::Black).bg(Color::Yellow)
                } else if c == self.cursor_col {
                    Style::default().fg(Color::Gray)
                } else {
                    Style::default().fg(Color::Gray)
                };

                spans.push(Span::styled(format!("{} ", display), style));
            }

            lines.push(Line::from(spans));
        }

        let para = Paragraph::new(lines);
        frame.render_widget(para, area);
    }

    fn draw_status(&self, frame: &mut Frame, area: Rect) {
        let data = self.data();
        if let Some(p) = &self.prompt {
            let label = match p.kind {
                PromptKind::Search => "/".to_string(),
                PromptKind::Pivot => "pivot: ".to_string(),
            };
            let para = Paragraph::new(Line::from(vec![
                Span::styled(format!(" {label}{}", p.input), Style::default().fg(Color::White)),
                Span::styled("█", Style::default().fg(Color::Yellow)),
            ]))
            .style(Style::default().bg(Color::Black));
            frame.render_widget(para, area);
            return;
        }
        let cell_value = self.cell(self.cursor_row, self.cursor_col);

        let col_name = data
            .col_names
            .get(self.cursor_col)
            .map(|s| s.as_str())
            .unwrap_or("?");

        let file_row = data.file_row(self.row_at(self.cursor_row));
        let total = data.total_data_rows();

        // Column locator: show visible column range
        let grid_available =
            (area.width as usize).saturating_sub(self.row_num_width + 1);
        let vis_cols = self.visible_columns(self.scroll_col, grid_available);
        let col_range = if vis_cols.is_empty() {
            String::new()
        } else {
            let first = self.col_letter(*vis_cols.first().unwrap());
            let last = self.col_letter(*vis_cols.last().unwrap());
            if first == last {
                format!("Col {}", first)
            } else {
                format!("Cols {}..{}", first, last)
            }
        };

        let sheet_info = if self.multi_sheet {
            let name = &self.sheets[self.active_sheet].name;
            format!("  sheet: {} ({}/{})", name, self.active_sheet + 1, self.sheets.len())
        } else {
            String::new()
        };

        // Show formula when raw differs from display (i.e. cell contains a formula)
        let formula_info = data.raw.as_ref()
            .and_then(|raw| raw.get(self.row_at(self.cursor_row)))
            .and_then(|row| row.get(self.cursor_col))
            .filter(|raw| raw.starts_with('='))
            .map(|raw| format!("  {}", raw))
            .unwrap_or_default();

        let left = match &self.message {
            Some(m) => format!(" {m}"),
            None => format!(" {}{} = {:?}{}{}", col_name, file_row, cell_value, formula_info, sheet_info),
        };
        let right = format!(
            "Row {}/{}  {}  ?: help ",
            file_row, total, col_range
        );

        let room = (area.width as usize).saturating_sub(right.chars().count() + 1);
        let left = if util::display_width(&left) > room { util::truncate_display(&left, room) } else { left };
        let padding = (area.width as usize)
            .saturating_sub(util::display_width(&left) + right.chars().count());
        let status = format!("{}{:pad$}{}", left, "", right, pad = padding);

        let para = Paragraph::new(Line::from(vec![Span::styled(
            status,
            Style::default().fg(Color::Black).bg(Color::DarkGray),
        )]))
        .style(Style::default().bg(Color::DarkGray));
        frame.render_widget(para, area);
    }

    fn draw_help(&self, frame: &mut Frame, area: Rect) {
        let mut help_lines = vec![
            "",
            "  Navigation",
            "  ----------",
            "  arrows / hjkl    Move cursor",
            "  PgUp / PgDn      Page up/down",
            "  Home / g          First row",
            "  End  / G          Last row",
            "  0                 First column",
            "  $                 Last column",
        ];

        if self.multi_sheet {
            help_lines.extend_from_slice(&[
                "",
                "  Sheets",
                "  ------",
                "  Tab / Shift+Tab     Next/prev sheet",
                "  Ctrl+PgDn/PgUp     Next/prev sheet",
                "  1..9                Jump to sheet",
            ]);
        } else {
            help_lines.push("  Tab / Shift+Tab   Next/prev column");
        }

        help_lines.extend_from_slice(&[
            "",
            "  Find and summarize",
            "  ------------------",
            "  /                 Search all cells",
            "  n / N             Next / previous match",
            "  [ / ]             Sort column asc / desc",
            "  R                 Restore file order",
            "  F                 Frequency of column",
            "  P                 Pivot (rows= column=",
            "                      values=sum:Field)",
            "",
            "  General",
            "  -------",
            "  q / Esc           Close tab / quit",
            "  ?                 Toggle this help",
            "",
        ]);
        let help_width: u16 = 44;
        let help_height: u16 = help_lines.len() as u16;

        let x = area
            .width
            .saturating_sub(help_width)
            / 2;
        let y = area
            .height
            .saturating_sub(help_height)
            / 2;
        let popup = Rect::new(
            area.x + x,
            area.y + y,
            help_width.min(area.width),
            help_height.min(area.height),
        );

        let lines: Vec<Line> = help_lines
            .iter()
            .map(|s| {
                Line::from(Span::styled(
                    *s,
                    Style::default().fg(Color::White),
                ))
            })
            .collect();

        let block = Block::default()
            .borders(Borders::ALL)
            .border_style(Style::default().fg(Color::Cyan))
            .title(" Keybindings ")
            .title_style(
                Style::default()
                    .fg(Color::Cyan)
                    .add_modifier(Modifier::BOLD),
            )
            .style(Style::default().bg(Color::Black));

        frame.render_widget(Clear, popup);
        let para = Paragraph::new(lines).block(block);
        frame.render_widget(para, popup);
    }
}

/// Run the interactive TUI viewer for a single table.
pub fn run(data: PeekData, file_name: String) -> Result<(), String> {
    let app = TuiApp::new(data, file_name);
    run_app(app)
}

/// Run the interactive TUI viewer for a multi-sheet workbook.
pub fn run_multi(sheets: Vec<SheetData>, file_name: String, initial_sheet: usize) -> Result<(), String> {
    let app = TuiApp::new_multi(sheets, file_name, initial_sheet);
    run_app(app)
}

fn run_app(mut app: TuiApp) -> Result<(), String> {
    terminal::enable_raw_mode()
        .map_err(|e| format!("failed to enable raw mode: {}", e))?;
    stdout()
        .execute(EnterAlternateScreen)
        .map_err(|e| format!("failed to enter alternate screen: {}", e))?;

    struct Cleanup;
    impl Drop for Cleanup {
        fn drop(&mut self) {
            let _ = stdout().execute(LeaveAlternateScreen);
            let _ = terminal::disable_raw_mode();
        }
    }
    let _cleanup = Cleanup;

    let backend = CrosstermBackend::new(stdout());
    let mut terminal =
        Terminal::new(backend).map_err(|e| format!("failed to create terminal: {}", e))?;

    loop {
        let term_size = terminal
            .size()
            .map(|s| Rect::new(0, 0, s.width, s.height))
            .unwrap_or_default();
        let chrome = if app.multi_sheet { 4u16 } else { 3u16 };
        let visible_rows = term_size.height.saturating_sub(chrome) as usize;
        app.ensure_visible(visible_rows, term_size.width);

        terminal
            .draw(|frame| app.draw(frame))
            .map_err(|e| format!("draw error: {}", e))?;

        if event::poll(Duration::from_millis(100))
            .map_err(|e| format!("event poll error: {}", e))?
        {
            if let Event::Key(key) =
                event::read().map_err(|e| format!("event read error: {}", e))?
            {
                app.handle_key(key);
            }
        }

        if app.should_quit {
            break;
        }
    }

    Ok(())
}

/// Print data as a plain text table to stdout (no TUI, no raw mode).
pub fn print_plain(data: &PeekData, max_rows: usize) -> Result<(), String> {
    let out = io::stdout();
    let mut w = out.lock();
    let row_num_width = 6;
    let limit = if max_rows == 0 { data.num_rows } else { max_rows.min(data.num_rows) };

    // Header
    write!(w, "{:>width$} ", "", width = row_num_width)
        .map_err(|e| e.to_string())?;
    for c in 0..data.num_cols {
        let name = data.col_names.get(c).map(|s| s.as_str()).unwrap_or("?");
        let cw = data.col_widths.get(c).copied().unwrap_or(3);
        write!(w, "{} ", util::pad_right(&util::truncate_display(name, cw), cw))
            .map_err(|e| e.to_string())?;
    }
    writeln!(w).map_err(|e| e.to_string())?;

    // Separator
    write!(w, "{:->width$}-", "", width = row_num_width)
        .map_err(|e| e.to_string())?;
    for c in 0..data.num_cols {
        let cw = data.col_widths.get(c).copied().unwrap_or(3);
        write!(w, "{}-", "-".repeat(cw)).map_err(|e| e.to_string())?;
    }
    writeln!(w).map_err(|e| e.to_string())?;

    // Rows
    for r in 0..limit {
        let file_row = data.file_row(r);
        let row_data = &data.rows[r];
        write!(w, "{:>width$} ", file_row, width = row_num_width)
            .map_err(|e| e.to_string())?;
        for c in 0..data.num_cols {
            let value = row_data.get(c).map(|s| s.as_str()).unwrap_or("");
            let cw = data.col_widths.get(c).copied().unwrap_or(3);
            write!(w, "{} ", util::pad_right(&util::truncate_display(value, cw), cw))
                .map_err(|e| e.to_string())?;
        }
        writeln!(w).map_err(|e| e.to_string())?;
    }

    if limit < data.total_data_rows() {
        writeln!(w, "... (showing {} of {} rows)", limit, data.total_data_rows())
            .map_err(|e| e.to_string())?;
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use ratatui::backend::TestBackend;

    #[test]
    fn parquet_preview_renders_and_navigates_with_schema_headers() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/parquet_orders.parquet");
        let data = data::load_parquet(&path, 2, false, 0, false).unwrap();
        let mut app = TuiApp::new(data, "orders.parquet".into());
        let mut terminal = Terminal::new(TestBackend::new(120, 30)).unwrap();
        terminal.draw(|frame| app.draw(frame)).unwrap();
        let screen: String = terminal.backend().buffer().content.iter()
            .map(|cell| cell.symbol()).collect();
        assert!(screen.contains("order_id"));
        assert!(screen.contains("007"));
        assert!(screen.contains("2026-09-01 14:02:00"));
        assert!(screen.contains("3 rows x 5 cols (showing 2)"));

        app.handle_key(KeyEvent::new(KeyCode::Char('G'), KeyModifiers::NONE));
        assert_eq!(app.data().file_row(app.cursor_row), 2);
        app.handle_key(KeyEvent::new(KeyCode::Char('?'), KeyModifiers::NONE));
        terminal.draw(|frame| app.draw(frame)).unwrap();
        let screen: String = terminal.backend().buffer().content.iter()
            .map(|cell| cell.symbol()).collect();
        assert!(screen.contains("Keybindings"));
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        assert!(!app.should_quit, "first Escape dismisses help");
        app.handle_key(KeyEvent::new(KeyCode::Char('q'), KeyModifiers::NONE));
        assert!(app.should_quit);
    }

    fn press(app: &mut TuiApp, keys: &str) {
        for ch in keys.chars() {
            let code = match ch {
                '\n' => KeyCode::Enter,
                '\x1b' => KeyCode::Esc,
                c => KeyCode::Char(c),
            };
            app.handle_key(KeyEvent::new(code, KeyModifiers::NONE));
        }
    }

    fn screen(app: &TuiApp) -> String {
        let mut terminal = Terminal::new(TestBackend::new(100, 20)).unwrap();
        terminal.draw(|frame| app.draw(frame)).unwrap();
        let buf = terminal.backend().buffer();
        (0..buf.area.height)
            .map(|y| (0..buf.area.width).map(|x| buf[(x, y)].symbol()).collect::<String>())
            .collect::<Vec<_>>()
            .join("\n")
    }

    #[test]
    fn keys_sort_search_frequency_and_pivot_tabs() {
        let f = tempfile::Builder::new().suffix(".csv").tempfile().unwrap();
        std::fs::write(f.path(), "Region,Amount\nWest,10\nEast,5\nWest,2.5\nNorth,40\n").unwrap();
        let data = data::load_csv(f.path(), b',', true, 0, 0).unwrap();
        let mut app = TuiApp::new(data, "sales.csv".into());

        // Sort by Amount descending: North (file row 5) comes first.
        press(&mut app, "l]");
        assert_eq!(app.cell(0, 0), "North");
        assert!(screen(&app).contains("sorted by Amount descending"));
        press(&mut app, "R");
        assert_eq!(app.cell(0, 0), "West");

        // Search lands on the next match and n wraps.
        press(&mut app, "g0/east\n");
        assert_eq!((app.cursor_row, app.cursor_col), (1, 0));
        assert!(screen(&app).contains("/east  1 match"));
        press(&mut app, "/nope\n");
        assert_eq!((app.cursor_row, app.cursor_col), (1, 0), "a miss leaves the cursor");

        // Frequency tab, then q returns to the source instead of quitting.
        press(&mut app, "0F");
        assert_eq!(app.sheets.len(), 2);
        assert_eq!(app.cell(0, 0), "West");
        assert_eq!(app.cell(0, 1), "2");
        assert!(screen(&app).contains("freq Region"));
        press(&mut app, "q");
        assert_eq!((app.sheets.len(), app.active_sheet, app.should_quit), (1, 0, false));
        assert_eq!((app.cursor_row, app.cursor_col), (1, 0), "returning keeps the source cursor");

        // Pivot prompt starts prefilled from the cursor column; edit and run.
        press(&mut app, "P");
        app.prompt.as_mut().unwrap().input = "rows=Region values=sum:Amount".into();
        press(&mut app, "\n");
        assert_eq!(app.sheets[1].name, "Sum of Amount");
        assert_eq!(app.cell(2, 1), "12.50");
        assert_eq!(app.cell(3, 0), "Grand Total");
        assert_eq!(app.cell(3, 1), "57.50");
        press(&mut app, "P\x1b");
        assert!(app.prompt.is_none());
        press(&mut app, "qq");
        assert!(app.should_quit);
    }
}
