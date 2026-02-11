pub mod data;

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
    /// Whether this is a multi-sheet workbook
    multi_sheet: bool,
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
        }
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
        self.active_sheet = idx;
        self.cursor_row = 0;
        self.cursor_col = 0;
        self.scroll_row = 0;
        self.scroll_col = 0;
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

        match key.code {
            KeyCode::Char('q') | KeyCode::Esc => self.should_quit = true,
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

        for r in self.scroll_row..end_row {
            let row_data = &data.rows[r];
            let is_cursor_row = r == self.cursor_row;
            let file_row = data.file_row(r);

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
        let cell_value = data
            .rows
            .get(self.cursor_row)
            .and_then(|row| row.get(self.cursor_col))
            .map(|s| s.as_str())
            .unwrap_or("");

        let col_name = data
            .col_names
            .get(self.cursor_col)
            .map(|s| s.as_str())
            .unwrap_or("?");

        let file_row = data.file_row(self.cursor_row);
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

        let left = format!(" {}{} = {:?}{}", col_name, file_row, cell_value, sheet_info);
        let right = format!(
            "Row {}/{}  {}  ?: help ",
            file_row, total, col_range
        );

        let padding = (area.width as usize)
            .saturating_sub(left.chars().count() + right.chars().count());
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
            "  General",
            "  -------",
            "  q / Esc           Quit",
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

/// Run the interactive TUI viewer for a single CSV/TSV file.
pub fn run(data: PeekData, file_name: String) -> Result<(), String> {
    let app = TuiApp::new(data, file_name);
    run_app(app)
}

/// Run the interactive TUI viewer for a multi-sheet .sheet workbook.
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

    if limit < data.num_rows {
        writeln!(w, "... ({} more rows)", data.num_rows - limit)
            .map_err(|e| e.to_string())?;
    }

    Ok(())
}
