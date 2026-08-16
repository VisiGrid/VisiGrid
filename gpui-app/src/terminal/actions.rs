//! Terminal panel actions on Spreadsheet: spawn, focus guards, AI terminal
//! launch, and importing terminal output into the grid.
//! Extracted from app.rs 2026-07-30 (pure move).

use gpui::*;

use crate::ai_cli::{detect_ai_cli, shell_quote, ALL_AI_CLIS};
use crate::app::{BottomPanelTab, Spreadsheet};
use crate::settings::{update_user_settings, user_settings, TipId};

impl Spreadsheet {
    /// Returns true if the terminal panel currently has keyboard focus.
    /// Used by action handlers to bail out so keys go to the PTY instead.
    pub fn terminal_has_focus(&self, window: &Window) -> bool {
        self.terminal_focus_handle.is_focused(window) || self.terminal_focused
    }
    /// Guard for grid actions: if terminal has focus, propagate and return true.
    /// In debug builds, also asserts the seq-stamped focus invariant to catch
    /// dual-dispatch bugs where both terminal and grid handle the same key event.
    ///
    /// Use: `if this.guard_terminal_focus(window, cx, "ActionName") { return; }`
    #[inline]
    pub fn guard_terminal_focus(
        &self,
        window: &Window,
        cx: &mut Context<Self>,
        _action_name: &str,
    ) -> bool {
        if self.terminal_has_focus(window) {
            cx.propagate();
            return true;
        }
        #[cfg(debug_assertions)]
        crate::views::terminal_panel::assert_grid_not_conflicting(_action_name);
        false
    }
    /// Spawn a new PTY terminal session.
    pub fn spawn_terminal(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        use crate::terminal::pty::{TerminalEventProxy, spawn_pty};
        use crate::terminal::resolve_workspace_root;
        use alacritty_terminal::event::Event as TermEvent;

        // App Sandbox (Mac App Store build): the shell can't run sandboxed.
        // Mark the session exited so the panel shows its explanation instead
        // of spawn/exit looping on every Enter.
        if std::env::var_os("APP_SANDBOX_CONTAINER_ID").is_some() {
            self.terminal.exited = true;
            self.terminal.exit_code = None;
            cx.notify();
            return;
        }

        // Resolve workspace root and use it as CWD
        let root = resolve_workspace_root(self.current_file.as_deref());
        self.terminal.set_workspace_root(root.clone());
        self.terminal.last_sent_cwd = Some(root.clone());
        let cwd = Some(root);

        self.terminal.cwd = cwd.clone();

        // Compute initial cols/rows from panel size
        let cols = 80u16;
        let rows = ((self.terminal.height - 32.0) / 18.0).max(2.0) as u16; // 32px for header+padding

        // Create event channel — std::sync::mpsc for the PTY proxy (requires Sync sender)
        let (event_tx, event_rx) = std::sync::mpsc::channel();
        let proxy = TerminalEventProxy::new(event_tx);

        match spawn_pty(cwd, cols, rows, proxy) {
            Ok((term, sender, _join_handle)) => {
                self.terminal.term = Some(term);
                self.terminal.event_loop_sender = Some(sender);
                self.terminal.exited = false;
                self.terminal.exit_code = None;

                // Bridge: blocking std::sync::mpsc → async smol::channel
                // The PTY I/O thread sends events via std::sync::mpsc.
                // We cannot call blocking recv() in a GPUI async task (it blocks the UI thread).
                // Instead, a dedicated bridge thread drains the blocking receiver and forwards
                // events to an async-compatible smol channel.
                let (async_tx, async_rx) = smol::channel::unbounded();

                std::thread::Builder::new()
                    .name("terminal-event-bridge".to_string())
                    .spawn(move || {
                        while let Ok(event) = event_rx.recv() {
                            if async_tx.send_blocking(event).is_err() {
                                break; // async receiver dropped
                            }
                        }
                    })
                    .expect("failed to spawn terminal event bridge thread");

                // Async task drains the smol channel without blocking the UI
                cx.spawn(async move |this, cx| {
                    loop {
                        // Await next event — non-blocking on the executor
                        let event = match async_rx.recv().await {
                            Ok(e) => e,
                            Err(_) => break, // Channel closed
                        };

                        let mut should_notify = false;
                        let mut exit_code = None;

                        // Process this event and drain any additional pending events
                        let mut process = |event: TermEvent| {
                            match event {
                                TermEvent::Wakeup => {
                                    should_notify = true;
                                }
                                TermEvent::ChildExit(code) => {
                                    exit_code = Some(code);
                                    should_notify = true;
                                }
                                TermEvent::Exit => {
                                    exit_code = Some(0);
                                    should_notify = true;
                                }
                                _ => {}
                            }
                        };

                        process(event);
                        while let Ok(e) = async_rx.try_recv() {
                            process(e);
                        }

                        if let Some(code) = exit_code {
                            let _ = this.update(cx, |app, cx| {
                                app.terminal.exited = true;
                                app.terminal.exit_code = Some(code);
                                app.terminal.event_loop_sender = None;
                                app.terminal.term = None;
                                cx.notify();
                            });
                            break;
                        }

                        if should_notify {
                            let _ = this.update(cx, |app, cx| {
                                app.terminal.bump_output_epoch();
                                // Phase 5: debounce settle timer for auto-detect
                                if app.terminal.watching_for_result {
                                    app.terminal.result_settle_task = None;
                                    let gen = app.terminal.watch_generation;
                                    let task = cx.spawn(async move |this, cx| {
                                        smol::Timer::after(std::time::Duration::from_millis(500)).await;
                                        let _ = this.update(cx, |app, cx| {
                                            if app.terminal.watch_generation != gen { return; }
                                            app.try_extract_structured_result(cx);
                                        });
                                    });
                                    app.terminal.result_settle_task = Some(task);
                                }
                                cx.notify();
                            });
                        }
                    }
                }).detach();
            }
            Err(e) => {
                self.terminal.exited = true;
                self.status_message = Some(format!("Failed to spawn terminal: {}", e));
                cx.notify();
            }
        }
    }
    /// Launch a detected AI CLI in the terminal panel.
    pub fn launch_ai_terminal(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let pref = crate::settings::user_settings(cx).terminal.preferred_ai_cli;
        let Some(cli) = detect_ai_cli(pref) else {
            let mut msg = "No AI CLI found. Install one: npm i -g @anthropic-ai/claude-code | @openai/codex | @google/gemini-cli".to_string();
            crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::Error { category: "no_cli" });
            // One-time extra hint
            if !user_settings(cx).is_tip_dismissed(TipId::AiExplainHint) {
                msg.push_str(" Then run \"Generate AI Context Files\" to set up persistent instructions.");
                update_user_settings(cx, |s| s.dismiss_tip(TipId::AiExplainHint));
            }
            self.status_message = Some(msg);
            cx.notify();
            return;
        };

        self.open_terminal(window, cx);
        self.terminal.write_to_pty(b"\x15"); // Ctrl+U clear line

        // cd to workbook directory so the CLI discovers root context files
        // (.visigrid/CLAUDE.md, AGENTS.md, GEMINI.md)
        if let Some(ref p) = self.current_file {
            if let Some(dir) = p.parent() {
                let cd_cmd = format!("cd {} && ", shell_quote(&dir.display().to_string()));
                self.terminal.write_to_pty(cd_cmd.as_bytes());
            }
        }
        self.terminal.write_to_pty(cli.binary().as_bytes());
        self.terminal.write_to_pty(b"\n");
        crate::ai_metrics::record(crate::ai_metrics::AiMetricEvent::LaunchAi { cli: cli.binary() });
        self.status_message = Some(format!(
            "Starting {}… If it needs login, run: {} login",
            cli.display_name(), cli.binary()
        ));
        cx.notify();
    }
    /// Open terminal panel (show + focus, never toggles/hides).
    pub fn open_terminal(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        self.bottom_panel_visible = true;
        self.bottom_panel_tab = BottomPanelTab::Terminal;
        self.lua_console.visible = false;
        self.terminal.visible = true;
        if self.terminal.term.is_none() && !self.terminal.exited {
            self.spawn_terminal(window, cx);
        } else {
            self.terminal.ensure_cwd();
        }
        // One-time tip: suggest AI features when terminal first opens
        if !user_settings(cx).is_tip_dismissed(TipId::AiTerminal) {
            self.status_message = Some(
                "Tip: Ctrl+K \u{2192} \"AI: Explain Selection\" to analyze data with an AI CLI.".into()
            );
            update_user_settings(cx, |s| s.dismiss_tip(TipId::AiTerminal));
        }
        self.terminal_focused = true;
        window.focus(&self.terminal_focus_handle, cx);
        cx.notify();
    }
    /// Manual import: extract structured result from terminal (no watching check).
    pub fn import_terminal_output(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        self.try_extract_structured_result(cx);
    }
}

impl Spreadsheet {
    /// Paste the current selection as delimited TSV into the terminal PTY.
    pub fn paste_selection_context(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        if !self.terminal.visible || self.terminal.term.is_none() {
            self.status_message = Some("Open the terminal first.".into());
            cx.notify();
            return;
        }

        let wb = self.workbook.read(cx);
        let sheet = wb.active_sheet();
        let ((top_row, left_col), (bottom_row, right_col)) = self.selection_range();

        let sel_rows = bottom_row - top_row + 1;
        let sel_cols = right_col - left_col + 1;
        let capped_rows = sel_rows.min(Self::PASTE_MAX_ROWS);
        let capped_cols = sel_cols.min(Self::PASTE_MAX_COLS);
        let truncated = capped_rows < sel_rows || capped_cols < sel_cols;

        let mut lines = Vec::new();
        let mut total_chars = 0usize;
        let mut char_truncated = false;

        for r in top_row..top_row + capped_rows {
            let mut cells = Vec::new();
            for c in left_col..left_col + capped_cols {
                cells.push(sheet.get_display(r, c));
            }
            let line = cells.join("\t");
            total_chars += line.len() + 1;
            if total_chars > Self::PASTE_MAX_CHARS {
                char_truncated = true;
                break;
            }
            lines.push(line);
        }
        let tsv = lines.join("\n");

        if tsv.trim().is_empty() {
            self.status_message = Some("Selection is empty.".into());
            cx.notify();
            return;
        }

        // Build delimited block
        let mut block = String::new();
        block.push_str("# VisiGrid selection (TSV)\n");
        if truncated || char_truncated {
            block.push_str(&format!(
                "# Truncated: showing {}×{} of {}×{} (cap: {}×{} or {}k chars)\n",
                capped_rows.min(lines.len()), capped_cols,
                sel_rows, sel_cols,
                Self::PASTE_MAX_ROWS, Self::PASTE_MAX_COLS, Self::PASTE_MAX_CHARS / 1000
            ));
            block.push_str("# Tip: export to file with vgrid export, then paste the path.\n");
        }
        block.push_str(&tsv);
        block.push('\n');

        self.write_to_pty_bracketed(&block, cx);
        let actual_rows = lines.len();
        let mut msg = format!(
            "Pasted {}×{} selection into terminal.{}",
            actual_rows, capped_cols,
            if truncated || char_truncated { " (truncated)" } else { "" }
        );
        self.maybe_append_ai_paste_tip(&mut msg, cx);
        self.status_message = Some(msg);
        cx.notify();
    }
    /// Paste detected header row into the terminal PTY.
    pub fn paste_headers_context(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        if !self.terminal.visible || self.terminal.term.is_none() {
            self.status_message = Some("Open the terminal first.".into());
            cx.notify();
            return;
        }

        let wb = self.workbook.read(cx);
        let sheet = wb.active_sheet();
        let (start_row, start_col, end_row, end_col) = crate::ai::find_used_range(sheet);

        // Find the header row: start from top of used range, scan down up to 20 rows
        let mut header_row = start_row;
        for r in start_row..(start_row + 20).min(end_row + 1) {
            let row_vals: Vec<String> = (start_col..=end_col)
                .map(|c| sheet.get_display(r, c))
                .collect();
            if crate::ai::looks_like_header_row(&row_vals) {
                header_row = r;
                break;
            }
        }

        let headers: Vec<String> = (start_col..=end_col)
            .map(|c| sheet.get_display(header_row, c))
            .filter(|s| !s.is_empty())
            .collect();

        if headers.is_empty() {
            self.status_message = Some("No headers found.".into());
            cx.notify();
            return;
        }

        let block = format!(
            "# VisiGrid headers (row {})\n{}\n",
            header_row + 1, // 1-indexed for display
            headers.join(", ")
        );
        self.write_to_pty_bracketed(&block, cx);
        let mut msg = format!(
            "Pasted {} headers (row {}) into terminal.",
            headers.len(), header_row + 1
        );
        self.maybe_append_ai_paste_tip(&mut msg, cx);
        self.status_message = Some(msg);
        cx.notify();
    }
    /// Paste the current workbook file path into the terminal PTY.
    pub fn paste_file_path_context(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        if !self.terminal.visible || self.terminal.term.is_none() {
            self.status_message = Some("Open the terminal first.".into());
            cx.notify();
            return;
        }

        let path = match &self.current_file {
            Some(p) => p.display().to_string(),
            None => {
                self.status_message = Some("Workbook not saved yet — no file path.".into());
                cx.notify();
                return;
            }
        };

        let block = format!("# VisiGrid workbook path\n{}\n", path);
        self.write_to_pty_bracketed(&block, cx);
        let mut msg = "Pasted file path into terminal.".to_string();
        self.maybe_append_ai_paste_tip(&mut msg, cx);
        self.status_message = Some(msg);
        cx.notify();
    }
    /// Paste full VisiGrid context into the terminal in one shot:
    /// file path, sheet name, selection shape, headers, and selection TSV (capped).
    pub fn paste_full_context(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        if !self.terminal.visible || self.terminal.term.is_none() {
            self.status_message = Some("Open the terminal first.".into());
            cx.notify();
            return;
        }

        let wb = self.workbook.read(cx);
        let sheet = wb.active_sheet();
        let sheet_name = sheet.name.clone();
        let (used_start_row, used_start_col, used_end_row, used_end_col) =
            crate::ai::find_used_range(sheet);

        // Detect headers
        let mut header_row = used_start_row;
        for r in used_start_row..(used_start_row + 20).min(used_end_row + 1) {
            let row_vals: Vec<String> = (used_start_col..=used_end_col)
                .map(|c| sheet.get_display(r, c))
                .collect();
            if crate::ai::looks_like_header_row(&row_vals) {
                header_row = r;
                break;
            }
        }
        let headers: Vec<String> = (used_start_col..=used_end_col)
            .map(|c| sheet.get_display(header_row, c))
            .filter(|s| !s.is_empty())
            .collect();

        // Extract selection TSV (capped)
        let ((top_row, left_col), (bottom_row, right_col)) = self.selection_range();
        let sel_rows = bottom_row - top_row + 1;
        let sel_cols = right_col - left_col + 1;
        let capped_rows = sel_rows.min(Self::PASTE_MAX_ROWS);
        let capped_cols = sel_cols.min(Self::PASTE_MAX_COLS);
        let truncated = capped_rows < sel_rows || capped_cols < sel_cols;

        let mut lines = Vec::new();
        let mut total_chars = 0usize;
        let mut char_truncated = false;
        for r in top_row..top_row + capped_rows {
            let mut cells = Vec::new();
            for c in left_col..left_col + capped_cols {
                cells.push(sheet.get_display(r, c));
            }
            let line = cells.join("\t");
            total_chars += line.len() + 1;
            if total_chars > Self::PASTE_MAX_CHARS {
                char_truncated = true;
                break;
            }
            lines.push(line);
        }
        drop(wb);

        let tsv = lines.join("\n");
        let actual_rows = lines.len();

        // Build combined block
        let mut block = String::new();
        block.push_str("# VisiGrid context\n");

        // File path
        if let Some(p) = &self.current_file {
            block.push_str(&format!("# File: {}\n", p.display()));
        }

        // Sheet + used range
        block.push_str(&format!(
            "# Sheet: \"{}\"  Used range: {}x{}\n",
            sheet_name, used_end_row + 1, used_end_col + 1,
        ));

        // Headers
        if !headers.is_empty() {
            block.push_str(&format!(
                "# Headers (row {}): {}\n",
                header_row + 1,
                headers.join(", ")
            ));
        }

        // Selection info
        block.push_str(&format!(
            "# Selection: {}x{} starting at {}\n",
            sel_rows, sel_cols,
            crate::ai::cell_ref(top_row, left_col),
        ));

        if truncated || char_truncated {
            block.push_str(&format!(
                "# Truncated: showing {}x{} of {}x{} (cap: {}x{} or {}k chars)\n",
                actual_rows, capped_cols,
                sel_rows, sel_cols,
                Self::PASTE_MAX_ROWS, Self::PASTE_MAX_COLS, Self::PASTE_MAX_CHARS / 1000
            ));
        }

        // TSV data
        if !tsv.trim().is_empty() {
            block.push_str(&tsv);
            block.push('\n');
        }

        self.write_to_pty_bracketed(&block, cx);
        self.status_message = Some(format!(
            "Pasted full context ({}x{}{}) into terminal.",
            actual_rows, capped_cols,
            if truncated || char_truncated { ", truncated" } else { "" }
        ));
        cx.notify();
    }
    /// Paste full VisiGrid context + analysis prompt into the terminal.
    ///
    /// Writes instructions to `.visigrid/CLAUDE.md` for the AI CLI to discover.
    /// Only pastes the selection data and a short prompt into the terminal.
    pub(crate) fn paste_full_context_with_prompt(&mut self, cx: &mut Context<Self>) {
        if !self.terminal.visible || self.terminal.term.is_none() {
            return;
        }

        let analysis_instructions = "\
## Task: Analyze Selection

Summarize what this data shows. Detect anomalies, outliers, or suspicious patterns.
Suggest formulas or next steps to validate it.
";

        let Some((sheet_name, headers, tsv, sel_rows, sel_cols, ..)) =
            self.write_ai_context_and_collect_data(cx, analysis_instructions)
        else {
            return;
        };

        // Paste a visible prompt (seen by the user AND sent to the AI CLI).
        // Full analysis instructions are in .visigrid/CLAUDE.md.
        let mut block = String::new();
        block.push_str(&format!(
            "# VisiGrid: Analyze Selection\n# Sheet: \"{}\" | Selection: {}x{}",
            sheet_name, sel_rows, sel_cols,
        ));
        if !headers.is_empty() {
            block.push_str(&format!(" | Headers: {}", headers.join(", ")));
        }
        block.push_str("\n# Full context and instructions are in .visigrid/CLAUDE.md\n");
        if !tsv.trim().is_empty() {
            block.push_str(&tsv);
            block.push('\n');
        }
        block.push_str("\nAnalyze this data. Detect anomalies, outliers, or suspicious patterns.\n");

        // Write directly — AI CLI doesn't support bracketed paste escapes.
        self.terminal.write_to_pty(block.as_bytes());
        self.status_message = Some("Sent selection to AI for analysis.".into());
        cx.notify();
    }
}
