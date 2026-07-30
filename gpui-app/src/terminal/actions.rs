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
