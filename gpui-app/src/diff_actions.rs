//! Running and refreshing `vgrid diff` from the app: command capture,
//! re-runs, and result loading.
//! Extracted from app.rs 2026-07-30 (pure move).

use gpui::*;

use crate::app::Spreadsheet;

impl Spreadsheet {
    /// Palette command: Run vgrid diff --json — prompts for the comparison file.
    pub fn run_vgrid_diff_json(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let Some(ref current) = self.current_file else {
            self.status_message = Some("Save workbook first.".into());
            cx.notify();
            return;
        };
        let left = current.display().to_string();

        let options = PathPromptOptions {
            files: true,
            directories: false,
            multiple: false,
            prompt: Some("Select file to compare against".into()),
        };
        let future = cx.prompt_for_paths(options);

        // Open terminal while the picker is up so it's ready when we get the path
        self.open_terminal(window, cx);

        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(paths))) = future.await {
                if let Some(right_path) = paths.first() {
                    let right = right_path.display().to_string();
                    let _ = this.update(cx, |this: &mut Self, cx| {
                        // Build diff command
                        let cmd = format!("vgrid diff \"{}\" \"{}\" --json\n", left, right);
                        this.terminal.write_to_pty(cmd.as_bytes());

                        this.terminal.watching_for_result = true;
                        this.terminal.watch_generation += 1;
                        this.terminal.pending_result = None;
                        this.terminal.result_settle_task = None;
                        this.terminal.last_injected_command = Some(cmd.trim().to_string());

                        // Explicit Left/Right in status — no ambiguity
                        let left_name = std::path::Path::new(&left)
                            .file_name()
                            .map(|n| n.to_string_lossy().to_string())
                            .unwrap_or_else(|| left.clone());
                        let right_name = right_path.file_name()
                            .map(|n| n.to_string_lossy().to_string())
                            .unwrap_or_else(|| right.clone());
                        this.status_message = Some(format!(
                            "Comparing: Left: {} | Right: {}", left_name, right_name
                        ));
                        cx.notify();
                    });
                }
            }
        }).detach();
    }
    /// Open the newest diff-*.json in the workspace as a "Diff Results" sheet.
    pub fn open_diff_results(&mut self, _window: &mut Window, cx: &mut Context<Self>) {
        let workspace = match self.terminal.workspace_root.as_ref() {
            Some(ws) => ws.clone(),
            None => {
                self.status_message = Some(
                    "No workspace root set. Open a file or start the terminal first.".to_string(),
                );
                cx.notify();
                return;
            }
        };

        let diff_path = match crate::diff_view::find_latest_diff_file(&workspace) {
            Some(p) => p,
            None => {
                self.status_message = Some(format!(
                    "No diff results found in {}. Run vgrid diff with --output first.",
                    workspace.display(),
                ));
                cx.notify();
                return;
            }
        };

        let bytes = match std::fs::read(&diff_path) {
            Ok(b) => b,
            Err(e) => {
                self.status_message = Some(format!(
                    "Could not read {}: {}",
                    diff_path.display(),
                    e,
                ));
                cx.notify();
                return;
            }
        };

        let parsed = match crate::diff_view::parse_diff_json(&bytes) {
            Ok(p) => p,
            Err(e) => {
                self.status_message = Some(format!(
                    "Error parsing {}: {}",
                    diff_path.display(),
                    e,
                ));
                cx.notify();
                return;
            }
        };

        crate::diff_view::populate_diff_sheet(self, parsed, &diff_path, &workspace, cx);
    }
    /// Validate and extract the diff command from the active Diff Results sheet.
    /// Returns `Some(command)` on success, or sets status_message and returns `None`.
    fn extract_diff_command(&mut self, cx: &mut Context<Self>) -> Option<String> {
        use crate::diff_view::{ROW_TITLE, ROW_WORKSPACE, ROW_COMMAND};

        let (is_diff_sheet, command) = {
            let wb = self.workbook.read(cx);
            let sheet = wb.active_sheet();
            let is_diff = sheet.get_display(ROW_TITLE, 0) == "Diff Report"
                && sheet.get_display(ROW_WORKSPACE, 0) == "Workspace";
            let cmd = sheet.get_display(ROW_COMMAND, 1);
            (is_diff, cmd)
        };

        if !is_diff_sheet {
            self.status_message = Some("Switch to a Diff Results sheet first.".into());
            cx.notify();
            return None;
        }

        let command = command.trim().to_string();
        if command.is_empty() {
            self.status_message = Some(
                "No command stored. This diff was created before Re-run support.".into()
            );
            cx.notify();
            return None;
        }

        if !command.starts_with("vgrid diff") {
            self.status_message = Some(
                "Stored command isn't a vgrid diff invocation.".into()
            );
            cx.notify();
            return None;
        }

        Some(command)
    }
    /// Re-run Diff: insert command into terminal (user presses Enter).
    pub fn rerun_diff(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let Some(command) = self.extract_diff_command(cx) else { return };
        self.open_terminal(window, cx);
        self.terminal.write_to_pty(b"\x15");
        self.terminal.write_to_pty(command.as_bytes());
        self.status_message = Some("Diff command inserted \u{2014} press Enter to run.".into());
        cx.notify();
    }
    /// Re-run Diff (Run): insert command into terminal and execute immediately.
    pub fn rerun_diff_run(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let Some(command) = self.extract_diff_command(cx) else { return };
        self.open_terminal(window, cx);
        self.terminal.write_to_pty(b"\x15");
        self.terminal.write_to_pty(command.as_bytes());
        self.terminal.write_to_pty(b"\n");
        self.status_message = Some("Running diff\u{2026}".into());
        cx.notify();
    }
    /// Open the exact diff file recorded in this Diff Results sheet (deterministic reload).
    pub fn open_this_diff_file(&mut self, cx: &mut Context<Self>) {
        use crate::diff_view::{ROW_TITLE, ROW_WORKSPACE, ROW_DIFF_FILE};

        // 1. Verify we're on a Diff Results sheet
        let (is_diff_sheet, diff_file_val, workspace_val) = {
            let wb = self.workbook.read(cx);
            let sheet = wb.active_sheet();
            let is_diff = sheet.get_display(ROW_TITLE, 0) == "Diff Report"
                && sheet.get_display(ROW_WORKSPACE, 0) == "Workspace";
            let df = sheet.get_display(ROW_DIFF_FILE, 1);
            let ws = sheet.get_display(ROW_WORKSPACE, 1);
            (is_diff, df, ws)
        };

        if !is_diff_sheet {
            self.status_message = Some("Switch to a Diff Results sheet first.".into());
            cx.notify();
            return;
        }

        let diff_file_str = diff_file_val.trim();
        if diff_file_str.is_empty() {
            self.status_message = Some(
                "No diff file path stored. This sheet was created before file tracking.".into()
            );
            cx.notify();
            return;
        }

        // 2. Resolve path (relative paths resolve against stored workspace)
        let diff_path = {
            let p = std::path::PathBuf::from(diff_file_str);
            if p.is_absolute() {
                p
            } else {
                let ws = workspace_val.trim();
                if ws.is_empty() {
                    p
                } else {
                    std::path::PathBuf::from(ws).join(p)
                }
            }
        };

        // 3. Read and parse
        let bytes = match std::fs::read(&diff_path) {
            Ok(b) => b,
            Err(e) => {
                self.status_message = Some(format!(
                    "Could not read {}: {}",
                    diff_path.display(), e,
                ));
                cx.notify();
                return;
            }
        };

        let parsed = match crate::diff_view::parse_diff_json(&bytes) {
            Ok(p) => p,
            Err(e) => {
                self.status_message = Some(format!(
                    "Error parsing {}: {}",
                    diff_path.display(), e,
                ));
                cx.notify();
                return;
            }
        };

        // 4. Rebuild sheet from the exact file
        let workspace = std::path::PathBuf::from(workspace_val.trim());
        crate::diff_view::populate_diff_sheet(self, parsed, &diff_path, &workspace, cx);
    }
    /// Refresh Diff Results: re-run command, wait for output, rebuild sheet with summary delta.
    pub fn refresh_diff_results(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        use crate::diff_view::{ROW_TITLE, ROW_WORKSPACE, ROW_COMMAND, ROW_DIFF_FILE, SUMMARY_START};

        // 1. Validate diff sheet and extract all provenance
        let (is_diff_sheet, command, workspace_str, diff_file_str, old_summary) = {
            let wb = self.workbook.read(cx);
            let sheet = wb.active_sheet();
            let is_diff = sheet.get_display(ROW_TITLE, 0) == "Diff Report"
                && sheet.get_display(ROW_WORKSPACE, 0) == "Workspace";
            let cmd = sheet.get_display(ROW_COMMAND, 1);
            let ws = sheet.get_display(ROW_WORKSPACE, 1);
            let df = sheet.get_display(ROW_DIFF_FILE, 1);
            // Capture old summary: Matched, Only left, Only right, Differences
            let old = [
                (sheet.get_display(SUMMARY_START + 2, 0), sheet.get_display(SUMMARY_START + 2, 1)),
                (sheet.get_display(SUMMARY_START + 3, 0), sheet.get_display(SUMMARY_START + 3, 1)),
                (sheet.get_display(SUMMARY_START + 4, 0), sheet.get_display(SUMMARY_START + 4, 1)),
                (sheet.get_display(SUMMARY_START + 5, 0), sheet.get_display(SUMMARY_START + 5, 1)),
            ];
            (is_diff, cmd, ws, df, old)
        };

        if !is_diff_sheet {
            self.status_message = Some("Switch to a Diff Results sheet first.".into());
            cx.notify();
            return;
        }

        let command = command.trim().to_string();
        if command.is_empty() {
            self.status_message = Some(
                "No command stored. This diff was created before Re-run support.".into()
            );
            cx.notify();
            return;
        }
        if !command.starts_with("vgrid diff") {
            self.status_message = Some(
                "Stored command isn't a vgrid diff invocation.".into()
            );
            cx.notify();
            return;
        }

        let diff_file_str = diff_file_str.trim().to_string();
        if diff_file_str.is_empty() {
            self.status_message = Some(
                "No diff file path stored. Use Re-run Diff (Run) instead.".into()
            );
            cx.notify();
            return;
        }

        // 2. Resolve diff file path
        let workspace_str = workspace_str.trim().to_string();
        let diff_path = {
            let p = std::path::PathBuf::from(&diff_file_str);
            if p.is_absolute() {
                p
            } else if !workspace_str.is_empty() {
                std::path::PathBuf::from(&workspace_str).join(p)
            } else {
                p
            }
        };

        // 3. Record pre-run mtime (if file exists already)
        let pre_mtime = std::fs::metadata(&diff_path)
            .ok()
            .and_then(|m| m.modified().ok());

        // 4. Execute command in terminal
        self.open_terminal(window, cx);
        self.terminal.write_to_pty(b"\x15");
        self.terminal.write_to_pty(command.as_bytes());
        self.terminal.write_to_pty(b"\n");
        self.status_message = Some("Refreshing diff\u{2026}".into());
        cx.notify();

        // 5. Spawn async poller: wait for file to settle, then rebuild
        let workspace = std::path::PathBuf::from(&workspace_str);
        cx.spawn(async move |this, cx| {
            let poll_interval = std::time::Duration::from_millis(200);
            let timeout = std::time::Duration::from_secs(10);
            let start = std::time::Instant::now();
            let mut last_size: Option<u64> = None;
            let mut stable_count: u32 = 0;

            loop {
                cx.background_executor().timer(poll_interval).await;

                if start.elapsed() > timeout {
                    let _ = this.update(cx, |app, cx| {
                        app.status_message = Some(
                            "Diff still running. Use Open Diff Results (This Sheet) when done.".into()
                        );
                        cx.notify();
                    });
                    return;
                }

                // Check if file has been updated since we started
                let meta = match std::fs::metadata(&diff_path) {
                    Ok(m) => m,
                    Err(_) => continue,
                };
                let mtime = meta.modified().ok();

                // File must have a newer mtime than before we ran the command
                let is_newer = match (mtime, pre_mtime) {
                    (Some(new), Some(old)) => new > old,
                    (Some(_), None) => true, // File didn't exist before
                    _ => false,
                };
                if !is_newer {
                    continue;
                }

                // Check size stability (settled when size unchanged for 2 consecutive polls)
                let size = meta.len();
                if Some(size) == last_size && size > 0 {
                    stable_count += 1;
                    if stable_count >= 2 {
                        break; // File is settled
                    }
                } else {
                    stable_count = 0;
                    last_size = Some(size);
                }
            }

            // 6. File settled — read, parse, rebuild with delta
            let _ = this.update(cx, |app, cx| {
                let bytes = match std::fs::read(&diff_path) {
                    Ok(b) => b,
                    Err(e) => {
                        app.status_message = Some(format!(
                            "Could not read {}: {}", diff_path.display(), e,
                        ));
                        cx.notify();
                        return;
                    }
                };

                let parsed = match crate::diff_view::parse_diff_json(&bytes) {
                    Ok(p) => p,
                    Err(e) => {
                        app.status_message = Some(format!(
                            "Error parsing {}: {}", diff_path.display(), e,
                        ));
                        cx.notify();
                        return;
                    }
                };

                // Capture new summary values before populate overwrites them
                let new_summary = [
                    &parsed.summary.matched,
                    &parsed.summary.only_left,
                    &parsed.summary.only_right,
                    &parsed.summary.diffs,
                ];

                // Build delta string
                let delta_parts: Vec<String> = old_summary.iter()
                    .zip(new_summary.iter())
                    .filter_map(|((label, old_val), new_val)| {
                        let old_n: i64 = old_val.trim().parse().unwrap_or(0);
                        let new_n: i64 = new_val.trim().parse().unwrap_or(0);
                        let diff = new_n - old_n;
                        if diff == 0 {
                            None
                        } else {
                            let sign = if diff > 0 { "+" } else { "" };
                            Some(format!("{}: {} \u{2192} {} ({}{})", label, old_n, new_n, sign, diff))
                        }
                    })
                    .collect();

                crate::diff_view::populate_diff_sheet(app, parsed, &diff_path, &workspace, cx);

                // Override the default status message with delta info
                if delta_parts.is_empty() {
                    app.status_message = Some("Diff refreshed \u{2014} no changes in summary.".into());
                } else {
                    app.status_message = Some(format!(
                        "Diff refreshed \u{2014} {}",
                        delta_parts.join(", "),
                    ));
                }
                cx.notify();
            });
        }).detach();
    }
}
