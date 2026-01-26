use gpui::*;
use std::path::PathBuf;
use std::time::{Duration, Instant};
use visigrid_engine::workbook::Workbook;
use visigrid_io::{csv, json, native, xlsx};

use crate::app::{Spreadsheet, DocumentMeta};
use crate::settings::{load_doc_settings, save_doc_settings, DocumentSettings};

/// Delay before showing the import overlay (prevents flash for fast imports)
const OVERLAY_DELAY_MS: u64 = 150;

impl Spreadsheet {
    pub fn new_file(&mut self, cx: &mut Context<Self>) {
        self.workbook = Workbook::new();
        self.base_workbook = self.workbook.clone(); // Capture base state for replay
        self.rewind_preview = crate::app::RewindPreviewState::Off; // Reset preview state
        self.current_file = None;
        self.is_modified = false;
        self.doc_settings = DocumentSettings::default();  // Reset doc settings
        self.view_state.selected = (0, 0);
        self.view_state.selection_end = None;
        self.view_state.scroll_row = 0;
        self.view_state.scroll_col = 0;
        self.history.clear();
        self.bump_cells_rev();  // Invalidate cell search cache

        // Reset document meta for new document
        self.document_meta = DocumentMeta::default();
        self.request_title_refresh(cx);

        self.status_message = Some("New workbook created".to_string());
        cx.notify();
    }

    pub fn open_file(&mut self, cx: &mut Context<Self>) {
        let options = PathPromptOptions {
            files: true,
            directories: false,
            multiple: false,
            prompt: Some("Open".into()),
        };

        let future = cx.prompt_for_paths(options);
        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(paths))) = future.await {
                if let Some(path) = paths.first() {
                    let path = path.clone();
                    let _ = this.update(cx, |this, cx| {
                        this.load_file(&path, cx);
                    });
                }
            }
        })
        .detach();
    }

    pub fn load_file(&mut self, path: &PathBuf, cx: &mut Context<Self>) {
        let extension = path.extension().and_then(|s| s.to_str()).unwrap_or("");
        let ext_lower = extension.to_lowercase();

        // Excel files: background import for Pro, synchronous for Free
        if matches!(ext_lower.as_str(), "xlsx" | "xls" | "xlsb" | "xlsm" | "ods") {
            if visigrid_license::is_feature_enabled("fast_large_files") {
                self.start_excel_import(path, cx);
            } else {
                // Free users: synchronous import with upgrade hint
                self.status_message = Some("Importing... (upgrade to Pro for faster large file imports)".to_string());
                cx.notify();
                self.load_excel_sync(path, cx);
            }
            return;
        }

        // Non-Excel files: synchronous load (fast, no need for background)
        let result: Result<Workbook, String> = match ext_lower.as_str() {
            "csv" => csv::import(path)
                .map(|sheet| Workbook::from_sheets(vec![sheet], 0)),
            "tsv" => csv::import_tsv(path)
                .map(|sheet| Workbook::from_sheets(vec![sheet], 0)),
            "sheet" => native::load_workbook(path),
            _ => Err(format!("Unknown file type: {}", extension)),
        };

        match result {
            Ok(workbook) => {
                self.workbook = workbook;
                self.base_workbook = self.workbook.clone(); // Capture base state for replay
                self.rewind_preview = crate::app::RewindPreviewState::Off;
                self.import_result = None;
                self.import_filename = None;
                self.import_source_dir = None;
                // Load document settings from sidecar file
                self.doc_settings = load_doc_settings(path);
                self.view_state.selected = (0, 0);
                self.view_state.selection_end = None;
                self.view_state.scroll_row = 0;
                self.view_state.scroll_col = 0;
                self.history.clear();
                self.bump_cells_rev();
                self.add_recent_file(path);

                // Set up document identity (this also sets current_file, is_modified, save_point)
                self.finalize_load(path);
                self.request_title_refresh(cx);

                // Load VisiHub link if present (for .sheet files)
                if ext_lower == "sheet" {
                    match crate::hub::load_hub_link(path) {
                        Ok(Some(link)) => {
                            self.hub_link = Some(link);
                            self.hub_status = crate::hub::HubStatus::Idle; // Will check on first sync
                        }
                        Ok(None) => {
                            self.hub_link = None;
                            self.hub_status = crate::hub::HubStatus::Unlinked;
                        }
                        Err(_) => {
                            self.hub_link = None;
                            self.hub_status = crate::hub::HubStatus::Unlinked;
                        }
                    }
                } else {
                    self.hub_link = None;
                    self.hub_status = crate::hub::HubStatus::Unlinked;
                }

                // Update session with new file path
                self.update_session_cached(cx);

                let named_count = self.workbook.list_named_ranges().len();

                let status = if named_count > 0 {
                    format!("Opened: {} ({} named ranges)", path.display(), named_count)
                } else {
                    format!("Opened: {}", path.display())
                };
                self.status_message = Some(status);
            }
            Err(e) => {
                self.status_message = Some(format!("Error opening file: {}", e));
            }
        }
        cx.notify();
    }

    /// Start background Excel import with delayed overlay
    fn start_excel_import(&mut self, path: &PathBuf, cx: &mut Context<Self>) {
        let filename = path.file_name()
            .and_then(|n| n.to_str())
            .map(|s| s.to_string())
            .unwrap_or_else(|| "file".to_string());
        let source_dir = path.parent().map(|p| p.to_path_buf());

        // Set up import state
        self.import_in_progress = true;
        self.import_overlay_visible = false;
        self.import_started_at = Some(Instant::now());
        self.status_message = Some(format!("Importing {}...", filename));

        // Clone what we need for async tasks
        let path_for_import = path.clone();
        let path_for_recent = path.clone();
        let filename_for_completion = filename.clone();

        cx.notify();

        // Task A: Delayed overlay trigger
        // Show overlay only if import takes longer than OVERLAY_DELAY_MS
        cx.spawn(async move |this, cx| {
            cx.background_executor().timer(Duration::from_millis(OVERLAY_DELAY_MS)).await;
            let _ = this.update(cx, |this, cx| {
                if this.import_in_progress {
                    this.import_overlay_visible = true;
                    cx.notify();
                }
            });
        })
        .detach();

        // Task B: Actual import (runs in background)
        cx.spawn(async move |this, cx| {
            // Do the import on background thread
            let import_result = cx.background_executor()
                .spawn(async move {
                    xlsx::import(&path_for_import)
                })
                .await;

            // Update UI on main thread
            let _ = this.update(cx, |this, cx| {
                this.import_in_progress = false;
                this.import_overlay_visible = false;

                let duration_ms = this.import_started_at
                    .map(|t| t.elapsed().as_millis())
                    .unwrap_or(0);

                match import_result {
                    Ok((workbook, mut result)) => {
                        // Atomic swap: replace entire workbook
                        this.workbook = workbook;
                        this.base_workbook = this.workbook.clone(); // Capture base state for replay
                        this.rewind_preview = crate::app::RewindPreviewState::Off;
                        this.import_filename = Some(filename_for_completion.clone());
                        this.import_source_dir = source_dir;
                        this.doc_settings = DocumentSettings::default();
                        this.view_state.selected = (0, 0);
                        this.view_state.selection_end = None;
                        this.view_state.scroll_row = 0;
                        this.view_state.scroll_col = 0;
                        this.history.clear();
                        this.bump_cells_rev();
                        this.add_recent_file(&path_for_recent);

                        // Set up document identity (XLSX is native per spec)
                        this.finalize_load(&path_for_recent);
                        this.request_title_refresh(cx);

                        // Build status message with timing
                        let duration_str = if duration_ms >= 1000 {
                            format!("{:.2}s", duration_ms as f64 / 1000.0)
                        } else {
                            format!("{}ms", duration_ms)
                        };

                        let status = format!(
                            "Opened {} in {}",
                            filename_for_completion,
                            duration_str
                        );

                        // Store the duration we measured (more accurate than import's internal timing
                        // since it includes workbook construction)
                        result.import_duration_ms = duration_ms;
                        this.import_result = Some(result);
                        this.status_message = Some(status);
                    }
                    Err(e) => {
                        this.import_result = None;
                        this.import_filename = None;
                        this.import_source_dir = None;
                        this.status_message = Some(format!("Import failed: {}", e));
                    }
                }

                cx.notify();
            });
        })
        .detach();
    }

    /// Synchronous Excel import for Free users (no background processing)
    fn load_excel_sync(&mut self, path: &PathBuf, cx: &mut Context<Self>) {
        let filename = path.file_name()
            .and_then(|n| n.to_str())
            .map(|s| s.to_string())
            .unwrap_or_else(|| "file".to_string());
        let source_dir = path.parent().map(|p| p.to_path_buf());
        let start_time = std::time::Instant::now();

        match xlsx::import(path) {
            Ok((workbook, mut result)) => {
                let duration_ms = start_time.elapsed().as_millis();

                self.workbook = workbook;
                self.base_workbook = self.workbook.clone(); // Capture base state for replay
                self.rewind_preview = crate::app::RewindPreviewState::Off;
                self.import_filename = Some(filename.clone());
                self.import_source_dir = source_dir;
                self.doc_settings = DocumentSettings::default();
                self.view_state.selected = (0, 0);
                self.view_state.selection_end = None;
                self.view_state.scroll_row = 0;
                self.view_state.scroll_col = 0;
                self.history.clear();
                self.bump_cells_rev();
                self.add_recent_file(path);

                // Set up document identity (XLSX is native per spec)
                self.finalize_load(path);
                self.request_title_refresh(cx);

                let duration_str = if duration_ms >= 1000 {
                    format!("{:.2}s", duration_ms as f64 / 1000.0)
                } else {
                    format!("{}ms", duration_ms)
                };

                let status = format!(
                    "Opened {} in {}",
                    filename,
                    duration_str
                );

                result.import_duration_ms = duration_ms;
                self.import_result = Some(result);
                self.status_message = Some(status);
            }
            Err(e) => {
                self.import_result = None;
                self.import_filename = None;
                self.import_source_dir = None;
                self.status_message = Some(format!("Import failed: {}", e));
            }
        }
        cx.notify();
    }

    pub fn save(&mut self, cx: &mut Context<Self>) {
        // Commit any pending edit so it's included in the save
        self.commit_pending_edit();

        if let Some(path) = &self.current_file.clone() {
            self.save_to_path(path, cx);
        } else {
            self.save_as(cx);
        }
    }

    pub fn save_as(&mut self, cx: &mut Context<Self>) {
        // Commit any pending edit so it's included in the save
        self.commit_pending_edit();

        // For directory: prefer current file location, then import source, then current dir
        let directory = self.current_file.as_ref()
            .and_then(|p| p.parent())
            .map(|p| p.to_path_buf())
            .or_else(|| self.import_source_dir.clone())
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());

        // For filename: prefer current file name, then import filename (with .sheet extension)
        let suggested_name = self.current_file.as_ref()
            .and_then(|p| p.file_name())
            .and_then(|n| n.to_str())
            .map(|s| s.to_string())
            .or_else(|| {
                // Convert import filename to .sheet extension
                self.import_filename.as_ref().map(|name| {
                    let stem = std::path::Path::new(name)
                        .file_stem()
                        .and_then(|s| s.to_str())
                        .unwrap_or("untitled");
                    format!("{}.sheet", stem)
                })
            })
            .unwrap_or_else(|| "untitled.sheet".to_string());

        let future = cx.prompt_for_new_path(&directory, Some(&suggested_name));
        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(path))) = future.await {
                let _ = this.update(cx, |this, cx| {
                    this.save_to_path(&path, cx);
                });
            }
        })
        .detach();
    }

    fn save_to_path(&mut self, path: &PathBuf, cx: &mut Context<Self>) {
        let extension = path.extension().and_then(|s| s.to_str()).unwrap_or("sheet");

        // For .sheet files, use workbook-level save to preserve named ranges
        // For CSV, use sheet-level export (named ranges not supported in CSV)
        let result = match extension.to_lowercase().as_str() {
            "csv" => csv::export(self.sheet(), path),
            _ => native::save_workbook(&self.workbook, path),  // Default to .sheet format
        };

        match result {
            Ok(()) => {
                // Update document identity (handles current_file, is_modified, save_point)
                self.finalize_save(path);
                self.request_title_refresh(cx);

                // Save document settings to sidecar file
                // (best-effort - don't fail the whole save if sidecar fails)
                let _ = save_doc_settings(path, &self.doc_settings);

                // Update session with new file path
                self.update_session_cached(cx);

                let named_count = self.workbook.list_named_ranges().len();
                let status = if named_count > 0 {
                    format!("Saved: {} ({} named ranges)", path.display(), named_count)
                } else {
                    format!("Saved: {}", path.display())
                };
                self.status_message = Some(status);
            }
            Err(e) => {
                self.status_message = Some(format!("Error saving file: {}", e));
            }
        }
        cx.notify();
    }

    pub fn export_csv(&mut self, cx: &mut Context<Self>) {
        self.export_delimited(cx, "csv", csv::export);
    }

    pub fn export_tsv(&mut self, cx: &mut Context<Self>) {
        self.export_delimited(cx, "tsv", csv::export_tsv);
    }

    pub fn export_json(&mut self, cx: &mut Context<Self>) {
        self.export_delimited(cx, "json", json::export);
    }

    /// Export workbook to Excel (.xlsx) format
    /// This is a presentation snapshot - not a round-trip format.
    pub fn export_xlsx(&mut self, cx: &mut Context<Self>) {
        // Commit any pending edit so it's included in the export
        self.commit_pending_edit();

        let directory = self.current_file.as_ref()
            .and_then(|p| p.parent())
            .map(|p| p.to_path_buf())
            .or_else(|| self.import_source_dir.clone())
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());

        let base_name = self.current_file.as_ref()
            .and_then(|p| p.file_stem())
            .and_then(|n| n.to_str())
            .or_else(|| {
                self.import_filename.as_ref()
                    .and_then(|name| std::path::Path::new(name).file_stem())
                    .and_then(|s| s.to_str())
            })
            .unwrap_or("export");
        let suggested_name = format!("{}.xlsx", base_name);

        // Build layout information for each sheet
        let _layouts = self.build_export_layouts();

        let future = cx.prompt_for_new_path(&directory, Some(&suggested_name));
        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(path))) = future.await {
                let _ = this.update(cx, |this, cx| {
                    // Rebuild layouts in case data changed
                    let layouts = this.build_export_layouts();

                    match xlsx::export(&this.workbook, &path, Some(&layouts)) {
                        Ok(result) => {
                            let filename = path.file_name()
                                .and_then(|n| n.to_str())
                                .unwrap_or("file.xlsx")
                                .to_string();

                            let has_warnings = result.has_warnings();
                            let mut status = format!("Exported to {}", path.display());
                            if let Some(warning) = result.warning_summary() {
                                status.push_str(&format!(" ({})", warning));
                            }
                            this.status_message = Some(status);

                            // Store result and show dialog if there are warnings
                            if has_warnings {
                                this.export_result = Some(result);
                                this.export_filename = Some(filename);
                                this.show_export_report(cx);
                            } else {
                                this.export_result = None;
                                this.export_filename = None;
                            }
                        }
                        Err(e) => {
                            this.status_message = Some(format!("Export failed: {}", e));
                            this.export_result = None;
                            this.export_filename = None;
                        }
                    }
                    cx.notify();
                });
            }
        })
        .detach();
    }

    /// Export history as a deterministic Lua provenance script.
    /// Phase 9A: allows history to be replayed, audited, or shared.
    pub fn export_provenance(&mut self, cx: &mut Context<Self>) {
        use crate::provenance::{export_script, ExportOptions};

        let directory = self.current_file.as_ref()
            .and_then(|p| p.parent())
            .map(|p| p.to_path_buf())
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());

        let base_name = self.current_file.as_ref()
            .and_then(|p| p.file_stem())
            .and_then(|n| n.to_str())
            .unwrap_or("history");
        let suggested_name = format!("{}_provenance.lua", base_name);

        // Capture data needed for export
        let entries: Vec<_> = self.history.canonical_entries().to_vec();
        let fingerprint = self.history.fingerprint();
        let workbook_name = self.current_file.as_ref()
            .and_then(|p| p.file_name())
            .and_then(|n| n.to_str())
            .map(|s| s.to_string());

        let future = cx.prompt_for_new_path(&directory, Some(&suggested_name));
        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(path))) = future.await {
                let _ = this.update(cx, |this, cx| {
                    let options = ExportOptions {
                        include_header: true,
                        include_fingerprint: true,
                        sheet_filter: None, // Export all sheets
                    };

                    let script = export_script(
                        &entries,
                        fingerprint,
                        workbook_name.as_deref(),
                        &options,
                    );

                    match std::fs::write(&path, &script) {
                        Ok(()) => {
                            let action_count = entries.len();
                            this.status_message = Some(format!(
                                "Exported {} actions to {}",
                                action_count,
                                path.display()
                            ));
                        }
                        Err(e) => {
                            this.status_message = Some(format!("Export failed: {}", e));
                        }
                    }
                    cx.notify();
                });
            }
        })
        .detach();
    }

    /// Build ExportLayout for each sheet (column widths, row heights)
    /// Note: Frozen panes will be added when that feature is implemented (see roadmap)
    fn build_export_layouts(&self) -> Vec<xlsx::ExportLayout> {
        let mut layouts = Vec::new();

        for _ in 0..self.workbook.sheet_count() {
            let mut layout = xlsx::ExportLayout::default();

            // Copy column widths from the app state
            // Note: We store these per-sheet in the future, but for now use current sheet's widths
            for (col, width) in &self.col_widths {
                layout.col_widths.insert(*col, *width);
            }

            // Row heights (if we track them)
            for (row, height) in &self.row_heights {
                layout.row_heights.insert(*row, *height);
            }

            // Frozen panes: Not yet implemented in VisiGrid (see roadmap)
            // layout.frozen_rows = ...;
            // layout.frozen_cols = ...;

            layouts.push(layout);
        }

        layouts
    }

    fn export_delimited<F>(&mut self, cx: &mut Context<Self>, ext: &'static str, export_fn: F)
    where
        F: Fn(&visigrid_engine::sheet::Sheet, &std::path::Path) -> Result<(), String> + Send + 'static,
    {
        // Commit any pending edit so it's included in the export
        self.commit_pending_edit();

        let directory = self.current_file.as_ref()
            .and_then(|p| p.parent())
            .map(|p| p.to_path_buf())
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());

        let base_name = self.current_file.as_ref()
            .and_then(|p| p.file_stem())
            .and_then(|n| n.to_str())
            .unwrap_or("export");
        let suggested_name = format!("{}.{}", base_name, ext);

        let future = cx.prompt_for_new_path(&directory, Some(&suggested_name));
        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(path))) = future.await {
                let _ = this.update(cx, |this, cx| {
                    match export_fn(this.sheet(), &path) {
                        Ok(()) => {
                            this.status_message = Some(format!("Exported: {}", path.display()));
                        }
                        Err(e) => {
                            this.status_message = Some(format!("Error exporting: {}", e));
                        }
                    }
                    cx.notify();
                });
            }
        })
        .detach();
    }

    /// Add a file to the recent files list
    /// Moves to front if already present, limits to 10 entries
    pub fn add_recent_file(&mut self, path: &PathBuf) {
        const MAX_RECENT: usize = 10;

        // Remove if already present (we'll add to front)
        self.recent_files.retain(|p| p != path);

        // Add to front
        self.recent_files.insert(0, path.clone());

        // Limit size
        self.recent_files.truncate(MAX_RECENT);
    }
}
