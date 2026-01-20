use gpui::*;
use std::path::PathBuf;
use std::time::{Duration, Instant};
use visigrid_engine::workbook::Workbook;
use visigrid_io::{csv, native, xlsx};

use crate::app::Spreadsheet;
use crate::settings::{load_doc_settings, save_doc_settings, DocumentSettings};

/// Delay before showing the import overlay (prevents flash for fast imports)
const OVERLAY_DELAY_MS: u64 = 150;

impl Spreadsheet {
    pub fn new_file(&mut self, cx: &mut Context<Self>) {
        self.workbook = Workbook::new();
        self.current_file = None;
        self.is_modified = false;
        self.doc_settings = DocumentSettings::default();  // Reset doc settings
        self.selected = (0, 0);
        self.selection_end = None;
        self.scroll_row = 0;
        self.scroll_col = 0;
        self.history.clear();
        self.bump_cells_rev();  // Invalidate cell search cache
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

        // Excel files use background import
        if matches!(ext_lower.as_str(), "xlsx" | "xls" | "xlsb" | "xlsm" | "ods") {
            self.start_excel_import(path, cx);
            return;
        }

        // Non-Excel files: synchronous load (fast, no need for background)
        let result: Result<Workbook, String> = match ext_lower.as_str() {
            "csv" => csv::import(path)
                .map(|sheet| Workbook::from_sheets(vec![sheet], 0)),
            "sheet" => native::load_workbook(path),
            _ => Err(format!("Unknown file type: {}", extension)),
        };

        match result {
            Ok(workbook) => {
                self.workbook = workbook;
                self.current_file = Some(path.clone());
                self.is_modified = false;
                self.import_result = None;
                self.import_filename = None;
                self.import_source_dir = None;
                // Load document settings from sidecar file
                self.doc_settings = load_doc_settings(path);
                self.selected = (0, 0);
                self.selection_end = None;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.history.clear();
                self.bump_cells_rev();
                self.add_recent_file(path);
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
                        this.current_file = None;  // Force Save As for .sheet
                        this.is_modified = true;
                        this.import_filename = Some(filename_for_completion.clone());
                        this.import_source_dir = source_dir;
                        this.doc_settings = DocumentSettings::default();
                        this.selected = (0, 0);
                        this.selection_end = None;
                        this.scroll_row = 0;
                        this.scroll_col = 0;
                        this.history.clear();
                        this.bump_cells_rev();
                        this.add_recent_file(&path_for_recent);

                        // Build status message with timing
                        let duration_str = if duration_ms >= 1000 {
                            format!("{:.2}s", duration_ms as f64 / 1000.0)
                        } else {
                            format!("{}ms", duration_ms)
                        };

                        let mut status = format!(
                            "Imported {} in {} — Save As to keep changes",
                            filename_for_completion,
                            duration_str
                        );

                        // Append warning hint if present
                        if result.has_warnings() {
                            status.push_str(" (see Import Report)");
                        }

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

    pub fn save(&mut self, cx: &mut Context<Self>) {
        if let Some(path) = &self.current_file.clone() {
            self.save_to_path(path, cx);
        } else {
            self.save_as(cx);
        }
    }

    pub fn save_as(&mut self, cx: &mut Context<Self>) {
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
                self.current_file = Some(path.clone());
                self.is_modified = false;

                // Save document settings to sidecar file
                // (best-effort - don't fail the whole save if sidecar fails)
                let _ = save_doc_settings(path, &self.doc_settings);

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
        let directory = self.current_file.as_ref()
            .and_then(|p| p.parent())
            .map(|p| p.to_path_buf())
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());

        let base_name = self.current_file.as_ref()
            .and_then(|p| p.file_stem())
            .and_then(|n| n.to_str())
            .unwrap_or("export");
        let suggested_name = format!("{}.csv", base_name);

        let future = cx.prompt_for_new_path(&directory, Some(&suggested_name));
        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(path))) = future.await {
                let _ = this.update(cx, |this, cx| {
                    match csv::export(this.sheet(), &path) {
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
