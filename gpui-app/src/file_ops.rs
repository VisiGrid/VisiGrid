use gpui::*;
use std::collections::HashMap;
use std::path::PathBuf;
use std::time::{Duration, Instant};
use visigrid_engine::workbook::Workbook;
use visigrid_io::{csv, json, native, xlsx};

use crate::app::{Spreadsheet, DocumentMeta};
use crate::settings::{load_doc_settings, save_doc_settings, DocumentSettings};

/// Delay before showing the import overlay (prevents flash for fast imports)
const OVERLAY_DELAY_MS: u64 = 150;

impl Spreadsheet {
    /// Replace the current workbook in-place with a blank workbook.
    ///
    /// WARNING: This is destructive - it discards all unsaved changes without
    /// confirmation. Use `NewWindow` action (Ctrl+N) instead for safe behavior
    /// that opens a new window.
    ///
    /// This method exists for:
    /// - Internal use (e.g., after explicit user confirmation)
    /// - "New in This Window" menu item (if exposed)
    pub fn new_in_place(&mut self, cx: &mut Context<Self>) {
        self.wb_mut(cx, |wb| *wb = Workbook::new());
        self.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
        self.debug_assert_sheet_cache_sync(cx);
        self.base_workbook = self.wb(cx).clone(); // Capture base state for replay
        self.rewind_preview = crate::app::RewindPreviewState::Off; // Reset preview state
        self.cycle_banner.reset_for_new_file();
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

        // CSV/TSV: background import (large files can freeze the UI)
        if matches!(ext_lower.as_str(), "csv" | "tsv") {
            self.start_csv_import(path, &ext_lower, cx);
            return;
        }

        // Remaining formats: synchronous load
        let load_start = Instant::now();
        let result: Result<Workbook, String> = match ext_lower.as_str() {
            "sheet" => native::load_workbook(path),
            _ => Err(format!("Unknown file type: {}", extension)),
        };

        match result {
            Ok(workbook) => {
                self.wb_mut(cx, |wb| *wb = workbook);
                // Rebuild dependency graph and recompute all formulas
                // This ensures formula cells have computed values, not just raw text
                self.wb_mut(cx, |wb| {
                    wb.rebuild_dep_graph();
                    wb.recompute_full_ordered();
                });
                self.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
                self.debug_assert_sheet_cache_sync(cx);
                self.base_workbook = self.wb(cx).clone(); // Capture base state for replay
                self.rewind_preview = crate::app::RewindPreviewState::Off;
                self.import_result = None;
                self.import_filename = None;
                self.import_source_dir = None;
                // Load document settings from sidecar file
                self.doc_settings = load_doc_settings(path);

                // Load column widths and row heights for .sheet files
                if ext_lower == "sheet" {
                    let layout = native::load_layout(path);
                    self.apply_sheet_layout(layout, cx);
                } else {
                    self.col_widths.clear();
                    self.row_heights.clear();
                }

                // Load semantic verification info for .sheet files
                if extension == "sheet" {
                    self.semantic_verification = visigrid_io::native::load_semantic_verification(path)
                        .unwrap_or_default();
                    // If file has expected fingerprint, track history state for drift diff
                    if self.semantic_verification.fingerprint.is_some() {
                        self.approved_fingerprint = Some(self.history.fingerprint());
                        self.approval_history_len = 0;
                    }
                }

                // Wire calculation settings to engine
                let auto = self.doc_settings.calculation.mode
                    .resolve(crate::settings::CalculationMode::Automatic)
                    != crate::settings::CalculationMode::Manual;
                let iterative = self.doc_settings.calculation.enable_iterative_calc.resolve(false);
                let max_iters = self.doc_settings.calculation.max_iterations.resolve(100);
                let tolerance = self.doc_settings.calculation.iteration_tolerance.resolve(1e-9);
                self.wb_mut(cx, |wb| {
                    wb.set_auto_recalc(auto);
                    wb.set_iterative_enabled(iterative);
                    wb.set_iterative_max_iters(max_iters);
                    wb.set_iterative_tolerance(tolerance);
                });

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

                // Load VisiHub link and cell metadata (for .sheet files)
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

                    // Load cell metadata for role-based auto-styling
                    match visigrid_io::native::load_cell_metadata(path) {
                        Ok(meta) => {
                            // Convert BTreeMap to HashMap for faster lookup
                            self.cell_metadata = meta.into_iter()
                                .map(|(k, v)| (k, v.into_iter().collect()))
                                .collect();

                            // Auto-fit columns for agent-built sheets with metadata
                            if !self.cell_metadata.is_empty() {
                                self.auto_fit_all_data_columns(cx);
                            }
                        }
                        Err(_) => {
                            self.cell_metadata.clear();
                        }
                    }
                    // Load attached scripts and run records
                    self.attached_scripts = visigrid_io::native::load_scripts(path)
                        .unwrap_or_default();
                    self.loaded_run_records = visigrid_io::native::load_run_records(path)
                        .unwrap_or_default();
                    self.pending_run_records.clear();
                } else {
                    self.hub_link = None;
                    self.hub_status = crate::hub::HubStatus::Unlinked;
                    self.cell_metadata.clear();
                    self.attached_scripts.clear();
                    self.loaded_run_records.clear();
                    self.pending_run_records.clear();
                }
                // Update session with new file path
                self.update_session_cached(cx);

                let named_count = self.wb(cx).list_named_ranges().len();

                let duration_ms = load_start.elapsed().as_millis();
                let duration_str = if duration_ms >= 1000 {
                    format!("{:.2}s", duration_ms as f64 / 1000.0)
                } else {
                    format!("{}ms", duration_ms)
                };
                let filename = path.file_name()
                    .and_then(|n| n.to_str())
                    .unwrap_or("file");
                let status = if named_count > 0 {
                    format!("Opened {} in {} ({} named ranges)", filename, duration_str, named_count)
                } else {
                    format!("Opened {} in {}", filename, duration_str)
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
        self.cycle_banner.reset_for_new_file();
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
                        // Atomic swap: replace entire workbook (wrap in Entity)
                        this.workbook = cx.new(|_| workbook);
                        this.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
                        this.debug_assert_sheet_cache_sync(cx);
                        this.base_workbook = this.wb(cx).clone(); // Capture base state for replay
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

                        let total_errors = result.recalc_errors + result.recalc_circular;
                        let status = if total_errors > 0 {
                            format!(
                                "Opened {} in {} \u{2014} {} errors (Import Report)",
                                filename_for_completion, duration_str, total_errors
                            )
                        } else {
                            format!(
                                "Opened {} in {} \u{2014} 0 errors",
                                filename_for_completion, duration_str
                            )
                        };

                        // Store the duration we measured (more accurate than import's internal timing
                        // since it includes workbook construction)
                        result.import_duration_ms = duration_ms;
                        let has_recalc_errors = result.recalc_errors > 0 || result.recalc_circular > 0;

                        // Apply imported column widths and row heights
                        this.apply_imported_layouts(&result, cx);

                        this.import_result = Some(result);
                        this.status_message = Some(status);

                        // Show cycle banner if applicable
                        if this.should_show_cycle_banner(cx) {
                            this.cycle_banner.show_force();
                        }

                        // Auto-show import report when recalc errors are detected
                        if has_recalc_errors {
                            this.show_import_report(cx);
                        }
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

    /// Start background CSV/TSV import with delayed overlay
    fn start_csv_import(&mut self, path: &PathBuf, ext: &str, cx: &mut Context<Self>) {
        let filename = path.file_name()
            .and_then(|n| n.to_str())
            .map(|s| s.to_string())
            .unwrap_or_else(|| "file".to_string());

        // Set up import state
        self.import_in_progress = true;
        self.import_overlay_visible = false;
        self.import_started_at = Some(Instant::now());
        self.import_filename = Some(filename.clone());
        self.status_message = Some(format!("Importing {}...", filename));

        let path_for_import = path.clone();
        let path_for_recent = path.clone();
        let filename_for_completion = filename.clone();
        let is_tsv = ext == "tsv";

        cx.notify();

        // Task A: Delayed overlay trigger
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
            let import_result = cx.background_executor()
                .spawn(async move {
                    let sheet = if is_tsv {
                        csv::import_tsv(&path_for_import)
                    } else {
                        csv::import(&path_for_import)
                    }?;
                    let mut workbook = Workbook::from_sheets(vec![sheet], 0);
                    workbook.rebuild_dep_graph();
                    workbook.recompute_full_ordered();
                    Ok::<Workbook, String>(workbook)
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
                    Ok(workbook) => {
                        this.workbook = cx.new(|_| workbook);
                        this.update_cached_sheet_id(cx);
                        this.debug_assert_sheet_cache_sync(cx);
                        this.base_workbook = this.wb(cx).clone();
                        this.rewind_preview = crate::app::RewindPreviewState::Off;
                        this.import_result = None;
                        this.import_filename = Some(filename_for_completion.clone());
                        this.import_source_dir = None;
                        this.doc_settings = DocumentSettings::default();
                        this.col_widths.clear();
                        this.row_heights.clear();
                        this.view_state.selected = (0, 0);
                        this.view_state.selection_end = None;
                        this.view_state.scroll_row = 0;
                        this.view_state.scroll_col = 0;
                        this.history.clear();
                        this.bump_cells_rev();
                        this.add_recent_file(&path_for_recent);

                        this.finalize_load(&path_for_recent);
                        this.request_title_refresh(cx);

                        // Clear hub link (CSV has no hub metadata)
                        this.hub_link = None;
                        this.hub_status = crate::hub::HubStatus::Unlinked;
                        this.cell_metadata.clear();

                        // Update session with new file path
                        this.update_session_cached(cx);

                        let duration_str = if duration_ms >= 1000 {
                            format!("{:.2}s", duration_ms as f64 / 1000.0)
                        } else {
                            format!("{}ms", duration_ms)
                        };
                        this.status_message = Some(
                            format!("Opened {} in {}", filename_for_completion, duration_str)
                        );
                    }
                    Err(e) => {
                        this.import_result = None;
                        this.import_filename = None;
                        this.import_source_dir = None;
                        this.status_message = Some(format!("Error opening file: {}", e));
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
        self.cycle_banner.reset_for_new_file();
        let start_time = std::time::Instant::now();

        match xlsx::import(path) {
            Ok((workbook, mut result)) => {
                let duration_ms = start_time.elapsed().as_millis();

                self.wb_mut(cx, |wb| *wb = workbook);
                self.update_cached_sheet_id(cx);  // Keep per-sheet sizing cache in sync
                self.debug_assert_sheet_cache_sync(cx);
                self.base_workbook = self.wb(cx).clone(); // Capture base state for replay
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

                let total_errors = result.recalc_errors + result.recalc_circular;
                let status = if total_errors > 0 {
                    format!(
                        "Opened {} in {} \u{2014} {} errors (Import Report)",
                        filename, duration_str, total_errors
                    )
                } else {
                    format!(
                        "Opened {} in {} \u{2014} 0 errors",
                        filename, duration_str
                    )
                };

                result.import_duration_ms = duration_ms;
                let has_recalc_errors = result.recalc_errors > 0 || result.recalc_circular > 0;

                // Apply imported column widths and row heights
                self.apply_imported_layouts(&result, cx);

                self.import_result = Some(result);
                self.status_message = Some(status);

                // Show cycle banner if applicable
                if self.should_show_cycle_banner(cx) {
                    self.cycle_banner.show_force();
                }

                // Auto-show import report when recalc errors are detected
                if has_recalc_errors {
                    self.show_import_report(cx);
                }
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

    /// Re-import the current file with freeze_cycles enabled.
    /// Called from the import report dialog's "Freeze Cycle Values" button.
    pub fn reimport_with_freeze(&mut self, cx: &mut Context<Self>) {
        let Some(path) = self.current_file.clone() else { return; };
        let current_sheet = self.wb(cx).active_sheet_index();
        self.import_result = None; // hide dialog
        self.mode = crate::mode::Mode::Navigation;
        self.status_message = Some("Re-importing with Freeze Cycle Values...".to_string());
        cx.notify();

        let options = xlsx::ImportOptions { freeze_cycles: true, ..Default::default() };

        if visigrid_license::is_feature_enabled("fast_large_files") {
            self.start_excel_import_with_options(&path, options, current_sheet, cx);
        } else {
            self.load_excel_sync_with_options(&path, options, current_sheet, cx);
        }
    }

    /// Enable iterative calculation and recompute all formulas in-place.
    /// No reimport — just toggles the workbook setting and recalcs.
    pub fn enable_iteration_and_recalc(&mut self, cx: &mut Context<Self>) {
        let max_iters = self.doc_settings.calculation.max_iterations.resolve(100);
        let tolerance = self.doc_settings.calculation.iteration_tolerance.resolve(1e-9);

        let report = self.wb_mut(cx, |wb| {
            wb.set_iterative_enabled(true);
            wb.set_iterative_max_iters(max_iters);
            wb.set_iterative_tolerance(tolerance);
            wb.rebuild_dep_graph();
            wb.recompute_full_ordered()
        });
        self.last_recalc_report = Some(report);

        // Persist to doc settings
        self.doc_settings.calculation.enable_iterative_calc = crate::settings::Setting::Value(true);
        if let Some(path) = &self.current_file {
            let _ = save_doc_settings(path, &self.doc_settings);
        }

        // Show banner with iteration result (force — user just acted)
        self.cycle_banner.show_force();

        // Close import report if open
        if self.mode == crate::mode::Mode::ImportReport {
            self.mode = crate::mode::Mode::Navigation;
        }

        self.status_message = Some("Iterative calculation enabled".to_string());
        self.is_modified = true;
        cx.notify();
    }

    /// Start background Excel import with options and optional sheet restore
    fn start_excel_import_with_options(
        &mut self,
        path: &PathBuf,
        options: xlsx::ImportOptions,
        restore_sheet: usize,
        cx: &mut Context<Self>,
    ) {
        let filename = path.file_name()
            .and_then(|n| n.to_str())
            .map(|s| s.to_string())
            .unwrap_or_else(|| "file".to_string());
        let source_dir = path.parent().map(|p| p.to_path_buf());

        self.import_in_progress = true;
        self.import_overlay_visible = false;
        self.import_started_at = Some(Instant::now());
        self.cycle_banner.reset_for_new_file();
        self.status_message = Some(format!("Importing {}...", filename));

        let path_for_import = path.clone();
        let path_for_recent = path.clone();
        let filename_for_completion = filename.clone();

        cx.notify();

        // Delayed overlay trigger
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

        // Actual import
        cx.spawn(async move |this, cx| {
            let import_result = cx.background_executor()
                .spawn(async move {
                    xlsx::import_with_options(&path_for_import, &options)
                })
                .await;

            let _ = this.update(cx, |this, cx| {
                this.import_in_progress = false;
                this.import_overlay_visible = false;

                let duration_ms = this.import_started_at
                    .map(|t| t.elapsed().as_millis())
                    .unwrap_or(0);

                match import_result {
                    Ok((workbook, mut result)) => {
                        this.workbook = cx.new(|_| workbook);
                        this.update_cached_sheet_id(cx);
                        this.debug_assert_sheet_cache_sync(cx);
                        this.base_workbook = this.wb(cx).clone();
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

                        this.finalize_load(&path_for_recent);
                        this.request_title_refresh(cx);

                        // Restore sheet selection
                        if restore_sheet < this.wb(cx).sheet_count() {
                            this.wb_mut(cx, |wb| { wb.set_active_sheet(restore_sheet); });
                            this.update_cached_sheet_id(cx);
                        }

                        let duration_str = if duration_ms >= 1000 {
                            format!("{:.2}s", duration_ms as f64 / 1000.0)
                        } else {
                            format!("{}ms", duration_ms)
                        };

                        let total_errors = result.recalc_errors + result.recalc_circular;
                        let freeze_note = if result.freeze_applied {
                            format!(" ({} cycles frozen)", result.cycles_frozen)
                        } else {
                            String::new()
                        };
                        let status = if total_errors > 0 {
                            format!("Opened {} in {} \u{2014} {} errors{}",
                                filename_for_completion, duration_str, total_errors, freeze_note)
                        } else {
                            format!("Opened {} in {} \u{2014} 0 errors{}",
                                filename_for_completion, duration_str, freeze_note)
                        };

                        result.import_duration_ms = duration_ms;
                        let show_report = result.recalc_errors > 0
                            || result.recalc_circular > 0
                            || result.freeze_applied;

                        this.apply_imported_layouts(&result, cx);

                        this.import_result = Some(result);
                        this.status_message = Some(status);

                        // Show cycle banner if applicable
                        if this.should_show_cycle_banner(cx) {
                            this.cycle_banner.show_force();
                        }

                        if show_report {
                            this.show_import_report(cx);
                        }
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

    /// Synchronous Excel import with options and optional sheet restore
    fn load_excel_sync_with_options(
        &mut self,
        path: &PathBuf,
        options: xlsx::ImportOptions,
        restore_sheet: usize,
        cx: &mut Context<Self>,
    ) {
        let filename = path.file_name()
            .and_then(|n| n.to_str())
            .map(|s| s.to_string())
            .unwrap_or_else(|| "file".to_string());
        let source_dir = path.parent().map(|p| p.to_path_buf());
        self.cycle_banner.reset_for_new_file();
        let start_time = std::time::Instant::now();

        match xlsx::import_with_options(path, &options) {
            Ok((workbook, mut result)) => {
                let duration_ms = start_time.elapsed().as_millis();

                self.wb_mut(cx, |wb| *wb = workbook);
                self.update_cached_sheet_id(cx);
                self.debug_assert_sheet_cache_sync(cx);
                self.base_workbook = self.wb(cx).clone();
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

                self.finalize_load(path);
                self.request_title_refresh(cx);

                // Restore sheet selection
                if restore_sheet < self.wb(cx).sheet_count() {
                    self.wb_mut(cx, |wb| { wb.set_active_sheet(restore_sheet); });
                    self.update_cached_sheet_id(cx);
                }

                let duration_str = if duration_ms >= 1000 {
                    format!("{:.2}s", duration_ms as f64 / 1000.0)
                } else {
                    format!("{}ms", duration_ms)
                };

                let total_errors = result.recalc_errors + result.recalc_circular;
                let freeze_note = if result.freeze_applied {
                    format!(" ({} cycles frozen)", result.cycles_frozen)
                } else {
                    String::new()
                };
                let status = if total_errors > 0 {
                    format!("Opened {} in {} \u{2014} {} errors{}",
                        filename, duration_str, total_errors, freeze_note)
                } else {
                    format!("Opened {} in {} \u{2014} 0 errors{}",
                        filename, duration_str, freeze_note)
                };

                result.import_duration_ms = duration_ms;
                let show_report = result.recalc_errors > 0
                    || result.recalc_circular > 0
                    || result.freeze_applied;

                self.apply_imported_layouts(&result, cx);

                self.import_result = Some(result);
                self.status_message = Some(status);

                // Show cycle banner if applicable
                if self.should_show_cycle_banner(cx) {
                    self.cycle_banner.show_force();
                }

                if show_report {
                    self.show_import_report(cx);
                }
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

    /// Apply imported column widths and row heights from XLSX formatting.
    /// Converts raw Excel units to pixel values used by the app.
    fn apply_imported_layouts(&mut self, result: &xlsx::ImportResult, cx: &mut Context<Self>) {
        for (sheet_idx, layout) in result.imported_layouts.iter().enumerate() {
            if layout.col_widths.is_empty() && layout.row_heights.is_empty() {
                continue;
            }
            // Get the SheetId for this sheet index
            let sheet_id = match self.wb(cx).sheet(sheet_idx) {
                Some(s) => s.id,
                None => continue,
            };

            // Apply column widths: Excel character width → pixels
            // Excel formula: pixels = width * max_digit_width + padding
            // For Calibri 11pt at 96 DPI: max_digit_width ≈ 7, padding ≈ 5
            if !layout.col_widths.is_empty() {
                let widths = self.col_widths.entry(sheet_id).or_insert_with(HashMap::new);
                for (&col, &excel_width) in &layout.col_widths {
                    let px_width = (excel_width * 7.0 + 5.0) as f32;
                    let clamped = px_width.max(20.0).min(500.0);
                    if (clamped - crate::app::CELL_WIDTH).abs() >= 1.0 {
                        widths.insert(col, clamped);
                    }
                }
            }

            // Apply row heights: Excel points * (96/72) → pixels at 96 DPI
            if !layout.row_heights.is_empty() {
                let heights = self.row_heights.entry(sheet_id).or_insert_with(HashMap::new);
                for (&row, &excel_height) in &layout.row_heights {
                    let px_height = (excel_height / 0.75) as f32;
                    let clamped = px_height.max(12.0).min(200.0);
                    if (clamped - crate::app::CELL_HEIGHT).abs() >= 1.0 {
                        heights.insert(row, clamped);
                    }
                }
            }
        }
    }

    pub fn save(&mut self, cx: &mut Context<Self>) {
        // Commit any pending edit so it's included in the save
        self.commit_pending_edit(cx);

        if let Some(path) = &self.current_file.clone() {
            self.save_to_path(path, cx);
        } else {
            self.save_as(cx);
        }
    }

    /// Save the workbook and return whether the save can proceed synchronously.
    /// Returns true if file was saved (has existing path), false if Save As dialog is needed.
    /// Used by close-with-save flow to know if window can be closed immediately.
    pub fn save_and_close(&mut self, cx: &mut Context<Self>) -> bool {
        // Commit any pending edit so it's included in the save
        self.commit_pending_edit(cx);

        if let Some(path) = &self.current_file.clone() {
            // File has a path - save synchronously
            self.save_to_path(path, cx);
            // Check if save succeeded by looking at is_modified flag
            // (save_to_path sets is_modified = false on success via finalize_save)
            !self.is_modified
        } else {
            // File is untitled - need Save As dialog
            // Set flag so window closes after save completes
            self.close_after_save = true;
            self.save_as(cx);
            false // Can't close immediately, async save in progress
        }
    }

    pub fn save_as(&mut self, cx: &mut Context<Self>) {
        // Commit any pending edit so it's included in the save
        self.commit_pending_edit(cx);

        // For directory: prefer current file location, then import source, then current dir
        let directory = self.current_file.as_ref()
            .and_then(|p| p.parent())
            .map(|p| p.to_path_buf())
            .or_else(|| self.import_source_dir.clone())
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());

        // For filename: use current file name but default to .sheet for non-native formats
        let suggested_name = self.current_file.as_ref()
            .and_then(|p| {
                let name = p.file_name()?.to_str()?;
                let ext = p.extension().and_then(|e| e.to_str()).unwrap_or("");
                // If already a VisiGrid native format, keep the name as-is
                if matches!(ext.to_lowercase().as_str(), "sheet" | "vgrid") {
                    Some(name.to_string())
                } else {
                    // Non-native format (xlsx, csv, etc.) — suggest .sheet
                    let stem = p.file_stem().and_then(|s| s.to_str()).unwrap_or("untitled");
                    Some(format!("{}.sheet", stem))
                }
            })
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
                let close_result = this.update(cx, |this, cx| {
                    this.save_to_path(&path, cx);
                    // Check if we should close after save (from save_and_close flow)
                    let should_close = this.close_after_save && !this.is_modified;
                    this.close_after_save = false;  // Reset flag
                    if should_close {
                        this.prepare_close(cx);
                        Some(this.window_handle)
                    } else {
                        None
                    }
                }).ok().flatten();

                // Close window if save_and_close was in progress
                if let Some(window_handle) = close_result {
                    let _ = window_handle.update(cx, |_, window, _| {
                        window.remove_window();
                    });
                }
            } else {
                // User cancelled Save As - reset close_after_save flag
                let _ = this.update(cx, |this, _cx| {
                    this.close_after_save = false;
                });
            }
        })
        .detach();
    }

    fn save_to_path(&mut self, path: &PathBuf, cx: &mut Context<Self>) {
        let extension = path.extension().and_then(|s| s.to_str()).unwrap_or("sheet");

        // For .sheet files, use full save (workbook + scripts + run records)
        // For CSV, use sheet-level export (named ranges not supported in CSV)
        let result = match extension.to_lowercase().as_str() {
            "csv" => csv::export(self.sheet(cx), path),
            _ => {
                // Convert HashMap<String, HashMap<String, String>> → BTreeMap<String, BTreeMap<String, String>>
                let metadata: visigrid_io::native::CellMetadata = self.cell_metadata.iter()
                    .map(|(k, v)| (k.clone(), v.iter().map(|(k2, v2)| (k2.clone(), v2.clone())).collect()))
                    .collect();
                let all_run_records: Vec<_> = self.loaded_run_records.iter()
                    .chain(self.pending_run_records.iter())
                    .cloned()
                    .collect();
                native::save_workbook_full(
                    self.wb(cx),
                    &metadata,
                    &self.attached_scripts,
                    &all_run_records,
                    path,
                )
            }
        };

        match result {
            Ok(()) => {
                // Save column widths and row heights for .sheet files
                if extension.to_lowercase() != "csv" {
                    let layout = self.build_sheet_layout(cx);
                    if let Err(e) = native::save_layout(path, &layout) {
                        eprintln!("Warning: failed to save layout: {}", e);
                    }
                }

                // Update document identity (handles current_file, is_modified, save_point)
                self.finalize_save(path);
                self.request_title_refresh(cx);

                // Merge pending run records into loaded (now persisted)
                if !self.pending_run_records.is_empty() {
                    self.loaded_run_records.append(&mut self.pending_run_records);
                }

                // Re-save hub_link if present (save_workbook recreates file fresh)
                if let Some(ref link) = self.hub_link {
                    if let Err(e) = crate::hub::save_hub_link(path, link) {
                        // Log but don't fail the save
                        eprintln!("Warning: failed to preserve hub link: {}", e);
                    }
                }

                // Save document settings to sidecar file
                // (best-effort - don't fail the whole save if sidecar fails)
                let _ = save_doc_settings(path, &self.doc_settings);

                // Update session with new file path
                self.update_session_cached(cx);

                let named_count = self.wb(cx).list_named_ranges().len();
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
        self.commit_pending_edit(cx);

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
        let _layouts = self.build_export_layouts(cx);

        let future = cx.prompt_for_new_path(&directory, Some(&suggested_name));
        cx.spawn(async move |this, cx| {
            if let Ok(Ok(Some(path))) = future.await {
                let _ = this.update(cx, |this, cx| {
                    // Rebuild layouts in case data changed
                    let layouts = this.build_export_layouts(cx);

                    match xlsx::export(this.wb(cx), &path, Some(&layouts)) {
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
    fn build_export_layouts(&self, cx: &App) -> Vec<xlsx::ExportLayout> {
        let mut layouts = Vec::new();
        let wb = self.wb(cx);

        for sheet_idx in 0..wb.sheet_count() {
            let mut layout = xlsx::ExportLayout::default();

            // Get the sheet's ID for per-sheet storage lookup
            if let Some(sheet) = wb.sheets().get(sheet_idx) {
                let sheet_id = sheet.id;

                // Copy column widths for this specific sheet
                if let Some(sheet_widths) = self.col_widths.get(&sheet_id) {
                    for (col, width) in sheet_widths {
                        layout.col_widths.insert(*col, *width);
                    }
                }

                // Copy row heights for this specific sheet
                if let Some(sheet_heights) = self.row_heights.get(&sheet_id) {
                    for (row, height) in sheet_heights {
                        layout.row_heights.insert(*row, *height);
                    }
                }
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
        self.commit_pending_edit(cx);

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
                    match export_fn(this.sheet(cx), &path) {
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

    /// Build a SheetLayout from the app's col_widths/row_heights (SheetId-keyed)
    /// mapped to sheet_idx for database storage.
    fn build_sheet_layout(&self, cx: &App) -> native::SheetLayout {
        let mut layout = native::SheetLayout {
            col_widths: HashMap::new(),
            row_heights: HashMap::new(),
            hidden_rows: HashMap::new(),
            hidden_cols: HashMap::new(),
        };

        let sheets = self.wb(cx).sheets();
        for (idx, sheet) in sheets.iter().enumerate() {
            if let Some(widths) = self.col_widths.get(&sheet.id) {
                if !widths.is_empty() {
                    layout.col_widths.insert(idx, widths.clone());
                }
            }
            if let Some(heights) = self.row_heights.get(&sheet.id) {
                if !heights.is_empty() {
                    layout.row_heights.insert(idx, heights.clone());
                }
            }
            if let Some(rows) = self.hidden_rows.get(&sheet.id) {
                if !rows.is_empty() {
                    layout.hidden_rows.insert(idx, rows.iter().copied().collect());
                }
            }
            if let Some(cols) = self.hidden_cols.get(&sheet.id) {
                if !cols.is_empty() {
                    layout.hidden_cols.insert(idx, cols.iter().copied().collect());
                }
            }
        }

        layout
    }

    /// Apply a loaded SheetLayout (sheet_idx-keyed) to the app's col_widths/row_heights
    /// by mapping sheet_idx back to SheetId.
    fn apply_sheet_layout(&mut self, layout: native::SheetLayout, cx: &App) {
        // Collect sheet IDs first to avoid borrow conflict
        let sheet_ids: Vec<(usize, visigrid_engine::sheet::SheetId)> = self.wb(cx).sheets()
            .iter().enumerate().map(|(i, s)| (i, s.id)).collect();

        self.col_widths.clear();
        self.row_heights.clear();
        self.hidden_rows.clear();
        self.hidden_cols.clear();

        for (idx, sheet_id) in sheet_ids {
            if let Some(widths) = layout.col_widths.get(&idx) {
                if !widths.is_empty() {
                    self.col_widths.insert(sheet_id, widths.clone());
                }
            }
            if let Some(heights) = layout.row_heights.get(&idx) {
                if !heights.is_empty() {
                    self.row_heights.insert(sheet_id, heights.clone());
                }
            }
            if let Some(rows) = layout.hidden_rows.get(&idx) {
                if !rows.is_empty() {
                    self.hidden_rows.insert(sheet_id, rows.iter().copied().collect());
                }
            }
            if let Some(cols) = layout.hidden_cols.get(&idx) {
                if !cols.is_empty() {
                    self.hidden_cols.insert(sheet_id, cols.iter().copied().collect());
                }
            }
        }
    }
}
