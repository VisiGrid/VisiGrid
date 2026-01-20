use gpui::*;
use std::path::PathBuf;
use visigrid_engine::workbook::Workbook;
use visigrid_io::{csv, native};

use crate::app::Spreadsheet;

impl Spreadsheet {
    pub fn new_file(&mut self, cx: &mut Context<Self>) {
        self.workbook = Workbook::new();
        self.current_file = None;
        self.is_modified = false;
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

        // For .sheet files, use workbook-level load to preserve named ranges
        // For CSV, use sheet-level import
        let result: Result<Workbook, String> = match extension.to_lowercase().as_str() {
            "csv" => csv::import(path).map(|sheet| Workbook::from_sheets(vec![sheet], 0)),
            "sheet" => native::load_workbook(path),
            _ => Err(format!("Unknown file type: {}", extension)),
        };

        match result {
            Ok(workbook) => {
                self.workbook = workbook;
                self.current_file = Some(path.clone());
                self.is_modified = false;
                self.selected = (0, 0);
                self.selection_end = None;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.history.clear();
                self.bump_cells_rev();  // Invalidate cell search cache
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

    pub fn save(&mut self, cx: &mut Context<Self>) {
        if let Some(path) = &self.current_file.clone() {
            self.save_to_path(path, cx);
        } else {
            self.save_as(cx);
        }
    }

    pub fn save_as(&mut self, cx: &mut Context<Self>) {
        let directory = self.current_file.as_ref()
            .and_then(|p| p.parent())
            .map(|p| p.to_path_buf())
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());

        let suggested_name = self.current_file.as_ref()
            .and_then(|p| p.file_name())
            .and_then(|n| n.to_str())
            .unwrap_or("untitled.sheet");

        let future = cx.prompt_for_new_path(&directory, Some(suggested_name));
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
