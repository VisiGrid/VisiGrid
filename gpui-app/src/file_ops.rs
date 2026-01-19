use gpui::*;
use std::path::PathBuf;
use visigrid_engine::sheet::Sheet;
use visigrid_io::{csv, native};

use crate::app::{Spreadsheet, NUM_ROWS, NUM_COLS};

impl Spreadsheet {
    pub fn new_file(&mut self, cx: &mut Context<Self>) {
        self.sheet = Sheet::new(NUM_ROWS, NUM_COLS);
        self.current_file = None;
        self.is_modified = false;
        self.selected = (0, 0);
        self.selection_end = None;
        self.scroll_row = 0;
        self.scroll_col = 0;
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

        let result = match extension.to_lowercase().as_str() {
            "csv" => csv::import(path),
            "sheet" => native::load(path),
            _ => Err(format!("Unknown file type: {}", extension)),
        };

        match result {
            Ok(sheet) => {
                self.sheet = sheet;
                self.current_file = Some(path.clone());
                self.is_modified = false;
                self.selected = (0, 0);
                self.selection_end = None;
                self.scroll_row = 0;
                self.scroll_col = 0;
                self.status_message = Some(format!("Opened: {}", path.display()));
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

        let result = match extension.to_lowercase().as_str() {
            "csv" => csv::export(&self.sheet, path),
            _ => native::save(&self.sheet, path),  // Default to .sheet format
        };

        match result {
            Ok(()) => {
                self.current_file = Some(path.clone());
                self.is_modified = false;
                self.status_message = Some(format!("Saved: {}", path.display()));
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
                    match csv::export(&this.sheet, &path) {
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
}
