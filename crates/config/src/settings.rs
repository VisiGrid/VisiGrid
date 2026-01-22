// Application settings
// Loaded from ~/.config/visigrid/settings.json

use serde::{Deserialize, Serialize};
use std::fs;
use std::path::PathBuf;
use crate::theme::ThemeSource;

/// Keyboard modifier style preference (primarily for macOS users)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum ModifierStyle {
    /// Use platform-native modifier (Cmd on macOS, Ctrl on Windows/Linux)
    #[default]
    Platform,
    /// Always use Ctrl (for users who prefer Windows-style shortcuts on Mac)
    Ctrl,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct Settings {
    // Grid appearance
    #[serde(rename = "grid.defaultColumnWidth")]
    pub default_column_width: f32,

    #[serde(rename = "grid.rowHeight")]
    pub row_height: f32,

    #[serde(rename = "grid.rowHeaderWidth")]
    pub row_header_width: f32,

    #[serde(rename = "grid.showGridLines")]
    pub show_grid_lines: bool,

    // Editor
    #[serde(rename = "editor.fontSize")]
    pub font_size: f32,

    #[serde(rename = "editor.fontFamily")]
    pub font_family: String,

    // Formula
    #[serde(rename = "formula.autoRecalc")]
    pub auto_recalc: bool,

    // File
    #[serde(rename = "file.autoSaveInterval")]
    pub auto_save_interval: Option<u32>,  // seconds, None = disabled

    #[serde(rename = "file.recentFilesLimit")]
    pub recent_files_limit: usize,

    // UI
    #[serde(rename = "ui.showFormulaBar")]
    pub show_formula_bar: bool,

    #[serde(rename = "ui.showStatusBar")]
    pub show_status_bar: bool,

    #[serde(rename = "ui.showSheetTabs")]
    pub show_sheet_tabs: bool,

    // Navigation
    #[serde(rename = "editor.vimMode")]
    pub vim_mode: bool,

    // Theme
    #[serde(rename = "theme.source")]
    pub theme_source: ThemeSource,

    // Keyboard
    #[serde(rename = "keyboard.modifierStyle")]
    pub modifier_style: ModifierStyle,
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            // Grid
            default_column_width: 80.0,
            row_height: 24.0,
            row_header_width: 50.0,
            show_grid_lines: true,
            // Editor
            font_size: 13.0,
            font_family: String::new(),  // Empty = system default
            // Formula
            auto_recalc: true,
            // File
            auto_save_interval: None,
            recent_files_limit: 10,
            // UI
            show_formula_bar: true,
            show_status_bar: true,
            show_sheet_tabs: true,
            // Navigation
            vim_mode: false,
            // Theme
            theme_source: ThemeSource::Auto,
            // Keyboard
            modifier_style: ModifierStyle::default(),
        }
    }
}

impl Settings {
    /// Get the settings file path
    pub fn config_path() -> PathBuf {
        let config_dir = dirs::config_dir()
            .unwrap_or_else(|| PathBuf::from("."))
            .join("visigrid");
        config_dir.join("settings.json")
    }

    /// Load settings from disk, falling back to defaults
    pub fn load() -> Self {
        let path = Self::config_path();

        if !path.exists() {
            let settings = Self::default();
            settings.create_default_file();
            return settings;
        }

        match fs::read_to_string(&path) {
            Ok(contents) => {
                // Strip comments (lines starting with //)
                let cleaned: String = contents
                    .lines()
                    .filter(|line| !line.trim().starts_with("//"))
                    .collect::<Vec<_>>()
                    .join("\n");

                match serde_json::from_str(&cleaned) {
                    Ok(settings) => settings,
                    Err(e) => {
                        eprintln!("Error parsing settings.json: {}", e);
                        eprintln!("Using default settings");
                        Self::default()
                    }
                }
            }
            Err(e) => {
                eprintln!("Error reading settings.json: {}", e);
                Self::default()
            }
        }
    }

    /// Save current settings to disk
    pub fn save(&self) -> Result<(), String> {
        let path = Self::config_path();

        // Ensure directory exists
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent).map_err(|e| e.to_string())?;
        }

        let json = serde_json::to_string_pretty(self)
            .map_err(|e| e.to_string())?;

        fs::write(&path, json).map_err(|e| e.to_string())
    }

    /// Create default settings file with comments
    fn create_default_file(&self) {
        let path = Self::config_path();

        // Ensure directory exists
        if let Some(parent) = path.parent() {
            if let Err(e) = fs::create_dir_all(parent) {
                eprintln!("Error creating config directory: {}", e);
                return;
            }
        }

        let default_config = r#"{
    // Grid appearance
    "grid.defaultColumnWidth": 80,
    "grid.rowHeight": 24,
    "grid.rowHeaderWidth": 50,
    "grid.showGridLines": true,

    // Editor
    "editor.fontSize": 13,
    "editor.fontFamily": "",

    // Formula calculation
    "formula.autoRecalc": true,

    // File handling
    "file.autoSaveInterval": null,
    "file.recentFilesLimit": 10,

    // UI elements
    "ui.showFormulaBar": true,
    "ui.showStatusBar": true,
    "ui.showSheetTabs": true,

    // Keyboard (macOS only: "platform" = Cmd, "ctrl" = Ctrl)
    "keyboard.modifierStyle": "platform"
}
"#;

        if let Err(e) = fs::write(&path, default_config) {
            eprintln!("Error writing default settings.json: {}", e);
        }
    }

    /// Get the config file path for display/opening
    pub fn config_path_display() -> String {
        Self::config_path().to_string_lossy().to_string()
    }
}
