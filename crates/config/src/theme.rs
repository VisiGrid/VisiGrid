// Theme configuration
// Supports: built-in themes, Omarchy system themes, and custom JSON themes

use crate::Color;
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::PathBuf;

/// Theme source - where to load theme from
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value")]
pub enum ThemeSource {
    /// Auto-detect: Use Omarchy if available, otherwise use dark
    Auto,
    /// Built-in dark theme
    Dark,
    /// Built-in light theme
    Light,
    /// Omarchy system theme (Linux only)
    System,
    /// Custom theme from file path
    Custom(String),
}

impl Default for ThemeSource {
    fn default() -> Self {
        ThemeSource::Auto
    }
}

/// JSON-serializable theme colors
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ThemeConfig {
    pub name: String,
    #[serde(default)]
    pub is_dark: bool,
    pub colors: ThemeColorsConfig,
}

/// JSON color definitions (hex strings)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ThemeColorsConfig {
    // Base colors
    pub background: String,
    #[serde(default = "default_background_secondary")]
    pub background_secondary: String,
    pub foreground: String,
    #[serde(default = "default_foreground_muted")]
    pub foreground_muted: String,

    // Grid
    #[serde(default = "default_gridline")]
    pub gridline: String,
    #[serde(default = "default_header_background")]
    pub header_background: String,

    // Selection
    pub accent: String,
    #[serde(default = "default_selection_alpha")]
    pub selection_alpha: f32,

    // UI elements
    #[serde(default = "default_border")]
    pub border: String,
}

fn default_background_secondary() -> String { "#0f172a".into() }
fn default_foreground_muted() -> String { "#64748b".into() }
fn default_gridline() -> String { "#334155".into() }
fn default_header_background() -> String { "#1e293b".into() }
fn default_border() -> String { "#334155".into() }
fn default_selection_alpha() -> f32 { 0.2 }

/// Runtime theme colors (iced::Color)
#[derive(Debug, Clone, Copy)]
pub struct ThemeColors {
    pub bg_dark: Color,
    pub bg_header: Color,
    pub bg_cell: Color,
    pub bg_input: Color,
    pub text: Color,
    pub text_dim: Color,
    pub gridline: Color,
    pub accent: Color,
    pub selected: Color,
    pub selected_border: Color,
    pub border: Color,
}

impl ThemeColors {
    /// Built-in dark theme
    pub fn dark() -> Self {
        ThemeColors {
            bg_dark: Color::from_rgb(0.008, 0.024, 0.090),        // #020617
            bg_header: Color::from_rgb(0.118, 0.161, 0.231),      // #1e293b
            bg_cell: Color::from_rgb(0.059, 0.090, 0.165),        // #0f172a
            bg_input: Color::from_rgb(0.059, 0.090, 0.165),       // #0f172a
            text: Color::from_rgb(0.945, 0.961, 0.976),           // #f1f5f9
            text_dim: Color::from_rgb(0.392, 0.439, 0.529),       // #64748b
            gridline: Color::from_rgb(0.200, 0.255, 0.333),       // #334155
            accent: Color::from_rgb(0.231, 0.510, 0.965),         // #3b82f6
            selected: Color::from_rgba(0.231, 0.510, 0.965, 0.2), // #3b82f6 @ 20%
            selected_border: Color::from_rgb(0.231, 0.510, 0.965),// #3b82f6
            border: Color::from_rgb(0.200, 0.255, 0.333),         // #334155
        }
    }

    /// Built-in light theme
    pub fn light() -> Self {
        ThemeColors {
            bg_dark: Color::from_rgb(0.973, 0.980, 0.988),        // #f8fafc
            bg_header: Color::from_rgb(0.886, 0.910, 0.941),      // #e2e8f0
            bg_cell: Color::from_rgb(0.945, 0.961, 0.976),        // #f1f5f9
            bg_input: Color::from_rgb(0.945, 0.961, 0.976),       // #f1f5f9
            text: Color::from_rgb(0.059, 0.090, 0.165),           // #0f172a
            text_dim: Color::from_rgb(0.278, 0.333, 0.412),       // #475569
            gridline: Color::from_rgb(0.796, 0.835, 0.882),       // #cbd5e1
            accent: Color::from_rgb(0.231, 0.510, 0.965),         // #3b82f6
            selected: Color::from_rgba(0.231, 0.510, 0.965, 0.15),// #3b82f6 @ 15%
            selected_border: Color::from_rgb(0.231, 0.510, 0.965),// #3b82f6
            border: Color::from_rgb(0.796, 0.835, 0.882),         // #cbd5e1
        }
    }

    /// Parse hex color to iced::Color
    pub fn hex_to_color(hex: &str) -> Option<Color> {
        let hex = hex.trim_start_matches('#');
        if hex.len() != 6 {
            return None;
        }
        let r = u8::from_str_radix(&hex[0..2], 16).ok()? as f32 / 255.0;
        let g = u8::from_str_radix(&hex[2..4], 16).ok()? as f32 / 255.0;
        let b = u8::from_str_radix(&hex[4..6], 16).ok()? as f32 / 255.0;
        Some(Color::from_rgb(r, g, b))
    }

    /// Lighten a color by mixing with white
    fn lighten(color: Color, amount: f32) -> Color {
        Color::from_rgb(
            color.r + (1.0 - color.r) * amount,
            color.g + (1.0 - color.g) * amount,
            color.b + (1.0 - color.b) * amount,
        )
    }

    /// Darken a color by mixing with black
    fn darken(color: Color, amount: f32) -> Color {
        Color::from_rgb(
            color.r * (1.0 - amount),
            color.g * (1.0 - amount),
            color.b * (1.0 - amount),
        )
    }

    /// Create ThemeColors from a ThemeConfig (JSON theme)
    pub fn from_config(config: &ThemeColorsConfig, is_dark: bool) -> Self {
        let bg = Self::hex_to_color(&config.background).unwrap_or(
            if is_dark { Color::from_rgb(0.008, 0.024, 0.090) }
            else { Color::from_rgb(0.973, 0.980, 0.988) }
        );
        let fg = Self::hex_to_color(&config.foreground).unwrap_or(
            if is_dark { Color::from_rgb(0.945, 0.961, 0.976) }
            else { Color::from_rgb(0.059, 0.090, 0.165) }
        );
        let accent = Self::hex_to_color(&config.accent).unwrap_or(
            Color::from_rgb(0.231, 0.510, 0.965)
        );
        let text_dim = Self::hex_to_color(&config.foreground_muted).unwrap_or(
            Color::from_rgb(0.392, 0.439, 0.529)
        );
        let gridline = Self::hex_to_color(&config.gridline).unwrap_or(
            Self::hex_to_color(&config.border).unwrap_or(
                if is_dark { Self::lighten(bg, 0.10) } else { Self::darken(bg, 0.10) }
            )
        );
        let header_bg = Self::hex_to_color(&config.header_background).unwrap_or(
            if is_dark { Self::lighten(bg, 0.06) } else { Self::darken(bg, 0.03) }
        );
        let border = Self::hex_to_color(&config.border).unwrap_or(gridline);

        // Derive cell background with subtle offset
        let cell_bg = if is_dark { Self::lighten(bg, 0.03) } else { Self::darken(bg, 0.01) };

        ThemeColors {
            bg_dark: bg,
            bg_header: header_bg,
            bg_cell: cell_bg,
            bg_input: header_bg,
            text: fg,
            text_dim,
            gridline,
            accent,
            selected: Color::from_rgba(accent.r, accent.g, accent.b, config.selection_alpha),
            selected_border: accent,
            border,
        }
    }
}

/// Theme manager - handles loading and switching themes
pub struct ThemeManager {
    source: ThemeSource,
    current: ThemeColors,
    current_name: String,
    omarchy_mtime: Option<std::time::SystemTime>,
}

impl ThemeManager {
    /// Create a new theme manager with the given source
    pub fn new(source: ThemeSource) -> Self {
        let (current, current_name, omarchy_mtime) = Self::load_theme(&source);
        ThemeManager {
            source,
            current,
            current_name,
            omarchy_mtime,
        }
    }

    /// Get current theme colors
    pub fn theme(&self) -> ThemeColors {
        self.current
    }

    /// Get current theme name
    pub fn name(&self) -> &str {
        &self.current_name
    }

    /// Get current source
    pub fn source(&self) -> &ThemeSource {
        &self.source
    }

    /// Set theme source and reload
    pub fn set_source(&mut self, source: ThemeSource) {
        self.source = source;
        let (theme, name, mtime) = Self::load_theme(&self.source);
        self.current = theme;
        self.current_name = name;
        self.omarchy_mtime = mtime;
    }

    /// Check if Omarchy theme changed and reload if needed
    pub fn check_omarchy_update(&mut self) -> bool {
        if !matches!(self.source, ThemeSource::Auto | ThemeSource::System) {
            return false;
        }

        let current_mtime = crate::omarchy::theme_mtime();
        if current_mtime != self.omarchy_mtime {
            self.omarchy_mtime = current_mtime;
            let (theme, name, _) = Self::load_theme(&self.source);
            self.current = theme;
            self.current_name = name;
            true
        } else {
            false
        }
    }

    /// Reload current theme
    pub fn reload(&mut self) {
        let (theme, name, mtime) = Self::load_theme(&self.source);
        self.current = theme;
        self.current_name = name;
        self.omarchy_mtime = mtime;
    }

    /// Load theme from source
    fn load_theme(source: &ThemeSource) -> (ThemeColors, String, Option<std::time::SystemTime>) {
        match source {
            ThemeSource::Auto => {
                // Try Omarchy first, then fall back to dark
                if crate::omarchy::is_omarchy() {
                    let theme = crate::omarchy::load_theme();
                    let name = crate::omarchy::current_theme_name()
                        .unwrap_or_else(|| "System".into());
                    let mtime = crate::omarchy::theme_mtime();
                    (theme, name, mtime)
                } else {
                    (ThemeColors::dark(), "Dark".into(), None)
                }
            }
            ThemeSource::Dark => (ThemeColors::dark(), "Dark".into(), None),
            ThemeSource::Light => (ThemeColors::light(), "Light".into(), None),
            ThemeSource::System => {
                if crate::omarchy::is_omarchy() {
                    let theme = crate::omarchy::load_theme();
                    let name = crate::omarchy::current_theme_name()
                        .unwrap_or_else(|| "System".into());
                    let mtime = crate::omarchy::theme_mtime();
                    (theme, name, mtime)
                } else {
                    // Fall back to dark if no system theme
                    (ThemeColors::dark(), "Dark (no system theme)".into(), None)
                }
            }
            ThemeSource::Custom(path) => {
                match Self::load_custom_theme(path) {
                    Some((theme, name)) => (theme, name, None),
                    None => {
                        eprintln!("Failed to load custom theme: {}", path);
                        (ThemeColors::dark(), "Dark (fallback)".into(), None)
                    }
                }
            }
        }
    }

    /// Load a custom theme from JSON file
    fn load_custom_theme(path: &str) -> Option<(ThemeColors, String)> {
        // Expand ~ to home directory
        let expanded = if path.starts_with("~/") {
            let home = std::env::var("HOME").ok()?;
            path.replacen("~", &home, 1)
        } else {
            path.to_string()
        };

        let content = fs::read_to_string(&expanded).ok()?;
        let config: ThemeConfig = serde_json::from_str(&content).ok()?;
        let theme = ThemeColors::from_config(&config.colors, config.is_dark);
        Some((theme, config.name))
    }

    /// List available themes (built-in + custom)
    pub fn list_themes() -> Vec<(String, ThemeSource)> {
        let mut themes = vec![
            ("Dark".into(), ThemeSource::Dark),
            ("Light".into(), ThemeSource::Light),
        ];

        // Add system theme if available
        if crate::omarchy::is_omarchy() {
            if let Some(name) = crate::omarchy::current_theme_name() {
                themes.insert(0, (format!("System: {}", name), ThemeSource::System));
            } else {
                themes.insert(0, ("System".into(), ThemeSource::System));
            }
        }

        // Add custom themes from config directory
        if let Some(theme_dir) = Self::custom_themes_dir() {
            if let Ok(entries) = fs::read_dir(&theme_dir) {
                for entry in entries.flatten() {
                    let path = entry.path();
                    if path.extension().map(|e| e == "json").unwrap_or(false) {
                        if let Ok(content) = fs::read_to_string(&path) {
                            if let Ok(config) = serde_json::from_str::<ThemeConfig>(&content) {
                                let path_str = path.to_string_lossy().to_string();
                                themes.push((config.name, ThemeSource::Custom(path_str)));
                            }
                        }
                    }
                }
            }
        }

        themes
    }

    /// Get the custom themes directory path
    pub fn custom_themes_dir() -> Option<PathBuf> {
        let config_dir = dirs::config_dir()?;
        Some(config_dir.join("visigrid").join("themes"))
    }

    /// Create example custom theme file
    pub fn create_example_theme() -> Result<PathBuf, String> {
        let theme_dir = Self::custom_themes_dir()
            .ok_or_else(|| "Could not determine config directory".to_string())?;

        fs::create_dir_all(&theme_dir)
            .map_err(|e| format!("Failed to create themes directory: {}", e))?;

        let example_path = theme_dir.join("example.json");

        let example = ThemeConfig {
            name: "Example Theme".into(),
            is_dark: true,
            colors: ThemeColorsConfig {
                background: "#1a1b26".into(),
                background_secondary: "#24283b".into(),
                foreground: "#c0caf5".into(),
                foreground_muted: "#565f89".into(),
                gridline: "#3b4261".into(),
                header_background: "#1f2335".into(),
                accent: "#7aa2f7".into(),
                selection_alpha: 0.25,
                border: "#3b4261".into(),
            },
        };

        let json = serde_json::to_string_pretty(&example)
            .map_err(|e| format!("Failed to serialize theme: {}", e))?;

        fs::write(&example_path, json)
            .map_err(|e| format!("Failed to write theme file: {}", e))?;

        Ok(example_path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hex_to_color() {
        let color = ThemeColors::hex_to_color("#3b82f6").unwrap();
        assert!((color.r - 0.231).abs() < 0.01);
        assert!((color.g - 0.510).abs() < 0.01);
        assert!((color.b - 0.965).abs() < 0.01);
    }

    #[test]
    fn test_dark_theme() {
        let theme = ThemeColors::dark();
        assert!(theme.bg_dark.r < 0.1); // Should be dark
    }

    #[test]
    fn test_light_theme() {
        let theme = ThemeColors::light();
        assert!(theme.bg_dark.r > 0.9); // Should be light
    }
}
