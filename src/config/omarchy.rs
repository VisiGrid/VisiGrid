// Omarchy theme integration
// Reads the current Omarchy theme and maps it to VisiGrid colors

use std::fs;
use std::path::PathBuf;
use std::time::SystemTime;

use serde::Deserialize;

use crate::app::ThemeColors;
use iced::Color;

/// Omarchy colors.toml structure
#[derive(Debug, Deserialize)]
pub struct OmarchyColors {
    pub accent: String,
    pub cursor: String,
    pub foreground: String,
    pub background: String,
    pub selection_foreground: String,
    pub selection_background: String,
    // ANSI colors 0-15
    pub color0: String,
    pub color1: String,
    pub color2: String,
    pub color3: String,
    pub color4: String,
    pub color5: String,
    pub color6: String,
    pub color7: String,
    pub color8: String,
    pub color9: String,
    pub color10: String,
    pub color11: String,
    pub color12: String,
    pub color13: String,
    pub color14: String,
    pub color15: String,
}

impl OmarchyColors {
    /// Path to current Omarchy theme colors
    pub fn config_path() -> Option<PathBuf> {
        let home = std::env::var("HOME").ok()?;
        let path = PathBuf::from(home)
            .join(".config/omarchy/current/theme/colors.toml");
        if path.exists() {
            Some(path)
        } else {
            None
        }
    }

    /// Load colors from the current Omarchy theme
    pub fn load() -> Option<Self> {
        let path = Self::config_path()?;
        let content = fs::read_to_string(path).ok()?;
        toml::from_str(&content).ok()
    }

    /// Parse hex color to RGB floats (0.0-1.0)
    fn hex_to_rgb(hex: &str) -> Option<(f32, f32, f32)> {
        let hex = hex.trim_start_matches('#');
        if hex.len() != 6 {
            return None;
        }
        let r = u8::from_str_radix(&hex[0..2], 16).ok()? as f32 / 255.0;
        let g = u8::from_str_radix(&hex[2..4], 16).ok()? as f32 / 255.0;
        let b = u8::from_str_radix(&hex[4..6], 16).ok()? as f32 / 255.0;
        Some((r, g, b))
    }

    /// Lighten a color by mixing with white
    fn lighten(r: f32, g: f32, b: f32, amount: f32) -> (f32, f32, f32) {
        (
            r + (1.0 - r) * amount,
            g + (1.0 - g) * amount,
            b + (1.0 - b) * amount,
        )
    }

    /// Convert Omarchy colors to VisiGrid theme
    pub fn to_theme_colors(&self) -> ThemeColors {
        // Parse base colors
        let (bg_r, bg_g, bg_b) = Self::hex_to_rgb(&self.background)
            .unwrap_or((0.008, 0.024, 0.090));
        let (fg_r, fg_g, fg_b) = Self::hex_to_rgb(&self.foreground)
            .unwrap_or((0.945, 0.961, 0.976));
        let (accent_r, accent_g, accent_b) = Self::hex_to_rgb(&self.accent)
            .unwrap_or((0.231, 0.510, 0.965));
        let (dim_r, dim_g, dim_b) = Self::hex_to_rgb(&self.color8)
            .unwrap_or((0.392, 0.439, 0.529));

        // Derive secondary colors with subtle lightening for depth
        let (cell_r, cell_g, cell_b) = Self::lighten(bg_r, bg_g, bg_b, 0.03);  // Subtle lift for cells
        let (header_r, header_g, header_b) = Self::lighten(bg_r, bg_g, bg_b, 0.06);
        let (border_r, border_g, border_b) = Self::lighten(bg_r, bg_g, bg_b, 0.10);

        ThemeColors {
            bg_dark: Color::from_rgb(bg_r, bg_g, bg_b),
            bg_header: Color::from_rgb(header_r, header_g, header_b),
            bg_cell: Color::from_rgb(cell_r, cell_g, cell_b),  // Slightly lighter than background
            bg_input: Color::from_rgb(header_r, header_g, header_b),
            text: Color::from_rgb(fg_r, fg_g, fg_b),
            text_dim: Color::from_rgb(dim_r, dim_g, dim_b),
            gridline: Color::from_rgb(border_r, border_g, border_b),
            accent: Color::from_rgb(accent_r, accent_g, accent_b),
            selected: Color::from_rgba(accent_r, accent_g, accent_b, 0.2),
            selected_border: Color::from_rgb(accent_r, accent_g, accent_b),
            border: Color::from_rgb(border_r, border_g, border_b),
        }
    }
}

/// Check if running on Omarchy (has theme config)
pub fn is_omarchy() -> bool {
    OmarchyColors::config_path().is_some()
}

/// Load theme from Omarchy, falling back to default dark theme
pub fn load_theme() -> ThemeColors {
    OmarchyColors::load()
        .map(|c| c.to_theme_colors())
        .unwrap_or_else(ThemeColors::dark)
}

/// Get current Omarchy theme name
pub fn current_theme_name() -> Option<String> {
    let home = std::env::var("HOME").ok()?;
    let path = PathBuf::from(home)
        .join(".config/omarchy/current/theme.name");
    fs::read_to_string(path).ok().map(|s| s.trim().to_string())
}

/// Get the modification time of the theme files
pub fn theme_mtime() -> Option<SystemTime> {
    let colors_path = OmarchyColors::config_path()?;
    let colors_mtime = fs::metadata(&colors_path).ok()?.modified().ok()?;

    // Also check theme.name file
    let home = std::env::var("HOME").ok()?;
    let name_path = PathBuf::from(home).join(".config/omarchy/current/theme.name");
    let name_mtime = fs::metadata(&name_path).ok()?.modified().ok();

    // Return the most recent modification time
    match name_mtime {
        Some(t) if t > colors_mtime => Some(t),
        _ => Some(colors_mtime),
    }
}
