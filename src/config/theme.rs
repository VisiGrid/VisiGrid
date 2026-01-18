// Theme configuration

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ThemeColors {
    // Base colors
    pub background: String,
    pub background_secondary: String,
    pub foreground: String,
    pub foreground_muted: String,

    // Grid
    pub grid_lines: String,
    pub header_background: String,
    pub header_foreground: String,

    // Selection
    pub selection_fill: String,
    pub selection_border: String,
    pub selection_fill_alpha: f32,

    // Accent
    pub accent: String,
    pub accent_hover: String,

    // UI elements
    pub border: String,
    pub border_muted: String,
    pub formula_bar_background: String,

    // Command palette
    pub palette_background: String,
    pub palette_input_background: String,
    pub palette_selected: String,
    pub palette_hint: String,

    // Border radius
    pub border_radius: f32,
    pub border_radius_lg: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Theme {
    pub name: String,
    pub colors: ThemeColors,
}

impl Default for Theme {
    fn default() -> Self {
        dark_theme()
    }
}

pub fn dark_theme() -> Theme {
    Theme {
        name: "Dark".into(),
        colors: ThemeColors {
            // Base colors (grid-950, grid-900, grid-100, grid-400)
            background: "#020617".into(),
            background_secondary: "#0f172a".into(),
            foreground: "#f1f5f9".into(),
            foreground_muted: "#94a3b8".into(),

            // Grid
            grid_lines: "#334155".into(),
            header_background: "#1e293b".into(),
            header_foreground: "#cbd5e1".into(),

            // Selection (accent blue with transparency)
            selection_fill: "#3b82f6".into(),
            selection_border: "#3b82f6".into(),
            selection_fill_alpha: 0.2,

            // Accent
            accent: "#3b82f6".into(),
            accent_hover: "#2563eb".into(),

            // UI elements
            border: "#334155".into(),
            border_muted: "#1e293b".into(),
            formula_bar_background: "#0f172a".into(),

            // Command palette
            palette_background: "#1e293b".into(),
            palette_input_background: "#0f172a".into(),
            palette_selected: "#3b82f6".into(),
            palette_hint: "#64748b".into(),

            // Border radius
            border_radius: 4.0,
            border_radius_lg: 8.0,
        },
    }
}

pub fn light_theme() -> Theme {
    Theme {
        name: "Light".into(),
        colors: ThemeColors {
            // Base colors (grid-50, grid-100, grid-900, grid-600)
            background: "#f8fafc".into(),
            background_secondary: "#f1f5f9".into(),
            foreground: "#0f172a".into(),
            foreground_muted: "#475569".into(),

            // Grid
            grid_lines: "#e2e8f0".into(),
            header_background: "#e2e8f0".into(),
            header_foreground: "#334155".into(),

            // Selection
            selection_fill: "#3b82f6".into(),
            selection_border: "#3b82f6".into(),
            selection_fill_alpha: 0.15,

            // Accent
            accent: "#3b82f6".into(),
            accent_hover: "#2563eb".into(),

            // UI elements
            border: "#cbd5e1".into(),
            border_muted: "#e2e8f0".into(),
            formula_bar_background: "#f1f5f9".into(),

            // Command palette
            palette_background: "#ffffff".into(),
            palette_input_background: "#f1f5f9".into(),
            palette_selected: "#3b82f6".into(),
            palette_hint: "#64748b".into(),

            // Border radius
            border_radius: 4.0,
            border_radius_lg: 8.0,
        },
    }
}

impl ThemeColors {
    /// Parse a hex color string to RGB components (0-255)
    pub fn hex_to_rgb(hex: &str) -> Option<(u8, u8, u8)> {
        let hex = hex.trim_start_matches('#');
        if hex.len() != 6 {
            return None;
        }
        let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
        let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
        let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
        Some((r, g, b))
    }

    /// Get selection fill color with alpha as RGBA
    pub fn selection_fill_rgba(&self) -> Option<(u8, u8, u8, f32)> {
        let (r, g, b) = Self::hex_to_rgb(&self.selection_fill)?;
        Some((r, g, b, self.selection_fill_alpha))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hex_to_rgb() {
        assert_eq!(ThemeColors::hex_to_rgb("#3b82f6"), Some((59, 130, 246)));
        assert_eq!(ThemeColors::hex_to_rgb("#000000"), Some((0, 0, 0)));
        assert_eq!(ThemeColors::hex_to_rgb("#ffffff"), Some((255, 255, 255)));
        assert_eq!(ThemeColors::hex_to_rgb("3b82f6"), Some((59, 130, 246)));
    }

    #[test]
    fn test_selection_fill_rgba() {
        let theme = dark_theme();
        let rgba = theme.colors.selection_fill_rgba().unwrap();
        assert_eq!(rgba.0, 59);  // R
        assert_eq!(rgba.1, 130); // G
        assert_eq!(rgba.2, 246); // B
        assert!((rgba.3 - 0.2).abs() < 0.001); // Alpha
    }
}
