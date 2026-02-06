//! Color palette data and utilities for the color picker.
//!
//! Provides theme colors, standard colors, tint/shade computation,
//! hex parsing, RGBA→HSLA conversion, and the picker's UI state.

use gpui::{FocusHandle, Hsla};

/// Which color property the picker is editing.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ColorTarget {
    Fill,
    Text,
    Border,
}

impl ColorTarget {
    pub fn title(self) -> &'static str {
        match self {
            Self::Fill => "Fill Color",
            Self::Text => "Text Color",
            Self::Border => "Border Color",
        }
    }

    /// Whether the "No Fill" / "Automatic" clear option is available for this target.
    /// For Text, None means Automatic (use theme default text color).
    /// For Border, None means Automatic (use theme default border color).
    pub fn allow_none(self) -> bool {
        match self {
            Self::Fill => true,
            Self::Text => true,
            Self::Border => true,
        }
    }
}

/// Transient UI state for the color picker modal.
///
/// Lives on `Spreadsheet` as a bundled struct — not serialized, not
/// part of the document model.
pub struct ColorPickerState {
    pub target: ColorTarget,
    pub hex_input: String,
    pub all_selected: bool,
    pub focus: FocusHandle,
    pub recent_colors: Vec<[u8; 4]>,
}

impl ColorPickerState {
    pub fn new(focus: FocusHandle) -> Self {
        Self {
            target: ColorTarget::Fill,
            hex_input: String::new(),
            all_selected: false,
            focus,
            recent_colors: Vec::new(),
        }
    }

    /// Reset input state (called on show/hide).
    pub fn reset(&mut self) {
        self.hex_input.clear();
        self.all_selected = false;
    }

    /// Push a color to the recent list (deduplicate, move-to-front, cap 10).
    pub fn push_recent(&mut self, color: [u8; 4]) {
        self.recent_colors.retain(|c| *c != color);
        self.recent_colors.insert(0, color);
        self.recent_colors.truncate(10);
    }
}

/// Theme base colors (Excel-aligned, 10 colors)
pub const THEME_COLORS: [[u8; 4]; 10] = [
    [255, 255, 255, 255], // White
    [0, 0, 0, 255],       // Black
    [68, 84, 106, 255],   // Dark Gray
    [132, 151, 176, 255], // Gray
    [214, 220, 228, 255], // Light Gray
    [68, 114, 196, 255],  // Blue
    [237, 125, 49, 255],  // Orange
    [165, 165, 165, 255], // Gray-Green
    [255, 192, 0, 255],   // Gold
    [91, 155, 213, 255],  // Teal
];

/// Standard colors row (10 colors)
pub const STANDARD_COLORS: [[u8; 4]; 10] = [
    [192, 0, 0, 255],     // Dark Red
    [255, 0, 0, 255],     // Red
    [255, 192, 0, 255],   // Orange
    [255, 255, 0, 255],   // Yellow
    [146, 208, 80, 255],  // Light Green
    [0, 176, 80, 255],    // Green
    [0, 176, 240, 255],   // Light Blue
    [0, 112, 192, 255],   // Blue
    [0, 32, 96, 255],     // Dark Blue
    [112, 48, 160, 255],  // Purple
];

/// Mix a color toward white by `factor` (0.0 = original, 1.0 = white).
/// sRGB linear interpolation per channel.
pub fn tint(color: [u8; 4], factor: f32) -> [u8; 4] {
    [
        (color[0] as f32 + (255.0 - color[0] as f32) * factor).round() as u8,
        (color[1] as f32 + (255.0 - color[1] as f32) * factor).round() as u8,
        (color[2] as f32 + (255.0 - color[2] as f32) * factor).round() as u8,
        color[3],
    ]
}

/// Mix a color toward black by `factor` (0.0 = original, 1.0 = black).
/// sRGB linear interpolation per channel.
pub fn shade(color: [u8; 4], factor: f32) -> [u8; 4] {
    [
        (color[0] as f32 * (1.0 - factor)).round() as u8,
        (color[1] as f32 * (1.0 - factor)).round() as u8,
        (color[2] as f32 * (1.0 - factor)).round() as u8,
        color[3],
    ]
}

/// Generate the 6×10 theme grid (tints → base → shades).
///
/// Row 0: 80% tint (lightest)
/// Row 1: 60% tint
/// Row 2: 40% tint
/// Row 3: base color
/// Row 4: 25% shade
/// Row 5: 50% shade (darkest)
///
/// Returns 60 colors in row-major order: grid[row * 10 + col]
pub fn theme_grid() -> [[u8; 4]; 60] {
    let mut grid = [[0u8; 4]; 60];
    for col in 0..10 {
        let base = THEME_COLORS[col];
        grid[0 * 10 + col] = tint(base, 0.80);
        grid[1 * 10 + col] = tint(base, 0.60);
        grid[2 * 10 + col] = tint(base, 0.40);
        grid[3 * 10 + col] = base;
        grid[4 * 10 + col] = shade(base, 0.25);
        grid[5 * 10 + col] = shade(base, 0.50);
    }
    grid
}

/// Parse a hex color string into [R, G, B, A].
///
/// Supported formats:
/// - `#RRGGBB` or `RRGGBB`
/// - `#RGB` (expanded to RRGGBB)
/// - `rgb(R, G, B)` with decimal values
pub fn parse_hex_color(input: &str) -> Option<[u8; 4]> {
    let trimmed = input.trim();

    // Handle rgb(r, g, b) format
    if let Some(inner) = trimmed.strip_prefix("rgb(").and_then(|s| s.strip_suffix(')')) {
        let parts: Vec<&str> = inner.split(',').collect();
        if parts.len() == 3 {
            let r = parts[0].trim().parse::<u8>().ok()?;
            let g = parts[1].trim().parse::<u8>().ok()?;
            let b = parts[2].trim().parse::<u8>().ok()?;
            return Some([r, g, b, 255]);
        }
        return None;
    }

    // Strip leading # if present
    let hex = trimmed.strip_prefix('#').unwrap_or(trimmed);

    match hex.len() {
        3 => {
            // #RGB → #RRGGBB
            let r = u8::from_str_radix(&hex[0..1], 16).ok()?;
            let g = u8::from_str_radix(&hex[1..2], 16).ok()?;
            let b = u8::from_str_radix(&hex[2..3], 16).ok()?;
            Some([r * 17, g * 17, b * 17, 255])
        }
        6 => {
            let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
            let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
            let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
            Some([r, g, b, 255])
        }
        _ => None,
    }
}

/// Try to extract a color token from arbitrary pasted text.
///
/// First tries `parse_hex_color` on the whole string. If that fails,
/// scans for the first `#RRGGBB`, `#RGB`, or `rgb(...)` substring.
/// Returns the cleaned token string (not the color value) so the
/// caller can display it in the hex input field.
pub fn extract_color_token(input: &str) -> Option<String> {
    let trimmed = input.trim();
    // Fast path: whole string is a valid color
    if parse_hex_color(trimmed).is_some() {
        return Some(trimmed.to_string());
    }
    // Scan for #RRGGBB or #RGB
    if let Some(hash_pos) = trimmed.find('#') {
        let after = &trimmed[hash_pos..];
        // Try 7-char (#RRGGBB) then 4-char (#RGB)
        for len in [7, 4] {
            if after.len() >= len {
                let candidate = &after[..len];
                if parse_hex_color(candidate).is_some() {
                    return Some(candidate.to_string());
                }
            }
        }
    }
    // Scan for rgb(...)
    if let Some(rgb_pos) = trimmed.find("rgb(") {
        if let Some(close) = trimmed[rgb_pos..].find(')') {
            let candidate = &trimmed[rgb_pos..rgb_pos + close + 1];
            if parse_hex_color(candidate).is_some() {
                return Some(candidate.to_string());
            }
        }
    }
    None
}

/// Convert [R, G, B, A] to "#RRGGBB" hex string.
pub fn to_hex(color: [u8; 4]) -> String {
    format!("#{:02X}{:02X}{:02X}", color[0], color[1], color[2])
}

/// Convert [R, G, B, A] to gpui `Hsla` for rendering.
pub fn rgba_to_hsla(color: [u8; 4]) -> Hsla {
    Hsla::from(gpui::Rgba {
        r: color[0] as f32 / 255.0,
        g: color[1] as f32 / 255.0,
        b: color[2] as f32 / 255.0,
        a: color[3] as f32 / 255.0,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tint() {
        // Pure black tinted 50% → [128, 128, 128]
        let result = tint([0, 0, 0, 255], 0.50);
        assert_eq!(result, [128, 128, 128, 255]);
    }

    #[test]
    fn test_shade() {
        // Pure white shaded 50% → [128, 128, 128]
        let result = shade([255, 255, 255, 255], 0.50);
        assert_eq!(result, [128, 128, 128, 255]);
    }

    #[test]
    fn test_parse_hex_6() {
        assert_eq!(parse_hex_color("#FF6600"), Some([255, 102, 0, 255]));
        assert_eq!(parse_hex_color("FF6600"), Some([255, 102, 0, 255]));
    }

    #[test]
    fn test_parse_hex_3() {
        assert_eq!(parse_hex_color("#F60"), Some([255, 102, 0, 255]));
    }

    #[test]
    fn test_parse_rgb() {
        assert_eq!(parse_hex_color("rgb(100, 200, 50)"), Some([100, 200, 50, 255]));
    }

    #[test]
    fn test_parse_invalid() {
        assert_eq!(parse_hex_color("not-a-color"), None);
        assert_eq!(parse_hex_color("#GGHHII"), None);
    }

    #[test]
    fn test_to_hex() {
        assert_eq!(to_hex([255, 102, 0, 255]), "#FF6600");
    }

    #[test]
    fn test_theme_grid_size() {
        let grid = theme_grid();
        assert_eq!(grid.len(), 60);
        // Row 3 should be the base colors
        for col in 0..10 {
            assert_eq!(grid[3 * 10 + col], THEME_COLORS[col]);
        }
    }

    #[test]
    fn test_extract_whole_hex() {
        assert_eq!(extract_color_token("#FF6600"), Some("#FF6600".to_string()));
        assert_eq!(extract_color_token("FF6600"), Some("FF6600".to_string()));
        assert_eq!(extract_color_token("#F60"), Some("#F60".to_string()));
    }

    #[test]
    fn test_extract_hex_from_text() {
        assert_eq!(
            extract_color_token("color: #AA33BB; font-size: 12px"),
            Some("#AA33BB".to_string())
        );
    }

    #[test]
    fn test_extract_short_hex_from_text() {
        assert_eq!(
            extract_color_token("background: #F60"),
            Some("#F60".to_string())
        );
    }

    #[test]
    fn test_extract_rgb_from_text() {
        assert_eq!(
            extract_color_token("background-color: rgb(100, 200, 50)"),
            Some("rgb(100, 200, 50)".to_string())
        );
    }

    #[test]
    fn test_extract_no_color() {
        assert_eq!(extract_color_token("no color here"), None);
        assert_eq!(extract_color_token(""), None);
    }
}
