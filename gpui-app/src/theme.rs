//! Theme system for VisiGrid
//!
//! Themes are defined using semantic tokens that map to colors.
//! This allows consistent theming across the entire application.

use gpui::Hsla;
use std::collections::HashMap;

/// Theme appearance - light or dark base
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Appearance {
    Light,
    #[default]
    Dark,
}

/// All semantic color tokens used in the application.
/// Strongly typed for compile-time safety and IDE support.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TokenKey {
    // App surfaces
    AppBg,
    PanelBg,
    PanelBorder,
    TextPrimary,
    TextMuted,
    TextDisabled,
    TextInverse,

    // Grid surfaces
    GridBg,
    GridLines,
    GridLinesBold,

    // Headers
    HeaderBg,
    HeaderText,
    HeaderTextMuted,
    HeaderBorder,
    HeaderHoverBg,
    HeaderActiveBg,

    // Cells
    CellBg,
    CellBgAlt,
    CellText,
    CellTextMuted,
    CellBorderFocus,
    CellHoverBg,

    // Selection + cursor
    SelectionBg,
    SelectionBorder,
    SelectionHandle,
    CursorBg,
    CursorText,

    // Formula bar + editor fields
    EditorBg,
    EditorBorder,
    EditorText,
    EditorPlaceholder,
    EditorSelectionBg,
    EditorSelectionText,

    // Status + chrome
    StatusBg,
    StatusText,
    StatusTextMuted,
    ToolbarBg,
    ToolbarBorder,
    ToolbarButtonHoverBg,
    ToolbarButtonActiveBg,

    // Semantic feedback
    Accent,
    Ok,
    Warn,
    Error,
    ErrorBg,
    Link,

    // Spreadsheet semantics
    FormulaText,
    RefHighlight1,
    RefHighlight2,
    RefHighlight3,
    SpillBorder,
    CommentIndicator,
}

impl TokenKey {
    /// Get all token keys for validation
    pub const ALL: &'static [TokenKey] = &[
        // App surfaces
        TokenKey::AppBg,
        TokenKey::PanelBg,
        TokenKey::PanelBorder,
        TokenKey::TextPrimary,
        TokenKey::TextMuted,
        TokenKey::TextDisabled,
        TokenKey::TextInverse,
        // Grid surfaces
        TokenKey::GridBg,
        TokenKey::GridLines,
        TokenKey::GridLinesBold,
        // Headers
        TokenKey::HeaderBg,
        TokenKey::HeaderText,
        TokenKey::HeaderTextMuted,
        TokenKey::HeaderBorder,
        TokenKey::HeaderHoverBg,
        TokenKey::HeaderActiveBg,
        // Cells
        TokenKey::CellBg,
        TokenKey::CellBgAlt,
        TokenKey::CellText,
        TokenKey::CellTextMuted,
        TokenKey::CellBorderFocus,
        TokenKey::CellHoverBg,
        // Selection + cursor
        TokenKey::SelectionBg,
        TokenKey::SelectionBorder,
        TokenKey::SelectionHandle,
        TokenKey::CursorBg,
        TokenKey::CursorText,
        // Formula bar + editor fields
        TokenKey::EditorBg,
        TokenKey::EditorBorder,
        TokenKey::EditorText,
        TokenKey::EditorPlaceholder,
        TokenKey::EditorSelectionBg,
        TokenKey::EditorSelectionText,
        // Status + chrome
        TokenKey::StatusBg,
        TokenKey::StatusText,
        TokenKey::StatusTextMuted,
        TokenKey::ToolbarBg,
        TokenKey::ToolbarBorder,
        TokenKey::ToolbarButtonHoverBg,
        TokenKey::ToolbarButtonActiveBg,
        // Semantic feedback
        TokenKey::Accent,
        TokenKey::Ok,
        TokenKey::Warn,
        TokenKey::Error,
        TokenKey::ErrorBg,
        TokenKey::Link,
        // Spreadsheet semantics
        TokenKey::FormulaText,
        TokenKey::RefHighlight1,
        TokenKey::RefHighlight2,
        TokenKey::RefHighlight3,
        TokenKey::SpillBorder,
        TokenKey::CommentIndicator,
    ];
}

/// Theme metadata
#[derive(Debug, Clone)]
pub struct ThemeMeta {
    pub id: &'static str,
    pub name: &'static str,
    pub author: &'static str,
    pub appearance: Appearance,
}

/// Typography settings for a theme
#[derive(Debug, Clone)]
pub struct ThemeTypography {
    pub font_family: Option<String>,
    pub font_size: f32,
    pub mono_family: Option<String>,
}

impl Default for ThemeTypography {
    fn default() -> Self {
        Self {
            font_family: None,  // Use system default
            font_size: 12.0,
            mono_family: None,
        }
    }
}

/// A complete theme definition
#[derive(Debug, Clone)]
pub struct Theme {
    pub meta: ThemeMeta,
    pub tokens: HashMap<TokenKey, Hsla>,
    pub typography: ThemeTypography,
}

impl Theme {
    /// Get a token color, panics if not found (should never happen with resolved themes)
    pub fn get(&self, key: TokenKey) -> Hsla {
        *self.tokens.get(&key).unwrap_or_else(|| {
            panic!("Missing theme token: {:?}", key)
        })
    }

    /// Validate theme has all required tokens
    pub fn validate(&self) -> Vec<String> {
        let mut warnings = Vec::new();
        for key in TokenKey::ALL {
            if !self.tokens.contains_key(key) {
                warnings.push(format!("Missing token: {:?}", key));
            }
        }
        warnings
    }
}

/// Helper to create Hsla from hex RGB
pub fn rgb(hex: u32) -> Hsla {
    gpui::rgb(hex).into()
}

/// Helper to create Hsla from hex RGBA
pub fn rgba(hex: u32) -> Hsla {
    gpui::rgba(hex).into()
}

// ============================================================================
// Built-in Themes
// ============================================================================

/// VisiGrid default theme - modern dark with slate-blue tones
/// Colors based on the VisiGrid landing page palette
pub fn visigrid_theme() -> Theme {
    let mut tokens = HashMap::new();

    // Slate-blue palette from landing page (Tailwind slate)
    // Full palette for reference:
    // 950: #020617, 900: #0f172a, 800: #1e293b, 700: #334155
    // 600: #475569, 500: #64748b, 400: #94a3b8, 300: #cbd5e1
    // 200: #e2e8f0, 100: #f1f5f9, 50: #f8fafc
    let grid_900 = rgb(0x0f172a);  // Main background
    let grid_800 = rgb(0x1e293b);  // Panels, headers
    let grid_700 = rgb(0x334155);  // Borders, subtle elements
    let grid_600 = rgb(0x475569);  // Disabled text
    let grid_500 = rgb(0x64748b);  // Muted elements
    let grid_400 = rgb(0x94a3b8);  // Secondary text
    let grid_200 = rgb(0xe2e8f0);  // Bright text
    let grid_100 = rgb(0xf1f5f9);  // Brightest
    let accent = rgb(0x3b82f6);    // Blue accent
    let accent_hover = rgb(0x2563eb);

    // App surfaces
    tokens.insert(TokenKey::AppBg, grid_900);
    tokens.insert(TokenKey::PanelBg, grid_800);
    tokens.insert(TokenKey::PanelBorder, grid_700);
    tokens.insert(TokenKey::TextPrimary, grid_200);
    tokens.insert(TokenKey::TextMuted, grid_400);
    tokens.insert(TokenKey::TextDisabled, grid_600);
    tokens.insert(TokenKey::TextInverse, grid_900);

    // Grid surfaces
    tokens.insert(TokenKey::GridBg, grid_900);
    tokens.insert(TokenKey::GridLines, rgba(0x33415540));  // Subtle gridlines (25% opacity)
    tokens.insert(TokenKey::GridLinesBold, grid_700);

    // Headers
    tokens.insert(TokenKey::HeaderBg, grid_800);
    tokens.insert(TokenKey::HeaderText, grid_200);
    tokens.insert(TokenKey::HeaderTextMuted, grid_400);
    tokens.insert(TokenKey::HeaderBorder, grid_700);
    tokens.insert(TokenKey::HeaderHoverBg, grid_700);
    tokens.insert(TokenKey::HeaderActiveBg, accent_hover);

    // Cells
    tokens.insert(TokenKey::CellBg, grid_900);
    tokens.insert(TokenKey::CellBgAlt, rgb(0x131c2e));  // Slightly lighter than grid_900
    tokens.insert(TokenKey::CellText, grid_200);
    tokens.insert(TokenKey::CellTextMuted, grid_400);
    tokens.insert(TokenKey::CellBorderFocus, accent);
    tokens.insert(TokenKey::CellHoverBg, grid_800);

    // Selection + cursor
    tokens.insert(TokenKey::SelectionBg, rgba(0x3b82f625));  // Range selection - subtle (15% opacity)
    tokens.insert(TokenKey::SelectionBorder, rgba(0x3b82f680));  // Selection border - visible but not harsh
    tokens.insert(TokenKey::SelectionHandle, accent);
    tokens.insert(TokenKey::CursorBg, grid_100);
    tokens.insert(TokenKey::CursorText, grid_900);

    // Formula bar + editor (subtle highlight within the blue theme)
    tokens.insert(TokenKey::EditorBg, rgb(0x1a2744));  // Slightly blue-tinted
    tokens.insert(TokenKey::EditorBorder, accent);
    tokens.insert(TokenKey::EditorText, grid_100);
    tokens.insert(TokenKey::EditorPlaceholder, grid_500);
    tokens.insert(TokenKey::EditorSelectionBg, rgba(0x3b82f660));
    tokens.insert(TokenKey::EditorSelectionText, grid_100);

    // Status + chrome
    tokens.insert(TokenKey::StatusBg, grid_800);
    tokens.insert(TokenKey::StatusText, grid_200);
    tokens.insert(TokenKey::StatusTextMuted, grid_400);
    tokens.insert(TokenKey::ToolbarBg, grid_800);
    tokens.insert(TokenKey::ToolbarBorder, grid_700);
    tokens.insert(TokenKey::ToolbarButtonHoverBg, grid_700);
    tokens.insert(TokenKey::ToolbarButtonActiveBg, accent_hover);

    // Semantic feedback
    tokens.insert(TokenKey::Accent, accent);
    tokens.insert(TokenKey::Ok, rgb(0x22c55e));       // Green
    tokens.insert(TokenKey::Warn, rgb(0xeab308));     // Yellow
    tokens.insert(TokenKey::Error, rgb(0xef4444));    // Red
    tokens.insert(TokenKey::ErrorBg, rgba(0xef444420));
    tokens.insert(TokenKey::Link, rgb(0x60a5fa));     // Light blue

    // Spreadsheet semantics
    tokens.insert(TokenKey::FormulaText, rgb(0x93c5fd));  // Light blue
    tokens.insert(TokenKey::RefHighlight1, rgb(0x22c55e)); // Green
    tokens.insert(TokenKey::RefHighlight2, rgb(0xfbbf24)); // Amber
    tokens.insert(TokenKey::RefHighlight3, rgb(0xa855f7)); // Purple
    tokens.insert(TokenKey::SpillBorder, rgb(0x60a5fa));   // Light blue
    tokens.insert(TokenKey::CommentIndicator, rgb(0xef4444));

    Theme {
        meta: ThemeMeta {
            id: "visigrid",
            name: "VisiGrid",
            author: "VisiGrid",
            appearance: Appearance::Dark,
        },
        tokens,
        typography: ThemeTypography::default(),
    }
}

/// Classic theme - Excel-inspired light theme
pub fn classic_theme() -> Theme {
    let mut tokens = HashMap::new();

    // App surfaces
    tokens.insert(TokenKey::AppBg, rgb(0xf0f0f0));
    tokens.insert(TokenKey::PanelBg, rgb(0xf5f5f5));
    tokens.insert(TokenKey::PanelBorder, rgb(0xd0d0d0));
    tokens.insert(TokenKey::TextPrimary, rgb(0x000000));
    tokens.insert(TokenKey::TextMuted, rgb(0x666666));
    tokens.insert(TokenKey::TextDisabled, rgb(0x999999));
    tokens.insert(TokenKey::TextInverse, rgb(0xffffff));

    // Grid surfaces
    tokens.insert(TokenKey::GridBg, rgb(0xffffff));
    tokens.insert(TokenKey::GridLines, rgba(0xc0c0c050));  // Subtle gridlines (30% opacity)
    tokens.insert(TokenKey::GridLinesBold, rgb(0xc0c0c0));

    // Headers
    tokens.insert(TokenKey::HeaderBg, rgb(0xe8e8e8));
    tokens.insert(TokenKey::HeaderText, rgb(0x000000));
    tokens.insert(TokenKey::HeaderTextMuted, rgb(0x666666));
    tokens.insert(TokenKey::HeaderBorder, rgb(0xc0c0c0));
    tokens.insert(TokenKey::HeaderHoverBg, rgb(0xd8d8d8));
    tokens.insert(TokenKey::HeaderActiveBg, rgb(0xb4d6fa));

    // Cells
    tokens.insert(TokenKey::CellBg, rgb(0xffffff));
    tokens.insert(TokenKey::CellBgAlt, rgb(0xf8f8f8));
    tokens.insert(TokenKey::CellText, rgb(0x000000));
    tokens.insert(TokenKey::CellTextMuted, rgb(0x666666));
    tokens.insert(TokenKey::CellBorderFocus, rgb(0x217346));
    tokens.insert(TokenKey::CellHoverBg, rgb(0xf0f8ff));

    // Selection + cursor
    tokens.insert(TokenKey::SelectionBg, rgba(0xb4d6fa40));  // Range selection - subtle (25% opacity)
    tokens.insert(TokenKey::SelectionBorder, rgba(0x21734680));  // Selection border - visible but softer
    tokens.insert(TokenKey::SelectionHandle, rgb(0x217346));
    tokens.insert(TokenKey::CursorBg, rgb(0x000000));
    tokens.insert(TokenKey::CursorText, rgb(0xffffff));

    // Formula bar + editor
    tokens.insert(TokenKey::EditorBg, rgb(0xffffff));
    tokens.insert(TokenKey::EditorBorder, rgb(0xc0c0c0));
    tokens.insert(TokenKey::EditorText, rgb(0x000000));
    tokens.insert(TokenKey::EditorPlaceholder, rgb(0x999999));
    tokens.insert(TokenKey::EditorSelectionBg, rgb(0x0078d4));
    tokens.insert(TokenKey::EditorSelectionText, rgb(0xffffff));

    // Status + chrome
    tokens.insert(TokenKey::StatusBg, rgb(0x217346));
    tokens.insert(TokenKey::StatusText, rgb(0xffffff));
    tokens.insert(TokenKey::StatusTextMuted, rgba(0xffffffcc));
    tokens.insert(TokenKey::ToolbarBg, rgb(0xf0f0f0));
    tokens.insert(TokenKey::ToolbarBorder, rgb(0xd0d0d0));
    tokens.insert(TokenKey::ToolbarButtonHoverBg, rgb(0xe0e0e0));
    tokens.insert(TokenKey::ToolbarButtonActiveBg, rgb(0xb4d6fa));

    // Semantic feedback
    tokens.insert(TokenKey::Accent, rgb(0x217346));
    tokens.insert(TokenKey::Ok, rgb(0x107c10));
    tokens.insert(TokenKey::Warn, rgb(0xca5010));
    tokens.insert(TokenKey::Error, rgb(0xa80000));
    tokens.insert(TokenKey::ErrorBg, rgba(0xa8000020));
    tokens.insert(TokenKey::Link, rgb(0x0066cc));

    // Spreadsheet semantics
    tokens.insert(TokenKey::FormulaText, rgb(0x0000ff));
    tokens.insert(TokenKey::RefHighlight1, rgb(0x0000ff));
    tokens.insert(TokenKey::RefHighlight2, rgb(0xff0000));
    tokens.insert(TokenKey::RefHighlight3, rgb(0x800080));
    tokens.insert(TokenKey::SpillBorder, rgb(0x0066cc));
    tokens.insert(TokenKey::CommentIndicator, rgb(0xff0000));

    Theme {
        meta: ThemeMeta {
            id: "classic",
            name: "Classic",
            author: "VisiGrid",
            appearance: Appearance::Light,
        },
        tokens,
        typography: ThemeTypography::default(),
    }
}

/// VisiCalc theme - retro black and green
pub fn visicalc_theme() -> Theme {
    let mut tokens = HashMap::new();

    let black = rgb(0x000000);
    let green = rgb(0x00ff66);
    let green_dim = rgb(0x009944);
    let green_bright = rgb(0x00ffaa);
    let green_bg = rgb(0x003300);

    // App surfaces
    tokens.insert(TokenKey::AppBg, black);
    tokens.insert(TokenKey::PanelBg, black);
    tokens.insert(TokenKey::PanelBorder, green_dim);
    tokens.insert(TokenKey::TextPrimary, green);
    tokens.insert(TokenKey::TextMuted, green_dim);
    tokens.insert(TokenKey::TextDisabled, rgb(0x004400));
    tokens.insert(TokenKey::TextInverse, black);

    // Grid surfaces
    tokens.insert(TokenKey::GridBg, black);
    tokens.insert(TokenKey::GridLines, rgba(0x00990030));  // Faint green gridlines (18% opacity)
    tokens.insert(TokenKey::GridLinesBold, rgb(0x004400));

    // Headers
    tokens.insert(TokenKey::HeaderBg, black);
    tokens.insert(TokenKey::HeaderText, green);
    tokens.insert(TokenKey::HeaderTextMuted, green_dim);
    tokens.insert(TokenKey::HeaderBorder, green_dim);
    tokens.insert(TokenKey::HeaderHoverBg, green_bg);
    tokens.insert(TokenKey::HeaderActiveBg, green_bg);

    // Cells
    tokens.insert(TokenKey::CellBg, black);
    tokens.insert(TokenKey::CellBgAlt, black);
    tokens.insert(TokenKey::CellText, green);
    tokens.insert(TokenKey::CellTextMuted, green_dim);
    tokens.insert(TokenKey::CellBorderFocus, green);
    tokens.insert(TokenKey::CellHoverBg, rgb(0x001100));

    // Selection + cursor
    tokens.insert(TokenKey::SelectionBg, rgba(0x00330060));  // Subtle green for range (37% opacity)
    tokens.insert(TokenKey::SelectionBorder, rgba(0x00ff66a0));  // Selection border - visible but not harsh
    tokens.insert(TokenKey::SelectionHandle, green);
    tokens.insert(TokenKey::CursorBg, green);
    tokens.insert(TokenKey::CursorText, black);

    // Formula bar + editor (slightly lifted from pure black to indicate editing)
    tokens.insert(TokenKey::EditorBg, rgb(0x0a1a0a));  // Very dark green tint
    tokens.insert(TokenKey::EditorBorder, green);
    tokens.insert(TokenKey::EditorText, green_bright);
    tokens.insert(TokenKey::EditorPlaceholder, green_dim);
    tokens.insert(TokenKey::EditorSelectionBg, green_bg);
    tokens.insert(TokenKey::EditorSelectionText, green_bright);

    // Status + chrome
    tokens.insert(TokenKey::StatusBg, black);
    tokens.insert(TokenKey::StatusText, green);
    tokens.insert(TokenKey::StatusTextMuted, green_dim);
    tokens.insert(TokenKey::ToolbarBg, black);
    tokens.insert(TokenKey::ToolbarBorder, green_dim);
    tokens.insert(TokenKey::ToolbarButtonHoverBg, green_bg);
    tokens.insert(TokenKey::ToolbarButtonActiveBg, green_bg);

    // Semantic feedback
    tokens.insert(TokenKey::Accent, green);
    tokens.insert(TokenKey::Ok, green);
    tokens.insert(TokenKey::Warn, rgb(0xffff00));
    tokens.insert(TokenKey::Error, rgb(0xff5555));
    tokens.insert(TokenKey::ErrorBg, rgba(0xff555520));
    tokens.insert(TokenKey::Link, green_bright);

    // Spreadsheet semantics
    tokens.insert(TokenKey::FormulaText, green_bright);
    tokens.insert(TokenKey::RefHighlight1, green_bright);
    tokens.insert(TokenKey::RefHighlight2, rgb(0xffff00));
    tokens.insert(TokenKey::RefHighlight3, rgb(0x00ffff));
    tokens.insert(TokenKey::SpillBorder, green_bright);
    tokens.insert(TokenKey::CommentIndicator, rgb(0xff5555));

    Theme {
        meta: ThemeMeta {
            id: "visicalc",
            name: "VisiCalc",
            author: "VisiGrid",
            appearance: Appearance::Dark,
        },
        tokens,
        typography: ThemeTypography {
            font_family: Some("IBM Plex Mono".to_string()),
            font_size: 12.0,
            mono_family: Some("IBM Plex Mono".to_string()),
        },
    }
}

// ============================================================================
// Theme Registry
// ============================================================================

/// All built-in themes
pub fn builtin_themes() -> Vec<Theme> {
    vec![
        visigrid_theme(),
        classic_theme(),
        visicalc_theme(),
    ]
}

/// Get a theme by ID
pub fn get_theme(id: &str) -> Option<Theme> {
    match id {
        "visigrid" => Some(visigrid_theme()),
        "classic" => Some(classic_theme()),
        "visicalc" => Some(visicalc_theme()),
        _ => None,
    }
}
