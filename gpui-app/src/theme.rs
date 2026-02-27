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
    HeaderActiveText,

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
    SelectionText,    // Text color inside selected cells (for opaque selection themes like VisiCalc)
    SelectionHandle,
    CursorBg,
    CursorText,

    // Formula bar + editor fields
    FormulaBarBg,
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
    SpillBorder,           // Spill parent border (blue)
    SpillReceiverBorder,   // Spill receiver border (light blue)
    SpillBlockedBorder,    // Blocked spill / #SPILL! error (red)
    CommentIndicator,

    // Formula syntax highlighting
    FormulaFunction,
    FormulaCellRef,
    FormulaNumber,
    FormulaString,
    FormulaBoolean,
    FormulaOperator,
    FormulaParens,
    FormulaError,

    // Keyboard hints (Vimium-style)
    HintBadgeBg,           // Background for non-matching hints (muted)
    HintBadgeText,         // Text for non-matching hints
    HintBadgeMatchBg,      // Background for matching hints (brighter)
    HintBadgeMatchText,    // Text for matching hints
    HintBadgeUniqueBg,     // Background for unique match (highlight)
    HintBadgeUniqueText,   // Text for unique match

    // Dependency tracing (Alt+T)
    TracePrecedentBg,      // Background tint for precedent cells (inputs)
    TraceDependentBg,      // Background tint for dependent cells (outputs)
    TraceSourceBorder,     // Border for the source cell being traced

    // User-defined cell borders
    UserBorder,            // Color for user-applied cell borders (themeable)

    // Cell styles (conditional formatting presets)
    CellStyleErrorBg,
    CellStyleErrorText,
    CellStyleErrorBorder,
    CellStyleWarningBg,
    CellStyleWarningText,
    CellStyleWarningBorder,
    CellStyleSuccessBg,
    CellStyleSuccessText,
    CellStyleSuccessBorder,
    CellStyleInputBg,
    CellStyleInputBorder,
    CellStyleTotalBorder,
    CellStyleNoteBg,
    CellStyleNoteText,
}

impl TokenKey {
    /// Get all token keys for validation (used for custom theme validation)
    #[allow(dead_code)]
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
        TokenKey::HeaderActiveText,
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
        TokenKey::SelectionText,
        TokenKey::SelectionHandle,
        TokenKey::CursorBg,
        TokenKey::CursorText,
        // Formula bar + editor fields
        TokenKey::FormulaBarBg,
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
        TokenKey::SpillReceiverBorder,
        TokenKey::SpillBlockedBorder,
        TokenKey::CommentIndicator,
        // Formula syntax highlighting
        TokenKey::FormulaFunction,
        TokenKey::FormulaCellRef,
        TokenKey::FormulaNumber,
        TokenKey::FormulaString,
        TokenKey::FormulaBoolean,
        TokenKey::FormulaOperator,
        TokenKey::FormulaParens,
        TokenKey::FormulaError,
        // Keyboard hints
        TokenKey::HintBadgeBg,
        TokenKey::HintBadgeText,
        TokenKey::HintBadgeMatchBg,
        TokenKey::HintBadgeMatchText,
        TokenKey::HintBadgeUniqueBg,
        TokenKey::HintBadgeUniqueText,
        // Dependency tracing
        TokenKey::TracePrecedentBg,
        TokenKey::TraceDependentBg,
        TokenKey::TraceSourceBorder,
        // User-defined cell borders
        TokenKey::UserBorder,
        // Cell styles
        TokenKey::CellStyleErrorBg,
        TokenKey::CellStyleErrorText,
        TokenKey::CellStyleErrorBorder,
        TokenKey::CellStyleWarningBg,
        TokenKey::CellStyleWarningText,
        TokenKey::CellStyleWarningBorder,
        TokenKey::CellStyleSuccessBg,
        TokenKey::CellStyleSuccessText,
        TokenKey::CellStyleSuccessBorder,
        TokenKey::CellStyleInputBg,
        TokenKey::CellStyleInputBorder,
        TokenKey::CellStyleTotalBorder,
        TokenKey::CellStyleNoteBg,
        TokenKey::CellStyleNoteText,
    ];
}

/// Theme metadata
#[derive(Debug, Clone)]
pub struct ThemeMeta {
    pub id: &'static str,
    pub name: &'static str,
    #[allow(dead_code)]
    pub author: &'static str,
    pub appearance: Appearance,
}

/// Typography settings for a theme (reserved for custom theme support)
#[derive(Debug, Clone)]
#[allow(dead_code)]
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
    #[allow(dead_code)]
    pub typography: ThemeTypography,
}

impl Theme {
    /// Get a token color, panics if not found (should never happen with resolved themes)
    pub fn get(&self, key: TokenKey) -> Hsla {
        *self.tokens.get(&key).unwrap_or_else(|| {
            panic!("Missing theme token: {:?}", key)
        })
    }

    /// Validate theme has all required tokens (used for custom theme validation)
    #[allow(dead_code)]
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

/// Ledger Dark - serious professional dark theme for spreadsheets
/// Neutral charcoal background, minimal accents, optimized for numeric scanning
pub fn ledger_dark_theme() -> Theme {
    let mut tokens = HashMap::new();

    // Core palette - neutral charcoal, NOT blue-biased
    let bg_darkest = rgb(0x0d1016);   // Editor bg (active surface)
    let bg_dark = rgb(0x0f1115);      // App/Grid background
    let bg_panel = rgb(0x141821);     // Panels
    let _bg_header = rgb(0x171c26);    // Headers
    let grid_subtle = rgb(0x1c2330);  // Subtle gridlines
    let grid_bold = rgb(0x273042);    // Bold gridlines / borders

    // Text hierarchy
    let text_primary = rgb(0xd6d9e0);
    let text_muted = rgb(0xa7aebb);
    let text_faint = rgb(0x6e7687);

    // Single accent (blue) - used sparingly
    let accent = rgb(0x4f8cff);
    let accent_bright = rgb(0x6aa0ff);

    // Semantic colors (muted, not neon)
    let ok = rgb(0x3ccb7f);
    let warn = rgb(0xf2b84b);
    let error = rgb(0xe4636a);

    // Reference highlights (3 only, muted)
    let ref1 = rgb(0x6aa0ff);  // Blue
    let ref2 = rgb(0xb48efc);  // Violet
    let ref3 = rgb(0x56c2d6);  // Cyan

    // App surfaces
    tokens.insert(TokenKey::AppBg, bg_dark);
    tokens.insert(TokenKey::PanelBg, bg_panel);
    tokens.insert(TokenKey::PanelBorder, grid_bold);
    tokens.insert(TokenKey::TextPrimary, text_primary);
    tokens.insert(TokenKey::TextMuted, text_muted);
    tokens.insert(TokenKey::TextDisabled, text_faint);
    tokens.insert(TokenKey::TextInverse, bg_dark);

    // Grid surfaces - grid must disappear
    tokens.insert(TokenKey::GridBg, bg_dark);
    tokens.insert(TokenKey::GridLines, rgba(0x1c233040));  // Very subtle (25% opacity)
    tokens.insert(TokenKey::GridLinesBold, grid_bold);

    // Headers - slightly lifted from cells, guides the eye
    let header_bg = rgb(0x1a1f2a);  // Lifted from 0x171c26
    let header_text = rgb(0xe2e5ec);  // Brighter than text_primary
    tokens.insert(TokenKey::HeaderBg, header_bg);
    tokens.insert(TokenKey::HeaderText, header_text);
    tokens.insert(TokenKey::HeaderTextMuted, text_muted);
    tokens.insert(TokenKey::HeaderBorder, grid_bold);
    tokens.insert(TokenKey::HeaderHoverBg, rgb(0x232a38));
    tokens.insert(TokenKey::HeaderActiveBg, accent);
    tokens.insert(TokenKey::HeaderActiveText, text_muted);

    // Cells
    tokens.insert(TokenKey::CellBg, bg_dark);
    tokens.insert(TokenKey::CellBgAlt, rgb(0x121620));
    tokens.insert(TokenKey::CellText, text_primary);
    tokens.insert(TokenKey::CellTextMuted, text_muted);
    tokens.insert(TokenKey::CellBorderFocus, accent);
    tokens.insert(TokenKey::CellHoverBg, bg_panel);

    // Selection + cursor - border-first, minimal fill
    tokens.insert(TokenKey::SelectionBg, rgba(0x4f8cff10));  // ~6% alpha - near transparent
    tokens.insert(TokenKey::SelectionBorder, accent);        // Solid border is the focus
    tokens.insert(TokenKey::SelectionText, text_primary);
    tokens.insert(TokenKey::SelectionHandle, accent_bright);
    tokens.insert(TokenKey::CursorBg, text_primary);
    tokens.insert(TokenKey::CursorText, bg_dark);

    // Formula bar + editor
    tokens.insert(TokenKey::FormulaBarBg, rgb(0x191e2a));  // 1 step lighter than panel
    tokens.insert(TokenKey::EditorBg, bg_darkest);
    tokens.insert(TokenKey::EditorBorder, accent);
    tokens.insert(TokenKey::EditorText, text_primary);
    tokens.insert(TokenKey::EditorPlaceholder, text_faint);
    tokens.insert(TokenKey::EditorSelectionBg, rgba(0x4f8cff40));
    tokens.insert(TokenKey::EditorSelectionText, text_primary);

    // Status + chrome
    tokens.insert(TokenKey::StatusBg, bg_panel);
    tokens.insert(TokenKey::StatusText, text_primary);
    tokens.insert(TokenKey::StatusTextMuted, text_muted);
    tokens.insert(TokenKey::ToolbarBg, bg_panel);
    tokens.insert(TokenKey::ToolbarBorder, grid_bold);
    tokens.insert(TokenKey::ToolbarButtonHoverBg, grid_subtle);
    tokens.insert(TokenKey::ToolbarButtonActiveBg, accent);

    // Semantic feedback
    tokens.insert(TokenKey::Accent, accent);
    tokens.insert(TokenKey::Ok, ok);
    tokens.insert(TokenKey::Warn, warn);
    tokens.insert(TokenKey::Error, error);
    tokens.insert(TokenKey::ErrorBg, rgba(0xe4636a20));
    tokens.insert(TokenKey::Link, accent);

    // Spreadsheet semantics - restrained
    tokens.insert(TokenKey::FormulaText, text_primary);  // Same as text, not colorful
    tokens.insert(TokenKey::RefHighlight1, ref1);
    tokens.insert(TokenKey::RefHighlight2, ref2);
    tokens.insert(TokenKey::RefHighlight3, ref3);
    tokens.insert(TokenKey::SpillBorder, accent);
    tokens.insert(TokenKey::SpillReceiverBorder, accent_bright);
    tokens.insert(TokenKey::SpillBlockedBorder, error);
    tokens.insert(TokenKey::CommentIndicator, error);

    // Formula syntax highlighting - muted, uses semantic colors
    tokens.insert(TokenKey::FormulaFunction, ref2);       // Violet
    tokens.insert(TokenKey::FormulaCellRef, ref1);        // Blue
    tokens.insert(TokenKey::FormulaNumber, warn);         // Amber
    tokens.insert(TokenKey::FormulaString, ok);           // Green
    tokens.insert(TokenKey::FormulaBoolean, ref3);        // Cyan
    tokens.insert(TokenKey::FormulaOperator, text_muted);
    tokens.insert(TokenKey::FormulaParens, text_muted);
    tokens.insert(TokenKey::FormulaError, error);

    // Keyboard hints
    tokens.insert(TokenKey::HintBadgeBg, grid_bold);
    tokens.insert(TokenKey::HintBadgeText, text_muted);
    tokens.insert(TokenKey::HintBadgeMatchBg, warn);
    tokens.insert(TokenKey::HintBadgeMatchText, bg_dark);
    tokens.insert(TokenKey::HintBadgeUniqueBg, ok);
    tokens.insert(TokenKey::HintBadgeUniqueText, bg_dark);

    // Dependency tracing - subtle, doesn't destroy readability
    tokens.insert(TokenKey::TracePrecedentBg, rgba(0x3ccb7f18));  // Green-500 at ~9% - inputs
    tokens.insert(TokenKey::TraceDependentBg, rgba(0xb48efc18));  // Violet at ~9% - outputs
    tokens.insert(TokenKey::TraceSourceBorder, accent);           // Blue accent for source

    // User-defined cell borders
    tokens.insert(TokenKey::UserBorder, rgb(0xd6d9e0));  // Light text color (visible on dark bg)

    // Cell styles
    tokens.insert(TokenKey::CellStyleErrorBg, rgb(0x5C1D18));
    tokens.insert(TokenKey::CellStyleErrorText, rgb(0xFDA29B));
    tokens.insert(TokenKey::CellStyleErrorBorder, rgb(0xF04438));
    tokens.insert(TokenKey::CellStyleWarningBg, rgb(0x5C3D10));
    tokens.insert(TokenKey::CellStyleWarningText, rgb(0xFEC84B));
    tokens.insert(TokenKey::CellStyleWarningBorder, rgb(0xFDB022));
    tokens.insert(TokenKey::CellStyleSuccessBg, rgb(0x1A4D35));
    tokens.insert(TokenKey::CellStyleSuccessText, rgb(0x6CE9A6));
    tokens.insert(TokenKey::CellStyleSuccessBorder, rgb(0x32D583));
    tokens.insert(TokenKey::CellStyleInputBg, rgb(0x2B5A8A));
    tokens.insert(TokenKey::CellStyleInputBorder, rgb(0x53B1FD));
    tokens.insert(TokenKey::CellStyleTotalBorder, rgb(0xEAECF0));
    tokens.insert(TokenKey::CellStyleNoteBg, rgb(0x2D3A47));
    tokens.insert(TokenKey::CellStyleNoteText, rgb(0xD0D5DD));

    Theme {
        meta: ThemeMeta {
            id: "ledger-dark",
            name: "Ledger Dark",
            author: "VisiGrid",
            appearance: Appearance::Dark,
        },
        tokens,
        typography: ThemeTypography::default(),
    }
}

/// Slate Dark - developer-style dark theme with blue tones
/// Based on Tailwind slate palette, editor-inspired aesthetic
pub fn slate_dark_theme() -> Theme {
    let mut tokens = HashMap::new();

    // Slate-blue palette (Tailwind slate)
    let grid_900 = rgb(0x0f172a);
    let grid_800 = rgb(0x1e293b);
    let grid_700 = rgb(0x334155);
    let grid_600 = rgb(0x475569);
    let grid_500 = rgb(0x64748b);
    let grid_400 = rgb(0x94a3b8);
    let grid_200 = rgb(0xe2e8f0);
    let grid_100 = rgb(0xf1f5f9);
    let accent = rgb(0x3b82f6);
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
    tokens.insert(TokenKey::GridLines, rgba(0x33415540));
    tokens.insert(TokenKey::GridLinesBold, grid_700);

    // Headers
    tokens.insert(TokenKey::HeaderBg, grid_800);
    tokens.insert(TokenKey::HeaderText, grid_200);
    tokens.insert(TokenKey::HeaderTextMuted, grid_400);
    tokens.insert(TokenKey::HeaderBorder, grid_700);
    tokens.insert(TokenKey::HeaderHoverBg, grid_700);
    tokens.insert(TokenKey::HeaderActiveBg, accent_hover);
    tokens.insert(TokenKey::HeaderActiveText, grid_400);

    // Cells
    tokens.insert(TokenKey::CellBg, grid_900);
    tokens.insert(TokenKey::CellBgAlt, rgb(0x131c2e));
    tokens.insert(TokenKey::CellText, grid_200);
    tokens.insert(TokenKey::CellTextMuted, grid_400);
    tokens.insert(TokenKey::CellBorderFocus, accent);
    tokens.insert(TokenKey::CellHoverBg, grid_800);

    // Selection + cursor
    tokens.insert(TokenKey::SelectionBg, rgba(0x3b82f625));
    tokens.insert(TokenKey::SelectionBorder, rgba(0x3b82f680));
    tokens.insert(TokenKey::SelectionText, grid_100);
    tokens.insert(TokenKey::SelectionHandle, accent);
    tokens.insert(TokenKey::CursorBg, grid_100);
    tokens.insert(TokenKey::CursorText, grid_900);

    // Formula bar + editor
    tokens.insert(TokenKey::FormulaBarBg, rgb(0x253248));  // 1 step lighter than panel
    tokens.insert(TokenKey::EditorBg, rgb(0x1a2744));
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
    tokens.insert(TokenKey::Ok, rgb(0x22c55e));
    tokens.insert(TokenKey::Warn, rgb(0xeab308));
    tokens.insert(TokenKey::Error, rgb(0xef4444));
    tokens.insert(TokenKey::ErrorBg, rgba(0xef444420));
    tokens.insert(TokenKey::Link, rgb(0x60a5fa));

    // Spreadsheet semantics
    tokens.insert(TokenKey::FormulaText, rgb(0x93c5fd));
    tokens.insert(TokenKey::RefHighlight1, rgb(0x22c55e));
    tokens.insert(TokenKey::RefHighlight2, rgb(0xfbbf24));
    tokens.insert(TokenKey::RefHighlight3, rgb(0xa855f7));
    tokens.insert(TokenKey::SpillBorder, rgb(0x3b82f6));
    tokens.insert(TokenKey::SpillReceiverBorder, rgb(0x93c5fd));
    tokens.insert(TokenKey::SpillBlockedBorder, rgb(0xef4444));
    tokens.insert(TokenKey::CommentIndicator, rgb(0xef4444));

    // Formula syntax highlighting
    tokens.insert(TokenKey::FormulaFunction, rgb(0xc084fc));
    tokens.insert(TokenKey::FormulaCellRef, rgb(0x22c55e));
    tokens.insert(TokenKey::FormulaNumber, rgb(0xfbbf24));
    tokens.insert(TokenKey::FormulaString, rgb(0xf97316));
    tokens.insert(TokenKey::FormulaBoolean, rgb(0x06b6d4));
    tokens.insert(TokenKey::FormulaOperator, rgb(0x94a3b8));
    tokens.insert(TokenKey::FormulaParens, rgb(0x94a3b8));
    tokens.insert(TokenKey::FormulaError, rgb(0xef4444));

    // Keyboard hints
    tokens.insert(TokenKey::HintBadgeBg, rgb(0x52525b));
    tokens.insert(TokenKey::HintBadgeText, rgb(0xa1a1aa));
    tokens.insert(TokenKey::HintBadgeMatchBg, rgb(0xfbbf24));
    tokens.insert(TokenKey::HintBadgeMatchText, rgb(0x18181b));
    tokens.insert(TokenKey::HintBadgeUniqueBg, rgb(0x22c55e));
    tokens.insert(TokenKey::HintBadgeUniqueText, rgb(0x18181b));

    // Dependency tracing
    tokens.insert(TokenKey::TracePrecedentBg, rgba(0x22c55e18));  // Green at ~9%
    tokens.insert(TokenKey::TraceDependentBg, rgba(0xa855f718));  // Purple at ~9%
    tokens.insert(TokenKey::TraceSourceBorder, accent);

    // User-defined cell borders
    tokens.insert(TokenKey::UserBorder, grid_200);  // Light text (visible on dark bg)

    // Cell styles
    tokens.insert(TokenKey::CellStyleErrorBg, rgb(0x5C1D18));
    tokens.insert(TokenKey::CellStyleErrorText, rgb(0xFDA29B));
    tokens.insert(TokenKey::CellStyleErrorBorder, rgb(0xF04438));
    tokens.insert(TokenKey::CellStyleWarningBg, rgb(0x5C3D10));
    tokens.insert(TokenKey::CellStyleWarningText, rgb(0xFEC84B));
    tokens.insert(TokenKey::CellStyleWarningBorder, rgb(0xFDB022));
    tokens.insert(TokenKey::CellStyleSuccessBg, rgb(0x1A4D35));
    tokens.insert(TokenKey::CellStyleSuccessText, rgb(0x6CE9A6));
    tokens.insert(TokenKey::CellStyleSuccessBorder, rgb(0x32D583));
    tokens.insert(TokenKey::CellStyleInputBg, rgb(0x2B5A8A));
    tokens.insert(TokenKey::CellStyleInputBorder, rgb(0x53B1FD));
    tokens.insert(TokenKey::CellStyleTotalBorder, rgb(0xEAECF0));
    tokens.insert(TokenKey::CellStyleNoteBg, rgb(0x2D3A47));
    tokens.insert(TokenKey::CellStyleNoteText, rgb(0xD0D5DD));

    Theme {
        meta: ThemeMeta {
            id: "slate-dark",
            name: "Slate Dark",
            author: "VisiGrid",
            appearance: Appearance::Dark,
        },
        tokens,
        typography: ThemeTypography::default(),
    }
}

/// Ledger Light - cool, flat, high-contrast light theme for precision work
/// "Accountant baseline" - boring on purpose, trustworthy
/// Pairs with Ledger Dark as the professional default
pub fn ledger_light_theme() -> Theme {
    let mut tokens = HashMap::new();

    // Cool white point - NOT warm/creamy
    let bg_cool = rgb(0xf7f8fa);      // Cool gray-white
    let bg_panel = rgb(0xeef0f4);     // Slightly darker cool
    let bg_header = rgb(0xe4e7ec);    // Header background
    let white = rgb(0xffffff);        // Pure white for cells

    // Text - high contrast
    let text_primary = rgb(0x1a1d24);  // Near black
    let text_muted = rgb(0x5c6370);    // Cool gray
    let text_faint = rgb(0x9ca3af);    // Light gray

    // Borders - more visible than Catppuccin (precision)
    let border = rgb(0xd1d5dc);
    let _grid_line = rgb(0xe5e7eb);

    // Conservative blue accent
    let accent = rgb(0x2563eb);        // Blue-600
    let accent_light = rgb(0x3b82f6);  // Blue-500

    // Semantic colors
    let ok = rgb(0x059669);            // Emerald-600
    let warn = rgb(0xd97706);          // Amber-600
    let error = rgb(0xdc2626);         // Red-600

    // App surfaces
    tokens.insert(TokenKey::AppBg, bg_cool);
    tokens.insert(TokenKey::PanelBg, bg_panel);
    tokens.insert(TokenKey::PanelBorder, border);
    tokens.insert(TokenKey::TextPrimary, text_primary);
    tokens.insert(TokenKey::TextMuted, text_muted);
    tokens.insert(TokenKey::TextDisabled, text_faint);
    tokens.insert(TokenKey::TextInverse, white);

    // Grid surfaces - slightly more visible gridlines than Catppuccin
    tokens.insert(TokenKey::GridBg, white);
    tokens.insert(TokenKey::GridLines, rgba(0xd1d5dc60));  // More visible (38% opacity)
    tokens.insert(TokenKey::GridLinesBold, border);

    // Headers
    tokens.insert(TokenKey::HeaderBg, bg_header);
    tokens.insert(TokenKey::HeaderText, text_primary);
    tokens.insert(TokenKey::HeaderTextMuted, text_muted);
    tokens.insert(TokenKey::HeaderBorder, border);
    tokens.insert(TokenKey::HeaderHoverBg, rgb(0xdce0e8));
    tokens.insert(TokenKey::HeaderActiveBg, accent_light);
    tokens.insert(TokenKey::HeaderActiveText, text_muted);

    // Cells
    tokens.insert(TokenKey::CellBg, white);
    tokens.insert(TokenKey::CellBgAlt, bg_cool);
    tokens.insert(TokenKey::CellText, text_primary);
    tokens.insert(TokenKey::CellTextMuted, text_muted);
    tokens.insert(TokenKey::CellBorderFocus, accent);
    tokens.insert(TokenKey::CellHoverBg, rgb(0xf3f4f6));

    // Selection + cursor - conservative blue
    tokens.insert(TokenKey::SelectionBg, rgba(0x2563eb20));  // Blue with low alpha
    tokens.insert(TokenKey::SelectionBorder, rgba(0x2563eb80));
    tokens.insert(TokenKey::SelectionText, text_primary);
    tokens.insert(TokenKey::SelectionHandle, accent);
    tokens.insert(TokenKey::CursorBg, text_primary);
    tokens.insert(TokenKey::CursorText, white);

    // Formula bar + editor
    tokens.insert(TokenKey::FormulaBarBg, rgb(0xf5f6f9));  // Elevated: between panel (#eef0f4) and grid (white)
    tokens.insert(TokenKey::EditorBg, white);
    tokens.insert(TokenKey::EditorBorder, accent);
    tokens.insert(TokenKey::EditorText, text_primary);
    tokens.insert(TokenKey::EditorPlaceholder, text_faint);
    tokens.insert(TokenKey::EditorSelectionBg, rgba(0x2563eb40));
    tokens.insert(TokenKey::EditorSelectionText, text_primary);

    // Status + chrome - neutral, not green like Excel
    tokens.insert(TokenKey::StatusBg, bg_panel);
    tokens.insert(TokenKey::StatusText, text_primary);
    tokens.insert(TokenKey::StatusTextMuted, text_muted);
    tokens.insert(TokenKey::ToolbarBg, bg_panel);
    tokens.insert(TokenKey::ToolbarBorder, border);
    tokens.insert(TokenKey::ToolbarButtonHoverBg, rgb(0xe5e7eb));
    tokens.insert(TokenKey::ToolbarButtonActiveBg, accent_light);

    // Semantic feedback
    tokens.insert(TokenKey::Accent, accent);
    tokens.insert(TokenKey::Ok, ok);
    tokens.insert(TokenKey::Warn, warn);
    tokens.insert(TokenKey::Error, error);
    tokens.insert(TokenKey::ErrorBg, rgba(0xdc262620));
    tokens.insert(TokenKey::Link, accent);

    // Spreadsheet semantics
    tokens.insert(TokenKey::FormulaText, accent);
    tokens.insert(TokenKey::RefHighlight1, rgb(0x2563eb));  // Blue
    tokens.insert(TokenKey::RefHighlight2, rgb(0xdc2626));  // Red
    tokens.insert(TokenKey::RefHighlight3, rgb(0x7c3aed));  // Violet
    tokens.insert(TokenKey::SpillBorder, accent);
    tokens.insert(TokenKey::SpillReceiverBorder, accent_light);
    tokens.insert(TokenKey::SpillBlockedBorder, error);
    tokens.insert(TokenKey::CommentIndicator, error);

    // Formula syntax highlighting
    tokens.insert(TokenKey::FormulaFunction, rgb(0x7c3aed));  // Violet
    tokens.insert(TokenKey::FormulaCellRef, rgb(0x059669));   // Green
    tokens.insert(TokenKey::FormulaNumber, text_primary);     // Black (flat)
    tokens.insert(TokenKey::FormulaString, rgb(0xb45309));    // Amber-700
    tokens.insert(TokenKey::FormulaBoolean, accent);          // Blue
    tokens.insert(TokenKey::FormulaOperator, text_muted);
    tokens.insert(TokenKey::FormulaParens, text_muted);
    tokens.insert(TokenKey::FormulaError, error);

    // Keyboard hints
    tokens.insert(TokenKey::HintBadgeBg, rgb(0xe5e7eb));
    tokens.insert(TokenKey::HintBadgeText, text_muted);
    tokens.insert(TokenKey::HintBadgeMatchBg, rgb(0xfbbf24));
    tokens.insert(TokenKey::HintBadgeMatchText, text_primary);
    tokens.insert(TokenKey::HintBadgeUniqueBg, ok);
    tokens.insert(TokenKey::HintBadgeUniqueText, white);

    // Dependency tracing - light mode needs slightly more saturation
    tokens.insert(TokenKey::TracePrecedentBg, rgba(0x05966920));  // Emerald-600 at ~12%
    tokens.insert(TokenKey::TraceDependentBg, rgba(0x7c3aed20));  // Violet-600 at ~12%
    tokens.insert(TokenKey::TraceSourceBorder, accent);

    // User-defined cell borders
    tokens.insert(TokenKey::UserBorder, rgb(0x000000));  // Black (standard on light bg)

    // Cell styles
    tokens.insert(TokenKey::CellStyleErrorBg, rgb(0xFEE4E2));
    tokens.insert(TokenKey::CellStyleErrorText, rgb(0xB42318));
    tokens.insert(TokenKey::CellStyleErrorBorder, rgb(0xD92D20));
    tokens.insert(TokenKey::CellStyleWarningBg, rgb(0xFEF0C7));
    tokens.insert(TokenKey::CellStyleWarningText, rgb(0xB54708));
    tokens.insert(TokenKey::CellStyleWarningBorder, rgb(0xF79009));
    tokens.insert(TokenKey::CellStyleSuccessBg, rgb(0xECFDF3));
    tokens.insert(TokenKey::CellStyleSuccessText, rgb(0x067647));
    tokens.insert(TokenKey::CellStyleSuccessBorder, rgb(0x12B76A));
    tokens.insert(TokenKey::CellStyleInputBg, rgb(0xEFF8FF));
    tokens.insert(TokenKey::CellStyleInputBorder, rgb(0x2E90FA));
    tokens.insert(TokenKey::CellStyleTotalBorder, rgb(0x101828));
    tokens.insert(TokenKey::CellStyleNoteBg, rgb(0xF9FAFB));
    tokens.insert(TokenKey::CellStyleNoteText, rgb(0x475467));

    Theme {
        meta: ThemeMeta {
            id: "ledger-light",
            name: "Ledger Light",
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
    tokens.insert(TokenKey::HeaderActiveBg, green);
    tokens.insert(TokenKey::HeaderActiveText, black);

    // Cells
    tokens.insert(TokenKey::CellBg, black);
    tokens.insert(TokenKey::CellBgAlt, black);
    tokens.insert(TokenKey::CellText, green);
    tokens.insert(TokenKey::CellTextMuted, green_dim);
    tokens.insert(TokenKey::CellBorderFocus, green);
    tokens.insert(TokenKey::CellHoverBg, rgb(0x001100));

    // Selection + cursor — original VisiCalc: inverse video (solid green block, black text)
    tokens.insert(TokenKey::SelectionBg, rgba(0x00ff66ff));  // Fully opaque green — true inverse video
    tokens.insert(TokenKey::SelectionBorder, green);
    tokens.insert(TokenKey::SelectionText, black);           // Black text on green (inverse video)
    tokens.insert(TokenKey::SelectionHandle, green);
    tokens.insert(TokenKey::CursorBg, green);
    tokens.insert(TokenKey::CursorText, black);

    // Formula bar + editor (slightly lifted from pure black to indicate editing)
    tokens.insert(TokenKey::FormulaBarBg, rgb(0x061206));  // Subtle green tint above black
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
    tokens.insert(TokenKey::SpillBorder, green_bright);           // Green - spill parent
    tokens.insert(TokenKey::SpillReceiverBorder, rgb(0x55aa55));  // Dim green - spill receiver
    tokens.insert(TokenKey::SpillBlockedBorder, rgb(0xff5555));   // Red - #SPILL! error
    tokens.insert(TokenKey::CommentIndicator, rgb(0xff5555));

    // Formula syntax highlighting (retro green monochrome style)
    tokens.insert(TokenKey::FormulaFunction, rgb(0x55ff55));  // Bright green - functions
    tokens.insert(TokenKey::FormulaCellRef, rgb(0x00ff00));   // Green - cell references
    tokens.insert(TokenKey::FormulaNumber, rgb(0x55ff55));    // Bright green - numbers
    tokens.insert(TokenKey::FormulaString, rgb(0x00ff00));    // Green - strings
    tokens.insert(TokenKey::FormulaBoolean, rgb(0x55ff55));   // Bright green - booleans
    tokens.insert(TokenKey::FormulaOperator, rgb(0x00ff00));  // Green - operators
    tokens.insert(TokenKey::FormulaParens, rgb(0x00ff00));    // Green - parentheses
    tokens.insert(TokenKey::FormulaError, rgb(0xff5555));     // Red - errors

    // Keyboard hints (retro style)
    tokens.insert(TokenKey::HintBadgeBg, rgb(0x003300));           // Dark green for non-matches
    tokens.insert(TokenKey::HintBadgeText, rgb(0x55aa55));         // Dim green text
    tokens.insert(TokenKey::HintBadgeMatchBg, rgb(0x00ff00));      // Bright green for matches
    tokens.insert(TokenKey::HintBadgeMatchText, rgb(0x000000));    // Black text on green
    tokens.insert(TokenKey::HintBadgeUniqueBg, rgb(0xffff00));     // Yellow for unique match
    tokens.insert(TokenKey::HintBadgeUniqueText, rgb(0x000000));   // Black text on yellow

    // Dependency tracing (retro green/yellow)
    tokens.insert(TokenKey::TracePrecedentBg, rgba(0x00ff0020));   // Bright green at ~12%
    tokens.insert(TokenKey::TraceDependentBg, rgba(0xffff0020));   // Yellow at ~12%
    tokens.insert(TokenKey::TraceSourceBorder, green);

    // User-defined cell borders
    tokens.insert(TokenKey::UserBorder, green);  // Green (consistent with retro style)

    // Cell styles (dark theme colors — VisiCalc has black bg)
    tokens.insert(TokenKey::CellStyleErrorBg, rgb(0x5C1D18));
    tokens.insert(TokenKey::CellStyleErrorText, rgb(0xFDA29B));
    tokens.insert(TokenKey::CellStyleErrorBorder, rgb(0xF04438));
    tokens.insert(TokenKey::CellStyleWarningBg, rgb(0x5C3D10));
    tokens.insert(TokenKey::CellStyleWarningText, rgb(0xFEC84B));
    tokens.insert(TokenKey::CellStyleWarningBorder, rgb(0xFDB022));
    tokens.insert(TokenKey::CellStyleSuccessBg, rgb(0x1A4D35));
    tokens.insert(TokenKey::CellStyleSuccessText, rgb(0x6CE9A6));
    tokens.insert(TokenKey::CellStyleSuccessBorder, rgb(0x32D583));
    tokens.insert(TokenKey::CellStyleInputBg, rgb(0x2B5A8A));
    tokens.insert(TokenKey::CellStyleInputBorder, rgb(0x53B1FD));
    tokens.insert(TokenKey::CellStyleTotalBorder, rgb(0xEAECF0));
    tokens.insert(TokenKey::CellStyleNoteBg, rgb(0x2D3A47));
    tokens.insert(TokenKey::CellStyleNoteText, rgb(0xD0D5DD));

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

/// Catppuccin Latte theme - warm pastel light theme
pub fn catppuccin_theme() -> Theme {
    let mut tokens = HashMap::new();

    // Catppuccin Latte palette (light variant)
    let base = rgb(0xeff1f5);      // Background
    let mantle = rgb(0xe6e9ef);    // Slightly darker
    let crust = rgb(0xdce0e8);     // Darkest background
    let surface0 = rgb(0xccd0da);  // Surface
    let surface1 = rgb(0xbcc0cc);  // Surface highlight
    let overlay0 = rgb(0x9ca0b0);  // Muted
    let text = rgb(0x4c4f69);      // Primary text
    let subtext0 = rgb(0x6c6f85);  // Secondary text
    let _subtext1 = rgb(0x5c5f77);  // Tertiary text

    // Accent colors (Latte variants - slightly more saturated for light bg)
    let blue = rgb(0x1e66f5);
    let lavender = rgb(0x7287fd);
    let sapphire = rgb(0x209fb5);
    let _teal = rgb(0x179299);
    let green = rgb(0x40a02b);
    let yellow = rgb(0xdf8e1d);
    let peach = rgb(0xfe640b);
    let maroon = rgb(0xe64553);
    let red = rgb(0xd20f39);
    let mauve = rgb(0x8839ef);

    // App surfaces
    tokens.insert(TokenKey::AppBg, base);
    tokens.insert(TokenKey::PanelBg, mantle);
    tokens.insert(TokenKey::PanelBorder, surface0);
    tokens.insert(TokenKey::TextPrimary, text);
    tokens.insert(TokenKey::TextMuted, subtext0);
    tokens.insert(TokenKey::TextDisabled, overlay0);
    tokens.insert(TokenKey::TextInverse, base);

    // Grid surfaces
    tokens.insert(TokenKey::GridBg, rgb(0xffffff));  // Pure white cells
    tokens.insert(TokenKey::GridLines, rgba(0xccd0da40));  // Subtle gridlines
    tokens.insert(TokenKey::GridLinesBold, surface0);

    // Headers
    tokens.insert(TokenKey::HeaderBg, mantle);
    tokens.insert(TokenKey::HeaderText, text);
    tokens.insert(TokenKey::HeaderTextMuted, subtext0);
    tokens.insert(TokenKey::HeaderBorder, surface0);
    tokens.insert(TokenKey::HeaderHoverBg, crust);
    tokens.insert(TokenKey::HeaderActiveBg, surface1);
    tokens.insert(TokenKey::HeaderActiveText, subtext0);

    // Cells
    tokens.insert(TokenKey::CellBg, rgb(0xffffff));
    tokens.insert(TokenKey::CellBgAlt, base);
    tokens.insert(TokenKey::CellText, text);
    tokens.insert(TokenKey::CellTextMuted, subtext0);
    tokens.insert(TokenKey::CellBorderFocus, lavender);
    tokens.insert(TokenKey::CellHoverBg, base);

    // Selection + cursor
    tokens.insert(TokenKey::SelectionBg, rgba(0x7287fd30));  // Lavender with alpha
    tokens.insert(TokenKey::SelectionBorder, rgba(0x7287fd80));
    tokens.insert(TokenKey::SelectionText, text);
    tokens.insert(TokenKey::SelectionHandle, lavender);
    tokens.insert(TokenKey::CursorBg, text);
    tokens.insert(TokenKey::CursorText, base);

    // Formula bar + editor
    tokens.insert(TokenKey::FormulaBarBg, rgb(0xebedf2));  // Elevated: between mantle and base
    tokens.insert(TokenKey::EditorBg, rgb(0xffffff));
    tokens.insert(TokenKey::EditorBorder, lavender);
    tokens.insert(TokenKey::EditorText, text);
    tokens.insert(TokenKey::EditorPlaceholder, overlay0);
    tokens.insert(TokenKey::EditorSelectionBg, rgba(0x7287fd40));
    tokens.insert(TokenKey::EditorSelectionText, text);

    // Status + chrome
    tokens.insert(TokenKey::StatusBg, mantle);
    tokens.insert(TokenKey::StatusText, text);
    tokens.insert(TokenKey::StatusTextMuted, subtext0);
    tokens.insert(TokenKey::ToolbarBg, mantle);
    tokens.insert(TokenKey::ToolbarBorder, surface0);
    tokens.insert(TokenKey::ToolbarButtonHoverBg, crust);
    tokens.insert(TokenKey::ToolbarButtonActiveBg, surface1);

    // Semantic feedback
    tokens.insert(TokenKey::Accent, lavender);
    tokens.insert(TokenKey::Ok, green);
    tokens.insert(TokenKey::Warn, yellow);
    tokens.insert(TokenKey::Error, red);
    tokens.insert(TokenKey::ErrorBg, rgba(0xd20f3920));
    tokens.insert(TokenKey::Link, sapphire);

    // Spreadsheet semantics
    tokens.insert(TokenKey::FormulaText, blue);
    tokens.insert(TokenKey::RefHighlight1, green);
    tokens.insert(TokenKey::RefHighlight2, peach);
    tokens.insert(TokenKey::RefHighlight3, mauve);
    tokens.insert(TokenKey::SpillBorder, sapphire);        // Blue - spill parent
    tokens.insert(TokenKey::SpillReceiverBorder, lavender); // Light blue - spill receiver
    tokens.insert(TokenKey::SpillBlockedBorder, red);       // Red - #SPILL! error
    tokens.insert(TokenKey::CommentIndicator, maroon);

    // Formula syntax highlighting (Catppuccin Latte colors)
    tokens.insert(TokenKey::FormulaFunction, mauve);      // Purple - functions
    tokens.insert(TokenKey::FormulaCellRef, green);       // Green - cell references
    tokens.insert(TokenKey::FormulaNumber, peach);        // Peach - numbers
    tokens.insert(TokenKey::FormulaString, green);        // Green - strings
    tokens.insert(TokenKey::FormulaBoolean, sapphire);    // Sapphire - booleans
    tokens.insert(TokenKey::FormulaOperator, text);       // Text - operators
    tokens.insert(TokenKey::FormulaParens, text);         // Text - parentheses
    tokens.insert(TokenKey::FormulaError, red);           // Red - errors

    // Keyboard hints (Catppuccin style)
    tokens.insert(TokenKey::HintBadgeBg, surface1);             // Surface for non-matches
    tokens.insert(TokenKey::HintBadgeText, subtext0);           // Muted text
    tokens.insert(TokenKey::HintBadgeMatchBg, yellow);          // Yellow for matches
    tokens.insert(TokenKey::HintBadgeMatchText, base);          // Dark text on yellow
    tokens.insert(TokenKey::HintBadgeUniqueBg, green);          // Green for unique match
    tokens.insert(TokenKey::HintBadgeUniqueText, base);         // Dark text on green

    // Dependency tracing (Catppuccin pastels)
    tokens.insert(TokenKey::TracePrecedentBg, rgba(0x40a02b20));  // Green at ~12%
    tokens.insert(TokenKey::TraceDependentBg, rgba(0x8839ef20));  // Mauve at ~12%
    tokens.insert(TokenKey::TraceSourceBorder, lavender);

    // User-defined cell borders (use text color, not pure black — softer on pastel palette)
    tokens.insert(TokenKey::UserBorder, text);

    // Cell styles (light theme colors — Catppuccin Latte is a light variant)
    tokens.insert(TokenKey::CellStyleErrorBg, rgb(0xFEE4E2));
    tokens.insert(TokenKey::CellStyleErrorText, rgb(0xB42318));
    tokens.insert(TokenKey::CellStyleErrorBorder, rgb(0xD92D20));
    tokens.insert(TokenKey::CellStyleWarningBg, rgb(0xFEF0C7));
    tokens.insert(TokenKey::CellStyleWarningText, rgb(0xB54708));
    tokens.insert(TokenKey::CellStyleWarningBorder, rgb(0xF79009));
    tokens.insert(TokenKey::CellStyleSuccessBg, rgb(0xECFDF3));
    tokens.insert(TokenKey::CellStyleSuccessText, rgb(0x067647));
    tokens.insert(TokenKey::CellStyleSuccessBorder, rgb(0x12B76A));
    tokens.insert(TokenKey::CellStyleInputBg, rgb(0xEFF8FF));
    tokens.insert(TokenKey::CellStyleInputBorder, rgb(0x2E90FA));
    tokens.insert(TokenKey::CellStyleTotalBorder, rgb(0x101828));
    tokens.insert(TokenKey::CellStyleNoteBg, rgb(0xF9FAFB));
    tokens.insert(TokenKey::CellStyleNoteText, rgb(0x475467));

    Theme {
        meta: ThemeMeta {
            id: "catppuccin",
            name: "Catppuccin",
            author: "Catppuccin",
            appearance: Appearance::Light,
        },
        tokens,
        typography: ThemeTypography::default(),
    }
}

// ============================================================================
// Theme Registry
// ============================================================================

/// All built-in themes
/// Order: Ledger Light (default), Ledger Dark, Catppuccin, VisiCalc, Slate Dark
pub fn builtin_themes() -> Vec<Theme> {
    vec![
        ledger_light_theme(),   // Default
        ledger_dark_theme(),    // Recommended dark
        catppuccin_theme(),     // Comfort light
        visicalc_theme(),       // Retro fun
        slate_dark_theme(),     // Developer/editor-style dark
    ]
}

/// Get a theme by ID
pub fn get_theme(id: &str) -> Option<Theme> {
    match id {
        "ledger-dark" => Some(ledger_dark_theme()),
        "ledger-light" => Some(ledger_light_theme()),
        "catppuccin" => Some(catppuccin_theme()),
        "visicalc" => Some(visicalc_theme()),
        "slate-dark" => Some(slate_dark_theme()),
        // Legacy aliases for migration
        "visigrid" => Some(ledger_dark_theme()),
        "classic" => Some(ledger_light_theme()),
        "neutral-light" => Some(ledger_light_theme()),
        _ => None,
    }
}

/// Default theme
pub fn default_theme() -> Theme {
    ledger_light_theme()
}

pub const SYSTEM_THEME_ID: &str = "system";

/// Returns the theme ID to use for the given OS appearance.
pub fn resolve_system_theme_id(appearance: gpui::WindowAppearance) -> &'static str {
    match appearance {
        gpui::WindowAppearance::Dark | gpui::WindowAppearance::VibrantDark => "ledger-dark",
        // Light, VibrantLight, and any future additions → light theme
        _ => "ledger-light",
    }
}

/// Placeholder Theme for the picker list. id="system", name="System".
/// Uses Ledger Light tokens so swatches are neutral.
pub fn system_placeholder_theme() -> Theme {
    let mut t = ledger_light_theme();
    t.meta = ThemeMeta {
        id: SYSTEM_THEME_ID,
        name: "System",
        author: "VisiGrid",
        appearance: Appearance::Light,
    };
    t
}
