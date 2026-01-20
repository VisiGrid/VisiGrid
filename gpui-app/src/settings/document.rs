//! Document settings (file-scoped, stored with document)
//!
//! These settings represent "How should this file behave and appear?"
//! They override user settings and are saved with the document.

use serde::{Deserialize, Serialize};

use super::types::{CalculationMode, Setting};

/// Document-level settings (per-file)
///
/// These settings are stored in sidecar files: `myfile.vgrid.settings.json`
/// They override user settings for this specific document.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct DocumentSettings {
    /// Display options for this document
    #[serde(default)]
    pub display: DocumentDisplaySettings,

    /// Calculation options for this document
    #[serde(default)]
    pub calculation: DocumentCalculationSettings,
}

// ============================================================================
// Display settings (document-level)
// ============================================================================

/// Display options that are document-specific
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DocumentDisplaySettings {
    /// Show formulas instead of calculated values
    #[serde(default, skip_serializing_if = "Setting::is_inherit")]
    pub show_formulas: Setting<bool>,

    /// Show zeros in cells (vs blank)
    #[serde(default = "default_show_zeros", skip_serializing_if = "Setting::is_inherit")]
    pub show_zeros: Setting<bool>,

    /// Show row and column headers
    #[serde(default = "default_show_headers", skip_serializing_if = "Setting::is_inherit")]
    pub show_headers: Setting<bool>,

    /// Show gridlines (overrides user preference for this doc)
    #[serde(default, skip_serializing_if = "Setting::is_inherit")]
    pub show_gridlines: Setting<bool>,
}

fn default_show_zeros() -> Setting<bool> {
    Setting::Value(true)
}

fn default_show_headers() -> Setting<bool> {
    Setting::Value(true)
}

impl Default for DocumentDisplaySettings {
    fn default() -> Self {
        Self {
            show_formulas: Setting::Inherit, // Inherit from user (which inherits default: false)
            show_zeros: Setting::Value(true),
            show_headers: Setting::Value(true),
            show_gridlines: Setting::Inherit, // Inherit from user preference
        }
    }
}

// ============================================================================
// Calculation settings (document-level)
// ============================================================================

/// Calculation options that are document-specific
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DocumentCalculationSettings {
    /// When formulas are recalculated
    #[serde(default = "default_calculation_mode", skip_serializing_if = "Setting::is_inherit")]
    pub mode: Setting<CalculationMode>,
}

fn default_calculation_mode() -> Setting<CalculationMode> {
    Setting::Value(CalculationMode::Automatic)
}

impl Default for DocumentCalculationSettings {
    fn default() -> Self {
        Self {
            mode: Setting::Value(CalculationMode::Automatic),
        }
    }
}
