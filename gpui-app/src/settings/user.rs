//! User settings (global, personal, persistent)
//!
//! These settings represent "How do I want spreadsheets to behave?"
//! They apply to every document and persist across sessions.

use serde::{Deserialize, Serialize};

use super::types::{DismissedTips, EnterBehavior, Setting};

/// User-level settings (global, persistent)
///
/// This is the top-level type for user preferences.
/// Stored in `~/.config/visigrid/settings.json`
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct UserSettings {
    /// Visual appearance preferences
    #[serde(default)]
    pub appearance: AppearanceSettings,

    /// Cell editing behavior
    #[serde(default)]
    pub editing: EditingSettings,

    /// Keyboard navigation preferences
    #[serde(default)]
    pub navigation: NavigationSettings,

    /// Tips and onboarding state
    #[serde(default)]
    pub tips: DismissedTips,
}

// ============================================================================
// Appearance settings
// ============================================================================

/// Visual appearance preferences
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AppearanceSettings {
    /// Color theme ID
    #[serde(default, skip_serializing_if = "Setting::is_inherit")]
    pub theme_id: Setting<String>,

    /// Show gridlines by default
    #[serde(default = "default_show_gridlines", skip_serializing_if = "Setting::is_inherit")]
    pub show_gridlines: Setting<bool>,
}

fn default_show_gridlines() -> Setting<bool> {
    Setting::Value(true)
}

impl Default for AppearanceSettings {
    fn default() -> Self {
        Self {
            theme_id: Setting::Inherit, // Use app default theme
            show_gridlines: Setting::Value(true),
        }
    }
}

// ============================================================================
// Editing settings
// ============================================================================

/// Cell editing behavior preferences
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EditingSettings {
    /// What happens after pressing Enter
    #[serde(default = "default_enter_behavior", skip_serializing_if = "Setting::is_inherit")]
    pub enter_behavior: Setting<EnterBehavior>,

    /// Allow typing directly to overwrite cell (vs requiring edit mode)
    #[serde(default = "default_allow_direct_edit", skip_serializing_if = "Setting::is_inherit")]
    pub allow_direct_edit: Setting<bool>,
}

fn default_enter_behavior() -> Setting<EnterBehavior> {
    Setting::Value(EnterBehavior::MoveDown)
}

fn default_allow_direct_edit() -> Setting<bool> {
    Setting::Value(true)
}

impl Default for EditingSettings {
    fn default() -> Self {
        Self {
            enter_behavior: Setting::Value(EnterBehavior::MoveDown),
            allow_direct_edit: Setting::Value(true),
        }
    }
}

// ============================================================================
// Navigation settings
// ============================================================================

/// Keyboard navigation preferences
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NavigationSettings {
    /// Tab key moves right and commits (standard spreadsheet behavior)
    /// This is "locked" but shown in UI for documentation
    #[serde(default = "default_tab_moves_right", skip_serializing_if = "Setting::is_inherit")]
    pub tab_moves_right: Setting<bool>,
}

fn default_tab_moves_right() -> Setting<bool> {
    Setting::Value(true)
}

impl Default for NavigationSettings {
    fn default() -> Self {
        Self {
            tab_moves_right: Setting::Value(true),
        }
    }
}
