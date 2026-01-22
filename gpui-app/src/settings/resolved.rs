//! Resolved settings (runtime truth)
//!
//! This module provides the merged view of settings that the app actually uses.
//! ResolvedSettings = DocumentSettings + UserSettings + AppDefaults
//!
//! Merge rule: Document.Value > User.Value > Default

use super::document::DocumentSettings;
use super::types::{CalculationMode, EnterBehavior, Setting, TipId};
use super::user::UserSettings;

/// The resolved settings that the grid actually uses.
///
/// All fields are concrete values (no Setting<T>) because inheritance
/// has been resolved. This is what you query at runtime.
#[derive(Debug, Clone)]
pub struct ResolvedSettings {
    pub appearance: ResolvedAppearance,
    pub editing: ResolvedEditing,
    pub navigation: ResolvedNavigation,
    pub display: ResolvedDisplay,
    pub calculation: ResolvedCalculation,
}

/// Resolved appearance settings
#[derive(Debug, Clone)]
pub struct ResolvedAppearance {
    /// Theme ID (or None for app default)
    pub theme_id: Option<String>,
    /// Show gridlines
    pub show_gridlines: bool,
}

/// Resolved editing settings
#[derive(Debug, Clone)]
pub struct ResolvedEditing {
    /// What happens after Enter
    pub enter_behavior: EnterBehavior,
    /// Allow direct cell editing by typing
    pub allow_direct_edit: bool,
}

/// Resolved navigation settings
#[derive(Debug, Clone)]
pub struct ResolvedNavigation {
    /// Tab moves right
    pub tab_moves_right: bool,
}

/// Resolved display settings (from document)
#[derive(Debug, Clone)]
pub struct ResolvedDisplay {
    /// Show formulas instead of values
    pub show_formulas: bool,
    /// Show zeros
    pub show_zeros: bool,
    /// Show row/column headers
    pub show_headers: bool,
}

/// Resolved calculation settings (from document)
#[derive(Debug, Clone)]
pub struct ResolvedCalculation {
    /// Calculation mode
    pub mode: CalculationMode,
}

// ============================================================================
// App defaults (hard-coded, not in JSON)
// ============================================================================

/// Hard-coded application defaults
///
/// These are the ultimate fallback values when both user and document
/// settings inherit. Defined in code, not configuration.
pub struct AppDefaults;

impl AppDefaults {
    pub const THEME_ID: Option<&'static str> = None; // Use built-in default
    pub const SHOW_GRIDLINES: bool = true;
    pub const ENTER_BEHAVIOR: EnterBehavior = EnterBehavior::MoveDown;
    pub const ALLOW_DIRECT_EDIT: bool = true;
    pub const TAB_MOVES_RIGHT: bool = true;
    pub const SHOW_FORMULAS: bool = false;
    pub const SHOW_ZEROS: bool = true;
    pub const SHOW_HEADERS: bool = true;
    pub const CALCULATION_MODE: CalculationMode = CalculationMode::Automatic;
}

// ============================================================================
// Merge logic
// ============================================================================

impl ResolvedSettings {
    /// Create resolved settings from user settings only (no document)
    pub fn from_user(user: &UserSettings) -> Self {
        Self::merge(user, &DocumentSettings::default())
    }

    /// Create resolved settings by merging document settings over user settings
    ///
    /// Merge rule: Document.Value > User.Value > AppDefault
    pub fn merge(user: &UserSettings, doc: &DocumentSettings) -> Self {
        Self {
            appearance: ResolvedAppearance {
                theme_id: resolve_option(&user.appearance.theme_id),
                show_gridlines: resolve_layered(
                    &doc.display.show_gridlines,
                    &user.appearance.show_gridlines,
                    AppDefaults::SHOW_GRIDLINES,
                ),
            },
            editing: ResolvedEditing {
                enter_behavior: resolve_value(
                    &user.editing.enter_behavior,
                    AppDefaults::ENTER_BEHAVIOR,
                ),
                allow_direct_edit: resolve_value(
                    &user.editing.allow_direct_edit,
                    AppDefaults::ALLOW_DIRECT_EDIT,
                ),
            },
            navigation: ResolvedNavigation {
                tab_moves_right: resolve_value(
                    &user.navigation.tab_moves_right,
                    AppDefaults::TAB_MOVES_RIGHT,
                ),
            },
            display: ResolvedDisplay {
                show_formulas: resolve_layered(
                    &doc.display.show_formulas,
                    &Setting::Value(AppDefaults::SHOW_FORMULAS), // No user setting for this
                    AppDefaults::SHOW_FORMULAS,
                ),
                show_zeros: resolve_value(&doc.display.show_zeros, AppDefaults::SHOW_ZEROS),
                show_headers: resolve_value(&doc.display.show_headers, AppDefaults::SHOW_HEADERS),
            },
            calculation: ResolvedCalculation {
                mode: resolve_value(&doc.calculation.mode, AppDefaults::CALCULATION_MODE),
            },
        }
    }
}

/// Resolve a Setting<T> against a default value
fn resolve_value<T: Clone>(setting: &Setting<T>, default: T) -> T {
    match setting {
        Setting::Value(v) => v.clone(),
        Setting::Inherit => default,
    }
}

/// Resolve Setting<String> to Option<String> (for theme_id)
fn resolve_option(setting: &Setting<String>) -> Option<String> {
    match setting {
        Setting::Value(v) => Some(v.clone()),
        Setting::Inherit => None,
    }
}

/// Resolve with two layers: doc > user > default
fn resolve_layered<T: Clone>(doc: &Setting<T>, user: &Setting<T>, default: T) -> T {
    match doc {
        Setting::Value(v) => v.clone(),
        Setting::Inherit => match user {
            Setting::Value(v) => v.clone(),
            Setting::Inherit => default,
        },
    }
}

// ============================================================================
// Tip checking (convenience methods)
// ============================================================================

impl UserSettings {
    /// Check if a tip has been dismissed
    pub fn is_tip_dismissed(&self, tip: TipId) -> bool {
        self.tips.is_dismissed(tip)
    }

    /// Dismiss a tip
    pub fn dismiss_tip(&mut self, tip: TipId) {
        self.tips.dismiss(tip);
    }

    /// Reset all tips
    pub fn reset_all_tips(&mut self) {
        self.tips.reset_all();
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::settings::{
        DocumentSettings, UserSettings, Setting, EnterBehavior, TipId,
    };

    /// Test 1: Document settings override user settings
    #[test]
    fn document_overrides_user() {
        let mut user = UserSettings::default();
        user.editing.enter_behavior = Setting::Value(EnterBehavior::MoveRight);

        let mut doc = DocumentSettings::default();
        // Document doesn't have enter_behavior, but show_gridlines demonstrates the pattern
        doc.display.show_gridlines = Setting::Value(false);

        // User wants gridlines on
        user.appearance.show_gridlines = Setting::Value(true);

        let resolved = ResolvedSettings::merge(&user, &doc);

        // Document's false should override user's true
        assert!(!resolved.appearance.show_gridlines);
    }

    /// Test 2: User settings override app defaults
    #[test]
    fn user_overrides_default() {
        let mut user = UserSettings::default();
        // User explicitly sets show_gridlines to false
        user.appearance.show_gridlines = Setting::Value(false);

        // Document inherits (no opinion)
        let doc = DocumentSettings::default();

        let resolved = ResolvedSettings::merge(&user, &doc);

        // Default is true, but user set false
        assert!(!resolved.appearance.show_gridlines);

        // Also verify enter_behavior override
        user.editing.enter_behavior = Setting::Value(EnterBehavior::Stay);
        let resolved = ResolvedSettings::merge(&user, &doc);
        assert_eq!(resolved.editing.enter_behavior, EnterBehavior::Stay);
    }

    /// Test 3: Inherit falls back to app defaults
    #[test]
    fn inherit_falls_back_to_default() {
        // User inherits everything
        let user = UserSettings {
            appearance: crate::settings::user::AppearanceSettings {
                theme_id: Setting::Inherit,
                show_gridlines: Setting::Inherit,
            },
            editing: crate::settings::user::EditingSettings {
                enter_behavior: Setting::Inherit,
                allow_direct_edit: Setting::Inherit,
            },
            navigation: crate::settings::user::NavigationSettings {
                tab_moves_right: Setting::Inherit,
                keyboard_hints: Setting::Inherit,
                vim_mode: Setting::Inherit,
                modifier_style: Setting::Inherit,
            },
            tips: Default::default(),
        };

        // Document also inherits
        let doc = DocumentSettings {
            display: crate::settings::document::DocumentDisplaySettings {
                show_formulas: Setting::Inherit,
                show_zeros: Setting::Inherit,
                show_headers: Setting::Inherit,
                show_gridlines: Setting::Inherit,
            },
            calculation: Default::default(),
        };

        let resolved = ResolvedSettings::merge(&user, &doc);

        // Should use AppDefaults
        assert_eq!(resolved.appearance.show_gridlines, AppDefaults::SHOW_GRIDLINES);
        assert_eq!(resolved.editing.enter_behavior, AppDefaults::ENTER_BEHAVIOR);
        assert_eq!(resolved.editing.allow_direct_edit, AppDefaults::ALLOW_DIRECT_EDIT);
        assert_eq!(resolved.navigation.tab_moves_right, AppDefaults::TAB_MOVES_RIGHT);
        assert_eq!(resolved.display.show_formulas, AppDefaults::SHOW_FORMULAS);
        assert_eq!(resolved.display.show_zeros, AppDefaults::SHOW_ZEROS);
        assert_eq!(resolved.display.show_headers, AppDefaults::SHOW_HEADERS);
    }

    /// Test 4: Serialization roundtrip preserves settings
    #[test]
    fn serialization_roundtrip() {
        let mut settings = UserSettings::default();
        settings.appearance.theme_id = Setting::Value("catppuccin-mocha".to_string());
        settings.editing.enter_behavior = Setting::Value(EnterBehavior::MoveRight);
        settings.appearance.show_gridlines = Setting::Value(false);
        settings.dismiss_tip(TipId::F2Edit);
        settings.dismiss_tip(TipId::RenameF12);

        // Serialize
        let json = serde_json::to_string_pretty(&settings).unwrap();

        // Deserialize
        let loaded: UserSettings = serde_json::from_str(&json).unwrap();

        // Verify equality
        assert!(matches!(
            loaded.appearance.theme_id,
            Setting::Value(ref s) if s == "catppuccin-mocha"
        ));
        assert!(matches!(
            loaded.editing.enter_behavior,
            Setting::Value(EnterBehavior::MoveRight)
        ));
        assert!(matches!(
            loaded.appearance.show_gridlines,
            Setting::Value(false)
        ));
        assert!(loaded.is_tip_dismissed(TipId::F2Edit));
        assert!(loaded.is_tip_dismissed(TipId::RenameF12));
        assert!(!loaded.is_tip_dismissed(TipId::NamedRanges));
    }

    /// Test 5: Unknown JSON fields are ignored (forward compatibility)
    #[test]
    fn unknown_fields_ignored() {
        // JSON with extra fields that don't exist in our struct
        let json = r#"{
            "appearance": {
                "theme_id": "dark-theme",
                "show_gridlines": true,
                "future_field": "ignored",
                "another_unknown": 42
            },
            "editing": {
                "enter_behavior": "move_right"
            },
            "unknown_section": {
                "nested": "data"
            }
        }"#;

        // Should parse without error
        let settings: UserSettings = serde_json::from_str(json).unwrap();

        // Known fields should work
        assert!(matches!(
            settings.appearance.theme_id,
            Setting::Value(ref s) if s == "dark-theme"
        ));
        assert!(matches!(
            settings.editing.enter_behavior,
            Setting::Value(EnterBehavior::MoveRight)
        ));
    }

    /// Bonus: Test TipId dismiss/reset cycle
    #[test]
    fn tip_dismiss_and_reset() {
        let mut settings = UserSettings::default();

        // Initially not dismissed
        assert!(!settings.is_tip_dismissed(TipId::F2Edit));
        assert!(!settings.is_tip_dismissed(TipId::NamedRanges));

        // Dismiss one
        settings.dismiss_tip(TipId::F2Edit);
        assert!(settings.is_tip_dismissed(TipId::F2Edit));
        assert!(!settings.is_tip_dismissed(TipId::NamedRanges));

        // Dismiss another
        settings.dismiss_tip(TipId::NamedRanges);
        assert!(settings.is_tip_dismissed(TipId::F2Edit));
        assert!(settings.is_tip_dismissed(TipId::NamedRanges));

        // Reset all
        settings.reset_all_tips();
        assert!(!settings.is_tip_dismissed(TipId::F2Edit));
        assert!(!settings.is_tip_dismissed(TipId::NamedRanges));
    }

    /// Test: Inherit values are omitted from serialization (clean JSON)
    #[test]
    fn inherit_values_omitted_from_json() {
        // Create settings with only one field explicitly set
        let mut settings = UserSettings::default();
        settings.appearance.theme_id = Setting::Value("my-theme".to_string());
        // Leave all other fields as Inherit
        settings.appearance.show_gridlines = Setting::Inherit;
        settings.editing.enter_behavior = Setting::Inherit;
        settings.editing.allow_direct_edit = Setting::Inherit;
        settings.navigation.tab_moves_right = Setting::Inherit;

        let json = serde_json::to_string(&settings).unwrap();

        // The JSON should contain theme_id but NOT the Inherit fields
        assert!(json.contains("theme_id"));
        assert!(json.contains("my-theme"));
        // Inherit fields should be omitted, not serialized as null
        assert!(!json.contains("show_gridlines"));
        assert!(!json.contains("enter_behavior"));
        assert!(!json.contains("null"));
    }
}
