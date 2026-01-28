//! Core settings types
//!
//! The `Setting<T>` enum is the foundation of the settings system.
//! It provides explicit three-state semantics without Option<Option<T>> confusion.

use serde::{Deserialize, Serialize};
use std::collections::HashSet;

/// A setting value with explicit inherit semantics.
///
/// This is the single most important type in the settings system.
/// It distinguishes between:
/// - `Inherit`: Use parent scope's value (user inherits from default, doc inherits from user)
/// - `Value(T)`: Explicitly set to this value
///
/// # Serialization
/// - `Inherit` is represented by the field being absent in JSON
/// - `Value(T)` is represented by the value being present
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Setting<T> {
    /// Use the parent scope's value
    Inherit,
    /// Explicitly set to this value
    Value(T),
}

impl<T> Setting<T> {
    /// Returns true if this setting is explicitly set
    pub fn is_set(&self) -> bool {
        matches!(self, Setting::Value(_))
    }

    /// Returns true if this setting inherits from parent
    /// (Used by serde skip_serializing_if)
    pub fn is_inherit(&self) -> bool {
        matches!(self, Setting::Inherit)
    }

    /// Returns the value if set, or None if inheriting
    pub fn as_value(&self) -> Option<&T> {
        match self {
            Setting::Value(v) => Some(v),
            Setting::Inherit => None,
        }
    }

    /// Resolves this setting against a fallback value
    pub fn resolve(&self, fallback: T) -> T
    where
        T: Clone,
    {
        match self {
            Setting::Value(v) => v.clone(),
            Setting::Inherit => fallback,
        }
    }

    /// Resolves this setting against another Setting (for layered inheritance)
    pub fn resolve_with(&self, parent: &Setting<T>) -> Setting<T>
    where
        T: Clone,
    {
        match self {
            Setting::Value(v) => Setting::Value(v.clone()),
            Setting::Inherit => parent.clone(),
        }
    }
}

impl<T: Default> Default for Setting<T> {
    fn default() -> Self {
        Setting::Inherit
    }
}

// Custom serialization: Inherit = absent, Value = present
impl<T: Serialize> Serialize for Setting<T> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        match self {
            Setting::Inherit => serializer.serialize_none(),
            Setting::Value(v) => v.serialize(serializer),
        }
    }
}

impl<'de, T: Deserialize<'de>> Deserialize<'de> for Setting<T> {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        // Use Option to handle null values: null -> Inherit, value -> Value(T)
        let opt = Option::<T>::deserialize(deserializer)?;
        Ok(match opt {
            None => Setting::Inherit,
            Some(v) => Setting::Value(v),
        })
    }
}

// ============================================================================
// Keyboard modifier types
// ============================================================================

/// Keyboard modifier style preference (primarily for macOS users)
///
/// On macOS, users can choose between platform-native shortcuts (Cmd+C, Cmd+V)
/// or Windows-style shortcuts (Ctrl+C, Ctrl+V) for familiarity.
/// On Windows/Linux, this setting has no effect (Ctrl is always used).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum ModifierStyle {
    /// Use platform-native modifier (Cmd on macOS, Ctrl on Windows/Linux)
    #[default]
    Platform,
    /// Always use Ctrl (for users who prefer Windows-style shortcuts on Mac)
    Ctrl,
}

/// Excel-style Alt menu accelerators (macOS only)
///
/// When enabled, Alt+F opens File commands, Alt+E opens Edit commands, etc.
/// These open the Command Palette pre-scoped to the menu namespace.
/// On Windows/Linux, this setting has no effect (native Alt menus exist).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum AltAccelerators {
    /// Alt accelerators disabled (default, preserves Option key for symbols)
    #[default]
    Disabled,
    /// Alt accelerators enabled (Alt+F, Alt+E, etc. open scoped palette)
    Enabled,
}

// ============================================================================
// Editing behavior types
// ============================================================================

/// What happens after pressing Enter in a cell
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "snake_case")]
pub enum EnterBehavior {
    /// Move selection down (Excel default)
    #[default]
    MoveDown,
    /// Move selection right (data entry style)
    MoveRight,
    /// Stay in current cell
    Stay,
}

// ============================================================================
// Calculation types
// ============================================================================

/// When formulas are recalculated
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "snake_case")]
pub enum CalculationMode {
    /// Recalculate automatically on any change
    #[default]
    Automatic,
    /// Only recalculate when explicitly requested
    Manual,
}

// ============================================================================
// Tips / onboarding types
// ============================================================================

/// Identifiers for dismissable tips
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum TipId {
    /// F2 edit key tip (macOS function keys)
    F2Edit,
    /// Named ranges tip (first range selection)
    NamedRanges,
    /// F12 rename symbol hint
    RenameF12,
    /// "Set as default app" prompt in title bar
    DefaultAppPrompt,
    /// Window switcher shortcut tip (shown when 2nd window opens)
    WindowSwitcher,
    /// Fill handle usage tip (shown on first drag)
    FillHandle,
}

/// Collection of dismissed tips
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DismissedTips {
    #[serde(default)]
    pub dismissed: HashSet<TipId>,

    /// Per-extension default app prompt state.
    /// Key is extension category: "csv", "xlsx", "tsv"
    #[serde(default, skip_serializing_if = "std::collections::HashMap::is_empty")]
    pub default_app_prompts: std::collections::HashMap<String, DefaultAppPromptExtState>,
}

/// Per-extension state for the default app prompt.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DefaultAppPromptExtState {
    /// User explicitly dismissed via ✕ button (permanent)
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub dismissed: bool,

    /// Timestamp when prompt was last shown (Unix epoch seconds).
    /// Used for cool-down when user ignores or clicks "Open Settings" but doesn't complete.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shown_at: Option<u64>,

    /// User successfully set VisiGrid as default (no need to prompt again)
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub is_default: bool,
}

/// Cool-down period for default app prompt (7 days in seconds)
pub const DEFAULT_APP_PROMPT_COOLDOWN_SECS: u64 = 7 * 24 * 60 * 60;

impl DismissedTips {
    /// Check if a tip has been dismissed
    pub fn is_dismissed(&self, tip: TipId) -> bool {
        self.dismissed.contains(&tip)
    }

    /// Dismiss a tip
    pub fn dismiss(&mut self, tip: TipId) {
        self.dismissed.insert(tip);
    }

    /// Reset all tips (show them again)
    pub fn reset_all(&mut self) {
        self.dismissed.clear();
        self.default_app_prompts.clear();
    }

    /// Get the state for a specific extension's default app prompt.
    pub fn get_default_app_prompt(&self, ext: &str) -> Option<&DefaultAppPromptExtState> {
        self.default_app_prompts.get(ext)
    }

    /// Get mutable state for a specific extension's default app prompt.
    pub fn get_default_app_prompt_mut(&mut self, ext: &str) -> &mut DefaultAppPromptExtState {
        self.default_app_prompts.entry(ext.to_string()).or_default()
    }

    /// Check if default app prompt for an extension is permanently dismissed (✕ clicked).
    pub fn is_default_app_prompt_dismissed(&self, ext: &str) -> bool {
        self.default_app_prompts.get(ext).is_some_and(|s| s.dismissed)
    }

    /// Check if default app prompt for an extension is in cool-down period.
    pub fn is_default_app_prompt_in_cooldown(&self, ext: &str) -> bool {
        if let Some(state) = self.default_app_prompts.get(ext) {
            if let Some(shown_at) = state.shown_at {
                let now = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_secs())
                    .unwrap_or(0);
                return now < shown_at + DEFAULT_APP_PROMPT_COOLDOWN_SECS;
            }
        }
        false
    }

    /// Check if we've recorded that VisiGrid is already default for this extension.
    pub fn is_default_app_prompt_completed(&self, ext: &str) -> bool {
        self.default_app_prompts.get(ext).is_some_and(|s| s.is_default)
    }

    /// Record that default app prompt was shown for an extension (for cool-down tracking).
    pub fn mark_default_app_prompt_shown(&mut self, ext: &str) {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_secs())
            .unwrap_or(0);
        self.get_default_app_prompt_mut(ext).shown_at = Some(now);
    }

    /// Permanently dismiss the default app prompt for an extension (user clicked ✕).
    pub fn dismiss_default_app_prompt(&mut self, ext: &str) {
        self.get_default_app_prompt_mut(ext).dismissed = true;
    }

    /// Mark that VisiGrid is now the default for this extension (success).
    pub fn mark_default_app_completed(&mut self, ext: &str) {
        let state = self.get_default_app_prompt_mut(ext);
        state.is_default = true;
        state.dismissed = true;  // Also dismiss so it doesn't come back
    }
}
