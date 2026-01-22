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
}

/// Collection of dismissed tips
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DismissedTips {
    #[serde(default)]
    pub dismissed: HashSet<TipId>,
}

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
    }
}
