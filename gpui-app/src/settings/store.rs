//! App-level settings store
//!
//! This module provides a centralized, app-wide UserSettings store.
//! All windows share this single source of truth, ensuring consistency
//! when settings change in any window.

use gpui::{App, BorrowAppContext, Global, Subscription};

use super::persistence::{load_user_settings, save_user_settings};
use super::user::UserSettings;

/// App-level settings store implementing GPUI's Global trait.
///
/// There is exactly one instance of this per application process.
/// All windows read from and write to this shared store.
pub struct SettingsStore {
    user_settings: UserSettings,
}

impl Global for SettingsStore {}

impl SettingsStore {
    /// Create a new settings store, loading from disk.
    pub fn new() -> Self {
        Self {
            user_settings: load_user_settings(),
        }
    }

    /// Get a reference to the user settings.
    pub fn user_settings(&self) -> &UserSettings {
        &self.user_settings
    }

    /// Get a mutable reference to the user settings.
    ///
    /// Note: After mutating, call `save()` to persist changes.
    pub fn user_settings_mut(&mut self) -> &mut UserSettings {
        &mut self.user_settings
    }

    /// Save the current settings to disk.
    pub fn save(&self) {
        save_user_settings(&self.user_settings);
    }
}

impl Default for SettingsStore {
    fn default() -> Self {
        Self::new()
    }
}

// ============================================================================
// Convenience functions for accessing the global store
// ============================================================================

/// Initialize the global settings store. Call this once at app startup.
pub fn init_settings_store(cx: &mut App) {
    cx.set_global(SettingsStore::new());
}

/// Get a reference to the user settings from the global store.
///
/// Panics if `init_settings_store` hasn't been called.
pub fn user_settings(cx: &App) -> &UserSettings {
    cx.global::<SettingsStore>().user_settings()
}

/// Update the user settings in the global store.
///
/// The closure receives mutable access to UserSettings.
/// Changes are automatically saved to disk and observers are notified.
pub fn update_user_settings<F, R>(cx: &mut App, f: F) -> R
where
    F: FnOnce(&mut UserSettings) -> R,
{
    let result = cx.update_global::<SettingsStore, _>(|store, _cx| {
        let result = f(store.user_settings_mut());
        store.save();
        result
    });
    result
}

/// Subscribe to settings changes.
///
/// The callback is invoked whenever any part of the global settings changes.
/// Returns a Subscription that must be held to keep the observer active.
pub fn observe_settings<F>(cx: &mut App, f: F) -> Subscription
where
    F: FnMut(&mut App) + 'static,
{
    cx.observe_global::<SettingsStore>(f)
}
