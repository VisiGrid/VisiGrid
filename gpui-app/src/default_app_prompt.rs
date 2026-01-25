//! Default App Prompt (macOS title bar chip)
//!
//! Contains Spreadsheet methods for the "Set as default app" prompt:
//! - Prompt visibility logic
//! - Setting VisiGrid as default handler
//! - System Settings integration
//! - Timer-based state transitions
//!
//! Core default app detection/registration is in default_app.rs.

use gpui::*;
use crate::app::{Spreadsheet, DefaultAppPromptState};
use crate::mode::Mode;
use crate::settings::{user_settings, update_user_settings};

impl Spreadsheet {
    // =========================================================================
    // Default App Prompt (macOS title bar chip)
    // =========================================================================

    /// Get the extension key for the current file (for per-extension state).
    fn get_prompt_ext_key(&self) -> Option<String> {
        use crate::default_app::SpreadsheetFileType;

        // Use cached file type if available
        if let Some(ft) = self.default_app_prompt_file_type {
            return Some(match ft {
                SpreadsheetFileType::Excel => "xlsx",
                SpreadsheetFileType::Csv => "csv",
                SpreadsheetFileType::Tsv => "tsv",
                SpreadsheetFileType::Native => "vgrid",
            }.to_string());
        }

        // Derive from current file
        let path = self.document_meta.path.as_ref()?;
        let ext = path.extension().and_then(|e| e.to_str())?;
        let file_type = SpreadsheetFileType::from_ext(ext)?;
        Some(match file_type {
            SpreadsheetFileType::Excel => "xlsx",
            SpreadsheetFileType::Csv => "csv",
            SpreadsheetFileType::Tsv => "tsv",
            SpreadsheetFileType::Native => "vgrid",
        }.to_string())
    }

    /// Check if the default app prompt should be shown.
    ///
    /// Returns true when ALL conditions are met:
    /// - macOS only
    /// - File successfully loaded (has path, no import errors showing)
    /// - Not a temporary file
    /// - File type is CSV/TSV/Excel (not native .vgrid)
    /// - User hasn't dismissed the prompt for THIS extension
    /// - Not in cool-down period for THIS extension (7 days after ignoring)
    /// - Not already shown this session
    /// - We haven't already marked this extension as completed
    /// - VisiGrid isn't already the default for this file type
    #[cfg(target_os = "macos")]
    pub fn should_show_default_app_prompt(&self, cx: &gpui::App) -> bool {
        use crate::default_app::{SpreadsheetFileType, is_default_handler, is_temporary_file, shown_this_session};

        // If we're showing success/needs-settings feedback, show that instead
        if self.default_app_prompt_state == DefaultAppPromptState::Success
            || self.default_app_prompt_state == DefaultAppPromptState::NeedsSettings
        {
            return true;
        }

        // Must have a file open
        let path = match &self.document_meta.path {
            Some(p) => p,
            None => return false,
        };

        // Skip if import report dialog is showing (don't prompt during error review)
        if self.mode == Mode::ImportReport {
            return false;
        }

        // Skip unsaved files (new documents)
        if !self.document_meta.is_saved && self.document_meta.source.is_none() {
            return false;
        }

        // Skip temporary files
        if is_temporary_file(path) {
            return false;
        }

        // Get file type from extension
        let ext = path.extension()
            .and_then(|e| e.to_str())
            .unwrap_or("");
        let file_type = match SpreadsheetFileType::from_ext(ext) {
            Some(ft) => ft,
            None => return false,
        };

        // Skip native VisiGrid files
        if file_type == SpreadsheetFileType::Native {
            return false;
        }

        // Get extension key for per-extension state
        let ext_key = match file_type {
            SpreadsheetFileType::Excel => "xlsx",
            SpreadsheetFileType::Csv => "csv",
            SpreadsheetFileType::Tsv => "tsv",
            SpreadsheetFileType::Native => return false,
        };

        let settings = user_settings(cx);

        // Check if user has permanently dismissed for THIS extension
        if settings.is_default_app_prompt_dismissed(ext_key) {
            return false;
        }

        // Check if we've already completed setup for this extension
        if settings.is_default_app_prompt_completed(ext_key) {
            return false;
        }

        // Check cool-down period for THIS extension (7 days after ignoring)
        if settings.is_default_app_prompt_in_cooldown(ext_key) {
            return false;
        }

        // Don't spam within same session
        if shown_this_session() {
            return false;
        }

        // Check if VisiGrid is already the default (do last - can be slow)
        if is_default_handler(file_type) {
            return false;
        }

        true
    }

    #[cfg(not(target_os = "macos"))]
    pub fn should_show_default_app_prompt(&self, _cx: &gpui::App) -> bool {
        false
    }

    /// Get the file type for the current prompt (for display).
    #[cfg(target_os = "macos")]
    pub fn get_prompt_file_type(&self) -> Option<crate::default_app::SpreadsheetFileType> {
        use crate::default_app::SpreadsheetFileType;

        // If we have a cached file type from when we showed the prompt, use that
        if let Some(ft) = self.default_app_prompt_file_type {
            return Some(ft);
        }

        // Otherwise derive from current file
        let path = self.document_meta.path.as_ref()?;
        let ext = path.extension().and_then(|e| e.to_str())?;
        SpreadsheetFileType::from_ext(ext)
    }

    #[cfg(not(target_os = "macos"))]
    pub fn get_prompt_file_type(&self) -> Option<crate::default_app::SpreadsheetFileType> {
        None
    }

    /// Called when the prompt becomes visible - marks session and records timestamp.
    pub fn on_default_app_prompt_shown(&mut self, cx: &mut Context<Self>) {
        use crate::default_app::{mark_shown_this_session, SpreadsheetFileType};

        // Cache the file type
        if let Some(path) = &self.document_meta.path {
            if let Some(ext) = path.extension().and_then(|e| e.to_str()) {
                self.default_app_prompt_file_type = SpreadsheetFileType::from_ext(ext);
            }
        }

        // Mark shown this session
        mark_shown_this_session();

        // Record timestamp for cool-down for THIS extension (in case they ignore)
        if let Some(ext_key) = self.get_prompt_ext_key() {
            update_user_settings(cx, |settings| {
                settings.mark_default_app_prompt_shown(&ext_key);
            });
        }

        self.default_app_prompt_state = DefaultAppPromptState::Showing;
    }

    /// Set VisiGrid as the default handler for the current file type.
    #[cfg(target_os = "macos")]
    pub fn set_as_default_app(&mut self, cx: &mut Context<Self>) {
        use crate::default_app::{set_as_default_handler, is_default_handler};

        let file_type = match self.get_prompt_file_type() {
            Some(ft) => ft,
            None => return,
        };

        let ext_key = self.get_prompt_ext_key();

        match set_as_default_handler(file_type) {
            Ok(()) => {
                // Check if it actually worked (duti succeeded)
                if is_default_handler(file_type) {
                    // Success! Show brief confirmation
                    self.default_app_prompt_state = DefaultAppPromptState::Success;
                    self.default_app_prompt_success_timer = Some(std::time::Instant::now());
                    // Mark completed for this extension (permanent)
                    if let Some(ext) = ext_key {
                        update_user_settings(cx, |settings| {
                            settings.mark_default_app_completed(&ext);
                        });
                    }
                } else {
                    // Needs manual completion in Settings
                    // Note: Do NOT permanently dismiss - just cool-down
                    // The cool-down timestamp is already set from when we showed the prompt
                    self.default_app_prompt_state = DefaultAppPromptState::NeedsSettings;
                    self.needs_settings_entered_at = Some(std::time::Instant::now());
                    self.needs_settings_check_count = 0;
                }
            }
            Err(_) => {
                // Failed - needs Settings
                // Note: Do NOT permanently dismiss - just cool-down
                self.default_app_prompt_state = DefaultAppPromptState::NeedsSettings;
                self.needs_settings_entered_at = Some(std::time::Instant::now());
                self.needs_settings_check_count = 0;
            }
        }

        cx.notify();
    }

    #[cfg(not(target_os = "macos"))]
    pub fn set_as_default_app(&mut self, _cx: &mut Context<Self>) {}

    /// Open System Settings to complete the default app setup.
    #[cfg(target_os = "macos")]
    pub fn open_default_app_settings(&mut self, cx: &mut Context<Self>) {
        use std::process::Command;

        let _ = Command::new("open")
            .args(["x-apple.systempreferences:com.apple.ExtensionsPreferences"])
            .spawn();

        // Keep in NeedsSettings state (don't hide) so we can re-check on focus
        // The prompt will be re-checked when the window regains focus
        cx.notify();
    }

    #[cfg(not(target_os = "macos"))]
    pub fn open_default_app_settings(&mut self, _cx: &mut Context<Self>) {}

    /// Dismiss the default app prompt permanently for this extension (user clicked ✕).
    pub fn dismiss_default_app_prompt(&mut self, cx: &mut Context<Self>) {
        if let Some(ext_key) = self.get_prompt_ext_key() {
            update_user_settings(cx, |settings| {
                settings.dismiss_default_app_prompt(&ext_key);
            });
        }
        self.default_app_prompt_state = DefaultAppPromptState::Hidden;
        self.needs_settings_entered_at = None;
        self.needs_settings_check_count = 0;
        cx.notify();
    }

    /// Re-check default handler after returning from System Settings.
    /// Called when window regains focus while in NeedsSettings state.
    /// Note: This is now mostly handled by check_default_app_prompt_timer(),
    /// but we keep this for explicit calls if needed.
    #[cfg(target_os = "macos")]
    pub fn recheck_default_app_handler(&mut self, cx: &mut Context<Self>) {
        use crate::default_app::is_default_handler;

        // Only re-check if we're in NeedsSettings state
        if self.default_app_prompt_state != DefaultAppPromptState::NeedsSettings {
            return;
        }

        let file_type = match self.get_prompt_file_type() {
            Some(ft) => ft,
            None => {
                self.default_app_prompt_state = DefaultAppPromptState::Hidden;
                self.needs_settings_entered_at = None;
                self.needs_settings_check_count = 0;
                cx.notify();
                return;
            }
        };

        // Check if they completed the setup in Settings
        if is_default_handler(file_type) {
            // Success! Show brief confirmation
            self.default_app_prompt_state = DefaultAppPromptState::Success;
            self.default_app_prompt_success_timer = Some(std::time::Instant::now());
            self.needs_settings_entered_at = None;
            self.needs_settings_check_count = 0;

            // Mark completed for this extension
            if let Some(ext_key) = self.get_prompt_ext_key() {
                update_user_settings(cx, |settings| {
                    settings.mark_default_app_completed(&ext_key);
                });
            }
        } else {
            // Still not default - hide for now (cool-down will handle re-show)
            self.default_app_prompt_state = DefaultAppPromptState::Hidden;
            self.needs_settings_entered_at = None;
            self.needs_settings_check_count = 0;
        }

        cx.notify();
    }

    #[cfg(not(target_os = "macos"))]
    pub fn recheck_default_app_handler(&mut self, _cx: &mut Context<Self>) {}

    /// Check if success timer has expired and hide the prompt.
    /// Also handles NeedsSettings state with exponential backoff:
    /// - Check 1: at 3 seconds
    /// - Check 2: at 8 seconds
    /// - Check 3: at 20 seconds
    /// - Then stop polling (rely on next file open or next session)
    pub fn check_default_app_prompt_timer(&mut self, cx: &mut Context<Self>) {
        if self.default_app_prompt_state == DefaultAppPromptState::Success {
            if let Some(started) = self.default_app_prompt_success_timer {
                if started.elapsed() > std::time::Duration::from_secs(2) {
                    self.default_app_prompt_state = DefaultAppPromptState::Hidden;
                    self.default_app_prompt_success_timer = None;
                    cx.notify();
                }
            }
        } else if self.default_app_prompt_state == DefaultAppPromptState::NeedsSettings {
            #[cfg(target_os = "macos")]
            {
                use crate::default_app::is_default_handler;
                use std::time::Instant;

                // Exponential backoff schedule (seconds since entered_at):
                // Check 0: 3s, Check 1: 8s, Check 2: 20s, then stop
                const CHECK_SCHEDULE: [u64; 3] = [3, 8, 20];

                let now = Instant::now();
                let entered_at = match self.needs_settings_entered_at {
                    Some(t) => t,
                    None => return, // No timestamp means we shouldn't be polling
                };

                let elapsed_secs = now.duration_since(entered_at).as_secs();
                let check_count = self.needs_settings_check_count as usize;

                // Already exhausted all checks? Stop polling.
                if check_count >= CHECK_SCHEDULE.len() {
                    return;
                }

                // Not yet time for the next check?
                let next_check_at = CHECK_SCHEDULE[check_count];
                if elapsed_secs < next_check_at {
                    return;
                }

                // Time for a check - increment counter first
                self.needs_settings_check_count += 1;

                // Re-check handler status
                if let Some(file_type) = self.get_prompt_file_type() {
                    if is_default_handler(file_type) {
                        // User completed setup! Show success briefly
                        self.default_app_prompt_state = DefaultAppPromptState::Success;
                        self.default_app_prompt_success_timer = Some(now);
                        self.needs_settings_entered_at = None;
                        self.needs_settings_check_count = 0;

                        // Mark completed for this extension
                        if let Some(ext_key) = self.get_prompt_ext_key() {
                            update_user_settings(cx, |settings| {
                                settings.mark_default_app_completed(&ext_key);
                            });
                        }

                        cx.notify();
                    }
                    // If not default yet, keep chip visible but stop polling after 3 checks
                }
            }
        }
    }
}
