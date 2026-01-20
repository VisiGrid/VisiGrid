//! Settings persistence (load/save)
//!
//! User settings are stored in `~/.config/visigrid/settings.json`
//! Document settings are stored in sidecar files: `myfile.vgrid.settings.json`

use std::fs;
use std::path::{Path, PathBuf};

use super::document::DocumentSettings;
use super::user::UserSettings;

/// Get the path to the user settings file
pub fn user_settings_path() -> Option<PathBuf> {
    dirs::config_dir().map(|p| p.join("visigrid").join("settings.json"))
}

/// Load user settings from disk
///
/// Returns default settings if file doesn't exist or can't be parsed.
/// This is intentional - we don't want settings errors to prevent app startup.
pub fn load_user_settings() -> UserSettings {
    user_settings_path()
        .and_then(|p| fs::read_to_string(p).ok())
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

/// Save user settings to disk
///
/// Creates the config directory if it doesn't exist.
/// Silently ignores errors (settings are not critical for operation).
pub fn save_user_settings(settings: &UserSettings) {
    if let Some(path) = user_settings_path() {
        if let Some(parent) = path.parent() {
            let _ = fs::create_dir_all(parent);
        }
        let _ = fs::write(path, serde_json::to_string_pretty(settings).unwrap_or_default());
    }
}

// ============================================================================
// Document settings persistence (sidecar files)
// ============================================================================

/// Get the sidecar settings path for a document
///
/// Format: `myfile.vgrid` → `myfile.vgrid.settings.json`
/// This works for any extension: .vgrid, .csv, .xlsx, etc.
pub fn doc_settings_path(doc_path: &Path) -> PathBuf {
    let mut sidecar = doc_path.as_os_str().to_owned();
    sidecar.push(".settings.json");
    PathBuf::from(sidecar)
}

/// Load document settings from sidecar file
///
/// Returns default settings if sidecar doesn't exist or can't be parsed.
/// This is intentional - missing/invalid sidecar should not prevent file open.
pub fn load_doc_settings(doc_path: &Path) -> DocumentSettings {
    let sidecar = doc_settings_path(doc_path);
    fs::read_to_string(&sidecar)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

/// Save document settings to sidecar file (atomic write)
///
/// Uses write-to-temp-then-rename pattern to prevent corruption on crash.
/// Returns Ok(()) on success, Err on failure (caller can decide whether to warn user).
pub fn save_doc_settings(doc_path: &Path, settings: &DocumentSettings) -> std::io::Result<()> {
    let sidecar = doc_settings_path(doc_path);
    let temp = doc_settings_path(doc_path).with_extension("json.tmp");

    // Write to temp file
    let json = serde_json::to_string_pretty(settings)
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
    fs::write(&temp, json)?;

    // Atomic rename
    fs::rename(&temp, &sidecar)?;

    Ok(())
}

/// Delete document settings sidecar file (optional cleanup)
///
/// Used when a file is deleted or moved. Silently ignores errors.
pub fn delete_doc_settings(doc_path: &Path) {
    let sidecar = doc_settings_path(doc_path);
    let _ = fs::remove_file(sidecar);
}

/// Copy document settings to a new path (for Save As)
///
/// Loads from old path, saves to new path. Returns Ok even if source doesn't exist.
pub fn copy_doc_settings(old_path: &Path, new_path: &Path) -> std::io::Result<()> {
    let settings = load_doc_settings(old_path);
    // Only save if there are non-default settings
    // (avoid creating empty sidecar files)
    if settings != DocumentSettings::default() {
        save_doc_settings(new_path, &settings)?;
    }
    Ok(())
}

/// Open the settings file in the system editor (power user escape hatch)
#[cfg(target_os = "linux")]
pub fn open_settings_file() -> std::io::Result<()> {
    use std::env;

    if let Some(path) = user_settings_path() {
        // Ensure file exists with defaults
        if !path.exists() {
            save_user_settings(&UserSettings::default());
        }

        // Try $VISUAL, then $EDITOR, then fall back to xdg-open
        // This respects user's preferred editor (e.g., vim, code, nano)
        if let Ok(editor) = env::var("VISUAL").or_else(|_| env::var("EDITOR")) {
            std::process::Command::new(&editor).arg(&path).spawn()?;
        } else {
            // Fall back to xdg-open (may open in browser for JSON)
            std::process::Command::new("xdg-open").arg(&path).spawn()?;
        }
    }
    Ok(())
}

#[cfg(target_os = "macos")]
pub fn open_settings_file() -> std::io::Result<()> {
    use std::env;

    if let Some(path) = user_settings_path() {
        // Ensure file exists with defaults
        if !path.exists() {
            save_user_settings(&UserSettings::default());
        }

        // Try $VISUAL, then $EDITOR, then fall back to `open`
        if let Ok(editor) = env::var("VISUAL").or_else(|_| env::var("EDITOR")) {
            std::process::Command::new(&editor).arg(&path).spawn()?;
        } else {
            // Fall back to macOS `open` (uses default app for JSON)
            std::process::Command::new("open").arg(&path).spawn()?;
        }
    }
    Ok(())
}

#[cfg(target_os = "windows")]
pub fn open_settings_file() -> std::io::Result<()> {
    use std::env;

    if let Some(path) = user_settings_path() {
        // Ensure file exists with defaults
        if !path.exists() {
            save_user_settings(&UserSettings::default());
        }

        // Try %EDITOR%, then fall back to `start`
        if let Ok(editor) = env::var("EDITOR") {
            std::process::Command::new(&editor).arg(&path).spawn()?;
        } else {
            // Fall back to Windows `start` (uses default app for JSON)
            std::process::Command::new("cmd")
                .args(["/C", "start", "", path.to_str().unwrap_or("")])
                .spawn()?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::settings::types::{EnterBehavior, Setting, TipId};

    #[test]
    fn test_default_settings_serialize() {
        let settings = UserSettings::default();
        let json = serde_json::to_string_pretty(&settings).unwrap();
        println!("Default settings JSON:\n{}", json);

        // Should be parseable
        let _parsed: UserSettings = serde_json::from_str(&json).unwrap();
    }

    #[test]
    fn test_settings_roundtrip() {
        let mut settings = UserSettings::default();
        settings.appearance.theme_id = Setting::Value("catppuccin-mocha".to_string());
        settings.editing.enter_behavior = Setting::Value(EnterBehavior::MoveRight);
        settings.tips.dismiss(TipId::F2Edit);

        let json = serde_json::to_string_pretty(&settings).unwrap();
        let parsed: UserSettings = serde_json::from_str(&json).unwrap();

        assert!(matches!(parsed.appearance.theme_id, Setting::Value(ref s) if s == "catppuccin-mocha"));
        assert!(matches!(parsed.editing.enter_behavior, Setting::Value(EnterBehavior::MoveRight)));
        assert!(parsed.tips.is_dismissed(TipId::F2Edit));
    }

    #[test]
    fn test_inherit_from_missing_field() {
        // When a field is missing from JSON, it should be Inherit
        let json = r#"{"appearance": {}}"#;
        let settings: UserSettings = serde_json::from_str(json).unwrap();

        // theme_id should be Inherit (missing from JSON)
        assert!(matches!(settings.appearance.theme_id, Setting::Inherit));
    }

    #[test]
    fn test_doc_settings_sidecar_roundtrip() {
        use std::fs;
        use tempfile::TempDir;
        use crate::settings::DocumentSettings;

        // Create a temp directory for our test files
        let temp_dir = TempDir::new().unwrap();
        let doc_path = temp_dir.path().join("test.sheet");

        // Create a document with explicit (non-Inherit) settings
        let mut settings = DocumentSettings::default();
        settings.display.show_formulas = Setting::Value(true);  // Non-default
        settings.display.show_zeros = Setting::Value(false);    // Non-default

        // Save to sidecar
        save_doc_settings(&doc_path, &settings).unwrap();

        // Verify sidecar file exists
        let sidecar = doc_settings_path(&doc_path);
        assert!(sidecar.exists(), "Sidecar file should exist");

        // Verify sidecar contains our settings
        let json = fs::read_to_string(&sidecar).unwrap();
        assert!(json.contains("show_formulas"), "JSON should contain show_formulas");
        assert!(json.contains("show_zeros"), "JSON should contain show_zeros");

        // Load back and verify values match
        let loaded = load_doc_settings(&doc_path);
        assert!(matches!(loaded.display.show_formulas, Setting::Value(true)));
        assert!(matches!(loaded.display.show_zeros, Setting::Value(false)));
    }

    #[test]
    fn test_doc_settings_missing_sidecar_returns_default() {
        use tempfile::TempDir;
        use crate::settings::DocumentSettings;

        let temp_dir = TempDir::new().unwrap();
        let doc_path = temp_dir.path().join("no-sidecar.sheet");

        // Load from non-existent sidecar - should return defaults, not error
        let loaded = load_doc_settings(&doc_path);

        // Should be default values
        assert_eq!(loaded, DocumentSettings::default());
    }

    #[test]
    fn test_doc_settings_sidecar_path() {
        use std::path::Path;

        // Test various file extensions
        assert_eq!(
            doc_settings_path(Path::new("/foo/bar.sheet")),
            PathBuf::from("/foo/bar.sheet.settings.json")
        );
        assert_eq!(
            doc_settings_path(Path::new("/foo/data.csv")),
            PathBuf::from("/foo/data.csv.settings.json")
        );
        assert_eq!(
            doc_settings_path(Path::new("relative/path.xlsx")),
            PathBuf::from("relative/path.xlsx.settings.json")
        );
    }
}
