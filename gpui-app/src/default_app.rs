//! Default application handler detection and registration.
//!
//! On macOS, this module provides functionality to:
//! - Check if VisiGrid is the default handler for spreadsheet file types
//! - Request to be set as the default handler via Launch Services
//!
//! On other platforms, these functions are no-ops.

use std::sync::atomic::{AtomicBool, Ordering};

/// Session-level flag: have we shown the prompt this session?
/// Prevents spamming if user opens multiple files without dismissing.
static SHOWN_THIS_SESSION: AtomicBool = AtomicBool::new(false);

/// File types that VisiGrid can be the default handler for.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SpreadsheetFileType {
    /// Excel spreadsheets (.xlsx, .xls, .xlsm, .xlsb)
    Excel,
    /// CSV files (.csv)
    Csv,
    /// Tab-separated values (.tsv)
    Tsv,
    /// Native VisiGrid format (.vgrid)
    Native,
}

impl SpreadsheetFileType {
    /// Get the file type from a file extension.
    pub fn from_ext(ext: &str) -> Option<Self> {
        match ext.to_lowercase().as_str() {
            "xlsx" | "xls" | "xlsm" | "xlsb" => Some(Self::Excel),
            "csv" => Some(Self::Csv),
            "tsv" | "tab" => Some(Self::Tsv),
            "vgrid" | "sheet" => Some(Self::Native),
            _ => None,
        }
    }

    /// Get the UTI (Uniform Type Identifier) for this file type on macOS.
    #[cfg(target_os = "macos")]
    pub fn uti(&self) -> &'static str {
        match self {
            Self::Excel => "org.openxmlformats.spreadsheetml.sheet",
            Self::Csv => "public.comma-separated-values-text",
            Self::Tsv => "public.tab-separated-values-text",
            Self::Native => "com.visigrid.sheet",
        }
    }

    /// Short name for the prompt (e.g., "CSV files")
    pub fn short_name(&self) -> &'static str {
        match self {
            Self::Excel => "Excel files",
            Self::Csv => "CSV files",
            Self::Tsv => "TSV files",
            Self::Native => "VisiGrid files",
        }
    }

    /// Human-readable name for display.
    pub fn display_name(&self) -> &'static str {
        match self {
            Self::Excel => "Excel spreadsheets",
            Self::Csv => "CSV files",
            Self::Tsv => "TSV files",
            Self::Native => "VisiGrid files",
        }
    }
}

/// Check if we've already shown the prompt this session.
pub fn shown_this_session() -> bool {
    SHOWN_THIS_SESSION.load(Ordering::Relaxed)
}

/// Mark that we've shown the prompt this session.
pub fn mark_shown_this_session() {
    SHOWN_THIS_SESSION.store(true, Ordering::Relaxed);
}

/// Reset session state (for testing).
#[cfg(test)]
pub fn reset_session_state() {
    SHOWN_THIS_SESSION.store(false, Ordering::Relaxed);
}

/// Check if VisiGrid is the default handler for a file type.
///
/// Returns `true` if VisiGrid is already the default, `false` otherwise.
/// On non-macOS platforms, always returns `true` (suppresses the prompt).
#[cfg(target_os = "macos")]
pub fn is_default_handler(file_type: SpreadsheetFileType) -> bool {
    use std::process::Command;

    // Use duti to check the default handler
    // duti is a command-line tool for managing default applications on macOS
    // If duti isn't available, we fall back to assuming we're not the default
    let output = Command::new("duti")
        .args(["-x", file_type.uti()])
        .output();

    match output {
        Ok(output) => {
            let stdout = String::from_utf8_lossy(&output.stdout);
            // duti -x returns lines like:
            // VisiGrid.app
            // /Applications/VisiGrid.app
            // com.visigrid.app
            stdout.to_lowercase().contains("visigrid")
        }
        Err(_) => {
            // duti not available - use Launch Services API directly
            // For now, assume we're not the default (shows the prompt)
            // This is conservative: user can dismiss if not wanted
            is_default_handler_lsinfo(file_type)
        }
    }
}

#[cfg(target_os = "macos")]
fn is_default_handler_lsinfo(file_type: SpreadsheetFileType) -> bool {
    use std::process::Command;

    // Fallback: use lsappinfo or defaults
    // Try to get the default handler using mdls on a temp file
    // This is a heuristic - may not always work

    // For simplicity, try mdls on a known extension
    let ext = match file_type {
        SpreadsheetFileType::Excel => "xlsx",
        SpreadsheetFileType::Csv => "csv",
        SpreadsheetFileType::Tsv => "tsv",
        SpreadsheetFileType::Native => "vgrid",
    };

    // Create a temp file with the extension
    let temp_path = format!("/tmp/visigrid_check.{}", ext);
    let _ = std::fs::write(&temp_path, "");

    // Use `open -R` to reveal and check, or use Launch Services
    // For now, we'll use a simpler heuristic: check if the app bundle ID is registered
    let output = Command::new("/usr/bin/defaults")
        .args(["read", "com.apple.LaunchServices/com.apple.launchservices.secure", "LSHandlers"])
        .output();

    let _ = std::fs::remove_file(&temp_path);

    match output {
        Ok(output) => {
            let stdout = String::from_utf8_lossy(&output.stdout);
            // Look for our bundle ID associated with the UTI
            let uti = file_type.uti();
            // This is a rough heuristic - plist format varies
            stdout.contains("com.visigrid") && stdout.contains(uti)
        }
        Err(_) => false,
    }
}

#[cfg(not(target_os = "macos"))]
pub fn is_default_handler(_file_type: SpreadsheetFileType) -> bool {
    // On non-macOS, always return true to suppress the prompt
    true
}

/// Request to set VisiGrid as the default handler for a file type.
///
/// On macOS, this opens System Preferences to the default apps section.
/// The user must manually complete the assignment.
///
/// Returns `Ok(())` if the request was made, `Err` if something failed.
#[cfg(target_os = "macos")]
pub fn set_as_default_handler(file_type: SpreadsheetFileType) -> Result<(), String> {
    use std::process::Command;

    // Option 1: Try duti (if available)
    let duti_result = Command::new("duti")
        .args(["-s", "com.visigrid.app", file_type.uti(), "all"])
        .output();

    if let Ok(output) = duti_result {
        if output.status.success() {
            return Ok(());
        }
    }

    // Option 2: Open System Settings to the Extensions/File Type pane
    // This lets the user manually set the default
    // On macOS 13+, the path is: x-apple.systempreferences:com.apple.preference.extensions
    // On older macOS, it's the Finder info panel or Get Info
    let _ = Command::new("open")
        .args(["x-apple.systempreferences:com.apple.ExtensionsPreferences"])
        .spawn();

    // Also show a hint about how to do it manually
    // The user will see our status message
    Ok(())
}

#[cfg(not(target_os = "macos"))]
pub fn set_as_default_handler(_file_type: SpreadsheetFileType) -> Result<(), String> {
    // On non-macOS, this is a no-op
    Ok(())
}

/// Check if a file path looks like a temporary file (skip prompts for these).
pub fn is_temporary_file(path: &std::path::Path) -> bool {
    let path_str = path.to_string_lossy();

    // Common temp directories
    if path_str.contains("/tmp/")
        || path_str.contains("/var/folders/")
        || path_str.contains("/.Trash/")
        || path_str.contains("/Temp/")
        || path_str.contains("\\Temp\\")
    {
        return true;
    }

    // Files with temp-like names
    if let Some(name) = path.file_name().and_then(|n| n.to_str()) {
        if name.starts_with("~$")  // Excel temp files
            || name.starts_with("._")  // macOS resource forks
            || name.ends_with(".tmp")
            || name.ends_with(".temp")
        {
            return true;
        }
    }

    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_file_type_from_ext() {
        assert_eq!(SpreadsheetFileType::from_ext("xlsx"), Some(SpreadsheetFileType::Excel));
        assert_eq!(SpreadsheetFileType::from_ext("XLSX"), Some(SpreadsheetFileType::Excel));
        assert_eq!(SpreadsheetFileType::from_ext("csv"), Some(SpreadsheetFileType::Csv));
        assert_eq!(SpreadsheetFileType::from_ext("vgrid"), Some(SpreadsheetFileType::Native));
        assert_eq!(SpreadsheetFileType::from_ext("txt"), None);
    }

    #[test]
    fn test_is_temporary_file() {
        use std::path::Path;

        assert!(is_temporary_file(Path::new("/tmp/test.xlsx")));
        assert!(is_temporary_file(Path::new("/var/folders/ab/cd/T/test.csv")));
        assert!(is_temporary_file(Path::new("~$Budget.xlsx")));
        assert!(is_temporary_file(Path::new("/Users/me/.Trash/old.xlsx")));

        assert!(!is_temporary_file(Path::new("/Users/me/Documents/Budget.xlsx")));
        assert!(!is_temporary_file(Path::new("/home/user/data.csv")));
    }
}
