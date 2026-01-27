//! Global window registry for multi-window management
//!
//! Tracks all open VisiGrid windows with their metadata (title, dirty state, path).
//! Used by the window switcher palette and other multi-window features.
//!
//! # Debug Invariants
//!
//! In debug builds, the registry validates:
//! - No duplicate handles
//! - Count changes correctly on open/close
//! - Registry is never in an inconsistent state

use gpui::{AnyWindowHandle, Global};
use std::path::PathBuf;

/// Metadata for a registered window
#[derive(Clone, Debug)]
pub struct WindowInfo {
    /// The window handle for focusing
    pub handle: AnyWindowHandle,
    /// Display title (e.g., "Untitled", "budget.xlsx")
    pub title: String,
    /// Whether the workbook has unsaved changes
    pub is_dirty: bool,
    /// Full path to the file, if saved
    pub path: Option<PathBuf>,
}

impl WindowInfo {
    /// Create a new window info entry
    pub fn new(handle: AnyWindowHandle, title: String, is_dirty: bool, path: Option<PathBuf>) -> Self {
        Self {
            handle,
            title,
            is_dirty,
            path,
        }
    }

    /// Display string with dirty indicator: "Untitled" or "budget.xlsx ●"
    pub fn display_title(&self) -> String {
        if self.is_dirty {
            format!("{} ●", self.title)
        } else {
            self.title.clone()
        }
    }

    /// Optional subtitle showing the path (dimmed in UI)
    pub fn subtitle(&self) -> Option<String> {
        self.path.as_ref().and_then(|p| {
            p.parent()
                .and_then(|parent| parent.to_str())
                .map(|s| s.to_string())
        })
    }
}

/// Global registry of all open windows
#[derive(Default)]
pub struct WindowRegistry {
    windows: Vec<WindowInfo>,
}

impl Global for WindowRegistry {}

impl WindowRegistry {
    /// Create an empty registry
    pub fn new() -> Self {
        Self { windows: vec![] }
    }

    /// Register a new window
    pub fn register(&mut self, info: WindowInfo) {
        #[cfg(debug_assertions)]
        let count_before = self.windows.len();

        // Don't register duplicates
        debug_assert!(
            !self.windows.iter().any(|w| w.handle == info.handle),
            "WindowRegistry: attempted to register duplicate handle"
        );

        if !self.windows.iter().any(|w| w.handle == info.handle) {
            self.windows.push(info);
        }

        #[cfg(debug_assertions)]
        {
            debug_assert_eq!(
                self.windows.len(),
                count_before + 1,
                "WindowRegistry: count should increase by 1 on register"
            );
            self.debug_check_invariants();
        }
    }

    /// Unregister a window by handle
    pub fn unregister(&mut self, handle: AnyWindowHandle) {
        #[cfg(debug_assertions)]
        let count_before = self.windows.len();
        #[cfg(debug_assertions)]
        let had_handle = self.windows.iter().any(|w| w.handle == handle);

        self.windows.retain(|w| w.handle != handle);

        #[cfg(debug_assertions)]
        {
            if had_handle {
                debug_assert_eq!(
                    self.windows.len(),
                    count_before - 1,
                    "WindowRegistry: count should decrease by 1 on unregister"
                );
            }
            self.debug_check_invariants();
        }
    }

    /// Update a window's metadata
    pub fn update(&mut self, handle: AnyWindowHandle, title: String, is_dirty: bool, path: Option<PathBuf>) {
        if let Some(info) = self.windows.iter_mut().find(|w| w.handle == handle) {
            info.title = title;
            info.is_dirty = is_dirty;
            info.path = path;
        }

        #[cfg(debug_assertions)]
        self.debug_check_invariants();
    }

    /// Get all registered windows
    pub fn windows(&self) -> &[WindowInfo] {
        &self.windows
    }

    /// Get window count
    pub fn count(&self) -> usize {
        self.windows.len()
    }

    /// Find a window by handle
    pub fn find(&self, handle: AnyWindowHandle) -> Option<&WindowInfo> {
        self.windows.iter().find(|w| w.handle == handle)
    }

    /// Check if a handle is registered
    pub fn contains(&self, handle: AnyWindowHandle) -> bool {
        self.windows.iter().any(|w| w.handle == handle)
    }

    /// Debug-only invariant check: no duplicate handles
    #[cfg(debug_assertions)]
    fn debug_check_invariants(&self) {
        // Check for duplicate handles
        let mut seen = std::collections::HashSet::new();
        for window in &self.windows {
            debug_assert!(
                seen.insert(window.handle),
                "WindowRegistry invariant violation: duplicate handle detected"
            );
        }
    }
}

// Note: Tests for WindowRegistry require real AnyWindowHandle which can't be created
// in unit tests. The registry logic is tested implicitly through integration tests.
