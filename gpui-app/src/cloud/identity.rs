// Cloud sync state — runtime-only state (not persisted).
//
// CloudIdentity is defined in visigrid_io::native and re-exported from cloud/mod.rs.

/// Transient cloud sync state (not persisted — computed at runtime).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CloudSyncState {
    /// No cloud identity attached
    Local,
    /// Last upload matches local file
    Synced,
    /// Local changes pending upload
    Dirty,
    /// Upload in progress
    Syncing,
    /// Can't reach API
    Offline,
    /// Upload failed
    Error,
}

impl Default for CloudSyncState {
    fn default() -> Self {
        Self::Local
    }
}

impl CloudSyncState {
    /// Short label for status bar display.
    pub fn label(&self) -> &'static str {
        match self {
            Self::Local => "Local",
            Self::Synced => "Synced",
            Self::Dirty => "Modified",
            Self::Syncing => "Syncing...",
            Self::Offline => "Offline",
            Self::Error => "Error",
        }
    }
}
