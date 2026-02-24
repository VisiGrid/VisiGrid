// Hub sync types

use visigrid_io::native::HubLink;

/// Internal activity state for debugging and status messages.
/// Used alongside HubStatus::Syncing to provide granular feedback.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HubActivity {
    /// Fetching remote status
    Checking,
    /// Downloading revision bytes
    Downloading,
    /// Verifying content hash
    Verifying,
    /// Writing file to disk
    Writing,
    /// Reloading workbook
    Reloading,
    /// Uploading file to cloud storage (Phase 3)
    Uploading,
    /// Finalizing revision after upload (Phase 3)
    Finalizing,
}

impl HubActivity {
    pub fn label(&self) -> &'static str {
        match self {
            HubActivity::Checking => "Checking...",
            HubActivity::Downloading => "Downloading...",
            HubActivity::Verifying => "Verifying...",
            HubActivity::Writing => "Writing...",
            HubActivity::Reloading => "Reloading...",
            HubActivity::Uploading => "Uploading...",
            HubActivity::Finalizing => "Finalizing...",
        }
    }
}

/// Status of the hub sync for a workbook.
///
/// This is derived from comparing local state (HubLink in .sheet) with
/// remote state (fetched from hub API).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HubStatus {
    /// Not linked to remote
    Unlinked,
    /// Linked and in sync (local_head matches remote current)
    Idle,
    /// Local has unpublished changes (local content hash differs from local_head_hash)
    Ahead,
    /// Remote has newer revision than local_head
    Behind,
    /// Both local changes and remote updates exist
    Diverged,
    /// Currently syncing (checking status, pulling, or publishing)
    Syncing,
    /// Cannot reach hub API
    Offline,
    /// User lacks permission to access the linked dataset
    Forbidden,
}

impl HubStatus {
    /// Returns a short display label for the status
    pub fn label(&self) -> &'static str {
        match self {
            HubStatus::Unlinked => "Not Linked",
            HubStatus::Idle => "Up to Date",
            HubStatus::Ahead => "Local Changes",
            HubStatus::Behind => "Update Available",
            HubStatus::Diverged => "Local & Remote Changed",
            HubStatus::Syncing => "Syncing...",
            HubStatus::Offline => "Offline",
            HubStatus::Forbidden => "Access Denied",
        }
    }

    /// Returns a hint about available actions for this status (Phase 1 honest)
    pub fn action_hint(&self) -> &'static str {
        match self {
            HubStatus::Unlinked => "",
            HubStatus::Idle => "Click to refresh",
            HubStatus::Ahead => "Open Remote as Copy available",
            HubStatus::Behind => "Click to update",
            HubStatus::Diverged => "Open Remote as Copy available",
            HubStatus::Syncing => "",
            HubStatus::Offline => "Click to retry",
            HubStatus::Forbidden => "",
        }
    }

    /// Returns true if the status indicates the workbook is linked to remote
    pub fn is_linked(&self) -> bool {
        !matches!(self, HubStatus::Unlinked)
    }

    /// Returns true if a pull operation is available
    pub fn can_pull(&self) -> bool {
        matches!(self, HubStatus::Behind | HubStatus::Diverged)
    }

    /// Returns true if syncing is in progress
    pub fn is_syncing(&self) -> bool {
        matches!(self, HubStatus::Syncing)
    }
}

/// Remote dataset status from hub API
#[derive(Debug, Clone)]
pub struct RemoteStatus {
    /// Current revision ID on remote
    pub current_revision_id: Option<String>,
    /// Content hash of the current revision
    pub content_hash: Option<String>,
    /// Byte size of the current revision
    pub byte_size: Option<u64>,
    /// When the revision was last updated (ISO 8601)
    pub updated_at: Option<String>,
    /// Who last updated it (user slug)
    pub updated_by: Option<String>,
}

/// Compute HubStatus from local and remote state.
///
/// # Arguments
/// * `hub_link` - The HubLink stored in the .sheet file (None if unlinked)
/// * `local_content_hash` - Blake3 hash of the current .sheet file contents
/// * `remote` - Status from hub API (None if offline/forbidden)
/// * `remote_error` - Error from fetching remote status, if any
pub fn compute_status(
    hub_link: Option<&HubLink>,
    local_content_hash: Option<&str>,
    remote: Option<&RemoteStatus>,
    remote_error: Option<&str>,
) -> HubStatus {
    // Not linked
    let Some(link) = hub_link else {
        return HubStatus::Unlinked;
    };

    // Check for access errors
    if let Some(err) = remote_error {
        if err.contains("403") || err.contains("forbidden") || err.contains("Forbidden") {
            return HubStatus::Forbidden;
        }
        return HubStatus::Offline;
    }

    // No remote status means offline
    let Some(remote) = remote else {
        return HubStatus::Offline;
    };

    // Check if local has uncommitted changes
    let local_dirty = match (&link.local_head_hash, local_content_hash) {
        (Some(head_hash), Some(current_hash)) => head_hash != current_hash,
        (None, Some(_)) => true, // Never synced, has content
        _ => false,
    };

    // Check if remote has updates
    let remote_ahead = match (&link.local_head_id, &remote.current_revision_id) {
        (Some(local_id), Some(remote_id)) => local_id != remote_id,
        (None, Some(_)) => true, // Never synced, remote has content
        _ => false,
    };

    match (local_dirty, remote_ahead) {
        (false, false) => HubStatus::Idle,
        (true, false) => HubStatus::Ahead,
        (false, true) => HubStatus::Behind,
        (true, true) => HubStatus::Diverged,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compute_status_unlinked() {
        assert_eq!(
            compute_status(None, None, None, None),
            HubStatus::Unlinked
        );
    }

    #[test]
    fn test_compute_status_idle() {
        let link = HubLink {
            repo_owner: "alice".to_string(),
            repo_slug: "budget".to_string(),
            dataset_id: "123".to_string(),
            local_head_id: Some("rev1".to_string()),
            local_head_hash: Some("abc123".to_string()),
            link_mode: "pull".to_string(),
            linked_at: "2024-01-01".to_string(),
            api_base: "https://api.visiapi.com".to_string(),
        };
        let remote = RemoteStatus {
            current_revision_id: Some("rev1".to_string()),
            content_hash: Some("abc123".to_string()),
            byte_size: Some(1000),
            updated_at: None,
            updated_by: None,
        };

        assert_eq!(
            compute_status(Some(&link), Some("abc123"), Some(&remote), None),
            HubStatus::Idle
        );
    }

    #[test]
    fn test_compute_status_behind() {
        let link = HubLink {
            repo_owner: "alice".to_string(),
            repo_slug: "budget".to_string(),
            dataset_id: "123".to_string(),
            local_head_id: Some("rev1".to_string()),
            local_head_hash: Some("abc123".to_string()),
            link_mode: "pull".to_string(),
            linked_at: "2024-01-01".to_string(),
            api_base: "https://api.visiapi.com".to_string(),
        };
        let remote = RemoteStatus {
            current_revision_id: Some("rev2".to_string()), // Different!
            content_hash: Some("def456".to_string()),
            byte_size: Some(1000),
            updated_at: None,
            updated_by: None,
        };

        assert_eq!(
            compute_status(Some(&link), Some("abc123"), Some(&remote), None),
            HubStatus::Behind
        );
    }

    #[test]
    fn test_compute_status_ahead() {
        let link = HubLink {
            repo_owner: "alice".to_string(),
            repo_slug: "budget".to_string(),
            dataset_id: "123".to_string(),
            local_head_id: Some("rev1".to_string()),
            local_head_hash: Some("abc123".to_string()),
            link_mode: "pull".to_string(),
            linked_at: "2024-01-01".to_string(),
            api_base: "https://api.visiapi.com".to_string(),
        };
        let remote = RemoteStatus {
            current_revision_id: Some("rev1".to_string()),
            content_hash: Some("abc123".to_string()),
            byte_size: Some(1000),
            updated_at: None,
            updated_by: None,
        };

        // Local has changed (different hash)
        assert_eq!(
            compute_status(Some(&link), Some("xyz789"), Some(&remote), None),
            HubStatus::Ahead
        );
    }

    #[test]
    fn test_compute_status_diverged() {
        let link = HubLink {
            repo_owner: "alice".to_string(),
            repo_slug: "budget".to_string(),
            dataset_id: "123".to_string(),
            local_head_id: Some("rev1".to_string()),
            local_head_hash: Some("abc123".to_string()),
            link_mode: "pull".to_string(),
            linked_at: "2024-01-01".to_string(),
            api_base: "https://api.visiapi.com".to_string(),
        };
        let remote = RemoteStatus {
            current_revision_id: Some("rev2".to_string()), // Remote updated
            content_hash: Some("def456".to_string()),
            byte_size: Some(1000),
            updated_at: None,
            updated_by: None,
        };

        // Local also changed
        assert_eq!(
            compute_status(Some(&link), Some("xyz789"), Some(&remote), None),
            HubStatus::Diverged
        );
    }

    #[test]
    fn test_compute_status_offline() {
        let link = HubLink {
            repo_owner: "alice".to_string(),
            repo_slug: "budget".to_string(),
            dataset_id: "123".to_string(),
            local_head_id: Some("rev1".to_string()),
            local_head_hash: Some("abc123".to_string()),
            link_mode: "pull".to_string(),
            linked_at: "2024-01-01".to_string(),
            api_base: "https://api.visiapi.com".to_string(),
        };

        assert_eq!(
            compute_status(Some(&link), Some("abc123"), None, Some("network error")),
            HubStatus::Offline
        );
    }

    #[test]
    fn test_compute_status_forbidden() {
        let link = HubLink {
            repo_owner: "alice".to_string(),
            repo_slug: "budget".to_string(),
            dataset_id: "123".to_string(),
            local_head_id: Some("rev1".to_string()),
            local_head_hash: Some("abc123".to_string()),
            link_mode: "pull".to_string(),
            linked_at: "2024-01-01".to_string(),
            api_base: "https://api.visiapi.com".to_string(),
        };

        assert_eq!(
            compute_status(Some(&link), Some("abc123"), None, Some("403 Forbidden")),
            HubStatus::Forbidden
        );
    }
}
