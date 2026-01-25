//! VisiHub sync operations for Spreadsheet.
//!
//! This module contains all hub_* methods for VisiGrid ↔ VisiHub synchronization:
//! - Status checking
//! - Pull (download from remote)
//! - Publish (upload to remote)
//! - Sign in/out authentication
//! - Link/unlink workbook to dataset

use gpui::*;
use crate::app::Spreadsheet;
use crate::hub::{
    HubStatus, HubActivity, HubClient, HubLink,
    compute_status, hash_file, hash_bytes, hashes_match,
    load_hub_link, save_hub_link, delete_hub_link,
    load_auth, save_auth, delete_auth, AuthCredentials,
};

/// Generate an ISO 8601 timestamp for the current time.
/// Format: "2024-01-15T14:30:00Z" (simplified UTC timestamp)
fn iso_timestamp_now() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};

    let duration = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default();

    let secs = duration.as_secs();

    // Calculate date/time components from unix timestamp
    // This is a simplified calculation that works for dates 1970-2099
    let days = secs / 86400;
    let time_of_day = secs % 86400;

    let hours = time_of_day / 3600;
    let minutes = (time_of_day % 3600) / 60;
    let seconds = time_of_day % 60;

    // Calculate year, month, day from days since epoch
    // Using a simplified algorithm
    let mut year = 1970;
    let mut remaining_days = days as i64;

    loop {
        let days_in_year = if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) {
            366
        } else {
            365
        };
        if remaining_days < days_in_year {
            break;
        }
        remaining_days -= days_in_year;
        year += 1;
    }

    let is_leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    let days_in_months = if is_leap {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month = 1;
    for &days_in_month in &days_in_months {
        if remaining_days < days_in_month as i64 {
            break;
        }
        remaining_days -= days_in_month as i64;
        month += 1;
    }
    let day = remaining_days + 1;

    format!("{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
            year, month, day, hours, minutes, seconds)
}

/// Generate a unique copy path for VisiHub downloads.
/// Format: "{stem} (from VisiHub).sheet", with (2), (3), etc. if exists.
fn generate_copy_path(original: &std::path::Path) -> std::path::PathBuf {
    let parent = original.parent().unwrap_or(std::path::Path::new("."));
    let stem = original.file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("workbook");

    // Try base name first
    let base_name = format!("{} (from VisiHub).sheet", stem);
    let mut candidate = parent.join(&base_name);

    if !candidate.exists() {
        return candidate;
    }

    // Find next available number
    for i in 2..=100 {
        let numbered_name = format!("{} (from VisiHub) ({}).sheet", stem, i);
        candidate = parent.join(&numbered_name);
        if !candidate.exists() {
            return candidate;
        }
    }

    // Fallback with timestamp (should never happen)
    let ts = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0);
    parent.join(format!("{} (from VisiHub) {}.sheet", stem, ts))
}

impl Spreadsheet {
    // ========================================================================
    // VisiHub Sync
    // ========================================================================

    /// Check hub status for the current file.
    /// Loads hub_link from file if needed, then queries VisiHub API.
    pub fn hub_check_status(&mut self, cx: &mut Context<Self>) {
        // Prevent concurrent checks
        if self.hub_check_in_progress {
            return;
        }

        // Need a saved file to check
        let Some(path) = self.current_file.clone() else {
            self.hub_status = HubStatus::Unlinked;
            self.hub_activity = None;
            cx.notify();
            return;
        };

        // Load hub link from file if not cached
        if self.hub_link.is_none() {
            match load_hub_link(&path) {
                Ok(Some(link)) => {
                    self.hub_link = Some(link);
                }
                Ok(None) => {
                    self.hub_status = HubStatus::Unlinked;
                    self.hub_activity = None;
                    cx.notify();
                    return;
                }
                Err(e) => {
                    self.status_message = Some(format!("Failed to load hub link: {}", e));
                    self.hub_last_error = Some(e.to_string());
                    cx.notify();
                    return;
                }
            }
        }

        let hub_link = self.hub_link.clone().unwrap();

        // Check if authenticated
        let client = match HubClient::from_saved_auth() {
            Ok(c) => c,
            Err(_) => {
                self.hub_status = HubStatus::Offline;
                self.hub_activity = None;
                self.hub_last_error = Some("Not authenticated".to_string());
                self.status_message = Some("Not signed in to VisiHub".to_string());
                cx.notify();
                return;
            }
        };

        // Compute local content hash
        let local_hash = hash_file(&path).ok();

        // Mark as syncing with Checking activity
        self.hub_status = HubStatus::Syncing;
        self.hub_activity = Some(HubActivity::Checking);
        self.hub_check_in_progress = true;
        self.hub_last_error = None;
        cx.notify();

        let dataset_id = hub_link.dataset_id.clone();

        // Spawn async task to check remote status
        cx.spawn(async move |this, cx| {
            let result = client.get_dataset_status(&dataset_id).await;

            let _ = this.update(cx, |this, cx| {
                this.hub_check_in_progress = false;
                this.hub_last_check = Some(std::time::Instant::now());
                this.hub_activity = None;

                match result {
                    Ok(remote) => {
                        this.hub_status = compute_status(
                            this.hub_link.as_ref(),
                            local_hash.as_deref(),
                            Some(&remote),
                            None,
                        );
                        this.hub_last_error = None;
                    }
                    Err(e) => {
                        let error_str = e.to_string();
                        this.hub_status = compute_status(
                            this.hub_link.as_ref(),
                            local_hash.as_deref(),
                            None,
                            Some(&error_str),
                        );
                        this.hub_last_error = Some(error_str.clone());
                        this.status_message = Some(format!("VisiHub: {}", error_str));
                    }
                }
                cx.notify();
            });
        }).detach();
    }

    /// Open remote version as a copy (always safe, never overwrites).
    /// Downloads the latest revision and saves to a new file.
    pub fn hub_open_remote_as_copy(&mut self, cx: &mut Context<Self>) {
        let Some(path) = self.current_file.clone() else {
            self.status_message = Some("No file open".to_string());
            cx.notify();
            return;
        };

        let Some(hub_link) = self.hub_link.clone() else {
            self.status_message = Some("Not linked to VisiHub".to_string());
            cx.notify();
            return;
        };

        let client = match HubClient::from_saved_auth() {
            Ok(c) => c,
            Err(_) => {
                self.hub_last_error = Some("Not authenticated".to_string());
                self.status_message = Some("Not signed in to VisiHub".to_string());
                cx.notify();
                return;
            }
        };

        self.hub_status = HubStatus::Syncing;
        self.hub_activity = Some(HubActivity::Checking);
        self.hub_last_error = None;
        cx.notify();

        let dataset_id = hub_link.dataset_id.clone();

        cx.spawn(async move |this, cx| {
            // Get current revision info
            let status = match client.get_dataset_status(&dataset_id).await {
                Ok(s) => s,
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_status = HubStatus::Offline;
                        this.hub_activity = None;
                        this.hub_last_error = Some(e.to_string());
                        this.status_message = Some(format!("Failed to get status: {}", e));
                        cx.notify();
                    });
                    return;
                }
            };

            let Some(revision_id) = status.current_revision_id.clone() else {
                let _ = this.update(cx, |this, cx| {
                    this.hub_activity = None;
                    this.status_message = Some("No revisions available".to_string());
                    cx.notify();
                });
                return;
            };

            let expected_hash = status.content_hash.clone();

            // Update activity to downloading
            let _ = this.update(cx, |this, cx| {
                this.hub_activity = Some(HubActivity::Downloading);
                cx.notify();
            });

            // Download the revision
            let content = match client.download_revision(&revision_id).await {
                Ok(c) => c,
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_status = HubStatus::Offline;
                        this.hub_activity = None;
                        this.hub_last_error = Some(e.to_string());
                        this.status_message = Some(format!("Download failed: {}", e));
                        cx.notify();
                    });
                    return;
                }
            };

            // Update activity to verifying
            let _ = this.update(cx, |this, cx| {
                this.hub_activity = Some(HubActivity::Verifying);
                cx.notify();
            });

            // Integrity check: verify hash matches
            let actual_hash = hash_bytes(&content);
            if let Some(expected) = &expected_hash {
                if !hashes_match(&actual_hash, expected) {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_status = HubStatus::Offline;
                        this.hub_activity = None;
                        this.hub_last_error = Some("Hash mismatch".to_string());
                        this.status_message = Some("Download failed integrity check. Please retry.".to_string());
                        cx.notify();
                    });
                    return;
                }
            }

            // Update activity to writing
            let _ = this.update(cx, |this, cx| {
                this.hub_activity = Some(HubActivity::Writing);
                cx.notify();
            });

            // Generate copy path: "{stem} (from VisiHub).sheet"
            let copy_path = generate_copy_path(&path);

            // Write to copy path
            if let Err(e) = std::fs::write(&copy_path, &content) {
                let _ = this.update(cx, |this, cx| {
                    this.hub_activity = None;
                    this.hub_last_error = Some(e.to_string());
                    this.status_message = Some(format!("Write failed: {}", e));
                    cx.notify();
                });
                return;
            }

            // Update hub_link in the NEW file with current head
            let mut updated_link = hub_link.clone();
            updated_link.local_head_id = Some(revision_id.clone());
            updated_link.local_head_hash = Some(actual_hash);

            if let Err(e) = save_hub_link(&copy_path, &updated_link) {
                // Non-fatal: file is saved, just link state is stale
                let _ = this.update(cx, |this, cx| {
                    this.status_message = Some(format!("Warning: could not update link: {}", e));
                    cx.notify();
                });
            }

            // Update activity to reloading
            let _ = this.update(cx, |this, cx| {
                this.hub_activity = Some(HubActivity::Reloading);
                cx.notify();
            });

            // Load the copy as the new workbook
            let _ = this.update(cx, |this, cx| {
                match visigrid_io::native::load_workbook(&copy_path) {
                    Ok(workbook) => {
                        this.workbook = workbook;
                        this.current_file = Some(copy_path.clone());
                        this.hub_link = Some(updated_link);
                        this.hub_status = HubStatus::Idle;
                        this.hub_activity = None;
                        this.hub_last_error = None;
                        this.is_modified = false;
                        this.history.clear();
                        this.document_meta.display_name = copy_path
                            .file_name()
                            .and_then(|n| n.to_str())
                            .unwrap_or("Untitled")
                            .to_string();
                        this.document_meta.is_saved = true;
                        this.document_meta.path = Some(copy_path);
                        this.request_title_refresh(cx);
                        this.status_message = Some("Opened remote copy from VisiHub".to_string());
                    }
                    Err(e) => {
                        this.hub_activity = None;
                        this.hub_last_error = Some(e.to_string());
                        this.status_message = Some(format!("Failed to open: {}", e));
                    }
                }
                cx.notify();
            });
        }).detach();
    }

    /// Pull latest version from VisiHub (update in place).
    /// Only allowed when local is clean (no uncommitted changes).
    /// If local is dirty, use hub_open_remote_as_copy() instead.
    pub fn hub_pull(&mut self, cx: &mut Context<Self>) {
        if !self.hub_status.can_pull() {
            self.status_message = Some("No updates available".to_string());
            cx.notify();
            return;
        }

        let Some(path) = self.current_file.clone() else {
            return;
        };

        let Some(hub_link) = self.hub_link.clone() else {
            return;
        };

        // SAFETY CHECK: Never overwrite dirty local changes
        // Compute current file hash and compare to local_head_hash
        let current_hash = match hash_file(&path) {
            Ok(h) => h,
            Err(e) => {
                self.hub_last_error = Some(e.to_string());
                self.status_message = Some(format!("Cannot verify local state: {}", e));
                cx.notify();
                return;
            }
        };

        let local_is_clean = hub_link.local_head_hash
            .as_ref()
            .map(|h| hashes_match(h, &current_hash))
            .unwrap_or(false);

        if !local_is_clean {
            // Local has changes - redirect to safe copy flow
            self.status_message = Some("Local changes detected. Opening remote as copy...".to_string());
            cx.notify();
            self.hub_open_remote_as_copy(cx);
            return;
        }

        // Local is clean - safe to update in place
        let client = match HubClient::from_saved_auth() {
            Ok(c) => c,
            Err(_) => {
                self.hub_last_error = Some("Not authenticated".to_string());
                self.status_message = Some("Not signed in to VisiHub".to_string());
                cx.notify();
                return;
            }
        };

        self.hub_status = HubStatus::Syncing;
        self.hub_activity = Some(HubActivity::Checking);
        self.hub_last_error = None;
        cx.notify();

        let dataset_id = hub_link.dataset_id.clone();

        cx.spawn(async move |this, cx| {
            // Get current revision info
            let status = match client.get_dataset_status(&dataset_id).await {
                Ok(s) => s,
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_status = HubStatus::Offline;
                        this.hub_activity = None;
                        this.hub_last_error = Some(e.to_string());
                        this.status_message = Some(format!("Failed to get status: {}", e));
                        cx.notify();
                    });
                    return;
                }
            };

            let Some(revision_id) = status.current_revision_id.clone() else {
                let _ = this.update(cx, |this, cx| {
                    this.hub_activity = None;
                    this.status_message = Some("No revisions available".to_string());
                    cx.notify();
                });
                return;
            };

            let expected_hash = status.content_hash.clone();

            // Update activity to downloading
            let _ = this.update(cx, |this, cx| {
                this.hub_activity = Some(HubActivity::Downloading);
                cx.notify();
            });

            // Download the revision
            let content = match client.download_revision(&revision_id).await {
                Ok(c) => c,
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_status = HubStatus::Offline;
                        this.hub_activity = None;
                        this.hub_last_error = Some(e.to_string());
                        this.status_message = Some(format!("Download failed: {}", e));
                        cx.notify();
                    });
                    return;
                }
            };

            // Update activity to verifying
            let _ = this.update(cx, |this, cx| {
                this.hub_activity = Some(HubActivity::Verifying);
                cx.notify();
            });

            // Integrity check: verify hash matches
            let actual_hash = hash_bytes(&content);
            if let Some(expected) = &expected_hash {
                if !hashes_match(&actual_hash, expected) {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_status = HubStatus::Offline;
                        this.hub_activity = None;
                        this.hub_last_error = Some("Hash mismatch".to_string());
                        this.status_message = Some("Download failed integrity check. Please retry.".to_string());
                        cx.notify();
                    });
                    return;
                }
            }

            // Update activity to writing
            let _ = this.update(cx, |this, cx| {
                this.hub_activity = Some(HubActivity::Writing);
                cx.notify();
            });

            // Write to temp file first, then atomic rename
            let temp_path = path.with_extension("sheet.tmp");
            if let Err(e) = std::fs::write(&temp_path, &content) {
                let _ = this.update(cx, |this, cx| {
                    this.hub_activity = None;
                    this.hub_last_error = Some(e.to_string());
                    this.status_message = Some(format!("Write failed: {}", e));
                    cx.notify();
                });
                return;
            }

            // Atomic rename
            if let Err(e) = std::fs::rename(&temp_path, &path) {
                let _ = std::fs::remove_file(&temp_path);
                let _ = this.update(cx, |this, cx| {
                    this.hub_activity = None;
                    this.hub_last_error = Some(e.to_string());
                    this.status_message = Some(format!("Save failed: {}", e));
                    cx.notify();
                });
                return;
            }

            // Update hub link with new head
            let mut updated_link = hub_link.clone();
            updated_link.local_head_id = Some(revision_id.clone());
            updated_link.local_head_hash = Some(actual_hash);

            if let Err(e) = save_hub_link(&path, &updated_link) {
                let _ = this.update(cx, |this, cx| {
                    this.hub_last_error = Some(e.to_string());
                    this.status_message = Some(format!("Failed to update link: {}", e));
                    cx.notify();
                });
                return;
            }

            // Update activity to reloading
            let _ = this.update(cx, |this, cx| {
                this.hub_activity = Some(HubActivity::Reloading);
                cx.notify();
            });

            // Reload the workbook
            let _ = this.update(cx, |this, cx| {
                match visigrid_io::native::load_workbook(&path) {
                    Ok(workbook) => {
                        this.workbook = workbook;
                        this.hub_link = Some(updated_link);
                        this.hub_status = HubStatus::Idle;
                        this.hub_activity = None;
                        this.hub_last_error = None;
                        this.is_modified = false;
                        this.history.clear();
                        this.status_message = Some("Updated from VisiHub".to_string());
                    }
                    Err(e) => {
                        this.hub_activity = None;
                        this.hub_last_error = Some(e.to_string());
                        this.status_message = Some(format!("Failed to reload: {}", e));
                    }
                }
                cx.notify();
            });
        }).detach();
    }

    /// Publish local changes to VisiHub.
    /// This is an explicit action, never automatic.
    /// If diverged, shows confirmation dialog first.
    pub fn hub_publish(&mut self, cx: &mut Context<Self>) {
        // Precondition: Must have a saved file
        if self.current_file.is_none() {
            self.status_message = Some("Save file first".to_string());
            cx.notify();
            return;
        }

        // Precondition: Must be linked
        let Some(ref hub_link) = self.hub_link else {
            self.status_message = Some("Link to VisiHub first".to_string());
            cx.notify();
            return;
        };

        // Precondition: Must be in publish mode (not pull-only)
        if hub_link.link_mode == "pull" {
            self.status_message = Some("This workbook is linked in Pull-only mode. Use 'VisiHub: Link to Dataset' to change to Pull & Publish.".to_string());
            cx.notify();
            return;
        }

        // Precondition: Must be signed in
        if load_auth().is_none() {
            self.status_message = Some("Sign in to VisiHub first".to_string());
            cx.notify();
            return;
        }

        // Check if we're in Diverged state - show confirmation dialog
        if self.hub_status == HubStatus::Diverged {
            self.mode = crate::mode::Mode::HubPublishConfirm;
            cx.notify();
            return;
        }

        // Not diverged - proceed with publish
        self.hub_publish_internal(cx);
    }

    /// Internal publish implementation (called after confirmation if diverged)
    pub(crate) fn hub_publish_internal(&mut self, cx: &mut Context<Self>) {
        let Some(path) = self.current_file.clone() else {
            return;
        };

        let Some(hub_link) = self.hub_link.clone() else {
            return;
        };

        let client = match HubClient::from_saved_auth() {
            Ok(c) => c,
            Err(_) => {
                self.hub_last_error = Some("Not authenticated".to_string());
                self.status_message = Some("Sign in to VisiHub first".to_string());
                cx.notify();
                return;
            }
        };

        // Read file and check size
        let file_bytes = match std::fs::read(&path) {
            Ok(b) => b,
            Err(e) => {
                self.hub_last_error = Some(e.to_string());
                self.status_message = Some(format!("Cannot read file: {}", e));
                cx.notify();
                return;
            }
        };

        const MAX_SIZE: u64 = 200 * 1024 * 1024; // 200 MB
        if file_bytes.len() as u64 > MAX_SIZE {
            self.status_message = Some("File too large for cloud sync (max 200 MB)".to_string());
            cx.notify();
            return;
        }

        // Compute content hash
        let content_hash = hash_file(&path).unwrap_or_default();
        let byte_size = file_bytes.len() as u64;

        // Start publish flow
        self.hub_status = HubStatus::Syncing;
        self.hub_activity = Some(HubActivity::Uploading);
        self.hub_last_error = None;
        cx.notify();

        let dataset_id = hub_link.dataset_id.clone();

        cx.spawn(async move |this, cx| {
            // Step 1: Create revision (get upload URL)
            let (revision_id, upload_url) = match client.create_revision(&dataset_id, &content_hash, byte_size).await {
                Ok(r) => r,
                Err(e) => {
                    let error_msg = e.to_string();
                    let _ = this.update(cx, |this, cx| {
                        this.hub_status = if error_msg.contains("403") || error_msg.contains("Forbidden") {
                            HubStatus::Forbidden
                        } else {
                            HubStatus::Offline
                        };
                        this.hub_activity = None;
                        this.hub_last_error = Some(error_msg.clone());
                        if error_msg.contains("403") || error_msg.contains("Forbidden") {
                            this.status_message = Some("You don't have permission to publish".to_string());
                        } else {
                            this.status_message = Some(format!("Failed to create revision: {}", error_msg));
                        }
                        cx.notify();
                    });
                    return;
                }
            };

            // Step 2: Upload to signed URL
            if let Err(e) = client.upload_to_signed_url(&upload_url, file_bytes).await {
                let _ = this.update(cx, |this, cx| {
                    this.hub_status = HubStatus::Offline;
                    this.hub_activity = None;
                    this.hub_last_error = Some(e.to_string());
                    this.status_message = Some(format!("Upload failed: {}", e));
                    cx.notify();
                });
                return;
            }

            // Update activity to Finalizing
            let _ = this.update(cx, |this, cx| {
                this.hub_activity = Some(HubActivity::Finalizing);
                cx.notify();
            });

            // Step 3: Complete revision
            if let Err(e) = client.complete_revision(&revision_id, &content_hash).await {
                let _ = this.update(cx, |this, cx| {
                    this.hub_status = HubStatus::Offline;
                    this.hub_activity = None;
                    this.hub_last_error = Some(e.to_string());
                    this.status_message = Some(format!("Failed to finalize: {}", e));
                    cx.notify();
                });
                return;
            }

            // SUCCESS: Update HubLink with new head
            let _ = this.update(cx, |this, cx| {
                if let Some(ref mut link) = this.hub_link {
                    link.local_head_id = Some(revision_id.clone());
                    link.local_head_hash = Some(content_hash.clone());

                    // Save updated link to file
                    if let Some(path) = &this.current_file {
                        let _ = save_hub_link(path, link);
                    }
                }

                this.hub_status = HubStatus::Idle;
                this.hub_activity = None;
                this.hub_last_error = None;
                this.status_message = Some("Published to VisiHub".to_string());
                cx.notify();
            });
        }).detach();
    }

    /// Unlink current file from VisiHub
    pub fn hub_unlink(&mut self, cx: &mut Context<Self>) {
        let Some(path) = &self.current_file else {
            return;
        };

        if let Err(e) = delete_hub_link(path) {
            self.hub_last_error = Some(e.to_string());
            self.status_message = Some(format!("Failed to unlink: {}", e));
        } else {
            self.hub_link = None;
            self.hub_status = HubStatus::Unlinked;
            self.hub_activity = None;
            self.hub_last_error = None;
            self.status_message = Some("Unlinked from VisiHub".to_string());
        }
        cx.notify();
    }

    /// Show hub sync diagnostics (debugging aid)
    pub fn hub_diagnostics(&mut self, cx: &mut Context<Self>) {
        let mut lines = vec!["=== VisiHub Diagnostics ===".to_string()];
        lines.push("Sync is manual. Nothing uploads automatically.".to_string());
        lines.push(String::new());

        // Auth status
        if let Some(creds) = load_auth() {
            let user = creds.user_slug.as_deref().unwrap_or("unknown");
            lines.push(format!("Signed in as: @{}", user));
        } else {
            lines.push("Not signed in".to_string());
        }

        // Status
        lines.push(format!("Status: {:?}", self.hub_status));
        if let Some(activity) = self.hub_activity {
            lines.push(format!("Activity: {:?}", activity));
        }

        // Link info
        if let Some(link) = &self.hub_link {
            lines.push(format!("Dataset: {}/{}/{}", link.repo_owner, link.repo_slug, link.dataset_id));
            lines.push(format!("API: {}", link.api_base));
            lines.push(format!("Linked at: {}", link.linked_at));
            lines.push(format!("Link mode: {}", link.link_mode));
            if let Some(head_id) = &link.local_head_id {
                lines.push(format!("Local head ID: {}", head_id));
            }
            if let Some(head_hash) = &link.local_head_hash {
                lines.push(format!("Local head hash: {}", head_hash));
            }
        } else {
            lines.push("Not linked".to_string());
        }

        // Timing
        if let Some(last_check) = self.hub_last_check {
            let elapsed = last_check.elapsed();
            lines.push(format!("Last check: {:.1}s ago", elapsed.as_secs_f64()));
        }

        // Errors
        if let Some(err) = &self.hub_last_error {
            lines.push(format!("Last error: {}", err));
        }

        // Current file
        if let Some(path) = &self.current_file {
            lines.push(format!("File: {}", path.display()));
            // Try to compute current hash
            if let Ok(hash) = hash_file(path) {
                lines.push(format!("Current hash: {}", hash));
            }
        }

        self.status_message = Some(lines.join("\n"));
        cx.notify();
    }

    /// Start VisiHub sign in flow.
    /// Opens browser to authorize, then shows paste token dialog as fallback.
    pub fn hub_sign_in(&mut self, cx: &mut Context<Self>) {
        // If already signed in, just show status
        if let Some(creds) = load_auth() {
            let user = creds.user_slug.as_deref().unwrap_or("unknown");
            self.status_message = Some(format!("Already signed in as @{}", user));
            cx.notify();
            return;
        }

        // Open browser to authorize
        let auth_url = "https://visihub.app/desktop/authorize";
        if let Err(e) = open::that(auth_url) {
            self.status_message = Some(format!("Failed to open browser: {}", e));
            cx.notify();
            return;
        }

        // Show paste token dialog as fallback
        self.hub_token_input.clear();
        self.mode = crate::mode::Mode::HubPasteToken;
        self.status_message = Some("Opening browser... Paste token below if callback fails.".to_string());
        cx.notify();
    }

    /// Complete sign in with pasted token
    pub fn hub_complete_sign_in(&mut self, cx: &mut Context<Self>) {
        let token = self.hub_token_input.trim().to_string();
        if token.is_empty() {
            self.status_message = Some("Token cannot be empty".to_string());
            cx.notify();
            return;
        }

        // Create credentials and verify with API
        let creds = AuthCredentials::new(
            token.clone(),
            "https://api.visihub.app".to_string(),
        );

        // Verify token by fetching user info
        let client = HubClient::new(creds.clone());

        cx.spawn(async move |this, cx| {
            match client.verify_token().await {
                Ok(user_info) => {
                    // Save credentials with user info
                    let mut full_creds = creds;
                    full_creds.user_slug = Some(user_info.slug.clone());
                    full_creds.email = Some(user_info.email.clone());

                    if let Err(e) = save_auth(&full_creds) {
                        let _ = this.update(cx, |this, cx| {
                            this.status_message = Some(format!("Failed to save credentials: {}", e));
                            cx.notify();
                        });
                        return;
                    }

                    let email = user_info.email.clone();
                    let slug = user_info.slug.clone();
                    let _ = this.update(cx, |this, cx| {
                        this.mode = crate::mode::Mode::Navigation;
                        this.hub_token_input.clear();
                        // Show verified identity with email for trust
                        this.status_message = Some(format!("Signed in as @{} ({})", slug, email));
                        cx.notify();
                    });
                }
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        // Better error message mentioning API base
                        this.status_message = Some(format!(
                            "Token could not be verified with api.visihub.app. Check you copied the full token. ({})",
                            e
                        ));
                        cx.notify();
                    });
                }
            }
        }).detach();
    }

    /// Cancel sign in dialog
    pub fn hub_cancel_sign_in(&mut self, cx: &mut Context<Self>) {
        self.mode = crate::mode::Mode::Navigation;
        self.hub_token_input.clear();
        self.status_message = None;
        cx.notify();
    }

    /// Insert character into hub token input
    pub fn hub_token_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.mode == crate::mode::Mode::HubPasteToken {
            self.hub_token_input.push(c);
            cx.notify();
        }
    }

    /// Backspace in hub token input
    pub fn hub_token_backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode == crate::mode::Mode::HubPasteToken {
            self.hub_token_input.pop();
            cx.notify();
        }
    }

    /// Paste text into hub token input
    pub fn hub_token_paste(&mut self, text: &str, cx: &mut Context<Self>) {
        if self.mode == crate::mode::Mode::HubPasteToken {
            // Clear and paste - tokens are usually pasted entirely
            self.hub_token_input = text.trim().to_string();
            cx.notify();
        }
    }

    /// Insert character into new dataset name
    pub fn hub_dataset_insert_char(&mut self, c: char, cx: &mut Context<Self>) {
        if self.mode == crate::mode::Mode::HubLink {
            self.hub_new_dataset_name.push(c);
            cx.notify();
        }
    }

    /// Backspace in new dataset name
    pub fn hub_dataset_backspace(&mut self, cx: &mut Context<Self>) {
        if self.mode == crate::mode::Mode::HubLink {
            self.hub_new_dataset_name.pop();
            cx.notify();
        }
    }

    /// Sign out from VisiHub
    pub fn hub_sign_out(&mut self, cx: &mut Context<Self>) {
        if let Err(e) = delete_auth() {
            self.status_message = Some(format!("Failed to sign out: {}", e));
        } else {
            self.status_message = Some("Signed out from VisiHub".to_string());
        }
        cx.notify();
    }

    /// Show link to VisiHub dialog
    pub fn hub_show_link_dialog(&mut self, cx: &mut Context<Self>) {
        // Must be signed in first
        if load_auth().is_none() {
            self.status_message = Some("Sign in to VisiHub first".to_string());
            cx.notify();
            return;
        }

        // Must have a saved file
        if self.current_file.is_none() {
            self.status_message = Some("Save the file first".to_string());
            cx.notify();
            return;
        }

        // Already linked?
        if self.hub_link.is_some() {
            self.status_message = Some("Already linked. Unlink first to change.".to_string());
            cx.notify();
            return;
        }

        // Reset dialog state and load repos
        self.hub_repos.clear();
        self.hub_selected_repo = None;
        self.hub_datasets.clear();
        self.hub_selected_dataset = None;
        self.hub_new_dataset_name.clear();
        self.hub_link_loading = true;
        self.mode = crate::mode::Mode::HubLink;
        cx.notify();

        // Fetch repos from API
        self.hub_fetch_repos(cx);
    }

    /// Fetch available repos from VisiHub
    fn hub_fetch_repos(&mut self, cx: &mut Context<Self>) {
        let client = match HubClient::from_saved_auth() {
            Ok(c) => c,
            Err(_) => {
                self.hub_link_loading = false;
                self.status_message = Some("Not signed in".to_string());
                cx.notify();
                return;
            }
        };

        cx.spawn(async move |this, cx| {
            match client.list_repos().await {
                Ok(repos) => {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_repos = repos;
                        this.hub_link_loading = false;
                        cx.notify();
                    });
                }
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_link_loading = false;
                        this.status_message = Some(format!("Failed to load repos: {}", e));
                        cx.notify();
                    });
                }
            }
        }).detach();
    }

    /// Select a repo in the link dialog
    pub fn hub_select_repo(&mut self, index: usize, cx: &mut Context<Self>) {
        if index >= self.hub_repos.len() {
            return;
        }

        self.hub_selected_repo = Some(index);
        self.hub_datasets.clear();
        self.hub_selected_dataset = None;
        self.hub_link_loading = true;
        cx.notify();

        // Fetch datasets for this repo
        let repo = self.hub_repos[index].clone();
        self.hub_fetch_datasets(&repo.owner, &repo.slug, cx);
    }

    /// Fetch datasets for a repo
    fn hub_fetch_datasets(&mut self, owner: &str, slug: &str, cx: &mut Context<Self>) {
        let client = match HubClient::from_saved_auth() {
            Ok(c) => c,
            Err(_) => {
                self.hub_link_loading = false;
                cx.notify();
                return;
            }
        };

        let owner = owner.to_string();
        let slug = slug.to_string();

        cx.spawn(async move |this, cx| {
            match client.list_datasets(&owner, &slug).await {
                Ok(datasets) => {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_datasets = datasets;
                        this.hub_link_loading = false;
                        cx.notify();
                    });
                }
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_link_loading = false;
                        this.status_message = Some(format!("Failed to load datasets: {}", e));
                        cx.notify();
                    });
                }
            }
        }).detach();
    }

    /// Select a dataset in the link dialog
    pub fn hub_select_dataset(&mut self, index: usize, cx: &mut Context<Self>) {
        if index >= self.hub_datasets.len() {
            return;
        }
        self.hub_selected_dataset = Some(index);
        cx.notify();
    }

    /// Confirm linking to selected dataset
    pub fn hub_confirm_link(&mut self, cx: &mut Context<Self>) {
        let Some(repo_idx) = self.hub_selected_repo else {
            self.status_message = Some("Select a repository".to_string());
            cx.notify();
            return;
        };

        let Some(dataset_idx) = self.hub_selected_dataset else {
            self.status_message = Some("Select a dataset".to_string());
            cx.notify();
            return;
        };

        let Some(path) = self.current_file.clone() else {
            return;
        };

        let repo = &self.hub_repos[repo_idx];
        let dataset = &self.hub_datasets[dataset_idx];

        // Create HubLink
        let link = HubLink {
            repo_owner: repo.owner.clone(),
            repo_slug: repo.slug.clone(),
            dataset_id: dataset.id.clone(),
            local_head_id: None,  // Will be set after first status check
            local_head_hash: None,
            link_mode: "pull".to_string(),
            linked_at: iso_timestamp_now(),
            api_base: "https://api.visihub.app".to_string(),
        };

        // Save to file
        if let Err(e) = save_hub_link(&path, &link) {
            self.status_message = Some(format!("Failed to save link: {}", e));
            cx.notify();
            return;
        }

        // Update state
        let dataset_name = dataset.name.clone();
        self.hub_link = Some(link);
        self.mode = crate::mode::Mode::Navigation;
        // Toast with full context: @owner/repo · dataset_name
        self.status_message = Some(format!("Linked to @{}/{} · {}", repo.owner, repo.slug, dataset_name));
        cx.notify();

        // Immediately check status to set baseline
        self.hub_check_status(cx);
    }

    /// Create a new dataset and link to it
    pub fn hub_create_and_link(&mut self, cx: &mut Context<Self>) {
        let Some(repo_idx) = self.hub_selected_repo else {
            self.status_message = Some("Select a repository first".to_string());
            cx.notify();
            return;
        };

        let name = self.hub_new_dataset_name.trim().to_string();
        if name.is_empty() {
            self.status_message = Some("Enter a dataset name".to_string());
            cx.notify();
            return;
        }

        let Some(path) = self.current_file.clone() else {
            return;
        };

        let repo = self.hub_repos[repo_idx].clone();
        self.hub_link_loading = true;
        cx.notify();

        let client = match HubClient::from_saved_auth() {
            Ok(c) => c,
            Err(_) => {
                self.hub_link_loading = false;
                self.status_message = Some("Not signed in".to_string());
                cx.notify();
                return;
            }
        };

        cx.spawn(async move |this, cx| {
            match client.create_dataset(&repo.owner, &repo.slug, &name).await {
                Ok(dataset_id) => {
                    // Create and save link
                    let link = HubLink {
                        repo_owner: repo.owner.clone(),
                        repo_slug: repo.slug.clone(),
                        dataset_id,
                        local_head_id: None,
                        local_head_hash: None,
                        link_mode: "pull".to_string(),
                        linked_at: iso_timestamp_now(),
                        api_base: "https://api.visihub.app".to_string(),
                    };

                    if let Err(e) = save_hub_link(&path, &link) {
                        let _ = this.update(cx, |this, cx| {
                            this.hub_link_loading = false;
                            this.status_message = Some(format!("Failed to save link: {}", e));
                            cx.notify();
                        });
                        return;
                    }

                    let _ = this.update(cx, |this, cx| {
                        this.hub_link = Some(link);
                        this.hub_link_loading = false;
                        this.mode = crate::mode::Mode::Navigation;
                        // Toast with full context: @owner/repo · dataset_name
                        this.status_message = Some(format!("Linked to @{}/{} · {}", repo.owner, repo.slug, name));
                        // Check status to establish baseline
                        this.hub_check_status(cx);
                    });
                }
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        this.hub_link_loading = false;
                        this.status_message = Some(format!("Failed to create dataset: {}", e));
                        cx.notify();
                    });
                }
            }
        }).detach();
    }

    /// Cancel link dialog
    pub fn hub_cancel_link(&mut self, cx: &mut Context<Self>) {
        self.mode = crate::mode::Mode::Navigation;
        self.hub_repos.clear();
        self.hub_datasets.clear();
        self.hub_selected_repo = None;
        self.hub_selected_dataset = None;
        self.hub_new_dataset_name.clear();
        self.hub_link_loading = false;
        cx.notify();
    }

    /// Cancel publish confirm dialog
    pub fn hub_cancel_publish_confirm(&mut self, cx: &mut Context<Self>) {
        self.mode = crate::mode::Mode::Navigation;
        cx.notify();
    }

    /// Confirm publish even when diverged (force publish)
    pub fn hub_confirm_publish_anyway(&mut self, cx: &mut Context<Self>) {
        self.mode = crate::mode::Mode::Navigation;
        cx.notify();
        // Proceed with actual publish (bypasses diverged check)
        self.hub_publish_internal(cx);
    }
}
