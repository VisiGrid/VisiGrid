// Cloud sync — post-save upload with debounce/coalescing.
//
// After each local save, if the file has a cloud identity, we schedule a
// background upload. Rapid saves are coalesced via a generation counter:
// only the latest generation actually uploads.

use crate::app::Spreadsheet;
use crate::cloud::identity::CloudSyncState;
use crate::cloud::sheets_client::SheetsClient;
use crate::hub::client::{hash_bytes, HubError};

impl Spreadsheet {
    /// Schedule a cloud upload after a local save.
    ///
    /// Increments the generation counter and spawns a debounced task.
    /// If a newer save arrives within the debounce window, the older
    /// task will see a stale generation and bail out.
    pub fn cloud_schedule_upload(&mut self, cx: &mut gpui::Context<Self>) {
        let identity = match &self.cloud_identity {
            Some(id) => id.clone(),
            None => return,
        };

        let path = match &self.current_file {
            Some(p) => p.clone(),
            None => return,
        };

        self.cloud_upload_generation += 1;
        let generation = self.cloud_upload_generation;
        self.cloud_sync_state = CloudSyncState::Dirty;
        cx.notify();

        cx.spawn(async move |this, cx| {
            // Debounce: wait 500ms so rapid saves coalesce
            smol::Timer::after(std::time::Duration::from_millis(500)).await;

            // Check if a newer save superseded us
            let is_current = this.update(cx, |this, _cx| {
                this.cloud_upload_generation == generation
            }).unwrap_or(false);
            if !is_current {
                return; // Superseded by a newer save
            }

            // Set state to Syncing
            let _ = this.update(cx, |this, cx| {
                this.cloud_sync_state = CloudSyncState::Syncing;
                cx.notify();
            });

            // Read file bytes on background thread
            let path_clone = path.clone();
            let file_bytes = match smol::unblock(move || std::fs::read(&path_clone)).await {
                Ok(bytes) => bytes,
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        this.cloud_sync_state = CloudSyncState::Error;
                        this.cloud_last_error = Some(format!("Failed to read file: {}", e));
                        cx.notify();
                    });
                    return;
                }
            };

            let content_hash = hash_bytes(&file_bytes);
            let byte_size = file_bytes.len() as u64;
            let sheet_id = identity.sheet_id;

            // Request presigned upload URL
            let save_result = {
                smol::unblock(move || {
                    let client = SheetsClient::from_saved_auth()?;
                    let save_resp = client.save_sheet(sheet_id, byte_size)?;
                    client.upload_to_url(&save_resp.upload_url, &save_resp.headers, file_bytes)?;
                    Ok::<_, HubError>(())
                }).await
            };

            match save_result {
                Ok(()) => {
                    let hash = content_hash.clone();
                    let synced_path = path.clone();
                    let _ = this.update(cx, |this, cx| {
                        if let Some(ref mut id) = this.cloud_identity {
                            id.last_synced_hash = Some(hash);
                            id.last_synced_at = Some(now_iso8601());

                            // Persist updated identity to .sheet file
                            if let Err(e) = crate::cloud::save_cloud_identity(&synced_path, id) {
                                eprintln!("Warning: failed to persist cloud identity: {}", e);
                            }
                        }
                        this.cloud_sync_state = CloudSyncState::Synced;
                        cx.notify();
                    });
                }
                Err(HubError::Network(_)) => {
                    let _ = this.update(cx, |this, cx| {
                        this.cloud_sync_state = CloudSyncState::Offline;
                        this.cloud_last_error = Some("Network unavailable".to_string());
                        cx.notify();
                    });
                }
                Err(e) => {
                    let msg = e.to_string();
                    let _ = this.update(cx, |this, cx| {
                        this.cloud_sync_state = CloudSyncState::Error;
                        this.cloud_last_error = Some(msg);
                        cx.notify();
                    });
                }
            }
        }).detach();
    }

    /// Retry a failed or offline cloud upload.
    pub fn cloud_retry_upload(&mut self, cx: &mut gpui::Context<Self>) {
        if self.cloud_identity.is_some() {
            self.cloud_schedule_upload(cx);
        }
    }
}

fn now_iso8601() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    format!("{}", secs)
}
