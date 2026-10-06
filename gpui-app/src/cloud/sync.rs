// Cloud sync — post-save upload with debounce/coalescing.
//
// After each local save, if the file has a cloud identity, we schedule a
// background upload. Rapid saves are coalesced via a generation counter:
// only the latest generation actually uploads.
//
// Every upload names the revision the file last synced at. If the cloud copy
// moved past it (an edit in the browser, another machine), the server refuses
// and the file enters Conflict: nothing is uploaded until the user chooses to
// overwrite the cloud copy or re-open it from the cloud.

use crate::app::Spreadsheet;
use crate::cloud::identity::CloudSyncState;
use crate::cloud::sheets_client::{conflict_revision, is_unauthorized, SheetsClient};
use crate::hub::client::{hash_bytes, HubError};

/// How a finished upload attempt ended.
enum UploadOutcome {
    /// Uploaded; the cloud copy is now at this revision.
    Synced(i64),
    /// The cloud copy changed since this file last synced.
    Conflict { current_revision: Option<i64> },
    /// Failed. `reserved` is a revision the save request took on the server
    /// before the upload failed, so a retry can recognise it as ours.
    Failed { error: HubError, reserved: Option<i64> },
}

impl Spreadsheet {
    /// Schedule a cloud upload after a local save.
    ///
    /// Increments the generation counter and spawns a debounced task.
    /// If a newer save arrives within the debounce window, the older
    /// task will see a stale generation and bail out.
    pub fn cloud_schedule_upload(&mut self, cx: &mut gpui::Context<Self>) {
        // A conflicted file keeps saving locally; it must not keep asking the
        // server to take it.
        if self.cloud_sync_state == CloudSyncState::Conflict {
            return;
        }
        self.cloud_start_upload(false, cx);
    }

    /// Retry a failed or offline cloud upload.
    pub fn cloud_retry_upload(&mut self, cx: &mut gpui::Context<Self>) {
        if self.cloud_identity.is_some() && self.cloud_sync_state != CloudSyncState::Conflict {
            self.cloud_start_upload(false, cx);
        }
    }

    /// Replace the cloud copy with this file, discarding whatever changed in
    /// the cloud since the last sync. The explicit way out of a conflict.
    pub fn cloud_overwrite_cloud_copy(&mut self, cx: &mut gpui::Context<Self>) {
        if self.cloud_identity.is_none() {
            self.status_message = Some("This file is not cloud-backed.".to_string());
            cx.notify();
            return;
        }
        self.cloud_start_upload(true, cx);
    }

    fn cloud_start_upload(&mut self, force: bool, cx: &mut gpui::Context<Self>) {
        // The live sequencer owns publication; REST saves would replace its
        // document and race the operation log.
        if self.cloud_live_enabled() { return; }
        let identity = match &self.cloud_identity {
            Some(id) => id.clone(),
            None => return,
        };

        let path = match &self.current_file {
            Some(p) => p.clone(),
            None => return,
        };

        // Files linked before conflict checks don't know which revision they
        // match, so any upload would be a blind overwrite. Ask instead.
        if !force && identity.last_synced_revision.is_none() {
            self.cloud_sync_state = CloudSyncState::Conflict;
            self.cloud_last_error = Some("This file was linked before conflict checks.".to_string());
            self.status_message = Some(
                "Cloud sync paused: can't tell whether the cloud copy changed. \
                 Run \"Cloud: Overwrite Cloud Copy\" to upload this file, or File > Open Cloud to take the cloud version."
                    .to_string(),
            );
            cx.notify();
            return;
        }

        self.cloud_upload_generation += 1;
        let generation = self.cloud_upload_generation;
        if self.cloud_sync_state != CloudSyncState::Syncing {
            self.cloud_sync_state = CloudSyncState::Dirty;
        }
        cx.notify();

        cx.spawn(async move |this, cx| {
            // Debounce: wait 500ms so rapid saves coalesce
            smol::Timer::after(std::time::Duration::from_millis(500)).await;

            // Bail if a newer save superseded us, or another upload is still
            // running — that one re-schedules when it sees the newer
            // generation, so the latest content still goes up.
            let start = this.update(cx, |this, cx| {
                if this.cloud_upload_generation != generation || this.cloud_upload_in_flight {
                    return None;
                }
                this.cloud_upload_in_flight = true;
                this.cloud_sync_state = CloudSyncState::Syncing;
                this.status_message = Some("Syncing to cloud...".to_string());
                cx.notify();

                // Read the revision now, not at schedule time: an upload that
                // finished during the debounce has already moved it.
                let expected = if force {
                    None
                } else {
                    this.cloud_identity.as_ref().and_then(|id| id.last_synced_revision)
                };
                let reserved = this.cloud_reserved_revision;

                // Canonical cloud storage is visigrid-json, not the SQLite .sheet
                // on disk. Reading the file would write a format the browser cannot
                // parse, and after any cross-client save the key's extension lies.
                // Persist pivots' stale flags in the synced copy too.
                this.workbook.update(cx, |wb, _| {
                    wb.update_pivot_staleness();
                });
                let wb = this.wb(cx).clone();
                let layouts = this.build_json_sheet_layouts(cx);
                let active = wb.active_sheet_index();
                Some((wb, layouts, active, expected, reserved))
            });
            let (wb, layouts, active, expected, reserved) = match start {
                Ok(Some(s)) => s,
                _ => return,
            };

            let file_bytes = match smol::unblock(move || visigrid_io::json::export_workbook(&wb, &layouts, active)).await {
                Ok(json) => json.into_bytes(),
                Err(e) => {
                    let _ = this.update(cx, |this, cx| {
                        this.cloud_upload_in_flight = false;
                        this.cloud_sync_state = CloudSyncState::Error;
                        this.cloud_last_error = Some(format!("Failed to export workbook: {}", e));
                        this.status_message = Some(format!("Cloud sync error: {}", e));
                        cx.notify();
                    });
                    return;
                }
            };

            let content_hash = hash_bytes(&file_bytes);
            let byte_size = file_bytes.len() as u64;
            let sheet_id = identity.sheet_id;

            let outcome = smol::unblock(move || upload(sheet_id, file_bytes, byte_size, expected, reserved)).await;

            let _ = this.update(cx, |this, cx| {
                this.cloud_upload_in_flight = false;

                // Another file was opened mid-upload; this result isn't about it.
                if this.cloud_identity.as_ref().map(|id| id.sheet_id) != Some(sheet_id) {
                    return;
                }

                match outcome {
                    UploadOutcome::Synced(revision) => {
                        this.cloud_reserved_revision = None;
                        if let Some(ref mut id) = this.cloud_identity {
                            id.last_synced_hash = Some(content_hash);
                            id.last_synced_at = Some(now_iso8601());
                            id.last_synced_revision = Some(revision);

                            // Persist updated identity to .sheet file
                            if let Err(e) = crate::cloud::save_cloud_identity(&path, id) {
                                eprintln!("Warning: failed to persist cloud identity: {}", e);
                            }
                        }
                        this.cloud_last_error = None;
                        this.cloud_sync_state = CloudSyncState::Synced;
                        this.status_message = Some("Synced to cloud".to_string());
                    }
                    UploadOutcome::Conflict { current_revision } => {
                        this.cloud_reserved_revision = None;
                        this.cloud_sync_state = CloudSyncState::Conflict;
                        this.cloud_last_error = Some(match current_revision {
                            Some(rev) => format!("Cloud copy is at revision {}", rev),
                            None => "Cloud copy changed".to_string(),
                        });
                        this.status_message = Some(
                            "The cloud copy changed since this file last synced. Your file is saved locally and was not uploaded. \
                             Run \"Cloud: Overwrite Cloud Copy\" to replace it, or File > Open Cloud to take the cloud version (your copy is kept)."
                                .to_string(),
                        );
                        cx.notify();
                        return;
                    }
                    UploadOutcome::Failed { error, reserved } => {
                        if reserved.is_some() {
                            this.cloud_reserved_revision = reserved;
                        }
                        let offline = matches!(error, HubError::Network(_));
                        let msg = error.to_string();
                        this.cloud_last_error = Some(msg.clone());
                        if is_unauthorized(&error) {
                            this.cloud_sync_state = CloudSyncState::Error;
                            this.cloud_last_error = Some("Sign-in expired".to_string());
                            this.cloud_sign_in_expired(false, cx);
                        } else if offline {
                            this.cloud_sync_state = CloudSyncState::Offline;
                            this.status_message = Some("Cloud sync: offline".to_string());
                        } else {
                            this.cloud_sync_state = CloudSyncState::Error;
                            this.status_message = Some(format!("Cloud sync error: {}", msg));
                        }
                        cx.notify();
                        return;
                    }
                }

                // A save landed while we were uploading: send the newer content.
                if this.cloud_upload_generation != generation {
                    this.cloud_start_upload(false, cx);
                }
                cx.notify();
            });
        }).detach();
    }
}

/// Request a presigned URL, PUT, then confirm so the pointer moves only after
/// the bytes exist. A drop between save() and the PUT used to leave the row
/// pointing at a blob that was never uploaded.
fn upload(
    sheet_id: i64,
    file_bytes: Vec<u8>,
    byte_size: u64,
    expected: Option<i64>,
    reserved: Option<i64>,
) -> UploadOutcome {
    let client = match SheetsClient::from_saved_auth() {
        Ok(c) => c,
        Err(error) => return UploadOutcome::Failed { error, reserved: None },
    };

    let save_resp = match client.save_sheet(sheet_id, byte_size, expected) {
        Ok(r) => r,
        Err(error) => match conflict_revision(&error) {
            // Each save request moves the server revision before the upload
            // completes. If ours failed after that, the server now sits at the
            // revision we reserved; nobody else has written since. Retry
            // against it rather than reporting our own attempt as a conflict.
            Some(current) if reserved == Some(current) => {
                match client.save_sheet(sheet_id, byte_size, Some(current)) {
                    Ok(r) => r,
                    Err(error) => return classify(error, None),
                }
            }
            _ => return classify(error, None),
        },
    };

    let reserved = Some(save_resp.revision);
    if let Err(error) = client.upload_to_url(&save_resp.upload_url, &save_resp.headers, file_bytes) {
        return UploadOutcome::Failed { error, reserved };
    }
    match client.complete_save(sheet_id, &save_resp.blob_key, byte_size) {
        Ok(revision) => UploadOutcome::Synced(revision),
        Err(error) => UploadOutcome::Failed { error, reserved },
    }
}

fn classify(error: HubError, reserved: Option<i64>) -> UploadOutcome {
    match conflict_revision(&error) {
        Some(current) => UploadOutcome::Conflict { current_revision: Some(current) },
        None => UploadOutcome::Failed { error, reserved },
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
