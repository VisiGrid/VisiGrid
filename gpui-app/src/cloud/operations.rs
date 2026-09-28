// Cloud operations — Open Cloud, Move to Cloud.
//
// These operations interact with the VisiHub Sheets API to create,
// list, and download cloud-backed sheets.

use crate::app::Spreadsheet;
use crate::cloud::{CloudIdentity, CloudSyncState};
use crate::cloud::sheets_client::SheetsClient;
use crate::hub::client::HubError;
use visigrid_io::json::{self, CloudBlobKind};
use visigrid_io::native;

impl Spreadsheet {
    /// Move the current local file to cloud.
    /// Creates a sheet on the server, attaches a CloudIdentity, and triggers initial upload.
    pub fn cloud_move_to_cloud(&mut self, cx: &mut gpui::Context<Self>) {
        let path = match &self.current_file {
            Some(p) => p.clone(),
            None => {
                self.status_message = Some("Save the file first before moving to cloud.".to_string());
                cx.notify();
                return;
            }
        };

        if self.cloud_identity.is_some() {
            self.status_message = Some("This file is already cloud-backed.".to_string());
            cx.notify();
            return;
        }

        // Use the file name (without extension) as the sheet name
        let sheet_name = path.file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("Untitled Sheet")
            .to_string();

        self.status_message = Some("Moving to cloud...".to_string());
        cx.notify();

        let name = sheet_name.clone();
        cx.spawn(async move |this, cx| {
            let result = smol::unblock(move || {
                let client = SheetsClient::from_saved_auth()?;
                client.create_sheet(&name)
            }).await;

            match result {
                Ok(sheet_info) => {
                    let synced_path = path.clone();
                    let _ = this.update(cx, |this, cx| {
                        let identity = CloudIdentity {
                            sheet_id: sheet_info.id,
                            public_id: sheet_info.public_id,
                            sheet_name: sheet_info.name,
                            api_base: crate::hub::auth::load_auth()
                                .map(|a| a.api_base)
                                .unwrap_or_else(|| "https://api.visiapi.com".to_string()),
                            last_synced_hash: None,
                            last_synced_at: None,
                            // A new sheet: the first upload is expected to land on it.
                            last_synced_revision: Some(sheet_info.revision.unwrap_or(0)),
                        };

                        // Persist identity to the .sheet file
                        if let Err(e) = crate::cloud::save_cloud_identity(&synced_path, &identity) {
                            eprintln!("Warning: failed to persist cloud identity: {}", e);
                        }

                        this.cloud_identity = Some(identity);
                        this.cloud_sync_state = CloudSyncState::Dirty;
                        this.status_message = Some("Moved to cloud. Syncing...".to_string());
                        cx.notify();

                        // Trigger initial upload
                        this.cloud_schedule_upload(cx);
                    });
                }
                Err(HubError::NotAuthenticated) => {
                    let _ = this.update(cx, |this, cx| {
                        this.status_message = Some("Sign in first to move to cloud.".to_string());
                        cx.notify();
                    });
                }
                Err(e) => {
                    let msg = e.to_string();
                    let _ = this.update(cx, |this, cx| {
                        this.status_message = Some(format!("Failed to move to cloud: {}", msg));
                        cx.notify();
                    });
                }
            }
        }).detach();
    }

    /// Open the cloud sheet picker: fetches sheet list, stores it, and shows the dialog.
    pub fn cloud_open(&mut self, cx: &mut gpui::Context<Self>) {
        self.cloud_sheets_loading = true;
        self.cloud_sheets_list = Vec::new();
        self.cloud_selected_sheet = None;
        cx.notify();

        cx.spawn(async move |this, cx| {
            let result = smol::unblock(|| {
                let client = SheetsClient::from_saved_auth()?;
                client.list_sheets()
            }).await;

            let _ = this.update(cx, |this, cx| {
                this.cloud_sheets_loading = false;
                match result {
                    Ok(sheets) => {
                        this.cloud_sheets_list = sheets;
                        this.cloud_selected_sheet = if this.cloud_sheets_list.is_empty() { None } else { Some(0) };
                        // Switch to GoTo mode to show the picker (reusing the dialog pattern)
                        this.mode = crate::mode::Mode::CloudOpen;
                    }
                    Err(HubError::NotAuthenticated) => {
                        this.status_message = Some("Sign in first to open cloud sheets.".to_string());
                    }
                    Err(e) => {
                        this.status_message = Some(format!("Failed to list cloud sheets: {}", e));
                    }
                }
                cx.notify();
            });
        }).detach();
    }

    /// Download and open the selected cloud sheet.
    pub fn cloud_open_selected(&mut self, cx: &mut gpui::Context<Self>) {
        let selected = match self.cloud_selected_sheet {
            Some(idx) if idx < self.cloud_sheets_list.len() => self.cloud_sheets_list[idx].clone(),
            _ => return,
        };

        self.mode = crate::mode::Mode::Navigation;
        self.status_message = Some(format!("Downloading {}...", selected.name));
        cx.notify();

        let sheet_id = selected.id;
        let public_id = selected.public_id.clone();
        let sheet_name = selected.name.clone();
        let slug = selected.slug.clone();

        cx.spawn(async move |this, cx| {
            let result: Result<(Option<Vec<u8>>, Option<i64>), HubError> = smol::unblock(move || {
                let client = SheetsClient::from_saved_auth()?;
                // Read the revision BEFORE the bytes. If a save lands in
                // between, we hold newer bytes under an older revision and the
                // first upload conflicts — safe. The other order would record
                // a revision newer than the bytes and overwrite that save.
                let revision = client.get_sheet(sheet_id)?.revision;
                let url = client.get_data_url(sheet_id)?;
                match url {
                    Some(download_url) => {
                        let bytes = client.download_from_url(&download_url)?;
                        Ok((Some(bytes), revision))
                    }
                    None => Ok((None, revision)), // New sheet with no data yet
                }
            }).await;

            match result {
                Ok((maybe_bytes, revision)) => {
                    // Write to cloud cache directory
                    let cache_dir = cloud_cache_dir();
                    let _ = smol::unblock(move || std::fs::create_dir_all(&cache_dir)).await;

                    let file_path = cloud_cache_dir().join(format!("{}.sheet", slug));

                    if let Some(bytes) = maybe_bytes {
                        let fp = file_path.clone();
                        if let Err(e) = smol::unblock(move || materialize_cloud_blob(&fp, &bytes)).await {
                            let _ = this.update(cx, |this, cx| {
                                this.status_message = Some(format!("Failed to write file: {}", e));
                                cx.notify();
                            });
                            return;
                        }
                    }

                    let _ = this.update(cx, |this, cx| {
                        // Load the downloaded file
                        this.load_file(&file_path, cx);

                        // Attach cloud identity
                        let identity = CloudIdentity {
                            sheet_id,
                            public_id,
                            sheet_name,
                            api_base: crate::hub::auth::load_auth()
                                .map(|a| a.api_base)
                                .unwrap_or_else(|| "https://api.visiapi.com".to_string()),
                            last_synced_hash: None,
                            last_synced_at: None,
                            last_synced_revision: revision,
                        };

                        if let Err(e) = crate::cloud::save_cloud_identity(&file_path, &identity) {
                            eprintln!("Warning: failed to persist cloud identity: {}", e);
                        }

                        this.cloud_identity = Some(identity);
                        this.cloud_sync_state = CloudSyncState::Synced;
                        cx.notify();
                    });
                }
                Err(e) => {
                    let msg = e.to_string();
                    let _ = this.update(cx, |this, cx| {
                        this.status_message = Some(format!("Failed to download sheet: {}", msg));
                        cx.notify();
                    });
                }
            }
        }).detach();
    }
}

/// Write a cloud blob to a local `.sheet` file, converting visigrid-json
/// when that's what arrived. The key's extension is not consulted — after a
/// cross-client save it lies.
fn materialize_cloud_blob(path: &std::path::Path, bytes: &[u8]) -> Result<(), String> {
    match json::sniff_cloud_blob(bytes) {
        CloudBlobKind::NativeSqlite => {
            std::fs::write(path, bytes).map_err(|e| e.to_string())
        }
        CloudBlobKind::VisigridJson => {
            let text = std::str::from_utf8(bytes).map_err(|e| e.to_string())?;
            let (wb, layouts, _) = json::import_any(text)?;
            native::save_workbook(&wb, path)?;
            native::save_layout(path, &json_layouts_to_native(&layouts))?;
            Ok(())
        }
        CloudBlobKind::Unknown => {
            Err("Cloud blob is neither visigrid-json nor a native .sheet file".to_string())
        }
    }
}

fn json_layouts_to_native(layouts: &[json::SheetLayout]) -> native::SheetLayout {
    let mut layout = native::SheetLayout {
        col_widths: std::collections::HashMap::new(),
        row_heights: std::collections::HashMap::new(),
        hidden_rows: std::collections::HashMap::new(),
        hidden_cols: std::collections::HashMap::new(),
    };
    for (idx, src) in layouts.iter().enumerate() {
        if !src.col_widths.is_empty() {
            layout
                .col_widths
                .insert(idx, src.col_widths.iter().map(|(&k, &v)| (k, v)).collect());
        }
        if !src.row_heights.is_empty() {
            layout
                .row_heights
                .insert(idx, src.row_heights.iter().map(|(&k, &v)| (k, v)).collect());
        }
        if !src.hidden_rows.is_empty() {
            layout
                .hidden_rows
                .insert(idx, src.hidden_rows.iter().copied().collect());
        }
        if !src.hidden_cols.is_empty() {
            layout
                .hidden_cols
                .insert(idx, src.hidden_cols.iter().copied().collect());
        }
    }
    layout
}

/// Path to the local cloud sheet cache directory.
fn cloud_cache_dir() -> std::path::PathBuf {
    dirs::data_dir()
        .unwrap_or_else(|| std::path::PathBuf::from("."))
        .join("visigrid")
        .join("cloud")
}
