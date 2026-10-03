//! Desktop DuckDB import session. The C connection never leaves its worker
//! thread; catalogue, previews and import share one read-only transaction.
use std::collections::BTreeSet;
use std::path::PathBuf;
use std::sync::{
    atomic::{AtomicBool, AtomicU64, Ordering},
    mpsc, Arc,
};

use gpui::*;
use smol::channel::{bounded, Receiver, Sender};
use visigrid_engine::workbook::Workbook;
use visigrid_io::duckdb::{Database, TableInfo, IMPORT_CELL_LIMIT};

use crate::{app::Spreadsheet, mode::Mode};

pub const PREVIEW_ROWS: usize = 8;
pub const PREVIEW_COLUMNS: usize = 12;

pub struct Entry {
    pub name: String,
    pub info: Option<TableInfo>,
    pub error: Option<String>,
}
impl Entry {
    pub fn problem(&self) -> Option<String> {
        self.error
            .clone()
            .or_else(|| self.info.as_ref().and_then(TableInfo::import_error))
    }
    pub fn importable(&self) -> bool {
        self.info.is_some() && self.problem().is_none()
    }
}

pub struct Preview {
    pub rows: Vec<Vec<String>>,
    pub columns: usize,
}

enum Request {
    Preview {
        index: usize,
        revision: u64,
        reply: Sender<Result<Preview, String>>,
    },
    Import {
        indices: Vec<usize>,
        reply: Sender<Result<Workbook, String>>,
    },
}

struct Worker {
    sender: mpsc::Sender<Request>,
    cancelled: Arc<AtomicBool>,
    preview_revision: Arc<AtomicU64>,
}
impl Drop for Worker {
    fn drop(&mut self) {
        self.cancelled.store(true, Ordering::Relaxed);
    }
}
impl Worker {
    fn start(path: PathBuf) -> (Self, Receiver<Result<Vec<Entry>, String>>) {
        let (sender, requests) = mpsc::channel();
        let (reply, catalog) = bounded(1);
        let cancelled = Arc::new(AtomicBool::new(false));
        let preview_revision = Arc::new(AtomicU64::new(0));
        let stop = cancelled.clone();
        let latest = preview_revision.clone();
        let spawn = std::thread::Builder::new()
            .name("duckdb-import".into())
            .spawn(move || {
                let db = match Database::open(&path) {
                    Ok(db) => db,
                    Err(error) => {
                        let _ = reply.send_blocking(Err(error));
                        return;
                    }
                };
                let mut entries = Vec::new();
                for (index, table) in db.tables().iter().enumerate() {
                    if stop.load(Ordering::Relaxed) {
                        return;
                    }
                    let (info, error) = match db.describe_table(index) {
                        Ok(info) => (Some(info), None),
                        Err(error) => (None, Some(error)),
                    };
                    entries.push(Entry {
                        name: table.name.clone(),
                        info,
                        error,
                    });
                }
                if reply.send_blocking(Ok(entries)).is_err() {
                    return;
                }
                while let Ok(request) = requests.recv() {
                    if stop.load(Ordering::Relaxed) {
                        break;
                    }
                    match request {
                        Request::Preview {
                            index,
                            revision,
                            reply,
                        } => {
                            // Skip previews superseded while an earlier query was running.
                            if revision != latest.load(Ordering::Relaxed) {
                                continue;
                            }
                            let result = db
                                .read_table(
                                    index,
                                    PREVIEW_ROWS,
                                    Some((PREVIEW_ROWS + 1) * visigrid_io::parquet::MAX_COLS),
                                    false,
                                )
                                .map(|data| Preview {
                                    columns: data.cols_loaded.min(PREVIEW_COLUMNS),
                                    rows: (1..=data.rows_loaded)
                                        .map(|r| {
                                            (0..data.cols_loaded.min(PREVIEW_COLUMNS))
                                                .map(|c| data.sheet.get_interchange_display(r, c))
                                                .collect()
                                        })
                                        .collect(),
                                });
                            let _ = reply.send_blocking(result);
                        }
                        Request::Import { indices, reply } => {
                            let _ = reply.send_blocking(db.import_tables(&indices));
                        }
                    }
                }
            });
        if let Err(error) = spawn {
            // The moved reply is dropped on spawn failure, so recv reports it.
            log::error!("Could not start DuckDB reader: {error}");
        }
        (
            Self {
                sender,
                cancelled,
                preview_revision,
            },
            catalog,
        )
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Focus {
    Search,
    Tables,
    All,
    Cancel,
    Import,
}

pub struct ImportDialog {
    pub id: u64,
    workbook_id: EntityId,
    pub path: PathBuf,
    worker: Worker,
    pub entries: Vec<Entry>,
    pub selected: BTreeSet<usize>,
    pub active: Option<usize>,
    pub search: String,
    pub search_selected: bool,
    pub focus: Focus,
    pub loading: bool,
    pub importing: bool,
    pub error: Option<String>,
    pub preview: Option<Preview>,
    pub preview_error: Option<String>,
    pub preview_loading: bool,
    preview_revision: u64,
    pub scroll: ScrollHandle,
}
impl ImportDialog {
    pub fn visible(&self) -> Vec<usize> {
        let query = self.search.to_lowercase();
        self.entries
            .iter()
            .enumerate()
            .filter(|(_, e)| e.name.to_lowercase().contains(&query))
            .map(|(i, _)| i)
            .collect()
    }
    pub fn selection_error(&self) -> Option<String> {
        if self.loading || self.importing {
            return Some("Please wait".into());
        }
        if self.selected.is_empty() {
            return Some("Choose at least one table".into());
        }
        let cells: u64 = self
            .selected
            .iter()
            .filter_map(|i| self.entries[*i].info.as_ref())
            .map(TableInfo::cell_count)
            .fold(0, u64::saturating_add);
        if cells > IMPORT_CELL_LIMIT as u64 {
            Some("Selected tables exceed 10 million cells. Select fewer tables.".into())
        } else {
            None
        }
    }
    pub fn selected_rows(&self) -> u64 {
        self.selected
            .iter()
            .filter_map(|i| self.entries[*i].info.as_ref())
            .map(|t| t.row_count)
            .fold(0, u64::saturating_add)
    }
}

impl Spreadsheet {
    pub fn start_duckdb_import(&mut self, path: &std::path::Path, cx: &mut Context<Self>) {
        self.commit_pending_edit(cx);
        // Open a separate window rather than replacing unsaved edits.
        if self.is_modified {
            crate::open_file_urls(vec![path.to_string_lossy().into_owned()], cx);
            return;
        }
        self.cancel_duckdb_import(cx);
        self.duckdb_import_id = self.duckdb_import_id.wrapping_add(1);
        let id = self.duckdb_import_id;
        let (worker, catalog) = Worker::start(path.to_owned());
        self.close_validation_dropdown(
            crate::validation_dropdown::DropdownCloseReason::ModalOpened,
            cx,
        );
        self.lua_console.visible = false;
        self.open_menu = None;
        self.tab_chain_origin_col = None;
        self.duckdb_import = Some(ImportDialog {
            id,
            workbook_id: self.workbook.entity_id(),
            path: path.to_owned(),
            worker,
            entries: Vec::new(),
            selected: BTreeSet::new(),
            active: None,
            search: String::new(),
            search_selected: false,
            focus: Focus::Tables,
            loading: true,
            importing: false,
            error: None,
            preview: None,
            preview_error: None,
            preview_loading: false,
            preview_revision: 0,
            scroll: ScrollHandle::new(),
        });
        self.mode = Mode::DuckdbImport;
        cx.notify();
        cx.spawn(async move |this, cx| {
            let result = catalog.recv().await.unwrap_or_else(|_| {
                Err("The database reader stopped. Close this dialog and reopen the file.".into())
            });
            let _ = this.update(cx, |this, cx| {
                let Some(dialog) = this.duckdb_import.as_mut().filter(|d| d.id == id) else {
                    return;
                };
                dialog.loading = false;
                match result {
                    Ok(entries) => {
                        let first = entries
                            .iter()
                            .position(|e| {
                                e.importable() && e.info.as_ref().is_some_and(|i| i.row_count > 0)
                            })
                            .or_else(|| entries.iter().position(Entry::importable));
                        if let Some(index) = first {
                            dialog.selected.insert(index);
                        }
                        dialog.entries = entries;
                        this.duckdb_preview(first.unwrap_or(0), cx);
                    }
                    Err(error) => dialog.error = Some(error),
                }
                cx.notify();
            });
        })
        .detach();
    }

    pub fn cancel_duckdb_import(&mut self, cx: &mut Context<Self>) {
        self.duckdb_import = None; // Drops sender; the worker closes its snapshot.
        if self.mode == Mode::DuckdbImport {
            self.mode = Mode::Navigation;
        }
        cx.notify();
    }

    pub fn duckdb_preview(&mut self, index: usize, cx: &mut Context<Self>) {
        let Some(dialog) = self.duckdb_import.as_mut() else {
            return;
        };
        if dialog.importing || index >= dialog.entries.len() {
            return;
        }
        dialog.active = Some(index);
        dialog.preview = None;
        dialog.preview_error = None;
        dialog.preview_loading = true;
        dialog.preview_revision = dialog.preview_revision.wrapping_add(1);
        let revision = dialog.preview_revision;
        dialog
            .worker
            .preview_revision
            .store(revision, Ordering::Relaxed);
        let id = dialog.id;
        let (reply, result) = bounded(1);
        if dialog
            .worker
            .sender
            .send(Request::Preview {
                index,
                revision,
                reply,
            })
            .is_err()
        {
            dialog.preview_loading = false;
            dialog.preview_error = Some("The database reader stopped. Reopen the file.".into());
            cx.notify();
            return;
        }
        cx.spawn(async move |this, cx| {
            let result = result
                .recv()
                .await
                .unwrap_or_else(|_| Err("The database reader stopped. Reopen the file.".into()));
            let _ = this.update(cx, |this, cx| {
                let Some(dialog) = this
                    .duckdb_import
                    .as_mut()
                    .filter(|d| d.id == id && d.preview_revision == revision)
                else {
                    return;
                };
                dialog.preview_loading = false;
                match result {
                    Ok(preview) => dialog.preview = Some(preview),
                    Err(error) => dialog.preview_error = Some(error),
                }
                cx.notify();
            });
        })
        .detach();
        cx.notify();
    }

    pub fn duckdb_toggle(&mut self, index: usize, cx: &mut Context<Self>) {
        if let Some(d) = self.duckdb_import.as_mut() {
            if !d.importing && d.entries.get(index).is_some_and(Entry::importable) {
                if !d.selected.remove(&index) {
                    d.selected.insert(index);
                }
                d.error = None;
            }
        }
        cx.notify();
    }

    pub fn duckdb_toggle_all(&mut self, cx: &mut Context<Self>) {
        if let Some(d) = self.duckdb_import.as_mut().filter(|d| !d.importing) {
            let all: BTreeSet<_> = d
                .entries
                .iter()
                .enumerate()
                .filter(|(_, e)| e.importable())
                .map(|(i, _)| i)
                .collect();
            d.selected = if d.selected == all {
                BTreeSet::new()
            } else {
                all
            };
            d.error = None;
        }
        cx.notify();
    }

    pub fn duckdb_confirm_import(&mut self, cx: &mut Context<Self>) {
        let Some(d) = self.duckdb_import.as_mut() else {
            return;
        };
        if d.selection_error().is_some() {
            return;
        }
        if self.is_modified || self.workbook.entity_id() != d.workbook_id {
            d.error =
                Some("The current workbook changed. Cancel and save it before importing.".into());
            cx.notify();
            return;
        }
        let id = d.id;
        let path = d.path.clone();
        let count = d.selected.len();
        let rows = d.selected_rows();
        let (reply, result) = bounded(1);
        if d.worker
            .sender
            .send(Request::Import {
                indices: d.selected.iter().copied().collect(),
                reply,
            })
            .is_err()
        {
            d.error = Some("The database reader stopped. Reopen the file.".into());
            cx.notify();
            return;
        }
        d.importing = true;
        d.error = None;
        cx.notify();
        cx.spawn(async move |this, cx| {
            let result = result.recv().await.unwrap_or_else(|_| Err("The database reader stopped. Reopen the file.".into()));
            let _ = this.update(cx, |this, cx| {
                let Some(d) = this.duckdb_import.as_mut().filter(|d| d.id == id) else { return; };
                d.importing = false;
                match result {
                    Ok(workbook) if !this.is_modified && this.workbook.entity_id() == d.workbook_id => {
                        this.cancel_duckdb_import(cx);
                        this.finish_table_import(workbook, &path, cx);
                        this.status_message = Some(format!("Imported {count} table{} ({} rows). Edits stay in this workbook; the database is unchanged.", if count == 1 { "" } else { "s" }, number(rows)));
                    }
                    Ok(_) => d.error = Some("The workbook changed during import. Cancel and save it before trying again.".into()),
                    Err(error) => d.error = Some(error),
                }
                cx.notify();
            });
        }).detach();
    }

    /// Capture modal keys before Spreadsheet bindings (including destructive
    /// shortcuts), scoped to this window. No grid actions leak through.
    pub(crate) fn intercept_duckdb_keys(
        window: &mut Window,
        cx: &mut Context<Self>,
    ) -> Subscription {
        let this = cx.entity().downgrade();
        let handle = window.window_handle();
        cx.intercept_keystrokes(move |event, window, cx| {
            if window.window_handle() != handle {
                return;
            }
            let Some(this) = this.upgrade() else {
                return;
            };
            let handled = this.update(cx, |this, cx| {
                if this.mode != Mode::DuckdbImport {
                    return false;
                }
                this.duckdb_key(&event.keystroke, cx);
                true
            });
            if handled {
                cx.stop_propagation();
            }
        })
    }

    fn duckdb_key(&mut self, key: &Keystroke, cx: &mut Context<Self>) {
        if key.key == "escape" {
            self.cancel_duckdb_import(cx);
            return;
        }
        let Some(d) = self.duckdb_import.as_mut() else {
            return;
        };
        if d.importing {
            return;
        }
        let command = key.modifiers.control || key.modifiers.platform;
        if key.key == "tab" {
            let order = [
                Focus::Search,
                Focus::All,
                Focus::Tables,
                Focus::Cancel,
                Focus::Import,
            ];
            let i = order.iter().position(|f| *f == d.focus).unwrap_or(0);
            d.focus = order[(i + if key.modifiers.shift { 4 } else { 1 }) % 5];
        } else if command && key.key == "f" {
            d.focus = Focus::Search;
        } else if command && key.key == "enter" {
            self.duckdb_confirm_import(cx);
        } else if key.key == "enter" || (key.key == "space" && d.focus != Focus::Search) {
            match d.focus {
                Focus::Cancel => self.cancel_duckdb_import(cx),
                Focus::Import => self.duckdb_confirm_import(cx),
                Focus::All => self.duckdb_toggle_all(cx),
                Focus::Tables => {
                    if let Some(index) = d.active {
                        self.duckdb_toggle(index, cx);
                    }
                }
                Focus::Search => d.focus = Focus::Tables,
            }
        } else if matches!(key.key.as_str(), "up" | "down" | "home" | "end") && !command {
            let indices = d.visible();
            if !indices.is_empty() {
                let current = indices
                    .iter()
                    .position(|i| Some(*i) == d.active)
                    .unwrap_or(0);
                let next = match key.key.as_str() {
                    "up" => current.saturating_sub(1),
                    "down" => (current + 1).min(indices.len() - 1),
                    "home" => 0,
                    _ => indices.len() - 1,
                };
                d.focus = Focus::Tables;
                d.scroll.scroll_to_item(next);
                self.duckdb_preview(indices[next], cx);
            }
        } else if d.focus == Focus::Search {
            if command && key.key == "a" {
                crate::ui::text_input::handle_input_select_all(&mut d.search_selected);
            } else if command && key.key == "v" {
                if let Some(text) = cx.read_from_clipboard().and_then(|item| item.text()) {
                    crate::ui::text_input::handle_input_paste(
                        &mut d.search,
                        &mut d.search_selected,
                        &text,
                    );
                }
            } else {
                crate::ui::text_input::handle_input_key(
                    &mut d.search,
                    &mut d.search_selected,
                    key.key.as_str(),
                    key.key_char.as_deref(),
                    command || key.modifiers.alt,
                );
            }
            d.search = d.search.chars().take(256).collect();
            let visible = d.visible();
            if !d.active.is_some_and(|index| visible.contains(&index)) {
                d.active = None;
                d.preview = None;
                d.preview_error = None;
                d.preview_loading = false;
                d.preview_revision = d.preview_revision.wrapping_add(1);
                d.worker
                    .preview_revision
                    .store(d.preview_revision, Ordering::Relaxed);
                if let Some(&index) = visible.first() {
                    self.duckdb_preview(index, cx);
                }
            }
        }
        cx.notify();
    }
}

pub fn number(value: u64) -> String {
    let digits = value.to_string();
    let mut out = String::new();
    for (i, c) in digits.chars().enumerate() {
        if i > 0 && (digits.len() - i) % 3 == 0 {
            out.push(',');
        }
        out.push(c);
    }
    out
}
