//! `vgrid serve` — headless session host (detach phase 1).
//!
//! Engine + session server, no window: every protocol client (`vgrid
//! apply/inspect/view`, `vgrid mcp`, agents) works against it unchanged.
//! The workbook loads through visigrid-io, requests are pumped on a plain
//! loop, and paired credentials work exactly as with the GUI (the store is
//! machine-wide). Pairing requests prompt y/N on this terminal.

use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::time::{Duration, Instant};

use visigrid_engine::workbook::Workbook;
use visigrid_io::json::SheetLayout;
use visigrid_session_host as host;

use crate::CliError;

/// (at, count, delete, is_row) for row/column ops; None for sheet ops.
fn structural_span(op: &visigrid_protocol::StructureOp) -> Option<(usize, usize, bool, bool)> {
    use visigrid_protocol::StructureOp as S;
    match op {
        S::InsertRows { at, count, .. } => Some((*at, *count, false, true)),
        S::DeleteRows { at, count, .. } => Some((*at, *count, true, true)),
        S::InsertCols { at, count, .. } => Some((*at, *count, false, false)),
        S::DeleteCols { at, count, .. } => Some((*at, *count, true, false)),
        S::AddSheet { .. } | S::RenameSheet { .. } => None,
    }
}

/// What `serve` persists to, if anything.
enum SaveTarget {
    Native(PathBuf),
    JsonFull(PathBuf),
    /// Read-only session (xlsx/csv input without --save-as).
    None,
}

impl SaveTarget {
    fn path(&self) -> Option<&PathBuf> {
        match self {
            SaveTarget::Native(p) | SaveTarget::JsonFull(p) => Some(p),
            SaveTarget::None => None,
        }
    }
}

fn target_for(path: &PathBuf) -> Result<SaveTarget, CliError> {
    match path.extension().and_then(|e| e.to_str()).map(|e| e.to_lowercase()).as_deref() {
        Some("sheet") | Some("vgrid") => Ok(SaveTarget::Native(path.clone())),
        Some("json") => Ok(SaveTarget::JsonFull(path.clone())),
        Some(other) => Err(CliError::args(format!(
            "cannot save to .{} — use --save-as with a .sheet or .json path", other
        ))),
        None => Err(CliError::args("save path needs a .sheet or .json extension")),
    }
}

pub fn cmd_serve(
    file: Option<PathBuf>,
    new: bool,
    save_as: Option<PathBuf>,
    autosave: Option<u64>,
) -> Result<(), CliError> {
    // ---- Load the workbook -------------------------------------------------
    let (mut wb, mut layouts, title): (Workbook, Vec<SheetLayout>, String) = match (&file, new) {
        (Some(path), _) => {
            if !path.exists() {
                return Err(CliError::io(format!("{} not found", path.display())));
            }
            let ext = path.extension().and_then(|e| e.to_str()).map(|e| e.to_lowercase());
            let title = path.file_name().map(|n| n.to_string_lossy().to_string())
                .unwrap_or_else(|| "workbook".to_string());
            match ext.as_deref() {
                Some("sheet") | Some("vgrid") => {
                    let wb = visigrid_io::native::load_workbook(path).map_err(CliError::io)?;
                    let n = wb.sheets().len();
                    (wb, vec![SheetLayout::default(); n], title)
                }
                Some("json") => {
                    let content = std::fs::read_to_string(path).map_err(|e| CliError::io(e.to_string()))?;
                    let (wb, layouts, _) = visigrid_io::json::import_any(&content).map_err(CliError::io)?;
                    (wb, layouts, title)
                }
                Some("xlsx") | Some("xls") | Some("xlsb") | Some("ods") => {
                    let (wb, _) = visigrid_io::xlsx::import(path).map_err(CliError::parse)?;
                    let n = wb.sheets().len();
                    (wb, vec![SheetLayout::default(); n], title)
                }
                _ => return Err(CliError::args("serve supports .sheet, .json (visigrid-json), and .xlsx inputs")
                    .with_hint("or start empty with: vgrid serve --new --save-as out.sheet")),
            }
        }
        (None, true) => (Workbook::new(), vec![SheetLayout::default()], "Untitled".to_string()),
        (None, false) => {
            return Err(CliError::args("provide a file to serve, or --new for an empty workbook")
                .with_hint("vgrid serve budget.sheet"));
        }
    };
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();

    // ---- Save target -------------------------------------------------------
    let save_target = match (&save_as, &file) {
        (Some(p), _) => target_for(p)?,
        (None, Some(p)) => target_for(p).unwrap_or(SaveTarget::None),
        (None, None) => SaveTarget::None,
    };
    if matches!(save_target, SaveTarget::None) {
        eprintln!("note: read-only session (no .sheet/.json save target) — edits are not persisted; pass --save-as to keep them");
    }

    // ---- Session server -----------------------------------------------------
    let (tx, rx) = std::sync::mpsc::channel::<host::SessionRequest>();
    let bridge = host::SessionBridgeHandle::new(tx);
    let mut server = host::SessionServer::new();
    server
        .start(host::SessionServerConfig {
            mode: host::ServerMode::Apply,
            workbook_path: file.clone(),
            workbook_title: title.clone(),
            bridge: Some(bridge),
            token_override: std::env::var("VISIGRID_SESSION_TOKEN").ok(),
            ..Default::default()
        })
        .map_err(|e| CliError::io(format!("failed to start session server: {}", e)))?;

    if let Some((session_id, port, discovery)) = server.ready_info() {
        eprintln!("READY session_id={} port={} discovery={}", session_id, port, discovery.display());
        eprintln!(
            "serving \"{}\" headless — agents and the CLI connect as usual (vgrid mcp, vgrid apply/inspect); Ctrl+C to stop",
            title
        );
    }

    let stop = Arc::new(AtomicBool::new(false));
    {
        let stop = stop.clone();
        let _ = ctrlc::set_handler(move || stop.store(true, Ordering::SeqCst));
    }

    // ---- Request loop -------------------------------------------------------
    let mut dirty = false;
    let mut last_save = Instant::now();

    let save = |wb: &Workbook, layouts: &[SheetLayout]| -> Result<Option<String>, String> {
        match &save_target {
            SaveTarget::Native(p) => {
                visigrid_io::native::save_workbook(wb, p).map_err(|e| e.to_string())?;
                Ok(Some(p.display().to_string()))
            }
            SaveTarget::JsonFull(p) => {
                let json = visigrid_io::json::export_workbook(wb, layouts, wb.active_sheet_index())?;
                std::fs::write(p, json).map_err(|e| e.to_string())?;
                Ok(Some(p.display().to_string()))
            }
            SaveTarget::None => Err("read-only session: no save target (start with --save-as)".to_string()),
        }
    };

    loop {
        if stop.load(Ordering::SeqCst) {
            break;
        }
        // Autosave
        if dirty {
            if let Some(secs) = autosave {
                if last_save.elapsed() >= Duration::from_secs(secs) && save_target.path().is_some() {
                    match save(&wb, &layouts) {
                        Ok(_) => {
                            dirty = false;
                            last_save = Instant::now();
                        }
                        Err(e) => eprintln!("autosave failed: {}", e),
                    }
                }
            }
        }

        let req = match rx.recv_timeout(Duration::from_millis(300)) {
            Ok(r) => r,
            Err(std::sync::mpsc::RecvTimeoutError::Timeout) => continue,
            Err(std::sync::mpsc::RecvTimeoutError::Disconnected) => break,
        };

        match req {
            host::SessionRequest::ApplyOps { req, reply } => {
                let outcome = host::apply_ops(&mut wb, &req);
                if outcome.response.error.is_none() && outcome.response.applied > 0 {
                    dirty = true;
                    server.broadcast_cells(outcome.response.current_revision, outcome.changed_cells);
                }
                let _ = reply.send(outcome.response);
            }
            host::SessionRequest::Inspect { req, reply } => {
                let _ = reply.send(host::inspect(&wb, &req, &title));
            }
            host::SessionRequest::Subscribe { req, reply } => {
                let _ = reply.send(host::SubscribeResponse {
                    topics: req.topics,
                    current_revision: wb.revision(),
                });
            }
            host::SessionRequest::Unsubscribe { req, reply } => {
                let _ = reply.send(host::UnsubscribeResponse { topics: req.topics });
            }
            host::SessionRequest::Pair { client_name, reply } => {
                eprintln!();
                eprintln!("Pairing request from \"{}\" — allow it to control this workbook? [y/N]", client_name);
                let mut line = String::new();
                let approved = std::io::stdin().read_line(&mut line).is_ok()
                    && matches!(line.trim().to_ascii_lowercase().as_str(), "y" | "yes");
                eprintln!("{}", if approved { "approved" } else { "denied" });
                let _ = reply.send(approved);
            }
            host::SessionRequest::Structure { op, reply, .. } => {
                let outcome = match host::validate_structure_op(&op, &wb) {
                    Some((code, message, suggestion)) => host::StructureOutcome {
                        revision: wb.revision(),
                        sheet_count: wb.sheets().len(),
                        active_sheet: wb.active_sheet_index(),
                        error: Some((
                            code.to_string(),
                            match suggestion {
                                Some(s) => format!("{} — {}", message, s),
                                None => message,
                            },
                        )),
                        ..Default::default()
                    },
                    None => match host::apply_structure(&mut wb, &op) {
                        Err(msg) => host::StructureOutcome {
                            revision: wb.revision(),
                            sheet_count: wb.sheets().len(),
                            active_sheet: wb.active_sheet_index(),
                            error: Some(("invalid_op".to_string(), msg)),
                            ..Default::default()
                        },
                        Ok(description) => {
                        dirty = true;
                        // Layout side-cars are keyed by index, so they must
                        // follow the edit — widths, frozen panes, filters, and
                        // chart ranges all describe rows that just moved.
                        let target = host::structure_target_sheet(&op, wb.active_sheet_index());
                        if let Some(l) = layouts.get_mut(target) {
                            if let Some((at, count, delete, is_row)) = structural_span(&op) {
                                l.shift_for_structural(at, count, delete, is_row);
                            }
                        }
                        while layouts.len() < wb.sheets().len() {
                            layouts.push(SheetLayout::default());
                        }
                        host::StructureOutcome {
                            description,
                            revision: wb.revision(),
                            sheet_count: wb.sheets().len(),
                            active_sheet: wb.active_sheet_index(),
                            error: None,
                        }
                        }
                    },
                };
                let _ = reply.send(outcome);
            }
            host::SessionRequest::History { reply, .. } => {
                // No undo stack headless (history is GUI state) — documented
                // in the detach design doc as a phase-1 limitation.
                let _ = reply.send(host::HistoryOutcome {
                    revision: wb.revision(),
                    error: Some((
                        "history_unavailable".to_string(),
                        "this is a headless session (vgrid serve) — it has no undo stack; attach a GUI window for undo/redo".to_string(),
                    )),
                    ..Default::default()
                });
            }
            host::SessionRequest::Save { reply, .. } => {
                let outcome = match save(&wb, &layouts) {
                    Ok(path) => {
                        dirty = false;
                        last_save = Instant::now();
                        host::SaveOutcome { path, revision: wb.revision(), error: None }
                    }
                    Err(msg) => host::SaveOutcome {
                        path: None,
                        revision: wb.revision(),
                        error: Some(("save_unsupported".to_string(), msg)),
                    },
                };
                let _ = reply.send(outcome);
            }
        }

        // Layouts don't change headless (no layout-mutating ops yet), but
        // keep the binding mutable for when they do.
        let _ = &mut layouts;
    }

    // ---- Shutdown -----------------------------------------------------------
    if dirty && save_target.path().is_some() {
        match save(&wb, &layouts) {
            Ok(Some(p)) => eprintln!("saved {}", p),
            Ok(None) => {}
            Err(e) => eprintln!("save on exit failed: {} (edits lost)", e),
        }
    }
    server.stop();
    eprintln!("session closed");
    Ok(())
}
