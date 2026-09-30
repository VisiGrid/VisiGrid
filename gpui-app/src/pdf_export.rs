//! Active-view capture and asynchronous PDF export. Export never changes save state.
use crate::{app::Spreadsheet, mode::Mode};
use gpui::{App, Context};
use std::sync::{
    atomic::{AtomicBool, Ordering},
    Arc,
};
use visigrid_print::snapshot::{self, Scope, SheetView, Snapshot};
use visigrid_print::{paginate, AxisItem, PageSettings, GRID_UNIT_PT};

pub struct PdfExportState {
    pub settings: PageSettings,
    pub selection_only: bool,
    pub snapshot: Result<Arc<Snapshot>, String>,
    pub summary: Result<String, String>,
    pub busy: bool,
    pub cancel: Arc<AtomicBool>,
    pub focus: usize,
    pub error: Option<String>,
    pub report: Option<String>,
    pub saved_path: Option<std::path::PathBuf>,
}

impl PdfExportState {
    pub fn update_summary(&mut self) {
        self.summary = self.snapshot.as_ref().map_err(Clone::clone).and_then(|s| {
            let plan = paginate(&s.layout, &self.settings).map_err(|e| e.to_string())?;
            let mut summary = format!("{} • {} page{} • {:.0}% scale", s.name, plan.pages().len(), if plan.pages().len() == 1 { "" } else { "s" }, plan.scale()*100.0);
            if let Some(size) = plan.readability().smallest_text_pt {
                summary.push_str(&format!(" • smallest text {:.1} pt", size));
            }
            if plan.readability().cells_below_threshold > 0 {
                summary.push_str("\nSome text is below 8 pt. Try landscape, actual size, or a narrower selection.");
            }
            Ok(summary)
        });
    }
}

impl Spreadsheet {
    pub fn show_pdf_export(&mut self, cx: &mut Context<Self>) {
        if self.pdf_export.as_ref().is_some_and(|s| s.busy) {
            return;
        }
        if self.is_previewing() || self.review_mode.is_some() {
            self.status_message =
                Some("Return to the current sheet before exporting a PDF.".into());
            cx.notify();
            return;
        }
        if self.mode.is_editing() && !self.commit_current_edit(cx) {
            return;
        }
        let snapshot = self.capture_pdf(false, cx).map(Arc::new);
        let mut state = PdfExportState {
            settings: PageSettings::default(),
            selection_only: false,
            snapshot,
            summary: Ok(String::new()),
            busy: false,
            cancel: Arc::new(AtomicBool::new(false)),
            focus: 0,
            error: None,
            report: None,
            saved_path: None,
        };
        state.update_summary();
        self.pdf_export = Some(state);
        self.mode = Mode::ExportPdf;
        cx.notify();
    }

    fn capture_pdf(&self, selection_only: bool, cx: &App) -> Result<Snapshot, String> {
        let ((min_r, min_c), (max_r, max_c)) = self.selection_range();
        if selection_only && !self.active_view_state().additional_selections.is_empty() {
            return Err("Select one rectangular range to export. Multiple selections are not supported yet.".into());
        }
        let mut selected_rows = Vec::new();
        let rows = self
            .row_view
            .visible_rows()
            .iter()
            .filter_map(|&view_row| {
                let data_row = self.row_view.view_to_data(view_row);
                if self.is_row_hidden(data_row) {
                    return None;
                }
                selected_rows.push(view_row >= min_r && view_row <= max_r);
                Some(AxisItem {
                    source_index: data_row,
                    size_pt: f64::from(self.row_height(view_row)) * GRID_UNIT_PT,
                })
            })
            .collect::<Vec<_>>();
        let columns = (0..self.sheet(cx).cols)
            .filter(|&c| !self.is_col_hidden(c))
            .map(|c| AxisItem {
                source_index: c,
                size_pt: f64::from(self.col_width(c)) * GRID_UNIT_PT,
            })
            .collect::<Vec<_>>();
        let selection = if selection_only {
            let first_r = selected_rows
                .iter()
                .position(|&yes| yes)
                .ok_or("The selection has no visible rows.")?;
            let last_r = selected_rows.iter().rposition(|&yes| yes).unwrap();
            let first_c = columns
                .iter()
                .position(|c| c.source_index >= min_c && c.source_index <= max_c)
                .ok_or("The selection has no visible columns.")?;
            let last_c = columns
                .iter()
                .rposition(|c| c.source_index >= min_c && c.source_index <= max_c)
                .unwrap();
            Some(Scope {
                rows: first_r..last_r + 1,
                columns: first_c..last_c + 1,
            })
        } else {
            None
        };
        let mut snapshot = snapshot::capture(
            self.sheet(cx),
            &SheetView { rows, columns },
            selection,
            &self.cell_font.family,
            self.cell_font.size,
        )?;
        // Apply presentation-only agent roles, as the grid does.
        for cell in &mut snapshot.cells {
            if let Some(role) = self.get_cell_role_style(cell.source.0, cell.source.1) {
                let f = &mut cell.format;
                if let Some(bg) = role.background {
                    f.background_color.get_or_insert(rgba(bg));
                }
                if let Some(fg) = role.text_color {
                    f.font_color.get_or_insert(rgba(fg));
                }
                f.bold |= role.bold.unwrap_or(false);
                f.italic |= role.italic.unwrap_or(false);
                if role.align_right {
                    f.alignment = visigrid_engine::cell::Alignment::Right;
                } else if role.align_center {
                    f.alignment = visigrid_engine::cell::Alignment::Center;
                }
                if role.border_top && !f.border_top.is_set() {
                    f.border_top = visigrid_engine::cell::CellBorder::thin();
                }
                if role.border_bottom && !f.border_bottom.is_set() {
                    f.border_bottom = visigrid_engine::cell::CellBorder::thin();
                }
                if let Some(format) = role.number_format {
                    if let Ok(n) = cell.text.parse::<f64>() {
                        cell.text = crate::role_styles::format_number_for_display(n, format);
                    }
                }
            }
        }
        Ok(snapshot)
    }

    pub fn pdf_cycle_option(&mut self, option: usize, cx: &mut Context<Self>) {
        if self.pdf_export.as_ref().is_none_or(|s| s.busy) {
            return;
        }
        if option == 0 {
            let selection_only = !self.pdf_export.as_ref().unwrap().selection_only;
            let snapshot = self.capture_pdf(selection_only, cx).map(Arc::new);
            let state = self.pdf_export.as_mut().unwrap();
            state.selection_only = selection_only;
            state.snapshot = snapshot;
        } else {
            use visigrid_print::{Paper, Scale};
            let state = self.pdf_export.as_mut().unwrap();
            match option {
                1 => {
                    state.settings.paper = match state.settings.paper {
                        Paper::A4 => Paper::Letter,
                        Paper::Letter => Paper::Legal,
                        Paper::Legal => Paper::A4,
                    }
                }
                2 => state.settings.landscape = !state.settings.landscape,
                3 => {
                    state.settings.scale = match state.settings.scale {
                        Scale::FitColumns => Scale::Actual,
                        Scale::Actual => Scale::FitSheet,
                        _ => Scale::FitColumns,
                    }
                }
                4 => state.settings.footer = !state.settings.footer,
                _ => {}
            }
        }
        let state = self.pdf_export.as_mut().unwrap();
        state.error = None;
        state.update_summary();
        cx.notify();
    }

    pub fn close_pdf_export(&mut self, cx: &mut Context<Self>) {
        if let Some(state) = self.pdf_export.take() {
            state.cancel.store(true, Ordering::Relaxed);
        }
        self.mode = Mode::Navigation;
        cx.notify();
    }

    pub fn save_pdf(&mut self, cx: &mut Context<Self>) {
        let Some(state) = self.pdf_export.as_mut() else {
            return;
        };
        if state.busy || state.summary.is_err() {
            return;
        }
        let Ok(snapshot) = state.snapshot.clone() else {
            return;
        };
        let settings = state.settings.clone();
        let cancel = state.cancel.clone();
        state.busy = true;
        state.error = None;
        state.report = None;
        let directory = self
            .current_file
            .as_ref()
            .and_then(|p| p.parent())
            .map(|p| p.to_path_buf())
            .or_else(|| self.import_source_dir.clone())
            .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());
        let name = snapshot.name.replace(['/', '\\', ':'], "-");
        let future = cx.prompt_for_new_path(&directory, Some(&format!("{name}.pdf")));
        cx.notify();
        cx.spawn(async move |this, cx| {
            let path = match future.await {
                Ok(Ok(Some(path))) => path,
                other => {
                    if !cancel.load(Ordering::Relaxed) {
                        let _ = this.update(cx, |this, cx| {
                            if let Some(state) = this.pdf_export.as_mut() {
                                state.busy = false;
                                if !matches!(other, Ok(Ok(None))) {
                                    state.error = Some(
                                        "Could not open the save dialog. Please try again.".into(),
                                    );
                                }
                            }
                            cx.notify();
                        });
                    }
                    return;
                }
            };
            if cancel.load(Ordering::Relaxed) {
                return;
            }
            if !path
                .extension()
                .is_some_and(|e| e.eq_ignore_ascii_case("pdf"))
            {
                let _ = this.update(cx, |this, cx| {
                    if let Some(state) = this.pdf_export.as_mut() {
                        state.busy = false;
                        state.error = Some("Choose a filename ending in .pdf.".into());
                    }
                    cx.notify();
                });
                return;
            }
            let work_cancel = cancel.clone();
            let saved_path = path.clone();
            let result = smol::unblock(move || {
                let output = visigrid_print::pdf::render_cancellable(&snapshot, &settings, || {
                    work_cancel.load(Ordering::Relaxed)
                })?;
                if work_cancel.load(Ordering::Relaxed) {
                    return Err("Export cancelled".to_string());
                }
                visigrid_print::pdf::save_atomic(&path, &output.bytes)?;
                let mut message =
                    format!("Exported {} page(s) to {}", output.pages, path.display());
                if output.clipped_cells > 0 {
                    message.push_str(&format!(
                        " • {} cell(s) clipped ({}); increase row heights or column widths",
                        output.clipped_cells,
                        output.clipped_addresses.join(", ")
                    ));
                }
                if output.small_text_cells > 0 {
                    message.push_str(&format!(
                        " • {} cell(s) below 8 pt",
                        output.small_text_cells
                    ));
                }
                if !output.substituted_fonts.is_empty() {
                    message.push_str(&format!(
                        " • font fallback: {}",
                        output.substituted_fonts.join(", ")
                    ));
                }
                Ok(message)
            })
            .await;
            if cancel.load(Ordering::Relaxed) {
                return;
            }
            let _ = this.update(cx, |this, cx| {
                match result {
                    Ok(message) => {
                        this.status_message =
                            Some(format!("PDF exported to {}", saved_path.display()));
                        if let Some(state) = this.pdf_export.as_mut() {
                            state.busy = false;
                            state.report = Some(message);
                            state.saved_path = Some(saved_path);
                        }
                    }
                    Err(error) => {
                        if let Some(state) = this.pdf_export.as_mut() {
                            state.busy = false;
                            state.error = Some(error);
                        }
                    }
                }
                cx.notify();
            });
        })
        .detach();
    }
}

fn rgba(color: gpui::Hsla) -> [u8; 4] {
    let c: gpui::Rgba = color.into();
    [c.r, c.g, c.b, c.a].map(|v| (v * 255.0).round() as u8)
}
