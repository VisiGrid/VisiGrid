//! Active-view capture and asynchronous PDF export. Export never changes save state.
use crate::{app::Spreadsheet, mode::Mode};
use gpui::{App, Context, RenderImage};
use std::sync::{
    atomic::{AtomicBool, Ordering},
    Arc,
};
use visigrid_engine::{
    print_setup::{PrintArea, PrintPaper, PrintRows, PrintScale, PrintSetup},
    sheet::SheetId,
};
use visigrid_print::snapshot::{self, Scope, SheetView, Snapshot};
use visigrid_print::{paginate, AxisItem, PageSettings, GRID_UNIT_PT};

pub struct PdfExportState {
    pub settings: PageSettings,
    pub selection_only: bool,
    pub use_print_area: bool,
    pub setup: PrintSetup,
    pub saved_setup: PrintSetup,
    pub sheet_id: SheetId,
    pub snapshot: Result<Arc<Snapshot>, String>,
    pub summary: Result<String, String>,
    pub busy: bool,
    pub choosing_path: bool,
    pub printing: bool,
    pub print_mode: bool,
    pub print_message: Option<String>,
    pub cancel: Arc<AtomicBool>,
    pub focus: usize,
    pub settings_scroll: gpui::ScrollHandle,
    pub error: Option<String>,
    pub report: Option<Vec<String>>,
    pub saved_path: Option<std::path::PathBuf>,
    pub output: Option<Arc<visigrid_print::pdf::PdfOutput>>,
    pub preview_cancel: Arc<AtomicBool>,
    pub preparing: bool,
    pub page_loading: bool,
    pub preview_error: Option<String>,
    pub image: Option<Arc<RenderImage>>,
    pub page: usize,
    pub page_count: usize,
    pub page_size: (f64, f64),
    /// None fits the page in the viewport. This never affects PDF scale.
    pub zoom: Option<f32>,
    pub notices: Vec<String>,
    pub captured_revision: u64,
}

impl PdfExportState {
    pub fn focus_option(&mut self, option: usize) {
        self.focus = option;
        self.settings_scroll.scroll_to_item(match option {
            0..=5 => option + 1,
            6..=8 => 7,
            9..=10 => 8,
            _ => 9,
        });
    }

    pub fn update_summary(&mut self) {
        self.page_count = 0;
        self.summary = self.snapshot.as_ref().map_err(Clone::clone).and_then(|s| {
            self.settings = visigrid_print::setup::page_settings(s, &self.setup)?;
            let plan = paginate(&s.layout, &self.settings).map_err(|e| e.to_string())?;
            self.page_count = plan.pages().len();
            self.page_size = plan.page_size();
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
        if self.block_read_only_recovery(cx) { return; }
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
        if let Some(old) = self.pdf_export.take() {
            old.preview_cancel.store(true, Ordering::Relaxed);
        }
        let setup = self.sheet(cx).print_setup.clone();
        let snapshot = self.capture_pdf(false, setup.area, cx).map(Arc::new);
        let mut state = PdfExportState {
            settings: PageSettings::default(),
            selection_only: false,
            use_print_area: setup.area.is_some(),
            saved_setup: setup.clone(),
            setup,
            sheet_id: self.sheet(cx).id,
            snapshot,
            summary: Ok(String::new()),
            busy: false,
            choosing_path: false,
            printing: false,
            print_mode: false,
            print_message: None,
            cancel: Arc::new(AtomicBool::new(false)),
            focus: 0,
            settings_scroll: gpui::ScrollHandle::new(),
            error: None,
            report: None,
            saved_path: None,
            output: None,
            preview_cancel: Arc::new(AtomicBool::new(false)),
            preparing: false,
            page_loading: false,
            preview_error: None,
            image: None,
            page: 0,
            page_count: 0,
            page_size: (595.0, 842.0),
            zoom: None,
            notices: Vec::new(),
            captured_revision: self.workbook.read(cx).revision(),
        };
        state.update_summary();
        self.pdf_export = Some(state);
        self.mode = Mode::ExportPdf;
        self.rebuild_pdf_preview(cx);
        cx.notify();
    }

    fn rebuild_pdf_preview(&mut self, cx: &mut Context<Self>) {
        let Some(state) = self.pdf_export.as_mut() else {
            return;
        };
        state.preview_cancel.store(true, Ordering::Relaxed);
        state.preview_cancel = Arc::new(AtomicBool::new(false));
        state.output = None;
        state.image = None;
        state.page = 0;
        state.page_loading = false;
        state.preparing = false;
        state.preview_error = None;
        state.print_message = None;
        state.notices.clear();
        if state.summary.is_err() {
            cx.notify();
            return;
        }
        let Ok(snapshot) = state.snapshot.clone() else {
            return;
        };
        let settings = state.settings.clone();
        let cancel = state.preview_cancel.clone();
        state.preparing = true;
        cx.spawn(async move |this, cx| {
            let worker_cancel = cancel.clone();
            let result = smol::unblock(move || {
                visigrid_print::pdf::render_cancellable(&snapshot, &settings, || {
                    worker_cancel.load(Ordering::Relaxed)
                })
            })
            .await;
            if cancel.load(Ordering::Relaxed) {
                return;
            }
            let _ = this.update(cx, |this, cx| {
                let Some(state) = this.pdf_export.as_mut() else {
                    return;
                };
                if !Arc::ptr_eq(&state.preview_cancel, &cancel) {
                    return;
                }
                state.preparing = false;
                match result {
                    Ok(output) => {
                        state.notices = pdf_notices(&output);
                        state.output = Some(Arc::new(output));
                    }
                    Err(error) => state.preview_error = Some(error),
                }
                this.load_pdf_preview_page(cx);
                cx.notify();
            });
        })
        .detach();
        cx.notify();
    }

    fn load_pdf_preview_page(&mut self, cx: &mut Context<Self>) {
        let Some(state) = self.pdf_export.as_mut() else {
            return;
        };
        // Coalesce page changes: at most one raster task for this document.
        if state.page_loading {
            return;
        }
        let Some(output) = state.output.clone() else {
            return;
        };
        let page = state.page;
        let cancel = state.preview_cancel.clone();
        state.page_loading = true;
        state.preview_error = None;
        cx.spawn(async move |this, cx| {
            let worker_cancel = cancel.clone();
            let result = smol::unblock(move || {
                if worker_cancel.load(Ordering::Relaxed) {
                    return Err("Preview cancelled".to_string());
                }
                let mut page = visigrid_print::preview::rasterize(output, page, 2400)?;
                // GPUI's RenderImage consumes BGRA, the rasterizer returns RGBA.
                for pixel in page.rgba.as_chunks_mut::<4>().0 {
                    pixel.swap(0, 2);
                }
                let buffer = image::RgbaImage::from_raw(page.width, page.height, page.rgba)
                    .ok_or("Invalid preview bitmap")?;
                Ok(Arc::new(RenderImage::new(vec![image::Frame::new(buffer)])))
            })
            .await;
            if cancel.load(Ordering::Relaxed) {
                return;
            }
            let _ = this.update(cx, |this, cx| {
                let Some(state) = this.pdf_export.as_mut() else {
                    return;
                };
                if !Arc::ptr_eq(&state.preview_cancel, &cancel) {
                    return;
                }
                state.page_loading = false;
                if state.page == page {
                    match result {
                        Ok(image) => state.image = Some(image),
                        Err(error) => state.preview_error = Some(error),
                    }
                } else {
                    this.load_pdf_preview_page(cx);
                }
                cx.notify();
            });
        })
        .detach();
    }

    pub fn pdf_preview_page(&mut self, delta: isize, cx: &mut Context<Self>) {
        let Some(state) = self.pdf_export.as_mut() else {
            return;
        };
        if state.page_count == 0 || state.output.is_none() {
            return;
        }
        let page = state
            .page
            .saturating_add_signed(delta)
            .min(state.page_count - 1);
        if state.page == page {
            return;
        }
        state.page = page;
        state.image = None;
        state.preview_error = None;
        self.load_pdf_preview_page(cx);
        cx.notify();
    }

    pub fn pdf_preview_zoom(&mut self, zoom: f32, cx: &mut Context<Self>) {
        if let Some(state) = self.pdf_export.as_mut() {
            state.zoom = if zoom == 0.0 {
                None
            } else {
                Some(zoom.clamp(0.5, 2.0))
            };
            cx.notify();
        }
    }

    pub fn refresh_pdf_preview(&mut self, cx: &mut Context<Self>) {
        let Some(state) = self.pdf_export.as_ref() else {
            return;
        };
        if state.busy {
            return;
        }
        let snapshot = if state.sheet_id != self.sheet(cx).id {
            Err("The active sheet changed. Close and reopen Print Preview.".into())
        } else {
            self.capture_pdf(
                state.selection_only,
                if state.use_print_area {
                    state.setup.area
                } else {
                    None
                },
                cx,
            )
            .map(Arc::new)
        };
        let state = self.pdf_export.as_mut().unwrap();
        state.snapshot = snapshot;
        state.captured_revision = self.workbook.read(cx).revision();
        state.report = None;
        state.saved_path = None;
        state.error = None;
        state.update_summary();
        self.rebuild_pdf_preview(cx);
    }

    fn capture_pdf(
        &self,
        selection_only: bool,
        area: Option<PrintArea>,
        cx: &App,
    ) -> Result<Snapshot, String> {
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
        let view = SheetView { rows, columns };
        let mut snapshot = if let Some(area) = area {
            visigrid_print::setup::capture_area(
                self.sheet(cx),
                &view,
                area,
                &self.cell_font.family,
                self.cell_font.size,
            )?
        } else {
            snapshot::capture(
                self.sheet(cx),
                &view,
                selection,
                &self.cell_font.family,
                self.cell_font.size,
            )?
        };
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
        if option == 11 {
            self.save_pdf_setup(cx);
            return;
        }
        if option == 9 {
            self.pdf_area_from_selection(cx);
            return;
        }
        let state = self.pdf_export.as_mut().unwrap();
        state.focus = option;
        match option {
            0 => {
                if state.use_print_area {
                    state.use_print_area = false;
                } else if !state.selection_only {
                    state.selection_only = true;
                } else {
                    state.selection_only = false;
                    state.use_print_area = state.setup.area.is_some();
                }
            }
            1 => {
                state.setup.paper = match state.setup.paper {
                    PrintPaper::A4 => PrintPaper::Letter,
                    PrintPaper::Letter => PrintPaper::Legal,
                    PrintPaper::Legal => PrintPaper::A4,
                }
            }
            2 => state.setup.landscape = !state.setup.landscape,
            3 => {
                state.setup.scale = match state.setup.scale {
                    PrintScale::FitColumns => PrintScale::Actual,
                    PrintScale::Actual => PrintScale::FitSheet,
                    PrintScale::FitSheet => PrintScale::FitColumns,
                }
            }
            4 => state.setup.page_numbers = !state.setup.page_numbers,
            5 => state.setup.gridlines = !state.setup.gridlines,
            6 | 7 => {
                let Ok(snapshot) = &state.snapshot else {
                    return;
                };
                let rows = &snapshot.layout.rows;
                let count = state.setup.repeat_rows.map_or(0, |r| {
                    rows.iter()
                        .take_while(|i| i.source_index >= r.start && i.source_index <= r.end)
                        .count()
                });
                let count = if option == 6 {
                    (count + 1).min(rows.len().saturating_sub(1))
                } else {
                    count.saturating_sub(1)
                };
                if rows[..count]
                    .windows(2)
                    .any(|r| r[0].source_index >= r[1].source_index)
                {
                    state.error =
                        Some("Clear the sort before choosing repeated header rows.".into());
                    cx.notify();
                    return;
                }
                state.setup.repeat_rows = (count > 0).then(|| PrintRows {
                    start: rows[0].source_index,
                    end: rows[count - 1].source_index,
                });
            }
            8 => state.setup.repeat_rows = None,
            10 => {
                state.setup.area = None;
                state.use_print_area = false;
                state.selection_only = false;
            }
            _ => return,
        }
        state.error = None;
        state.report = None;
        state.saved_path = None;
        if option == 0 || option == 10 {
            self.refresh_pdf_preview(cx);
        } else {
            state.update_summary();
            self.rebuild_pdf_preview(cx);
        }
        cx.notify();
    }

    fn pdf_area_from_selection(&mut self, cx: &mut Context<Self>) {
        if self
            .pdf_export
            .as_ref()
            .is_none_or(|s| s.sheet_id != self.sheet(cx).id)
        {
            if let Some(state) = self.pdf_export.as_mut() {
                state.error =
                    Some("The active sheet changed. Close and reopen Print Preview.".into());
            }
            cx.notify();
            return;
        }
        let result = self.capture_pdf(true, None, cx).and_then(|snapshot| {
            let rows = &snapshot.layout.rows;
            let cols = &snapshot.layout.columns;
            let area = PrintArea {
                start_row: rows.iter().map(|r| r.source_index).min().ok_or("No selected rows")?,
                end_row: rows.iter().map(|r| r.source_index).max().unwrap(),
                start_col: cols.iter().map(|c| c.source_index).min().ok_or("No selected columns")?,
                end_col: cols.iter().map(|c| c.source_index).max().unwrap(),
            };
            let sources: std::collections::HashSet<_> = rows.iter().map(|r| r.source_index).collect();
            if self.row_view.visible_rows().iter().map(|&v| self.row_view.view_to_data(v)).any(|r| r >= area.start_row && r <= area.end_row && !sources.contains(&r) && !self.is_row_hidden(r)) {
                return Err("This sorted selection is not one source rectangle. Clear the sort before setting a print area.".into());
            }
            Ok(area)
        });
        let state = self.pdf_export.as_mut().unwrap();
        state.focus = 9;
        match result {
            Ok(area) => {
                state.setup.area = Some(area);
                state.use_print_area = true;
                state.selection_only = false;
                self.refresh_pdf_preview(cx);
            }
            Err(error) => {
                state.error = Some(error);
                cx.notify();
            }
        }
    }

    fn save_pdf_setup(&mut self, cx: &mut Context<Self>) {
        let Some(state) = &self.pdf_export else {
            return;
        };
        if state.busy || state.summary.is_err() {
            return;
        }
        let (sheet_id, after) = (state.sheet_id, state.setup.clone());
        let previous_revision = self.workbook.read(cx).revision();
        let result = self
            .workbook
            .update(cx, |wb, _| wb.set_print_setup(sheet_id, after.clone()));
        match result {
            Ok(before) => {
                if before != after {
                    self.history.record_action_with_provenance(
                        crate::history::UndoAction::PrintSetupChanged {
                            sheet_id,
                            before,
                            after: after.clone(),
                        },
                        None,
                    );
                    self.is_modified = true;
                }
                let state = self.pdf_export.as_mut().unwrap();
                state.saved_setup = after;
                if state.captured_revision == previous_revision {
                    state.captured_revision = self.workbook.read(cx).revision();
                }
                state.focus = 11;
                state.error = None;
                self.status_message = Some(
                    "Print setup applied to this sheet. Save the workbook to keep it on disk."
                        .into(),
                );
            }
            Err(error) => self.pdf_export.as_mut().unwrap().error = Some(error),
        }
        cx.notify();
    }

    pub fn close_pdf_export(&mut self, cx: &mut Context<Self>) {
        if self.pdf_export.as_ref().is_some_and(|s| s.printing) {
            return; // Cancel in the system print dialog; do not orphan its request.
        }
        if let Some(state) = self.pdf_export.take() {
            state.cancel.store(true, Ordering::Relaxed);
            state.preview_cancel.store(true, Ordering::Relaxed);
        }
        self.mode = Mode::Navigation;
        cx.notify();
    }

    pub fn save_pdf(&mut self, cx: &mut Context<Self>) {
        let Some(state) = self.pdf_export.as_mut() else {
            return;
        };
        if state.busy || state.preparing || state.summary.is_err() {
            return;
        }
        let Ok(snapshot) = state.snapshot.clone() else {
            return;
        };
        let Some(output) = state.output.clone() else {
            return;
        };
        let cancel = state.cancel.clone();
        state.busy = true;
        state.choosing_path = true;
        state.print_message = None;
        state.error = None;
        state.report = None;
        state.saved_path = None;
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
                                state.choosing_path = false;
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
            let _ = this.update(cx, |this, cx| {
                if let Some(state) = this.pdf_export.as_mut() {
                    state.choosing_path = false;
                }
                cx.notify();
            });
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
                if work_cancel.load(Ordering::Relaxed) {
                    return Err("Export cancelled".to_string());
                }
                visigrid_print::pdf::save_atomic(&path, &output.bytes)?;
                let message = vec![format!(
                    "Saved {} page{} as PDF",
                    output.pages,
                    if output.pages == 1 { "" } else { "s" }
                )];
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

fn pdf_notices(output: &visigrid_print::pdf::PdfOutput) -> Vec<String> {
    let mut notices = Vec::new();
    if output.clipped_cells > 0 {
        notices.push(format!(
            "Clipped text in {} cell(s): {}{}. Increase row heights or column widths.",
            output.clipped_cells,
            output.clipped_addresses.join(", "),
            if output.clipped_cells > output.clipped_addresses.len() {
                " (first 10 shown)"
            } else {
                ""
            }
        ));
    }
    if !output.substituted_fonts.is_empty() {
        notices.push(format!(
            "Fonts substituted: {}",
            output.substituted_fonts.join(", ")
        ));
    }
    notices
}
