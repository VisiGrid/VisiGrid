//! System printing of the cached preview document.
use crate::{app::Spreadsheet, mode::Mode};
use gpui::Context;

#[cfg(not(target_os = "linux"))]
const EXPORT_TO_PRINT: &str =
    "System printing is currently available on Linux. Export PDF to print using your PDF viewer.";

impl Spreadsheet {
    pub fn show_print_preview(&mut self, cx: &mut Context<Self>) {
        if self.mode != Mode::ExportPdf {
            self.show_pdf_export(cx);
        }
        if let Some(state) = self.pdf_export.as_mut() {
            if !state.busy {
                state.print_mode = cfg!(target_os = "linux");
                #[cfg(not(target_os = "linux"))]
                {
                    state.print_message = Some(EXPORT_TO_PRINT.into());
                }
            }
        }
        cx.notify();
    }

    /// Ctrl+P opens preview first; from preview it opens the system print dialog.
    pub fn print_action(&mut self, cx: &mut Context<Self>) {
        if self.mode == Mode::ExportPdf {
            self.print_pdf(cx);
        } else {
            self.show_print_preview(cx);
        }
    }

    pub fn confirm_pdf_action(&mut self, cx: &mut Context<Self>) {
        if self.pdf_export.as_ref().is_some_and(|s| s.print_mode) {
            self.print_pdf(cx);
        } else {
            self.save_pdf(cx);
        }
    }

    #[cfg(not(target_os = "linux"))]
    pub fn print_pdf(&mut self, cx: &mut Context<Self>) {
        if let Some(state) = self.pdf_export.as_mut() {
            if state.busy {
                return;
            }
            state.print_message = Some(EXPORT_TO_PRINT.into());
        }
        cx.notify();
    }

    #[cfg(target_os = "linux")]
    pub fn print_pdf(&mut self, cx: &mut Context<Self>) {
        let Some(state) = self.pdf_export.as_mut() else {
            return;
        };
        if state.busy || state.preparing || state.summary.is_err() {
            return;
        }
        let (Some(output), Ok(snapshot)) = (state.output.clone(), &state.snapshot) else {
            return;
        };
        let title = format!("{} — VisiGrid", snapshot.name);
        let settings = state.settings.clone();
        state.busy = true;
        state.printing = true;
        state.print_message = None;
        state.error = None;
        // The preview cannot close or change while the system dialog owns this
        // request. Cancellation belongs to that dialog; there are no auto-retries.
        let task = cx.background_executor().spawn(async move {
            visigrid_print::native::print_pdf(&output.bytes, &settings, &title).await
        });
        cx.spawn(async move |this, cx| {
            let result = task.await;
            let _ = this.update(cx, |this, cx| {
                let Some(state) = this.pdf_export.as_mut() else {
                    return;
                };
                state.busy = false;
                state.printing = false;
                let message = match result {
                    Ok(visigrid_print::native::Outcome::Cancelled) => "Printing cancelled.",
                    Ok(visigrid_print::native::Outcome::Submitted) => {
                        "Print request submitted. Check your printer’s queue."
                    }
                    Err(error) => {
                        log::warn!("System printing failed: {error}");
                        "Print request wasn’t completed. Check system printing, or export PDF."
                    }
                };
                state.print_message = Some(message.into());
                this.status_message = Some(message.into());
                cx.notify();
            });
        })
        .detach();
        cx.notify();
    }
}
