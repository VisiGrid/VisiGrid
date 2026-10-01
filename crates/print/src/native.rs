//! Linux system-print dialog. The portal receives the exact, already paginated PDF.
use crate::{PageSettings, Paper};
use ashpd::desktop::{
    print::{Orientation, PageSetup, PrintOptions, PrintProxy, Settings},
    ResponseError,
};
use ashpd::zvariant::{OwnedObjectPath, OwnedValue, Value};
use futures_lite::StreamExt;
use std::{
    collections::HashMap,
    fs::File,
    io::{Seek, Write},
    os::unix::fs::PermissionsExt,
    sync::atomic::{AtomicU64, Ordering},
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Outcome {
    Cancelled,
    Submitted,
}

/// Paper dimensions are unrotated; orientation is a separate printer setting.
fn dialog_settings(page: &PageSettings) -> (Settings, PageSetup) {
    let (name, display_name, width, height) = match page.paper {
        Paper::A4 => ("iso_a4", "A4", 210.0, 297.0),
        Paper::Letter => ("na_letter", "Letter", 215.9, 279.4),
        Paper::Legal => ("na_legal", "Legal", 215.9, 355.6),
    };
    let orientation = if page.landscape {
        Orientation::Landscape
    } else {
        Orientation::Portrait
    };
    let settings = Settings::default()
        .set_orientation(orientation)
        .set_paper_format(name)
        .set_n_copies(1)
        .set_scale(100);
    let setup = PageSetup::default()
        .set_name(name)
        .set_display_name(display_name)
        .set_width(width)
        .set_height(height)
        .set_orientation(orientation)
        .set_margin_top(page.margins.top * 25.4 / 72.0)
        .set_margin_bottom(page.margins.bottom * 25.4 / 72.0)
        .set_margin_left(page.margins.left * 25.4 / 72.0)
        .set_margin_right(page.margins.right * 25.4 / 72.0);
    (settings, setup)
}

// Kept at the transport boundary so tests exercise cancellation and the bytes
// handed to the system without opening a dialog or submitting a real print job.
trait Portal {
    async fn prepare(
        &self,
        title: &str,
        settings: Settings,
        page: PageSetup,
    ) -> Result<u32, ashpd::Error>;
    async fn submit(&self, title: &str, file: &File, token: u32) -> Result<(), ashpd::Error>;
}
impl Portal for PrintProxy {
    async fn prepare(
        &self,
        title: &str,
        settings: Settings,
        page: PageSetup,
    ) -> Result<u32, ashpd::Error> {
        // ashpd 0.13.13 serializes these option names with hyphens and PDF
        // formats in uppercase. The portal requires underscores and "pdf".
        // Its typed response also rejects GTK's lowercase output-file-format.
        // Use the public D-Bus proxy for PreparePrint and decode only the fields
        // needed for submission.
        static NEXT_REQUEST: AtomicU64 = AtomicU64::new(1);
        let token = format!(
            "visigrid_print_{}",
            NEXT_REQUEST.fetch_add(1, Ordering::Relaxed)
        );
        let sender = self
            .connection()
            .unique_name()
            .ok_or(ashpd::Error::ParseError("Missing session bus name"))?
            .as_str()
            .trim_start_matches(':')
            .replace('.', "_");
        let path = format!("/org/freedesktop/portal/desktop/request/{sender}/{token}");
        let request = ashpd::zbus::Proxy::new(
            self.connection(),
            "org.freedesktop.portal.Desktop",
            path.as_str(),
            "org.freedesktop.portal.Request",
        )
        .await?;
        // Subscribe before calling: an immediate cancellation must not be lost.
        let mut responses = request.receive_signal("Response").await?;
        let actual: OwnedObjectPath = self
            .call(
                "PreparePrint",
                &(
                    "",
                    title,
                    settings,
                    page,
                    prepare_options(&token, self.version()),
                ),
            )
            .await?;
        if actual.as_str() != path {
            let unexpected = ashpd::zbus::Proxy::new(
                self.connection(),
                "org.freedesktop.portal.Desktop",
                actual,
                "org.freedesktop.portal.Request",
            )
            .await?;
            let _: () = unexpected.call("Close", &()).await?;
            return Err(ashpd::Error::ParseError("Unexpected print request handle"));
        }
        let message = responses.next().await.ok_or(ashpd::Error::NoResponse)?;
        let (status, results) = message.body().deserialize()?;
        prepared_token(status, results)
    }
    async fn submit(&self, title: &str, file: &File, token: u32) -> Result<(), ashpd::Error> {
        self.print(
            None,
            title,
            file,
            PrintOptions::default().set_token(token).set_modal(true),
        )
        .await?
        .response()
    }
}

fn prepare_options(token: &str, version: u32) -> HashMap<&str, Value<'_>> {
    let mut options = HashMap::from([
        ("handle_token", Value::from(token)),
        ("modal", Value::from(true)),
    ]);
    if version >= 3 {
        options.insert("supported_output_file_formats", Value::from(vec!["pdf"]));
    }
    if version >= 4 {
        options.insert("has_current_page", Value::from(false));
        options.insert("has_selected_pages", Value::from(false));
    }
    options
}

fn prepared_token(
    status: u32,
    mut results: HashMap<String, OwnedValue>,
) -> Result<u32, ashpd::Error> {
    match status {
        1 => return Err(ResponseError::Cancelled.into()),
        0 => {}
        _ => return Err(ResponseError::Other.into()),
    }
    // Older portals may offer other file formats. Never send PDF bytes to a
    // requested PostScript/SVG destination. Ignore unrelated vendor settings.
    if let Some(settings) = results.remove("settings") {
        let settings = HashMap::<String, OwnedValue>::try_from(settings)
            .map_err(|_| ashpd::Error::ParseError("Invalid print settings"))?;
        if let Some(format) = settings.get("output-file-format") {
            let format = <&str>::try_from(format)
                .map_err(|_| ashpd::Error::ParseError("Invalid print output format"))?;
            if !format.eq_ignore_ascii_case("pdf") {
                return Err(ashpd::Error::ParseError(
                    "Only PDF print-to-file output is supported",
                ));
            }
        }
    }
    results
        .remove("token")
        .and_then(|v| u32::try_from(v).ok())
        .ok_or(ashpd::Error::ParseError("Missing print request token"))
}

fn portal_result<T>(result: Result<T, ashpd::Error>) -> Result<Option<T>, String> {
    match result {
        Ok(value) => Ok(Some(value)),
        Err(ashpd::Error::Response(ResponseError::Cancelled)) => Ok(None),
        Err(error) => Err(error.to_string()),
    }
}

async fn print_with(
    portal: &impl Portal,
    bytes: &[u8],
    page: &PageSettings,
    title: &str,
) -> Result<Outcome, String> {
    if !bytes.starts_with(b"%PDF-") {
        return Err("The print document is not a PDF.".into());
    }
    // Anonymous, mode-0600 file: no lasting path and no cleanup race with the portal.
    let mut file =
        tempfile::tempfile().map_err(|e| format!("Could not stage print document: {e}"))?;
    file.set_permissions(std::fs::Permissions::from_mode(0o600))
        .map_err(|e| format!("Could not secure print document: {e}"))?;
    file.write_all(bytes)
        .and_then(|_| file.rewind())
        .map_err(|e| format!("Could not stage print document: {e}"))?;
    let (settings, setup) = dialog_settings(page);
    let Some(token) = portal_result(portal.prepare(title, settings, setup).await)? else {
        return Ok(Outcome::Cancelled);
    };
    // Keep the file open until the portal has accepted the document. The token
    // carries the user's printer, copies and page-range choices; never auto-retry.
    Ok(
        if portal_result(portal.submit(title, &file, token).await)?.is_some() {
            Outcome::Submitted
        } else {
            Outcome::Cancelled
        },
    )
}

pub async fn print_pdf(bytes: &[u8], page: &PageSettings, title: &str) -> Result<Outcome, String> {
    let portal = PrintProxy::new()
        .await
        .map_err(|e| format!("System print service is unavailable: {e}"))?;
    print_with(&portal, bytes, page, title).await
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::{
        cell::{Cell, RefCell},
        io::Read,
        os::unix::fs::PermissionsExt,
    };

    struct Fake {
        prepare_result: Option<ResponseError>,
        submit_result: Option<ResponseError>,
        prepared: Cell<usize>,
        submitted: Cell<usize>,
        sent: RefCell<Vec<u8>>,
    }
    impl Portal for Fake {
        async fn prepare(&self, _: &str, _: Settings, _: PageSetup) -> Result<u32, ashpd::Error> {
            self.prepared.set(self.prepared.get() + 1);
            self.prepare_result.map_or(Ok(42), |e| Err(e.into()))
        }
        async fn submit(&self, _: &str, mut file: &File, token: u32) -> Result<(), ashpd::Error> {
            self.submitted.set(self.submitted.get() + 1);
            assert_eq!(token, 42);
            assert_eq!(file.metadata().unwrap().permissions().mode() & 0o777, 0o600);
            file.read_to_end(&mut self.sent.borrow_mut()).unwrap();
            self.submit_result.map_or(Ok(()), |e| Err(e.into()))
        }
    }
    fn fake(prepare_result: Option<ResponseError>, submit_result: Option<ResponseError>) -> Fake {
        Fake {
            prepare_result,
            submit_result,
            prepared: Cell::new(0),
            submitted: Cell::new(0),
            sent: RefCell::new(Vec::new()),
        }
    }
    #[test]
    fn cancel_does_not_submit_and_invalid_document_never_opens_dialog() {
        smol::block_on(async {
            let fake = fake(Some(ResponseError::Cancelled), None);
            assert_eq!(
                print_with(
                    &fake,
                    b"%PDF-1.7\nexact bytes",
                    &PageSettings::default(),
                    "Report"
                )
                .await
                .unwrap(),
                Outcome::Cancelled
            );
            assert!(fake.sent.borrow().is_empty());
            assert_eq!(fake.submitted.get(), 0);
            assert!(
                print_with(&fake, b"bad", &PageSettings::default(), "Report")
                    .await
                    .is_err()
            );
            assert_eq!(fake.prepared.get(), 1);
        });
    }
    #[test]
    fn sends_exact_preview_bytes_and_does_not_report_failed_submission_as_success() {
        smol::block_on(async {
            let bytes = b"%PDF-1.7\nunchanged cached preview\0\xff";
            for error in [
                None,
                Some(ResponseError::Other),
                Some(ResponseError::Cancelled),
            ] {
                let fake = fake(None, error);
                let result = print_with(&fake, bytes, &PageSettings::default(), "Report").await;
                assert_eq!(fake.sent.borrow().as_slice(), bytes);
                match error {
                    None => assert_eq!(result.unwrap(), Outcome::Submitted),
                    Some(ResponseError::Cancelled) => {
                        assert_eq!(result.unwrap(), Outcome::Cancelled)
                    }
                    _ => assert!(result.is_err()),
                }
                assert_eq!(fake.prepared.get(), 1);
                assert_eq!(fake.submitted.get(), 1);
            }
        });
    }
    #[test]
    fn failed_dialog_does_not_submit() {
        smol::block_on(async {
            let fake = fake(Some(ResponseError::Other), None);
            assert!(
                print_with(&fake, b"%PDF-1.7\n", &PageSettings::default(), "Report")
                    .await
                    .is_err()
            );
            assert_eq!(fake.prepared.get(), 1);
            assert_eq!(fake.submitted.get(), 0);
        });
    }
    #[test]
    fn portal_wire_options_and_gtk_response_use_portal_spelling() {
        let mut options = prepare_options("test", 4);
        assert_eq!(
            Vec::<String>::try_from(options.remove("supported_output_file_formats").unwrap())
                .unwrap(),
            vec!["pdf"]
        );
        assert!(!bool::try_from(options.remove("has_current_page").unwrap()).unwrap());
        assert!(!bool::try_from(options.remove("has_selected_pages").unwrap()).unwrap());
        assert!(!prepare_options("test", 2).contains_key("supported_output_file_formats"));
        assert!(!prepare_options("test", 3).contains_key("has_current_page"));
        for format in ["pdf", "ps"] {
            let settings = Value::from(HashMap::from([
                ("output-file-format", Value::from(format)),
                ("vendor-specific-option", Value::from("unknown")),
            ]));
            let results = HashMap::from([
                ("token".into(), OwnedValue::from(42u32)),
                ("settings".into(), OwnedValue::try_from(settings).unwrap()),
            ]);
            let result = prepared_token(0, results);
            if format == "pdf" {
                assert_eq!(result.unwrap(), 42);
            } else {
                assert!(result.is_err());
            }
        }
        assert!(matches!(
            prepared_token(1, HashMap::new()),
            Err(ashpd::Error::Response(ResponseError::Cancelled))
        ));
        assert!(prepared_token(0, HashMap::new()).is_err());
    }
    #[test]
    fn paper_orientation_and_existing_pdf_margins_seed_dialog() {
        let page = PageSettings {
            paper: Paper::Legal,
            landscape: true,
            ..Default::default()
        };
        let (_, setup) = dialog_settings(&page);
        assert_eq!(setup.width, Some(215.9));
        assert_eq!(setup.height, Some(355.6));
        assert_eq!(setup.orientation, Some(Orientation::Landscape));
        assert!((setup.margin_top.unwrap() - 12.7).abs() < 0.0001);
    }
}
