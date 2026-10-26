// Grid (Loco) as the cloud-sheet backend.
//
// With a Grid sign-in on this computer (`grid-auth.json`), Open Cloud, Move
// to Cloud and every upload go to Grid's sheet routes, the contract the web
// uses. Without one, the Rails desktop sheets client keeps working for
// installed users. Hub publish, license checks and `vgrid login` never read
// the Grid credential; they stay on the Rails token in `auth.json`.

use crate::app::Spreadsheet;
use crate::cloud::SheetInfo;
use crate::hub::client::HubError;
use visigrid_hub_client::grid::{self, GridClient, GridCredentials, GridError, GridSheet};

/// The Grid this computer signs into: `VISIGRID_GRID_API_BASE`, or the one
/// it is already signed into. There is no default, and it is never the Rails
/// Hub base.
pub fn configured_base() -> Option<String> {
    std::env::var("VISIGRID_GRID_API_BASE")
        .ok()
        .map(|b| b.trim().trim_end_matches('/').to_string())
        .filter(|b| !b.is_empty())
        .or_else(|| grid::load_grid_auth().map(|c| c.api_base))
}

/// Whether cloud sheets go to Grid on this computer.
pub fn signed_in() -> bool {
    grid::load_grid_auth().is_some()
}

/// A Grid sheet as the Open Cloud picker shows it. The pid doubles as the
/// cache file name: it is stable and filesystem-safe.
pub fn sheet_info(sheet: GridSheet) -> SheetInfo {
    SheetInfo {
        id: sheet.id,
        public_id: sheet.pid.clone(),
        name: sheet.name,
        slug: sheet.pid.clone(),
        byte_size: None,
        last_edited_at: sheet.updated_at,
        revision: Some(sheet.revision),
        grid_pid: Some(sheet.pid),
    }
}

/// For the paths that report through `HubError`. A signed-out device is
/// handled before this, by the caller.
pub fn hub_error(error: GridError) -> HubError {
    match error {
        GridError::NotSignedIn => HubError::NotAuthenticated,
        GridError::Network(e) => HubError::Network(e),
        GridError::Http(status, body) => HubError::Http(status, body),
        GridError::SignedOut => HubError::Http(401, "Signed out of Grid".into()),
        other => HubError::Parse(other.to_string()),
    }
}

impl Spreadsheet {
    /// "Grid: Sign In": open Grid's authorize page and wait for its code.
    pub fn grid_sign_in(&mut self, cx: &mut gpui::Context<Self>) {
        let Some(base) = configured_base() else {
            self.status_message = Some(
                "Set VISIGRID_GRID_API_BASE to your Grid address, then run \"Grid: Sign In\".".to_string(),
            );
            cx.notify();
            return;
        };
        if signed_in() {
            self.status_message = Some(format!("Already signed in to Grid at {base}"));
            cx.notify();
            return;
        }
        let _ = open::that(format!("{base}/desktop/authorize"));
        self.grid_sign_in_base = Some(base);
        self.hub_token_input.clear();
        self.mode = crate::mode::Mode::HubPasteToken;
        self.status_message = Some("Authorize this computer in Grid, then paste the code below.".to_string());
        cx.notify();
    }

    /// Check the pasted device code with Grid, then keep it (0600).
    pub(crate) fn grid_complete_sign_in(&mut self, secret: String, base: String, cx: &mut gpui::Context<Self>) {
        let creds = GridCredentials { secret, api_base: base };
        let client = GridClient::new(creds.clone());
        self.status_message = Some("Checking the code with Grid…".to_string());
        cx.notify();
        cx.spawn(async move |this, cx| {
            let result = smol::unblock(move || client.verify()).await;
            let _ = this.update(cx, |this, cx| {
                match result.and_then(|()| grid::save_grid_auth(&creds).map_err(GridError::Parse)) {
                    Ok(()) => {
                        this.grid_sign_in_base = None;
                        this.mode = crate::mode::Mode::Navigation;
                        this.hub_token_input.clear();
                        this.status_message = Some(format!("Signed in to Grid. Cloud sheets now sync with {}", creds.api_base));
                    }
                    Err(GridError::SignedOut) => {
                        this.status_message = Some("Grid did not accept that code. Copy the whole code from the authorize page and try again.".to_string());
                    }
                    Err(e) => this.status_message = Some(format!("Could not sign in to Grid: {e}")),
                }
                cx.notify();
            });
        })
        .detach();
    }

    /// "Grid: Sign Out": revoke this device in Grid and forget the code.
    pub fn grid_sign_out(&mut self, cx: &mut gpui::Context<Self>) {
        let Some(creds) = grid::load_grid_auth() else {
            self.status_message = Some("Not signed in to Grid".to_string());
            cx.notify();
            return;
        };
        cx.spawn(async move |this, cx| {
            let revoked = smol::unblock(move || GridClient::new(creds).sign_out()).await;
            let deleted = grid::delete_grid_auth();
            let _ = this.update(cx, |this, cx| {
                this.status_message = Some(match (revoked, deleted) {
                    (_, Err(e)) => format!("Could not remove the Grid sign-in: {e}"),
                    (Ok(()), Ok(())) => "Signed out of Grid".to_string(),
                    (Err(e), Ok(())) => format!(
                        "Signed out of Grid on this computer, but Grid could not be told ({e}). Remove the device in Grid to revoke it."
                    ),
                });
                cx.notify();
            });
        })
        .detach();
    }

    /// Grid refused this device (revoked, expired, or the account's sign-ins
    /// were reset). Forget the code; the Rails sign-in is untouched.
    pub(crate) fn grid_sign_in_expired(&mut self, cx: &mut gpui::Context<Self>) {
        let _ = grid::delete_grid_auth();
        self.status_message = Some(
            "This computer was signed out of Grid. Run \"Grid: Sign In\" to resume cloud sync; your file is saved locally."
                .to_string(),
        );
        cx.notify();
    }
}
