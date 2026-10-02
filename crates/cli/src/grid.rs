//! `vgrid grid` — sign this computer into Grid (Loco) for cloud sheets.
//!
//! Separate from `vgrid login`, which stores the Rails Hub token in
//! `auth.json` for Hub publish and license checks. Grid's sign-in is a device
//! code from Grid's authorize page, stored in `grid-auth.json` (0600).

use std::io::{self, Write};

use visigrid_hub_client::grid::{self, GridClient, GridCredentials, GridError};

use crate::exit_codes::*;
use crate::CliError;

pub fn cmd_login(code: Option<String>, api_base: String) -> Result<(), CliError> {
    let api_base = api_base.trim().trim_end_matches('/').to_string();
    if !(api_base.starts_with("https://") || api_base.starts_with("http://")) {
        return Err(CliError { code: EXIT_USAGE, message: format!("Invalid Grid address: {api_base}"), hint: None });
    }
    // --code flag > VISIGRID_GRID_CODE env > prompt
    let code = match code.or_else(|| std::env::var("VISIGRID_GRID_CODE").ok()) {
        Some(c) => c,
        None if atty::is(atty::Stream::Stdin) => {
            eprintln!("Authorize this computer at {api_base}/desktop/authorize, then paste the code.");
            eprint!("Grid code: ");
            io::stderr().flush().ok();
            let mut buf = String::new();
            io::stdin()
                .read_line(&mut buf)
                .map_err(|e| CliError { code: EXIT_ERROR, message: e.to_string(), hint: None })?;
            buf
        }
        None => {
            return Err(CliError {
                code: EXIT_USAGE,
                message: "No Grid code provided and stdin is not a TTY".into(),
                hint: Some("pass --code or set VISIGRID_GRID_CODE".into()),
            })
        }
    };
    let code = code.trim().to_string();
    if code.is_empty() {
        return Err(CliError { code: EXIT_USAGE, message: "No Grid code provided".into(), hint: None });
    }
    let creds = GridCredentials { secret: code, api_base: api_base.clone() };
    GridClient::new(creds.clone()).verify().map_err(|e| match e {
        GridError::SignedOut => CliError {
            code: EXIT_HUB_NOT_AUTH,
            message: "Grid did not accept that code".into(),
            hint: Some(format!("authorize again at {api_base}/desktop/authorize")),
        },
        GridError::Network(msg) => CliError { code: EXIT_HUB_NETWORK, message: msg, hint: None },
        other => CliError { code: EXIT_ERROR, message: other.to_string(), hint: None },
    })?;
    grid::save_grid_auth(&creds).map_err(|e| CliError { code: EXIT_ERROR, message: e, hint: None })?;
    eprintln!("Signed in to Grid at {api_base}. Cloud sheets now sync with Grid; Hub stays on `vgrid login`.");
    Ok(())
}

pub fn cmd_logout() -> Result<(), CliError> {
    let Some(creds) = grid::load_grid_auth() else {
        eprintln!("Not signed in to Grid");
        return Ok(());
    };
    let revoked = GridClient::new(creds).sign_out();
    grid::delete_grid_auth().map_err(|e| CliError { code: EXIT_ERROR, message: e, hint: None })?;
    match revoked {
        Ok(()) => eprintln!("Signed out of Grid"),
        Err(e) => eprintln!("Signed out on this computer, but Grid could not be told ({e}). Remove the device in Grid to revoke it."),
    }
    Ok(())
}

pub fn cmd_status() -> Result<(), CliError> {
    match grid::load_grid_auth() {
        Some(creds) => println!("Signed in to Grid at {}", creds.api_base),
        None => println!("Not signed in to Grid"),
    }
    Ok(())
}
