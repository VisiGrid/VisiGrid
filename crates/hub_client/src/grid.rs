//! Grid (Loco) sign-in and cloud sheets — shared between desktop and CLI.
//!
//! Grid is a separate credential from `auth.json`, which holds the Rails Hub
//! token that Hub publish, license checks, `vgrid login` and `serve --share`
//! keep using. Grid's is a device sign-in minted by the Grid authorize page:
//! one opaque secret, good for 30 days, kept in `grid-auth.json` (0600). Each
//! sheet call carries a 10-minute access token refreshed from that secret, so
//! a leaked access token dies quickly and revoking the device in Grid signs
//! this computer out on its next request.
//!
//! The cloud document is visigrid-json (`visigrid_io::json::export_workbook`),
//! the same contract the web uses: never the SQLite `.sheet` file.

use std::path::PathBuf;
use std::sync::Mutex;
use std::time::{Duration, Instant};

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

/// Saves up to this size go inline; larger ones through a signed upload, as
/// the web does.
pub const INLINE_SAVE_BYTES: usize = 1024 * 1024;
/// Grid's upload limit.
pub const MAX_SAVE_BYTES: usize = 32 * 1024 * 1024;
/// Refresh this long before the access token runs out.
const REFRESH_MARGIN: Duration = Duration::from_secs(30);

/// The device secret and the Grid it belongs to.
#[derive(Clone, Serialize, Deserialize, PartialEq)]
pub struct GridCredentials {
    /// Opaque device secret, as the authorize page showed it once.
    pub secret: String,
    /// Grid API base, e.g. `https://grid.example`. Always explicit: there is
    /// no default, and it is never the Rails Hub base.
    pub api_base: String,
}

// The secret is a credential: keep it out of logs and panics.
impl std::fmt::Debug for GridCredentials {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("GridCredentials").field("api_base", &self.api_base).finish_non_exhaustive()
    }
}

/// `~/.config/visigrid/grid-auth.json`, beside (never inside) `auth.json`.
pub fn grid_auth_path() -> Option<PathBuf> {
    dirs::config_dir().map(|c| c.join("visigrid/grid-auth.json"))
}

pub fn load_grid_auth() -> Option<GridCredentials> {
    load_grid_auth_from(&grid_auth_path()?)
}

pub fn load_grid_auth_from(path: &std::path::Path) -> Option<GridCredentials> {
    serde_json::from_str(&std::fs::read_to_string(path).ok()?).ok()
}

pub fn save_grid_auth(creds: &GridCredentials) -> Result<(), String> {
    save_grid_auth_to(&grid_auth_path().ok_or("Could not determine config directory")?, creds)
}

/// Written to a 0600 temporary file first and renamed into place, so the
/// secret is never readable by others, even briefly.
pub fn save_grid_auth_to(path: &std::path::Path, creds: &GridCredentials) -> Result<(), String> {
    let parent = path.parent().ok_or("Invalid credential path")?;
    std::fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    let tmp = path.with_extension("json.tmp");
    let json = serde_json::to_string_pretty(creds).map_err(|e| e.to_string())?;
    {
        use std::io::Write;
        let mut options = std::fs::OpenOptions::new();
        options.write(true).create(true).truncate(true);
        #[cfg(unix)]
        {
            use std::os::unix::fs::OpenOptionsExt;
            options.mode(0o600);
        }
        let mut file = options.open(&tmp).map_err(|e| e.to_string())?;
        file.write_all(json.as_bytes()).map_err(|e| e.to_string())?;
        file.sync_all().map_err(|e| e.to_string())?;
    }
    std::fs::rename(&tmp, path).map_err(|e| e.to_string())
}

pub fn delete_grid_auth() -> Result<(), String> {
    match grid_auth_path() {
        Some(path) if path.exists() => std::fs::remove_file(path).map_err(|e| e.to_string()),
        _ => Ok(()),
    }
}

#[derive(Debug)]
pub enum GridError {
    /// No Grid credential on this computer.
    NotSignedIn,
    /// The device was revoked, expired, or its account's sign-ins were reset.
    SignedOut,
    /// The sheet changed since this copy last synced; reload before saving.
    Conflict,
    /// Larger than Grid accepts.
    TooLarge(usize),
    Http(u16, String),
    Network(String),
    Parse(String),
}

impl std::fmt::Display for GridError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NotSignedIn => write!(f, "Not signed in to Grid"),
            Self::SignedOut => write!(f, "This computer was signed out of Grid; sign in again"),
            Self::Conflict => write!(f, "The sheet changed in Grid since this copy last synced"),
            Self::TooLarge(n) => write!(f, "The sheet is {n} bytes; Grid accepts up to {MAX_SAVE_BYTES}"),
            Self::Http(status, body) => write!(f, "Grid returned {status}: {body}"),
            Self::Network(e) => write!(f, "Network error: {e}"),
            Self::Parse(e) => write!(f, "Unexpected response from Grid: {e}"),
        }
    }
}

impl std::error::Error for GridError {}

/// A sheet as Grid lists it. Grid sends more fields; only these are read.
#[derive(Debug, Clone, Deserialize, PartialEq)]
pub struct GridSheet {
    /// Grid's numeric row id (not an API identifier; the pid is).
    #[serde(default)]
    pub id: i64,
    pub pid: String,
    pub name: String,
    pub revision: i64,
    #[serde(default)]
    pub updated_at: Option<String>,
}

#[derive(Deserialize)]
struct Access {
    token: String,
    expires_in: u64,
}

#[derive(Deserialize)]
struct Upload {
    upload_id: String,
    sheet_id: String,
    expected_revision: i64,
    url: String,
}

pub struct GridClient {
    http: reqwest::blocking::Client,
    creds: GridCredentials,
    access: Mutex<Option<(String, Instant)>>,
}

impl GridClient {
    pub fn from_saved() -> Result<Self, GridError> {
        Ok(Self::new(load_grid_auth().ok_or(GridError::NotSignedIn)?))
    }

    pub fn new(creds: GridCredentials) -> Self {
        let http = reqwest::blocking::Client::builder()
            .user_agent(concat!("VisiGrid/", env!("CARGO_PKG_VERSION")))
            .timeout(Duration::from_secs(60))
            .build()
            .expect("Failed to create HTTP client");
        Self { http, creds, access: Mutex::new(None) }
    }

    pub fn api_base(&self) -> &str {
        self.creds.api_base.trim_end_matches('/')
    }

    fn url(&self, path: &str) -> String {
        format!("{}{}", self.api_base(), path)
    }

    /// A current access token, exchanging the secret when there is none or
    /// it is about to expire.
    fn access_token(&self, force: bool) -> Result<String, GridError> {
        let mut access = self.access.lock().expect("access token lock");
        if let Some((token, until)) = access.as_ref() {
            if !force && Instant::now() + REFRESH_MARGIN < *until {
                return Ok(token.clone());
            }
        }
        let response = self
            .http
            .post(self.url("/api/device-sessions/token"))
            .json(&serde_json::json!({ "secret": self.creds.secret }))
            .send()
            .map_err(|e| GridError::Network(e.to_string()))?;
        if response.status().as_u16() == 401 {
            *access = None;
            return Err(GridError::SignedOut);
        }
        let fresh: Access = read_json(response)?;
        *access = Some((fresh.token.clone(), Instant::now() + Duration::from_secs(fresh.expires_in)));
        Ok(fresh.token)
    }

    /// Send a sheet request with the access token, refreshing once if Grid
    /// rejects it (it may have expired, or been revoked a moment ago).
    fn send(
        &self,
        build: impl Fn(&str) -> reqwest::blocking::RequestBuilder,
    ) -> Result<reqwest::blocking::Response, GridError> {
        let token = self.access_token(false)?;
        let response = build(&token).send().map_err(|e| GridError::Network(e.to_string()))?;
        if response.status().as_u16() != 401 {
            return Ok(response);
        }
        let token = self.access_token(true)?;
        let response = build(&token).send().map_err(|e| GridError::Network(e.to_string()))?;
        if response.status().as_u16() == 401 {
            return Err(GridError::SignedOut);
        }
        Ok(response)
    }

    /// Prove the secret works (used when signing in).
    pub fn verify(&self) -> Result<(), GridError> {
        self.access_token(true).map(|_| ())
    }

    pub fn list(&self) -> Result<Vec<GridSheet>, GridError> {
        read_json(self.send(|t| self.http.get(self.url("/api/sheets")).bearer_auth(t))?)
    }

    pub fn get(&self, pid: &str) -> Result<GridSheet, GridError> {
        let url = self.url(&format!("/api/sheets/{}", checked_pid(pid)?));
        read_json(self.send(|t| self.http.get(&url).bearer_auth(t))?)
    }

    /// The sheet's visigrid-json document, as bytes.
    pub fn data(&self, pid: &str) -> Result<Vec<u8>, GridError> {
        let url = self.url(&format!("/api/sheets/{}/data", checked_pid(pid)?));
        let response = ok(self.send(|t| self.http.get(&url).bearer_auth(t))?)?;
        response.bytes().map(|b| b.to_vec()).map_err(|e| GridError::Network(e.to_string()))
    }

    /// Map an old Rails sheet id (numeric, or the 22-character public id) to
    /// its Grid pid. Only sheets this account can still open resolve.
    pub fn resolve_legacy(&self, identifier: &str) -> Result<String, GridError> {
        if identifier.is_empty() || !identifier.bytes().all(|b| b.is_ascii_alphanumeric() || b == b'_' || b == b'-') {
            return Err(GridError::Http(404, "Unknown sheet".into()));
        }
        let url = self.url(&format!("/api/sheets/legacy/{identifier}"));
        let body: serde_json::Value = read_json(self.send(|t| self.http.get(&url).bearer_auth(t))?)?;
        body["pid"]
            .as_str()
            .map(str::to_string)
            .ok_or_else(|| GridError::Parse("legacy response without pid".into()))
    }

    pub fn create(&self, name: &str, document: &serde_json::Value) -> Result<GridSheet, GridError> {
        let body = serde_json::json!({ "folder_id": null, "name": name, "document": document });
        read_json(self.send(|t| self.http.post(self.url("/api/sheets")).bearer_auth(t).json(&body))?)
    }

    /// Save a visigrid-json document over `expected_revision`. Inline up to
    /// `INLINE_SAVE_BYTES`, through a signed upload above it.
    pub fn save(&self, pid: &str, expected_revision: i64, document: &[u8]) -> Result<GridSheet, GridError> {
        let pid = checked_pid(pid)?;
        if document.len() > MAX_SAVE_BYTES {
            return Err(GridError::TooLarge(document.len()));
        }
        let saved: GridSheet = if document.len() <= INLINE_SAVE_BYTES {
            let parsed: serde_json::Value =
                serde_json::from_slice(document).map_err(|e| GridError::Parse(e.to_string()))?;
            let body = serde_json::json!({ "expected_revision": expected_revision, "document": parsed });
            let url = self.url(&format!("/api/sheets/{pid}/save"));
            read_json(self.send(|t| self.http.post(&url).bearer_auth(t).json(&body))?)?
        } else {
            self.save_upload(pid, expected_revision, document)?
        };
        if saved.pid != pid || saved.revision != expected_revision + 1 {
            return Err(GridError::Parse("save landed on another revision".into()));
        }
        Ok(saved)
    }

    fn save_upload(&self, pid: &str, expected_revision: i64, document: &[u8]) -> Result<GridSheet, GridError> {
        let sha256 = hex(&Sha256::digest(document));
        let body = serde_json::json!({ "expected_revision": expected_revision, "byte_size": document.len(), "sha256": sha256 });
        let url = self.url(&format!("/api/sheets/{pid}/uploads"));
        let upload: Upload = read_json(self.send(|t| self.http.post(&url).bearer_auth(t).json(&body))?)?;
        if upload.sheet_id != pid || upload.expected_revision != expected_revision || checked_pid(&upload.upload_id).is_err() {
            return Err(GridError::Parse("upload reserved for another sheet".into()));
        }
        let target = reqwest::Url::parse(&upload.url).map_err(|e| GridError::Parse(e.to_string()))?;
        if !matches!(target.scheme(), "https" | "http") {
            return Err(GridError::Parse("invalid upload address".into()));
        }
        // The signed URL is the credential here: no bearer token goes to storage.
        let put = self
            .http
            .put(target)
            .header("content-type", "application/json")
            .body(document.to_vec())
            .send()
            .map_err(|e| GridError::Network(e.to_string()))?;
        ok(put)?;
        let url = self.url(&format!("/api/sheets/{pid}/uploads/{}/complete", upload.upload_id));
        read_json(self.send(|t| self.http.post(&url).bearer_auth(t).json(&serde_json::json!({})))?)
    }

    /// Sign this computer out of Grid. Revokes the device server-side; the
    /// caller deletes the local credential.
    pub fn sign_out(&self) -> Result<(), GridError> {
        let response = self
            .http
            .post(self.url("/api/device-sessions/revoke"))
            .json(&serde_json::json!({ "secret": self.creds.secret }))
            .send()
            .map_err(|e| GridError::Network(e.to_string()))?;
        ok(response).map(|_| ())
    }
}

/// Grid pids are UUIDs; anything else never reaches a URL.
fn checked_pid(pid: &str) -> Result<&str, GridError> {
    let ok = pid.len() == 36
        && pid.bytes().enumerate().all(|(i, b)| match i {
            8 | 13 | 18 | 23 => b == b'-',
            _ => b.is_ascii_hexdigit(),
        });
    if ok { Ok(pid) } else { Err(GridError::Parse(format!("invalid sheet id {pid:?}"))) }
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

fn ok(response: reqwest::blocking::Response) -> Result<reqwest::blocking::Response, GridError> {
    let status = response.status().as_u16();
    if response.status().is_success() {
        return Ok(response);
    }
    let body = response.text().unwrap_or_default();
    Err(match status {
        401 => GridError::SignedOut,
        409 => GridError::Conflict,
        _ => GridError::Http(status, body),
    })
}

fn read_json<T: serde::de::DeserializeOwned>(response: reqwest::blocking::Response) -> Result<T, GridError> {
    ok(response)?.json().map_err(|e| GridError::Parse(e.to_string()))
}
