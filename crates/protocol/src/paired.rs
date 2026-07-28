//! Paired-client credential storage for the session protocol.
//!
//! Pairing replaces the copy-a-token dance: a client with no credential sends
//! `pair_request`, the GUI shows an approval dialog, and on approval the
//! server issues a token that persists across GUI restarts.
//!
//! Two files, both 0600 in the state dir (`~/.local/state/visigrid` on
//! Linux, platform-equivalent elsewhere):
//!
//! - `paired_clients.json` — server side: every approved client and its
//!   token. The server re-reads this on each authentication attempt, so
//!   `vgrid pair --revoke <name>` takes effect on the next connection with
//!   no GUI round-trip. (In-flight connections keep their session.)
//! - `client_token.json` — client side: the single credential this
//!   machine's tools present.
//!
//! Tokens are stored plaintext on both sides: the two files live in the same
//! trust domain (same user, same permissions), so hashing the server copy
//! would add ceremony, not security.

use std::fs;
use std::io;
use std::path::PathBuf;

use serde::{Deserialize, Serialize};

/// One approved client on the server side.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PairedClient {
    /// Human-readable name shown in the approval dialog (e.g. "Claude Code").
    pub name: String,
    /// The bearer token issued at approval.
    pub token: String,
    /// Unix seconds at approval time.
    pub paired_at: u64,
}

/// The credential a client stores after successful pairing.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ClientCredential {
    pub token: String,
    /// Name this client paired under.
    pub name: String,
    /// Unix seconds at approval time.
    pub paired_at: u64,
}

fn state_dir() -> io::Result<PathBuf> {
    #[cfg(target_os = "linux")]
    let base = std::env::var("XDG_STATE_HOME").map(PathBuf::from).unwrap_or_else(|_| {
        dirs::home_dir().unwrap_or_else(|| PathBuf::from("/tmp")).join(".local/state")
    });
    #[cfg(not(target_os = "linux"))]
    let base = dirs::data_local_dir()
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "no local data dir"))?;
    Ok(base.join("visigrid"))
}

pub fn paired_clients_path() -> io::Result<PathBuf> {
    Ok(state_dir()?.join("paired_clients.json"))
}

pub fn client_token_path() -> io::Result<PathBuf> {
    Ok(state_dir()?.join("client_token.json"))
}

fn write_private(path: &PathBuf, contents: &str) -> io::Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    // Write via a same-directory temp file + rename so a concurrent reader
    // never sees a partial file.
    let tmp = path.with_extension("tmp");
    fs::write(&tmp, contents)?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        fs::set_permissions(&tmp, fs::Permissions::from_mode(0o600))?;
    }
    fs::rename(&tmp, path)
}

fn now_unix() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0)
}

/// Generate a fresh 32-byte hex bearer token.
pub fn generate_token() -> String {
    use rand::RngCore;
    let mut bytes = [0u8; 32];
    rand::thread_rng().fill_bytes(&mut bytes);
    bytes.iter().map(|b| format!("{:02x}", b)).collect()
}

// ----------------------------------------------------------------------
// Server side
// ----------------------------------------------------------------------

/// Load the paired-clients list. Missing file = empty list.
pub fn load_paired_clients() -> Vec<PairedClient> {
    let Ok(path) = paired_clients_path() else { return Vec::new() };
    let Ok(contents) = fs::read_to_string(path) else { return Vec::new() };
    serde_json::from_str(&contents).unwrap_or_default()
}

fn save_paired_clients(clients: &[PairedClient]) -> io::Result<()> {
    let path = paired_clients_path()?;
    let contents = serde_json::to_string_pretty(clients)
        .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))?;
    write_private(&path, &contents)
}

/// True if `token` belongs to any paired client. Reads the file fresh each
/// call so revocation applies to the next connection without GUI involvement.
pub fn verify_paired_token(token: &str) -> Option<String> {
    load_paired_clients()
        .into_iter()
        .find(|c| constant_time_eq(c.token.as_bytes(), token.as_bytes()))
        .map(|c| c.name)
}

/// Approve a client: issue a token, persist it, and return it.
/// A repeated name replaces the earlier entry (re-pairing rotates the token).
pub fn approve_client(name: &str) -> io::Result<String> {
    let token = generate_token();
    let mut clients = load_paired_clients();
    clients.retain(|c| c.name != name);
    clients.push(PairedClient {
        name: name.to_string(),
        token: token.clone(),
        paired_at: now_unix(),
    });
    save_paired_clients(&clients)?;
    Ok(token)
}

/// Remove a paired client by name. Returns true if an entry was removed.
pub fn revoke_client(name: &str) -> io::Result<bool> {
    let mut clients = load_paired_clients();
    let before = clients.len();
    clients.retain(|c| c.name != name);
    let removed = clients.len() != before;
    if removed {
        save_paired_clients(&clients)?;
    }
    Ok(removed)
}

fn constant_time_eq(a: &[u8], b: &[u8]) -> bool {
    if a.len() != b.len() {
        return false;
    }
    a.iter().zip(b).fold(0u8, |acc, (x, y)| acc | (x ^ y)) == 0
}

// ----------------------------------------------------------------------
// Client side
// ----------------------------------------------------------------------

/// Load this machine's stored client credential, if any.
pub fn load_client_credential() -> Option<ClientCredential> {
    let path = client_token_path().ok()?;
    let contents = fs::read_to_string(path).ok()?;
    serde_json::from_str(&contents).ok()
}

/// Persist the credential received from a successful pairing.
pub fn store_client_credential(name: &str, token: &str) -> io::Result<()> {
    let cred = ClientCredential {
        token: token.to_string(),
        name: name.to_string(),
        paired_at: now_unix(),
    };
    let contents = serde_json::to_string_pretty(&cred)
        .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))?;
    write_private(&client_token_path()?, &contents)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn token_is_64_hex_chars_and_unique() {
        let a = generate_token();
        let b = generate_token();
        assert_eq!(a.len(), 64);
        assert!(a.chars().all(|c| c.is_ascii_hexdigit()));
        assert_ne!(a, b);
    }

    #[test]
    fn constant_time_eq_basics() {
        assert!(constant_time_eq(b"abc", b"abc"));
        assert!(!constant_time_eq(b"abc", b"abd"));
        assert!(!constant_time_eq(b"abc", b"ab"));
    }

    #[test]
    fn approve_verify_revoke_round_trip() {
        // Isolate the state dir so the test never touches real credentials.
        let dir = std::env::temp_dir().join(format!("vg-paired-test-{}", std::process::id()));
        std::env::set_var("XDG_STATE_HOME", &dir);

        let token = approve_client("Test Client").unwrap();
        assert_eq!(verify_paired_token(&token).as_deref(), Some("Test Client"));
        assert!(verify_paired_token("wrong").is_none());

        // Re-pairing rotates the token
        let token2 = approve_client("Test Client").unwrap();
        assert!(verify_paired_token(&token).is_none());
        assert_eq!(verify_paired_token(&token2).as_deref(), Some("Test Client"));
        assert_eq!(load_paired_clients().len(), 1);

        assert!(revoke_client("Test Client").unwrap());
        assert!(verify_paired_token(&token2).is_none());
        assert!(!revoke_client("Test Client").unwrap());

        let _ = std::fs::remove_dir_all(&dir);
        std::env::remove_var("XDG_STATE_HOME");
    }
}
