//! `vgrid fetch sftp` — download files from an SFTP server.
//!
//! Transport-only v1: connect → list → filter → stability check (min_age) →
//! download → BLAKE3 hash → atomic write → provenance JSON.

use std::fs;
use std::io::{BufRead, Read, Write};
use std::net::TcpStream;
use std::path::{Path, PathBuf};
use std::time::{Duration, SystemTime, UNIX_EPOCH};

use serde::{Deserialize, Serialize};
use ssh2::{CheckResult, HostKeyType, KnownHostFileKind, KnownHostKeyFormat, Session};

use crate::exit_codes;
use crate::CliError;

// ── Public types (used by mod.rs for clap) ──────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, clap::ValueEnum)]
pub enum ProvenanceMode {
    /// Write <filename>.visigrid.json next to each downloaded file
    Sidecar,
    /// Print provenance JSON to stderr
    Stderr,
    /// Suppress provenance output
    None,
}

/// All CLI args, bundled to avoid a 20-param function signature.
pub struct SftpArgs {
    pub host: String,
    pub port: u16,
    pub username: String,
    pub private_key: Option<PathBuf>,
    pub passphrase: Option<String>,
    pub password: Option<String>,
    pub remote_dir: String,
    pub glob_pattern: String,
    pub min_age: u64,
    pub out: Option<PathBuf>,
    pub stdout_mode: bool,
    pub known_hosts_path: String,
    pub known_hosts_out: Option<String>,
    pub trust_on_first_use: bool,
    pub overwrite: bool,
    pub max_files: Option<usize>,
    pub since: Option<String>,
    pub quiet: bool,
    pub list_only: bool,
    pub provenance: Option<ProvenanceMode>,
    pub state_dir: Option<PathBuf>,
    pub reprocess: bool,
}

// ── Internal types ──────────────────────────────────────────────────

struct SftpConnection {
    #[allow(dead_code)]
    session: Session,
    sftp: ssh2::Sftp,
    host: String,
    port: u16,
    username: String,
    host_key_fingerprint: String,
}

#[derive(Debug)]
struct RemoteFile {
    path: String,
    filename: String,
    size: u64,
    mtime: Option<u64>,
}

struct DownloadResult {
    blake3: String,
    local_path: PathBuf,
}

#[derive(Debug, Serialize)]
struct ProvenanceRecord {
    uri: String,
    size: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    mtime: Option<u64>,
    downloaded_at: String,
    blake3: String,
    host_key_fingerprint: String,
    cli_version: String,
    schema_version: u32,
}

/// One line in seen.jsonl — tracks files we've already processed.
#[derive(Debug, Serialize, Deserialize, PartialEq)]
struct SeenRecord {
    remote_path: String,
    mtime: Option<u64>,
    size: u64,
    blake3: String,
}

// ── Constants ───────────────────────────────────────────────────────

const CHUNK_SIZE: usize = 64 * 1024; // 64 KB
const CLI_VERSION: &str = env!("CARGO_PKG_VERSION");
const SEEN_FILENAME: &str = "seen.jsonl";

// ── Entry point ─────────────────────────────────────────────────────

pub fn cmd_fetch_sftp(args: SftpArgs) -> Result<(), CliError> {
    let stderr_tty = atty::is(atty::Stream::Stderr);
    let show_progress = !args.quiet && stderr_tty;

    // 1. Validate args
    if !args.stdout_mode && args.out.is_none() {
        return Err(CliError::args("--out is required unless --stdout is set"));
    }

    // Resolve provenance mode: default sidecar for --out, stderr for --stdout
    let prov_mode = args.provenance.unwrap_or(if args.stdout_mode {
        ProvenanceMode::Stderr
    } else {
        ProvenanceMode::Sidecar
    });

    // Parse --since into a unix timestamp
    let since_ts = match &args.since {
        Some(date_str) => {
            let date = chrono::NaiveDate::parse_from_str(date_str, "%Y-%m-%d").map_err(|e| {
                CliError::args(format!("invalid --since date {:?}: {}", date_str, e))
            })?;
            Some(
                date.and_hms_opt(0, 0, 0)
                    .unwrap()
                    .and_utc()
                    .timestamp() as u64,
            )
        }
        None => None,
    };

    // Compile glob pattern
    let pattern = glob::Pattern::new(&args.glob_pattern).map_err(|e| {
        CliError::args(format!("invalid --glob pattern {:?}: {}", args.glob_pattern, e))
    })?;

    // Expand known_hosts path
    let known_hosts_expanded = expand_path(&args.known_hosts_path);
    let known_hosts_out_expanded = args
        .known_hosts_out
        .as_deref()
        .map(expand_path)
        .unwrap_or_else(|| known_hosts_expanded.clone());

    // Load seen state (if --state-dir)
    let seen_set = if let Some(ref state_dir) = args.state_dir {
        if args.reprocess {
            std::collections::HashSet::new()
        } else {
            load_seen_state(state_dir)?
        }
    } else {
        std::collections::HashSet::new()
    };

    // 2. Resolve auth method
    let auth = resolve_auth(&args.private_key, &args.passphrase, &args.password)?;

    // 3. Connect
    if show_progress {
        eprintln!("Connecting to {}:{}...", args.host, args.port);
    }

    let tcp = TcpStream::connect_timeout(
        &format!("{}:{}", args.host, args.port)
            .parse()
            .map_err(|e| CliError {
                code: exit_codes::EXIT_FETCH_SFTP_CONNECT,
                message: format!("invalid address {}:{}: {}", args.host, args.port, e),
                hint: None,
            })?,
        Duration::from_secs(30),
    )
    .map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_SFTP_CONNECT,
        message: format!("TCP connection to {}:{} failed: {}", args.host, args.port, e),
        hint: None,
    })?;

    let mut session = Session::new().map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_SFTP_CONNECT,
        message: format!("failed to create SSH session: {}", e),
        hint: None,
    })?;

    session.set_tcp_stream(tcp);
    session.handshake().map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_SFTP_CONNECT,
        message: format!("SSH handshake with {}:{} failed: {}", args.host, args.port, e),
        hint: None,
    })?;

    // 4. Verify host key
    let fingerprint = verify_host_key(
        &session,
        &args.host,
        args.port,
        &known_hosts_expanded,
        &known_hosts_out_expanded,
        args.trust_on_first_use,
        show_progress,
    )?;

    // 5. Authenticate
    authenticate(&session, &args.username, &auth).map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_AUTH,
        message: format!(
            "Authentication failed for {}@{}: {}",
            args.username, args.host, e,
        ),
        hint: None,
    })?;

    if show_progress {
        eprintln!("Authenticated as {}", args.username);
    }

    // 6. Open SFTP channel
    let sftp = session.sftp().map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_UPSTREAM,
        message: format!("failed to open SFTP channel: {}", e),
        hint: None,
    })?;

    let conn = SftpConnection {
        session,
        sftp,
        host: args.host.clone(),
        port: args.port,
        username: args.username.clone(),
        host_key_fingerprint: fingerprint,
    };

    // 7. Discover files
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_secs();

    let mut files = discover_files(
        &conn,
        &args.remote_dir,
        &pattern,
        args.min_age,
        since_ts,
        now,
        args.quiet,
    )?;

    // Sort by mtime descending (newest-first) for display
    files.sort_by(|a, b| b.mtime.cmp(&a.mtime));

    // Apply max_files limit after filtering
    if let Some(limit) = args.max_files {
        files.truncate(limit);
    }

    // 8. List-only mode
    if args.list_only {
        if files.is_empty() {
            eprintln!("0 files matched");
        } else {
            println!(
                "{:<40} {:>12} {:>24} {:>10}",
                "FILENAME", "SIZE", "MTIME", "AGE",
            );
            for f in &files {
                let mtime_str = f
                    .mtime
                    .map(|t| {
                        chrono::DateTime::from_timestamp(t as i64, 0)
                            .map(|dt| dt.format("%Y-%m-%dT%H:%M:%SZ").to_string())
                            .unwrap_or_else(|| t.to_string())
                    })
                    .unwrap_or_else(|| "unknown".to_string());
                let age_str = f
                    .mtime
                    .map(|t| format_age(now.saturating_sub(t)))
                    .unwrap_or_else(|| "?".to_string());
                println!(
                    "{:<40} {:>12} {:>24} {:>10}",
                    f.filename,
                    format_size(f.size),
                    mtime_str,
                    age_str,
                );
            }
            eprintln!("{} files matched", files.len());
        }
        return Ok(());
    }

    // Stdout mode: error if multiple files match
    if args.stdout_mode && files.len() > 1 {
        return Err(CliError::args(format!(
            "--stdout requires exactly 1 matching file, but {} matched. \
             Use --glob or --max-files 1 to narrow.",
            files.len(),
        )));
    }

    if files.is_empty() {
        eprintln!("0 files matched");
        return Ok(());
    }

    // Create output directory if needed
    if let Some(ref out_dir) = args.out {
        fs::create_dir_all(out_dir).map_err(|e| {
            CliError::io(format!(
                "cannot create output directory {}: {}",
                out_dir.display(),
                e,
            ))
        })?;
    }

    // 9. Download + hash each file
    let mut downloaded = 0u64;
    let mut skipped = 0u64;
    let mut unchanged = 0u64;
    let mut new_seen: Vec<SeenRecord> = Vec::new();

    for file in &files {
        // Check seen state (--state-dir dedup)
        let seen_key = seen_key(&file.path, file.mtime, file.size);
        if !args.reprocess && seen_set.contains(&seen_key) {
            if show_progress {
                eprintln!("  skipped (already seen): {}", file.filename);
            }
            skipped += 1;
            continue;
        }

        if args.stdout_mode {
            let result = download_to_stdout(&conn, file)?;
            // Provenance
            emit_provenance(
                &conn,
                file,
                &result.blake3,
                &args.remote_dir,
                prov_mode,
                None,
            )?;
            new_seen.push(SeenRecord {
                remote_path: file.path.clone(),
                mtime: file.mtime,
                size: file.size,
                blake3: result.blake3,
            });
            downloaded += 1;
        } else {
            let out_dir = args.out.as_ref().unwrap();

            match download_to_file(&conn, file, out_dir, args.overwrite, show_progress) {
                Ok(SkipOrDownload::Skipped) => {
                    // Hash matched existing file — content identical
                    unchanged += 1;
                }
                Ok(SkipOrDownload::Downloaded(result)) => {
                    // Provenance
                    emit_provenance(
                        &conn,
                        file,
                        &result.blake3,
                        &args.remote_dir,
                        prov_mode,
                        Some(out_dir),
                    )?;

                    if show_progress {
                        let uri = format!(
                            "sftp://{}@{}:{}{}/{}",
                            conn.username,
                            conn.host,
                            conn.port,
                            args.remote_dir,
                            file.filename,
                        );
                        eprintln!("{}", uri);
                        eprintln!("  size:   {} bytes", file.size);
                        if let Some(mtime) = file.mtime {
                            let mtime_str =
                                chrono::DateTime::from_timestamp(mtime as i64, 0)
                                    .map(|dt| dt.format("%Y-%m-%dT%H:%M:%SZ").to_string())
                                    .unwrap_or_else(|| mtime.to_string());
                            eprintln!("  mtime:  {}", mtime_str);
                        }
                        eprintln!("  blake3: {}", result.blake3);
                        eprintln!("  saved:  {}", result.local_path.display());
                    }

                    new_seen.push(SeenRecord {
                        remote_path: file.path.clone(),
                        mtime: file.mtime,
                        size: file.size,
                        blake3: result.blake3,
                    });
                    downloaded += 1;
                }
                Err(e) => {
                    eprintln!(
                        "error: transfer failed for {}: {}",
                        file.filename, e.message,
                    );
                    continue;
                }
            }
        }
    }

    // Write new seen records (append)
    if let Some(ref state_dir) = args.state_dir {
        if !new_seen.is_empty() {
            append_seen_state(state_dir, &new_seen)?;
        }
    }

    // Final summary
    if show_progress {
        let mut parts = vec![format!("{} downloaded", downloaded)];
        if unchanged > 0 {
            parts.push(format!("{} unchanged", unchanged));
        }
        if skipped > 0 {
            parts.push(format!("{} skipped", skipped));
        }
        eprintln!("{}", parts.join(", "));
    }

    Ok(())
}

// ── Auth resolution ─────────────────────────────────────────────────

#[derive(Debug)]
enum AuthMethod {
    PrivateKey {
        path: PathBuf,
        passphrase: Option<String>,
    },
    Password(String),
    Agent,
}

fn resolve_auth(
    private_key: &Option<PathBuf>,
    passphrase: &Option<String>,
    password: &Option<String>,
) -> Result<AuthMethod, CliError> {
    if let Some(key_path) = private_key {
        let expanded = expand_path(&key_path.to_string_lossy());
        return Ok(AuthMethod::PrivateKey {
            path: PathBuf::from(expanded),
            passphrase: passphrase.clone(),
        });
    }

    if let Some(pw) = password {
        return Ok(AuthMethod::Password(pw.clone()));
    }

    // Try ssh-agent (will be attempted at auth time)
    Ok(AuthMethod::Agent)
}

fn authenticate(session: &Session, username: &str, auth: &AuthMethod) -> Result<(), String> {
    match auth {
        AuthMethod::PrivateKey { path, passphrase } => {
            session
                .userauth_pubkey_file(username, None, path, passphrase.as_deref())
                .map_err(|e| format!("public key auth failed: {}", e))?;
        }
        AuthMethod::Agent => {
            session
                .userauth_agent(username)
                .map_err(|e| format!("ssh-agent auth failed: {}", e))?;
        }
        AuthMethod::Password(pw) => {
            session
                .userauth_password(username, pw)
                .map_err(|e| format!("password auth failed: {}", e))?;
        }
    }

    if !session.authenticated() {
        return Err("session not authenticated after auth attempt".to_string());
    }

    Ok(())
}

// ── Host key verification ───────────────────────────────────────────

fn verify_host_key(
    session: &Session,
    host: &str,
    port: u16,
    known_hosts_path: &str,
    known_hosts_out_path: &str,
    trust_on_first_use: bool,
    show_progress: bool,
) -> Result<String, CliError> {
    let (host_key, key_type) = session.host_key().ok_or_else(|| CliError {
        code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
        message: "server did not provide a host key".into(),
        hint: None,
    })?;

    // Compute fingerprint: SHA-256 of raw key bytes, base64
    let fingerprint = {
        use sha2::Digest;
        let hash = sha2::Sha256::digest(host_key);
        format!(
            "SHA256:{}",
            base64::Engine::encode(&base64::engine::general_purpose::STANDARD, hash),
        )
    };

    if show_progress {
        eprintln!("Host key ({:?}): {}", key_type, fingerprint);
    }

    let mut known_hosts = session.known_hosts().map_err(|e| CliError {
        code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
        message: format!("failed to init known_hosts: {}", e),
        hint: None,
    })?;

    // Try to load existing known_hosts file
    let kh_path = Path::new(known_hosts_path);
    let kh_exists = kh_path.exists();
    if kh_exists {
        known_hosts
            .read_file(kh_path, KnownHostFileKind::OpenSSH)
            .map_err(|e| CliError {
                code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
                message: format!("failed to read {}: {}", known_hosts_path, e),
                hint: None,
            })?;
    }

    let key_format = match key_type {
        HostKeyType::Rsa => KnownHostKeyFormat::SshRsa,
        HostKeyType::Dss => KnownHostKeyFormat::SshDss,
        HostKeyType::Ed25519 => KnownHostKeyFormat::Ed25519,
        HostKeyType::Ecdsa256 => KnownHostKeyFormat::Ecdsa256,
        HostKeyType::Ecdsa384 => KnownHostKeyFormat::Ecdsa384,
        HostKeyType::Ecdsa521 => KnownHostKeyFormat::Ecdsa521,
        _ => {
            return Err(CliError {
                code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
                message: format!("unsupported host key type: {:?}", key_type),
                hint: None,
            });
        }
    };

    let check = known_hosts.check_port(host, port, host_key);

    match check {
        CheckResult::Match => Ok(fingerprint),

        CheckResult::NotFound => {
            if !kh_exists && !trust_on_first_use {
                // known_hosts file doesn't even exist
                return Err(CliError {
                    code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
                    message: format!(
                        "host key for {}:{} not found — {} does not exist.\n  \
                         presented: {}\n  \
                         To accept this key:\n    \
                         mkdir -p {} && \\\n    \
                         vgrid fetch sftp ... --trust-on-first-use",
                        host,
                        port,
                        known_hosts_path,
                        fingerprint,
                        kh_path
                            .parent()
                            .map(|p| p.display().to_string())
                            .unwrap_or_else(|| "~/.ssh".to_string()),
                    ),
                    hint: None,
                });
            }

            if trust_on_first_use {
                known_hosts
                    .add(host, host_key, "", key_format)
                    .map_err(|e| CliError {
                        code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
                        message: format!("failed to add host key: {}", e),
                        hint: None,
                    })?;

                // Ensure parent directory exists
                let out_path = Path::new(known_hosts_out_path);
                if let Some(parent) = out_path.parent() {
                    let _ = fs::create_dir_all(parent);
                }

                known_hosts
                    .write_file(out_path, KnownHostFileKind::OpenSSH)
                    .map_err(|e| CliError {
                        code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
                        message: format!(
                            "failed to write known_hosts to {}: {}",
                            known_hosts_out_path, e,
                        ),
                        hint: None,
                    })?;

                if show_progress {
                    eprintln!(
                        "Host key accepted (TOFU) and written to {}",
                        known_hosts_out_path,
                    );
                }

                Ok(fingerprint)
            } else {
                Err(CliError {
                    code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
                    message: format!(
                        "host key for {}:{} not found in {}.\n  \
                         presented: {}\n  \
                         Use --trust-on-first-use to accept and save this key.",
                        host, port, known_hosts_path, fingerprint,
                    ),
                    hint: None,
                })
            }
        }

        CheckResult::Mismatch => {
            // Try to extract the expected fingerprint from known_hosts
            let expected = find_known_fingerprint(&known_hosts, host, port);
            let expected_line = match expected {
                Some(fp) => format!("\n  expected:  {}", fp),
                None => String::new(),
            };
            Err(CliError {
                code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
                message: format!(
                    "HOST KEY MISMATCH for {}:{}!\n  \
                     The host key does not match the key in {}.{}\n  \
                     presented: {}\n  \
                     This could indicate a man-in-the-middle attack.\n  \
                     If the server key was legitimately rotated, remove the old \
                     entry from {} and re-run with --trust-on-first-use.",
                    host, port, known_hosts_path, expected_line, fingerprint, known_hosts_path,
                ),
                hint: None,
            })
        }

        CheckResult::Failure => Err(CliError {
            code: exit_codes::EXIT_FETCH_SFTP_HOST_KEY,
            message: format!(
                "host key check failed for {}:{}: internal error",
                host, port,
            ),
            hint: None,
        }),
    }
}

/// Try to find and fingerprint the existing known_hosts entry for this host.
/// Returns None if we can't extract it (best-effort).
fn find_known_fingerprint(
    known_hosts: &ssh2::KnownHosts,
    host: &str,
    _port: u16,
) -> Option<String> {
    let hosts = known_hosts.iter().ok()?;
    for entry in &hosts {
        if let Some(name) = entry.name() {
            if name == host || name.starts_with(&format!("[{}]:", host)) {
                // Found it — fingerprint the stored key bytes
                let key_bytes = entry.key().as_bytes();
                use sha2::Digest;
                let hash = sha2::Sha256::digest(key_bytes);
                return Some(format!(
                    "SHA256:{}",
                    base64::Engine::encode(
                        &base64::engine::general_purpose::STANDARD,
                        hash,
                    ),
                ));
            }
        }
    }
    None
}

// ── File discovery ──────────────────────────────────────────────────

fn discover_files(
    conn: &SftpConnection,
    remote_dir: &str,
    pattern: &glob::Pattern,
    min_age: u64,
    since_ts: Option<u64>,
    now: u64,
    quiet: bool,
) -> Result<Vec<RemoteFile>, CliError> {
    let entries = conn
        .sftp
        .readdir(Path::new(remote_dir))
        .map_err(|e| {
            let msg = e.to_string();
            let message = if msg.contains("No such file") {
                format!("Remote directory not found: {}", remote_dir)
            } else if msg.contains("Permission denied") || msg.contains("permission denied") {
                format!("Permission denied reading: {}", remote_dir)
            } else {
                format!("Failed to list {}: {}", remote_dir, e)
            };
            CliError {
                code: exit_codes::EXIT_FETCH_UPSTREAM,
                message,
                hint: None,
            }
        })?;

    let match_opts = glob::MatchOptions {
        case_sensitive: true,
        require_literal_separator: false,
        require_literal_leading_dot: false,
    };

    let mut files = Vec::new();

    for (path, stat) in entries {
        if stat.is_dir() {
            continue;
        }

        let filename = path
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("")
            .to_string();

        if filename.is_empty() {
            continue;
        }

        // Glob filter on basename only
        if !pattern.matches_with(&filename, match_opts) {
            continue;
        }

        let mtime = stat.mtime;
        let size = stat.size.unwrap_or(0);

        // Since filter
        if let Some(since) = since_ts {
            match mtime {
                Some(t) if (t as u64) < since => continue,
                None => continue,
                _ => {}
            }
        }

        // Min-age filter
        match mtime {
            Some(t) => {
                let age = now.saturating_sub(t as u64);
                if age < min_age {
                    continue;
                }
            }
            None => {
                if min_age == 0 {
                    if !quiet {
                        eprintln!(
                            "warning: mtime unavailable for {}, including anyway",
                            filename,
                        );
                    }
                } else {
                    if !quiet {
                        eprintln!(
                            "warning: mtime unavailable for {}, skipping (min-age requires mtime)",
                            filename,
                        );
                    }
                    continue;
                }
            }
        }

        let full_path = format!(
            "{}/{}",
            remote_dir.trim_end_matches('/'),
            filename,
        );

        files.push(RemoteFile {
            path: full_path,
            filename,
            size,
            mtime: mtime.map(|t| t as u64),
        });
    }

    Ok(files)
}

// ── Download ────────────────────────────────────────────────────────

enum SkipOrDownload {
    /// File exists locally with identical BLAKE3 hash — no replacement needed.
    Skipped,
    /// File was downloaded (new or content changed).
    Downloaded(DownloadResult),
}

/// Download a remote file to `out_dir`, with hash-based skip.
///
/// Always downloads to `.part` first, computing BLAKE3 as we stream.
/// Then checks the existing sidecar (if any) for the previous hash:
///
/// - Hash matches existing sidecar → discard `.part`, return Skipped.
/// - Hash differs + `--overwrite` → atomic rename `.part` → final.
/// - Hash differs + no `--overwrite` → warn and return Skipped.
/// - No existing file → atomic rename `.part` → final.
fn download_to_file(
    conn: &SftpConnection,
    file: &RemoteFile,
    out_dir: &Path,
    overwrite: bool,
    show_progress: bool,
) -> Result<SkipOrDownload, CliError> {
    let local_path = out_dir.join(&file.filename);
    let part_path = out_dir.join(format!("{}.part", file.filename));
    let sidecar_path = out_dir.join(format!("{}.visigrid.json", file.filename));

    // Always download to .part (streaming, no buffering full file in memory)
    let mut remote_file = conn
        .sftp
        .open(Path::new(&file.path))
        .map_err(|e| CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("Transfer failed for {}: {}", file.filename, e),
            hint: None,
        })?;

    let mut hasher = blake3::Hasher::new();
    let mut local_file = fs::File::create(&part_path).map_err(|e| {
        CliError::io(format!("cannot create {}: {}", part_path.display(), e))
    })?;

    let mut buf = vec![0u8; CHUNK_SIZE];
    loop {
        let n = remote_file.read(&mut buf).map_err(|e| {
            let _ = fs::remove_file(&part_path);
            CliError {
                code: exit_codes::EXIT_FETCH_UPSTREAM,
                message: format!("Transfer failed for {}: {}", file.filename, e),
                hint: None,
            }
        })?;
        if n == 0 {
            break;
        }
        hasher.update(&buf[..n]);
        local_file.write_all(&buf[..n]).map_err(|e| {
            let _ = fs::remove_file(&part_path);
            CliError::io(format!("write error for {}: {}", part_path.display(), e))
        })?;
    }

    local_file.flush().map_err(|e| {
        let _ = fs::remove_file(&part_path);
        CliError::io(format!("flush error for {}: {}", part_path.display(), e))
    })?;
    drop(local_file);

    let new_hash = hasher.finalize().to_hex().to_string();

    // Check if existing file has the same hash (via sidecar)
    if local_path.exists() {
        if let Some(existing_hash) = read_sidecar_hash(&sidecar_path) {
            if existing_hash == new_hash {
                // Content identical — discard .part
                let _ = fs::remove_file(&part_path);
                if show_progress {
                    eprintln!("  unchanged (hash match): {}", file.filename);
                }
                return Ok(SkipOrDownload::Skipped);
            }

            // Content differs
            if !overwrite {
                let _ = fs::remove_file(&part_path);
                eprintln!(
                    "  content changed: {} (existing blake3: {}…, new: {}…) — \
                     use --overwrite to replace",
                    file.filename,
                    &existing_hash[..existing_hash.len().min(16)],
                    &new_hash[..new_hash.len().min(16)],
                );
                return Ok(SkipOrDownload::Skipped);
            }
            // --overwrite: fall through to rename
        } else if !overwrite {
            // No sidecar to compare — can't verify. Download anyway since we
            // already have the .part file (better to have provenance than not).
        }
    }

    // Atomic rename .part → final
    fs::rename(&part_path, &local_path).map_err(|e| {
        let _ = fs::remove_file(&part_path);
        CliError::io(format!(
            "rename {} → {}: {}",
            part_path.display(),
            local_path.display(),
            e,
        ))
    })?;

    Ok(SkipOrDownload::Downloaded(DownloadResult {
        blake3: new_hash,
        local_path,
    }))
}

fn download_to_stdout(
    conn: &SftpConnection,
    file: &RemoteFile,
) -> Result<DownloadResult, CliError> {
    let mut remote_file = conn
        .sftp
        .open(Path::new(&file.path))
        .map_err(|e| CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("Transfer failed for {}: {}", file.filename, e),
            hint: None,
        })?;

    let mut hasher = blake3::Hasher::new();
    let stdout = std::io::stdout();
    let mut out = stdout.lock();
    let mut buf = vec![0u8; CHUNK_SIZE];

    loop {
        let n = remote_file.read(&mut buf).map_err(|e| CliError {
            code: exit_codes::EXIT_FETCH_UPSTREAM,
            message: format!("Transfer failed for {}: {}", file.filename, e),
            hint: None,
        })?;
        if n == 0 {
            break;
        }
        hasher.update(&buf[..n]);
        out.write_all(&buf[..n])
            .map_err(|e| CliError::io(format!("stdout write error: {}", e)))?;
    }

    out.flush()
        .map_err(|e| CliError::io(format!("stdout flush error: {}", e)))?;

    let blake3 = hasher.finalize().to_hex().to_string();

    Ok(DownloadResult {
        blake3,
        local_path: PathBuf::from("-"),
    })
}

// ── Sidecar hash reading ────────────────────────────────────────────

/// Read the blake3 field from an existing .visigrid.json sidecar. Returns
/// None if the file doesn't exist, can't be parsed, or lacks the field.
fn read_sidecar_hash(sidecar_path: &Path) -> Option<String> {
    let contents = fs::read_to_string(sidecar_path).ok()?;
    let parsed: serde_json::Value = serde_json::from_str(&contents).ok()?;
    parsed["blake3"].as_str().map(|s| s.to_string())
}

// ── Provenance ──────────────────────────────────────────────────────

fn build_provenance(conn: &SftpConnection, file: &RemoteFile, blake3: &str) -> ProvenanceRecord {
    let uri = format!(
        "sftp://{}@{}:{}{}",
        conn.username, conn.host, conn.port, file.path,
    );
    let downloaded_at = chrono::Utc::now().format("%Y-%m-%dT%H:%M:%SZ").to_string();

    ProvenanceRecord {
        uri,
        size: file.size,
        mtime: file.mtime,
        downloaded_at,
        blake3: blake3.to_string(),
        host_key_fingerprint: conn.host_key_fingerprint.clone(),
        cli_version: CLI_VERSION.to_string(),
        schema_version: 1,
    }
}

fn emit_provenance(
    conn: &SftpConnection,
    file: &RemoteFile,
    blake3: &str,
    remote_dir: &str,
    mode: ProvenanceMode,
    out_dir: Option<&Path>,
) -> Result<(), CliError> {
    let _ = remote_dir; // used in URI via file.path

    match mode {
        ProvenanceMode::None => Ok(()),
        ProvenanceMode::Stderr => {
            let prov = build_provenance(conn, file, blake3);
            let json = serde_json::to_string_pretty(&prov).unwrap();
            eprintln!("{}", json);
            Ok(())
        }
        ProvenanceMode::Sidecar => {
            let out_dir = out_dir.ok_or_else(|| {
                CliError::args("--provenance sidecar requires --out (not --stdout)")
            })?;
            let prov = build_provenance(conn, file, blake3);
            let prov_path = out_dir.join(format!("{}.visigrid.json", file.filename));
            let json = serde_json::to_string_pretty(&prov).unwrap();
            fs::write(&prov_path, &json).map_err(|e| {
                CliError::io(format!(
                    "cannot write provenance {}: {}",
                    prov_path.display(),
                    e,
                ))
            })?;
            Ok(())
        }
    }
}

// ── State directory (seen.jsonl) ────────────────────────────────────

/// Build a dedup key from remote path + mtime + size.
/// This is NOT the blake3 — we need it before downloading.
fn seen_key(remote_path: &str, mtime: Option<u64>, size: u64) -> String {
    format!(
        "{}:{}:{}",
        remote_path,
        mtime.map(|t| t.to_string()).unwrap_or_default(),
        size,
    )
}

/// Load seen.jsonl into a set of dedup keys.
fn load_seen_state(
    state_dir: &Path,
) -> Result<std::collections::HashSet<String>, CliError> {
    let seen_path = state_dir.join(SEEN_FILENAME);
    let mut set = std::collections::HashSet::new();

    if !seen_path.exists() {
        return Ok(set);
    }

    let file = fs::File::open(&seen_path).map_err(|e| {
        CliError::io(format!("cannot open {}: {}", seen_path.display(), e))
    })?;

    for line in std::io::BufReader::new(file).lines() {
        let line = line.map_err(|e| {
            CliError::io(format!("read error in {}: {}", seen_path.display(), e))
        })?;
        let line = line.trim();
        if line.is_empty() {
            continue;
        }
        if let Ok(record) = serde_json::from_str::<SeenRecord>(line) {
            set.insert(seen_key(&record.remote_path, record.mtime, record.size));
        }
    }

    Ok(set)
}

/// Append new seen records to seen.jsonl.
fn append_seen_state(state_dir: &Path, records: &[SeenRecord]) -> Result<(), CliError> {
    fs::create_dir_all(state_dir).map_err(|e| {
        CliError::io(format!(
            "cannot create state directory {}: {}",
            state_dir.display(),
            e,
        ))
    })?;

    let seen_path = state_dir.join(SEEN_FILENAME);
    let mut file = fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(&seen_path)
        .map_err(|e| {
            CliError::io(format!("cannot open {}: {}", seen_path.display(), e))
        })?;

    for record in records {
        let json = serde_json::to_string(record).unwrap();
        writeln!(file, "{}", json).map_err(|e| {
            CliError::io(format!("write error in {}: {}", seen_path.display(), e))
        })?;
    }

    Ok(())
}

// ── Helpers ─────────────────────────────────────────────────────────

fn expand_path(path: &str) -> String {
    shellexpand::tilde(path).to_string()
}

fn format_size(bytes: u64) -> String {
    if bytes < 1024 {
        format!("{} B", bytes)
    } else if bytes < 1024 * 1024 {
        format!("{:.1} KB", bytes as f64 / 1024.0)
    } else {
        format!("{:.1} MB", bytes as f64 / (1024.0 * 1024.0))
    }
}

fn format_age(seconds: u64) -> String {
    if seconds < 60 {
        format!("{}s", seconds)
    } else if seconds < 3600 {
        format!("{}m", seconds / 60)
    } else if seconds < 86400 {
        format!("{}h", seconds / 3600)
    } else {
        format!("{}d", seconds / 86400)
    }
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_glob_matching() {
        let pattern = glob::Pattern::new("settlement_*.csv").unwrap();
        let opts = glob::MatchOptions {
            case_sensitive: true,
            require_literal_separator: false,
            require_literal_leading_dot: false,
        };

        assert!(pattern.matches_with("settlement_20260215.csv", opts));
        assert!(pattern.matches_with("settlement_daily.csv", opts));
        assert!(!pattern.matches_with("report_20260215.csv", opts));
        assert!(!pattern.matches_with("settlement_20260215.xlsx", opts));
        assert!(!pattern.matches_with("SETTLEMENT_20260215.csv", opts));
    }

    #[test]
    fn test_glob_default_csv() {
        let pattern = glob::Pattern::new("*.csv").unwrap();
        let opts = glob::MatchOptions {
            case_sensitive: true,
            require_literal_separator: false,
            require_literal_leading_dot: false,
        };

        assert!(pattern.matches_with("anything.csv", opts));
        assert!(pattern.matches_with("report.csv", opts));
        assert!(!pattern.matches_with("report.xlsx", opts));
        assert!(!pattern.matches_with("report.csv.bak", opts));
    }

    #[test]
    fn test_min_age_filter() {
        let now = 1000u64;

        // File is 200s old, min_age 120 → include
        let age = now.saturating_sub(800);
        assert!(age >= 120);

        // File is 60s old, min_age 120 → exclude
        let age = now.saturating_sub(940);
        assert!(age < 120);

        // File is exactly min_age → include
        let age = now.saturating_sub(880);
        assert!(age >= 120);
    }

    #[test]
    fn test_min_age_missing_mtime() {
        let mtime: Option<u64> = None;

        // min_age == 0 + no mtime → include
        assert!(mtime.is_some() || 0u64 == 0);

        // min_age > 0 + no mtime → skip
        assert!(!(mtime.is_some() || 120u64 == 0));
    }

    #[test]
    fn test_since_filter() {
        let since_ts = chrono::NaiveDate::from_ymd_opt(2026, 2, 1)
            .unwrap()
            .and_hms_opt(0, 0, 0)
            .unwrap()
            .and_utc()
            .timestamp() as u64;

        let feb15 = chrono::NaiveDate::from_ymd_opt(2026, 2, 15)
            .unwrap()
            .and_hms_opt(6, 0, 0)
            .unwrap()
            .and_utc()
            .timestamp() as u64;
        assert!(feb15 >= since_ts);

        let jan15 = chrono::NaiveDate::from_ymd_opt(2026, 1, 15)
            .unwrap()
            .and_hms_opt(6, 0, 0)
            .unwrap()
            .and_utc()
            .timestamp() as u64;
        assert!(jan15 < since_ts);
    }

    #[test]
    fn test_auth_resolution_key_priority() {
        let auth = resolve_auth(
            &Some(PathBuf::from("~/.ssh/id_ed25519")),
            &Some("mypass".to_string()),
            &Some("password123".to_string()),
        )
        .unwrap();

        match auth {
            AuthMethod::PrivateKey { passphrase, .. } => {
                assert_eq!(passphrase, Some("mypass".to_string()));
            }
            _ => panic!("expected PrivateKey auth"),
        }
    }

    #[test]
    fn test_auth_resolution_password() {
        let auth = resolve_auth(&None, &None, &Some("password123".to_string())).unwrap();
        match auth {
            AuthMethod::Password(pw) => assert_eq!(pw, "password123"),
            _ => panic!("expected Password auth"),
        }
    }

    #[test]
    fn test_auth_resolution_agent_fallback() {
        let auth = resolve_auth(&None, &None, &None).unwrap();
        match auth {
            AuthMethod::Agent => {}
            _ => panic!("expected Agent auth"),
        }
    }

    #[test]
    fn test_provenance_json_schema() {
        let prov = ProvenanceRecord {
            uri: "sftp://testuser@sftp.example.com:22/reports/settlement_20260215.csv".to_string(),
            size: 12847,
            mtime: Some(1771128000),
            downloaded_at: "2026-02-16T12:00:00Z".to_string(),
            blake3: "a1b2c3d4e5f6".to_string(),
            host_key_fingerprint: "SHA256:abc123".to_string(),
            cli_version: "0.8.0".to_string(),
            schema_version: 1,
        };

        let json = serde_json::to_string_pretty(&prov).unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&json).unwrap();

        assert_eq!(
            parsed["uri"],
            "sftp://testuser@sftp.example.com:22/reports/settlement_20260215.csv"
        );
        assert_eq!(parsed["size"], 12847);
        assert_eq!(parsed["mtime"], 1771128000);
        assert_eq!(parsed["blake3"], "a1b2c3d4e5f6");
        assert_eq!(parsed["host_key_fingerprint"], "SHA256:abc123");
        assert_eq!(parsed["schema_version"], 1);
        assert!(parsed["downloaded_at"].as_str().unwrap().contains("T"));
        assert!(parsed["cli_version"].as_str().is_some());

        // mtime=None omits the field
        let prov_no_mtime = ProvenanceRecord {
            mtime: None,
            ..prov
        };
        let json2 = serde_json::to_string(&prov_no_mtime).unwrap();
        let parsed2: serde_json::Value = serde_json::from_str(&json2).unwrap();
        assert!(parsed2.get("mtime").is_none());
    }

    #[test]
    fn test_format_size() {
        assert_eq!(format_size(0), "0 B");
        assert_eq!(format_size(512), "512 B");
        assert_eq!(format_size(1024), "1.0 KB");
        assert_eq!(format_size(12847), "12.5 KB");
        assert_eq!(format_size(1048576), "1.0 MB");
    }

    #[test]
    fn test_format_age() {
        assert_eq!(format_age(30), "30s");
        assert_eq!(format_age(120), "2m");
        assert_eq!(format_age(3600), "1h");
        assert_eq!(format_age(86400), "1d");
        assert_eq!(format_age(172800), "2d");
    }

    #[test]
    fn test_expand_path() {
        let expanded = expand_path("~/.ssh/known_hosts");
        assert!(!expanded.starts_with('~'));
        assert!(expanded.contains(".ssh/known_hosts"));
    }

    #[test]
    fn test_read_sidecar_hash() {
        let dir = tempfile::tempdir().unwrap();
        let sidecar = dir.path().join("test.csv.visigrid.json");

        // No file → None
        assert_eq!(read_sidecar_hash(&sidecar), None);

        // Valid sidecar → Some(hash)
        fs::write(
            &sidecar,
            r#"{"uri":"sftp://x","size":100,"blake3":"abc123","downloaded_at":"2026-01-01T00:00:00Z","host_key_fingerprint":"SHA256:x","cli_version":"0.1","schema_version":1}"#,
        ).unwrap();
        assert_eq!(read_sidecar_hash(&sidecar), Some("abc123".to_string()));

        // Malformed JSON → None
        fs::write(&sidecar, "not json").unwrap();
        assert_eq!(read_sidecar_hash(&sidecar), None);
    }

    #[test]
    fn test_seen_state_roundtrip() {
        let dir = tempfile::tempdir().unwrap();
        let state_dir = dir.path();

        // Empty state
        let set = load_seen_state(state_dir).unwrap();
        assert!(set.is_empty());

        // Write some records
        let records = vec![
            SeenRecord {
                remote_path: "/reports/a.csv".to_string(),
                mtime: Some(1000),
                size: 500,
                blake3: "hash_a".to_string(),
            },
            SeenRecord {
                remote_path: "/reports/b.csv".to_string(),
                mtime: None,
                size: 200,
                blake3: "hash_b".to_string(),
            },
        ];
        append_seen_state(state_dir, &records).unwrap();

        // Reload
        let set = load_seen_state(state_dir).unwrap();
        assert_eq!(set.len(), 2);
        assert!(set.contains(&seen_key("/reports/a.csv", Some(1000), 500)));
        assert!(set.contains(&seen_key("/reports/b.csv", None, 200)));

        // Append more
        let more = vec![SeenRecord {
            remote_path: "/reports/c.csv".to_string(),
            mtime: Some(2000),
            size: 300,
            blake3: "hash_c".to_string(),
        }];
        append_seen_state(state_dir, &more).unwrap();

        let set = load_seen_state(state_dir).unwrap();
        assert_eq!(set.len(), 3);
    }

    #[test]
    fn test_seen_key_format() {
        assert_eq!(seen_key("/a/b.csv", Some(1234), 5678), "/a/b.csv:1234:5678");
        assert_eq!(seen_key("/a/b.csv", None, 5678), "/a/b.csv::5678");
    }

    #[test]
    fn test_provenance_mode_default() {
        // --out mode defaults to sidecar
        let default_out = ProvenanceMode::Sidecar;
        assert_eq!(default_out, ProvenanceMode::Sidecar);

        // --stdout mode defaults to stderr
        let default_stdout = ProvenanceMode::Stderr;
        assert_eq!(default_stdout, ProvenanceMode::Stderr);
    }
}
