//! A PostgreSQL table, view or query as a recipe source (Supabase included).
//!
//! Read-only is enforced by the database, not by VisiGrid: before every
//! read the role must pass a proof, and a role that could write is refused
//! with what let it write. The proof is a catalog check (no write privilege
//! on any table, no CREATE anywhere, no membership in another role, none of
//! superuser/createrole/createdb/bypassrls) and a real write that must fail
//! with "permission denied". The write runs under an explicit READ WRITE
//! transaction: with the role's `default_transaction_read_only` on, Postgres
//! reports the read-only error before checking privileges, which would hide
//! a role that can write. Reads then run in a READ ONLY transaction, a single
//! statement wrapped as a subquery, as protection on top of the proof.
//!
//! The password lives in the OS keychain under the exact user, host and
//! port, never in the recipe or the workbook. TLS is always verified, against
//! the public web roots plus Supabase's own root; only a server on this
//! computer may be read without TLS. The rows read, as text with each
//! column's Postgres type, are the run's snapshot.

use serde::{Deserialize, Serialize};

use super::{Frame, OutColumn, Snapshot, ValueKind};
use crate::csv_import::{ColumnRule, DateOrder};

pub const DEFAULT_PORT: u16 = 5432;
pub const DEFAULT_DATABASE: &str = "postgres";
/// The largest integer a number column holds exactly (2^53).
const MAX_EXACT_INTEGER: u64 = 1 << 53;
/// Significant digits a number column holds exactly.
const MAX_EXACT_DIGITS: usize = 15;
/// The largest snapshot kept: rows as text.
#[cfg(feature = "native")]
const MAX_SNAPSHOT_BYTES: usize = 512 * 1024 * 1024;
/// Write privileges named in a refusal before "and N more".
#[cfg(feature = "native")]
const MAX_FINDINGS_SHOWN: usize = 8;

/// Supabase's private root (Supabase Root 2021 CA, valid to 2031-04-26;
/// SHA-256 80:70:25:AD:50:D4:ED:21:9D:2C:9C:7D:29:9C:00:4F:82:4E:B0:0C:F7:F6:5A:FE:F6:07:D0:7B:72:E6:CA:FA,
/// the same as Supabase's published prod-ca-2021.crt). Its pooler and
/// database certificates chain to it, not to a public root.
#[cfg(feature = "native")]
const SUPABASE_ROOT_2021: &[u8] = include_bytes!("recipe_postgres_supabase_root_2021.crt");

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct PostgresSource {
    /// The server's host name, like `aws-0-us-east-1.pooler.supabase.com`.
    pub host: String,
    #[serde(default = "default_port", skip_serializing_if = "is_default_port")]
    pub port: u16,
    #[serde(default = "default_database", skip_serializing_if = "is_default_database")]
    pub database: String,
    /// The read-only role (on Supabase's pooler, `role.project_ref`).
    pub user: String,
    /// `schema.table` (a view works too), or a name unique in the database.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub table: String,
    /// One SELECT, instead of a table.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub query: String,
    /// Off only for a server on this computer; TLS is always verified.
    #[serde(default = "yes", skip_serializing_if = "is_true")]
    pub tls: bool,
    /// The column names when the recipe was saved.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
}

fn default_port() -> u16 {
    DEFAULT_PORT
}
fn is_default_port(p: &u16) -> bool {
    *p == DEFAULT_PORT
}
fn default_database() -> String {
    DEFAULT_DATABASE.to_string()
}
fn is_default_database(d: &String) -> bool {
    d == DEFAULT_DATABASE
}
fn yes() -> bool {
    true
}
fn is_true(b: &bool) -> bool {
    *b
}

impl PostgresSource {
    pub fn new(host: String, user: String) -> Self {
        PostgresSource {
            host,
            port: DEFAULT_PORT,
            database: default_database(),
            user,
            table: String::new(),
            query: String::new(),
            tls: true,
            columns: Vec::new(),
        }
    }

    /// What the recipe reads, for approvals and reports:
    /// `postgres://reader@db.example.com:5432/postgres/public.orders`, or
    /// `…/postgres?query=<hash>` so a changed query asks again.
    pub fn identity(&self) -> String {
        let base = format!("postgres://{}@{}:{}/{}", self.user.trim(), self.host.trim(), self.port, self.database.trim());
        if self.query.trim().is_empty() {
            format!("{base}/{}", self.table.trim())
        } else {
            format!("{base}?query={}", &blake3::hash(self.query.trim().as_bytes()).to_hex()[..16])
        }
    }

    /// In words, for the approval card and the strip.
    pub fn describe(&self) -> String {
        let what = if self.query.trim().is_empty() {
            format!("Table {}", self.table.trim())
        } else {
            "A query".to_string()
        };
        format!(
            "{what} in database {} on {}:{}, as role {}",
            self.database.trim(),
            self.host.trim(),
            self.port,
            self.user.trim()
        )
    }

    /// The keychain account the password is stored under.
    pub fn keychain_account(&self) -> String {
        keychain_account(&self.user, &self.host, self.port)
    }
}

/// `postgres/user@host:port`: a role's password, for that server only.
pub fn keychain_account(user: &str, host: &str, port: u16) -> String {
    format!("postgres/{}@{}:{port}", user.trim(), host.trim().to_lowercase())
}

/// Whether a host is this computer, where TLS may be off.
pub fn is_local(host: &str) -> bool {
    matches!(host.trim().to_lowercase().as_str(), "localhost" | "127.0.0.1" | "::1" | "[::1]")
}

/// The source's settings as given, before connecting.
pub fn check_settings(src: &PostgresSource) -> Result<(), String> {
    check_server(src)?;
    match (src.table.trim().is_empty(), src.query.trim().is_empty()) {
        (true, true) => Err("the recipe names no table or query; set source.table (like public.orders) or source.query".into()),
        (false, false) => Err("the recipe names both a table and a query; keep one".into()),
        _ => Ok(()),
    }
}

/// The server, role and database, before connecting (no table needed).
pub fn check_server(src: &PostgresSource) -> Result<(), String> {
    let host = src.host.trim();
    if host.is_empty() {
        return Err("the recipe names no PostgreSQL host; set source.host".into());
    }
    if host.contains("://") || host.contains('@') || host.contains('/') || host.chars().any(char::is_whitespace) {
        return Err(format!(
            "{host:?} isn't a host name: give the host alone, with user, port and database in their own fields; the password goes in the keychain"
        ));
    }
    if src.user.trim().is_empty() {
        return Err("the recipe names no PostgreSQL user; set source.user to the read-only role".into());
    }
    if src.database.trim().is_empty() {
        return Err("the recipe names no database; set source.database".into());
    }
    if src.port == 0 {
        return Err("the port must be a number from 1 to 65535 (PostgreSQL's is 5432)".into());
    }
    if !src.tls && !is_local(host) {
        return Err(format!("TLS can be off only for a server on this computer; {host} is read over verified TLS"));
    }
    Ok(())
}

// ============================================================================
// The snapshot: column names and Postgres types, then rows as text
// ============================================================================

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
struct Read {
    columns: Vec<ReadColumn>,
    /// NULL is None.
    rows: Vec<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    warnings: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
struct ReadColumn {
    name: String,
    /// The Postgres type name: int8, numeric, timestamptz…
    #[serde(rename = "type")]
    pg_type: String,
}

/// How a Postgres type comes into a column.
#[derive(Debug, Clone, Copy, PartialEq)]
enum PgKind {
    /// int2, int4, oid: always exact.
    SmallInt,
    /// int8: exact up to 2^53, else the column stays text.
    BigInt,
    Float,
    Numeric,
    Bool,
    Date,
    Timestamp,
    /// Read in UTC.
    Timestamptz,
    Time,
    Text,
}

fn pg_kind(type_name: &str) -> PgKind {
    match type_name {
        "int2" | "int4" | "oid" => PgKind::SmallInt,
        "int8" => PgKind::BigInt,
        "float4" | "float8" => PgKind::Float,
        "numeric" => PgKind::Numeric,
        "bool" => PgKind::Bool,
        "date" => PgKind::Date,
        "timestamp" => PgKind::Timestamp,
        "timestamptz" => PgKind::Timestamptz,
        "time" => PgKind::Time,
        _ => PgKind::Text,
    }
}

/// Whether a numeric's text holds exactly in a number column: finite, at
/// most 15 significant digits.
fn exact_decimal(s: &str) -> bool {
    let digits: String = s.trim_start_matches('-').chars().filter(|c| *c != '.').collect();
    if digits.is_empty() || !digits.chars().all(|c| c.is_ascii_digit()) {
        return false; // NaN, Infinity
    }
    let significant = digits.trim_start_matches('0');
    // Trailing zeros after the point are still digits the value states
    significant.len() <= MAX_EXACT_DIGITS
}

fn exact_integer(s: &str) -> bool {
    s.trim_start_matches('-').parse::<u64>().is_ok_and(|u| u <= MAX_EXACT_INTEGER)
}

/// YYYY-MM-DD with a four-digit year (not BC, not infinity).
fn plain_date(s: &str) -> bool {
    let b = s.as_bytes();
    b.len() == 10 && b[4] == b'-' && b[7] == b'-' && b.iter().enumerate().all(|(i, c)| i == 4 || i == 7 || c.is_ascii_digit())
}

/// A date-time as `YYYY-MM-DD HH:MM:SS[.ffffff]` (what Postgres writes in
/// ISO style), not BC or infinity.
fn plain_datetime(s: &str) -> bool {
    s.len() >= 19 && plain_date(&s[..10]) && s.as_bytes()[10] == b' ' && !s.ends_with(" BC")
}

/// The snapshot's rows as a recipe frame, and anything worth reporting.
pub(super) fn read_frame(snapshot: &Snapshot) -> Result<(Frame, Vec<String>), String> {
    let read: Read = serde_json::from_slice(&snapshot.bytes).map_err(|e| format!("the PostgreSQL snapshot can't be read: {e}"))?;
    let width = read.columns.len();
    let mut columns = Vec::with_capacity(width);
    let mut rows: Vec<Vec<String>> = vec![Vec::with_capacity(width); read.rows.len()];
    for (i, col) in read.columns.iter().enumerate() {
        let kind = pg_kind(&col.pg_type);
        let values = read.rows.iter().map(|r| r.get(i).and_then(|v| v.as_deref()));
        let all = |ok: fn(&str) -> bool| values.clone().flatten().all(ok);
        let (rule, value_kind) = match kind {
            PgKind::SmallInt => (ColumnRule::Number, ValueKind::Number),
            PgKind::BigInt if all(exact_integer) => (ColumnRule::Number, ValueKind::Number),
            PgKind::Float if all(|s| s.parse::<f64>().is_ok_and(f64::is_finite)) => (ColumnRule::Number, ValueKind::Number),
            PgKind::Numeric if all(exact_decimal) => (ColumnRule::Number, ValueKind::Number),
            PgKind::Date if all(plain_date) => (ColumnRule::Date(DateOrder::Ymd), ValueKind::Plain),
            PgKind::Timestamp | PgKind::Timestamptz if all(plain_datetime) => (ColumnRule::Text, ValueKind::DateTime),
            PgKind::Time => (ColumnRule::Text, ValueKind::Time),
            _ => (ColumnRule::Text, ValueKind::Plain),
        };
        columns.push(OutColumn { name: col.name.clone(), rule, kind: value_kind });
        for (r, value) in values.enumerate() {
            rows[r].push(match (kind, value) {
                (_, None) => String::new(),
                (PgKind::Bool, Some("true")) => "TRUE".to_string(),
                (PgKind::Bool, Some("false")) => "FALSE".to_string(),
                (_, Some(v)) => v.to_string(),
            });
        }
    }
    let lines = (1..=rows.len()).collect();
    Ok((Frame { columns, rows, lines, files: Vec::new(), file_names: Vec::new(), decimal_comma: false }, read.warnings))
}

// ============================================================================
// Connecting, the read-only proof, and the read
// ============================================================================

#[cfg(feature = "native")]
pub use native::*;

#[cfg(feature = "native")]
mod native {
    use super::*;
    use ::postgres::{Client, Config, Transaction};
    use rustls::pki_types::pem::PemObject;

    /// The environment variables a headless run (CLI, CI) may hold the
    /// password in. `PGPASSWORD` is used only when `PGHOST` and `PGUSER` are
    /// exactly the source's host and user (and `PGPORT`, if set, its port).
    pub const PASSWORD_ENV: &str = "PGPASSWORD";

    /// What the proof found: the role reads and cannot write.
    #[derive(Debug, Clone, PartialEq)]
    pub struct Verified {
        /// The role the server says this is (`current_user`).
        pub role: String,
        pub server_version: String,
        /// The table a write was tried on, or None when the role can read no
        /// table to try (the catalog check alone then).
        pub write_tried_on: Option<String>,
        /// A standby can't take a write at all: the catalog check alone.
        pub standby: bool,
    }

    impl Verified {
        /// For reports: `read-only: verified (role visigrid_reader)`.
        pub fn summary(&self) -> String {
            match (&self.write_tried_on, self.standby) {
                (_, true) => format!("read-only: verified by privileges (role {}, standby server)", self.role),
                (Some(t), _) => format!("read-only: verified (role {}; a write to {t} was refused)", self.role),
                (None, _) => format!("read-only: verified by privileges (role {}; it reads no table a write could be tried on)", self.role),
            }
        }
    }

    /// A table or view the role can read, for choosing one.
    #[derive(Debug, Clone, PartialEq)]
    pub struct TableInfo {
        /// `schema.table`
        pub name: String,
        /// "table", "view", "materialized view", "foreign table"
        pub kind: &'static str,
    }

    /// The password for a source: its keychain entry, else `PGPASSWORD`
    /// when `PGHOST`/`PGUSER` (and `PGPORT`, if set) are exactly this server
    /// and role.
    pub fn password(src: &PostgresSource) -> Result<String, String> {
        if let Some(pw) = visigrid_config::secrets::get(&src.keychain_account()) {
            return Ok(pw);
        }
        let env = |k: &str| std::env::var(k).ok().map(|v| v.trim().to_string());
        if let (Some(pw), Some(host), Some(user)) = (std::env::var(PASSWORD_ENV).ok(), env("PGHOST"), env("PGUSER")) {
            let port_ok = env("PGPORT").is_none_or(|p| p == src.port.to_string());
            if !pw.is_empty() && host.eq_ignore_ascii_case(src.host.trim()) && user == src.user.trim() && port_ok {
                return Ok(pw);
            }
        }
        Err(format!(
            "no password saved for {}@{}:{}; save one with `vgrid postgres password` (or in VisiGrid), or set PGHOST, PGUSER and {PASSWORD_ENV} for headless runs",
            src.user.trim(),
            src.host.trim(),
            src.port
        ))
    }

    fn tls_config() -> Result<rustls::ClientConfig, String> {
        let mut roots = rustls::RootCertStore::empty();
        roots.extend(webpki_roots::TLS_SERVER_ROOTS.iter().cloned());
        let supabase = rustls::pki_types::CertificateDer::from_pem_slice(SUPABASE_ROOT_2021)
            .map_err(|e| format!("the bundled Supabase root certificate can't be read: {e}"))?;
        roots.add(supabase).map_err(|e| format!("the bundled Supabase root certificate is invalid: {e}"))?;
        rustls::ClientConfig::builder_with_provider(rustls::crypto::ring::default_provider().into())
            .with_safe_default_protocol_versions()
            .map_err(|e| e.to_string())
            .map(|b| b.with_root_certificates(roots).with_no_client_auth())
    }

    /// Open a connection as the source's role, with its saved password.
    pub fn connect(src: &PostgresSource) -> Result<Client, String> {
        check_server(src)?;
        connect_with(src, password(src)?)
    }

    fn connect_with(src: &PostgresSource, pw: String) -> Result<Client, String> {
        check_server(src)?;
        let host = src.host.trim();
        let mut cfg = Config::new();
        cfg.host(host)
            .port(src.port)
            .dbname(src.database.trim())
            .user(src.user.trim())
            .password(pw)
            .application_name("VisiGrid")
            .connect_timeout(std::time::Duration::from_secs(10));
        let result = if src.tls {
            cfg.ssl_mode(::postgres::config::SslMode::Require);
            cfg.connect(tokio_postgres_rustls::MakeRustlsConnect::new(tls_config()?))
        } else {
            cfg.ssl_mode(::postgres::config::SslMode::Disable);
            cfg.connect(::postgres::NoTls)
        };
        result.map_err(|e| connect_error(src, &e))
    }

    fn connect_error(src: &PostgresSource, e: &::postgres::Error) -> String {
        let (user, host) = (src.user.trim(), src.host.trim());
        if let Some(db) = e.as_db_error() {
            return match db.code().code() {
                "28P01" => format!("{host} didn't accept the password for {user}; save a new one"),
                "28000" => format!("{host} refused {user}: {}", db.message()),
                "3D000" => format!("{host} has no database {:?}", src.database.trim()),
                _ => format!("{host}: {}", db.message()),
            };
        }
        let text = e.to_string();
        let detail = std::error::Error::source(e).map(|s| s.to_string()).unwrap_or_default();
        if detail.contains("InvalidCertificate") || detail.contains("certificate") {
            format!("couldn't verify {host}'s TLS certificate ({detail}); check the host name")
        } else if text.contains("server does not support TLS") || detail.contains("server does not support TLS") {
            format!("{host} doesn't offer TLS; VisiGrid reads a remote server only over verified TLS")
        } else {
            format!("couldn't connect to {host}:{}: {}", src.port, if detail.is_empty() { text } else { detail })
        }
    }

    /// The catalog half of the proof (no rows = nothing found). Counts
    /// membership, so a grant through another role is found too. Role-wide
    /// findings come first: a superuser is named as one, not by its first
    /// eight table privileges.
    const WRITE_PRIVILEGES: &str = "
SELECT 'role: ' || x FROM pg_roles r, LATERAL (VALUES
   (CASE WHEN r.rolsuper THEN 'superuser' END), (CASE WHEN r.rolcreaterole THEN 'createrole' END),
   (CASE WHEN r.rolcreatedb THEN 'createdb' END), (CASE WHEN r.rolbypassrls THEN 'bypassrls' END)) v(x)
 WHERE r.rolname = current_user AND x IS NOT NULL
UNION ALL
SELECT 'member of ' || quote_ident(g.rolname) FROM pg_roles g
 WHERE g.rolname <> current_user AND pg_has_role(current_user, g.oid, 'MEMBER')
UNION ALL
SELECT 'database: CREATE' WHERE has_database_privilege(current_database(), 'CREATE')
UNION ALL
SELECT 'schema ' || quote_ident(n.nspname) || ': CREATE' FROM pg_namespace n WHERE has_schema_privilege(n.oid, 'CREATE')
UNION ALL
SELECT 'table ' || c.oid::regclass || ': ' || p.priv
  FROM pg_class c
  CROSS JOIN (VALUES ('INSERT'), ('UPDATE'), ('DELETE'), ('TRUNCATE')) AS p(priv)
 WHERE c.relkind IN ('r', 'p', 'v', 'm', 'f')
   AND c.relnamespace NOT IN ('pg_catalog'::regnamespace, 'information_schema'::regnamespace)
   AND has_table_privilege(c.oid, p.priv)";

    /// A table the role can read, with a column a write can name: the
    /// source's own table when it is one, else the first such table.
    /// Generated and always-identity columns are skipped: naming them in an
    /// UPDATE fails before privileges are checked.
    const WRITE_TARGET: &str = "
SELECT c.oid::regclass::text,
       (SELECT quote_ident(a.attname) FROM pg_attribute a
         WHERE a.attrelid = c.oid AND a.attnum > 0 AND NOT a.attisdropped
           AND a.attgenerated = '' AND a.attidentity <> 'a'
         ORDER BY a.attnum LIMIT 1)
  FROM pg_class c
 WHERE c.relkind IN ('r', 'p')
   AND c.relnamespace NOT IN ('pg_catalog'::regnamespace, 'information_schema'::regnamespace)
   AND c.relnamespace::regnamespace::text NOT LIKE 'pg\\_%'
   AND has_table_privilege(c.oid, 'SELECT')
 ORDER BY c.oid = coalesce($1::oid, 0::oid) DESC, c.oid
 LIMIT 20";

    /// Prove the role can't write; the table to try a write on first is
    /// `prefer` (a table's oid).
    fn prove(client: &mut Client, prefer: Option<u32>) -> Result<Verified, String> {
        let row = client
            .query_one("SELECT current_user::text, current_setting('server_version'), pg_is_in_recovery()", &[])
            .map_err(|e| format!("checking the role: {e}"))?;
        let (role, server_version, standby): (String, String, bool) = (row.get(0), row.get(1), row.get(2));
        let findings: Vec<String> = client
            .query(WRITE_PRIVILEGES, &[])
            .map_err(|e| format!("checking {role}'s privileges: {e}"))?
            .iter()
            .map(|r| r.get(0))
            .collect();
        if !findings.is_empty() {
            let shown = findings.iter().take(MAX_FINDINGS_SHOWN).cloned().collect::<Vec<_>>().join("; ");
            let more = findings.len().saturating_sub(MAX_FINDINGS_SHOWN);
            let more = if more > 0 { format!("; and {more} more") } else { String::new() };
            return Err(format!(
                "role {role} isn't read-only, so VisiGrid won't read with it: {shown}{more}. Use a role that holds SELECT and nothing else"
            ));
        }
        if standby {
            return Ok(Verified { role, server_version, write_tried_on: None, standby: true });
        }
        let mut targets = client
            .query(WRITE_TARGET, &[&prefer])
            .map_err(|e| format!("choosing a table to try a write on: {e}"))?
            .into_iter()
            .filter_map(|r| Some((r.get::<_, String>(0), r.get::<_, Option<String>>(1)?)));
        let Some((table, column)) = targets.next() else {
            return Ok(Verified { role, server_version, write_tried_on: None, standby: false });
        };
        // Changes nothing even if it ran: no row matches, and it's rolled back
        let mut tx = client.build_transaction().read_only(false).start().map_err(|e| format!("starting the write check: {e}"))?;
        let tried = tx.execute(&format!("UPDATE {table} SET {column} = {column} WHERE false"), &[]);
        let _ = tx.rollback();
        match tried {
            Err(e) if e.code().is_some_and(|c| c.code() == "42501") => {
                Ok(Verified { role, server_version, write_tried_on: Some(table), standby: false })
            }
            Ok(_) => Err(format!("role {role} isn't read-only: a write to {table} was accepted. Use a role that holds SELECT and nothing else")),
            Err(e) => Err(format!(
                "couldn't prove role {role} is read-only: a write to {table} should fail with permission denied, but failed with: {}",
                e.as_db_error().map(|d| d.message().to_string()).unwrap_or_else(|| e.to_string())
            )),
        }
    }

    /// Connect and run the proof, without reading: for `vgrid postgres
    /// check` and the builder.
    pub fn verify(src: &PostgresSource) -> Result<Verified, String> {
        let mut client = connect(src)?;
        prove(&mut client, None)
    }

    /// Connect with `password` (not yet saved; None: the saved one), run
    /// the proof, and list the tables the role reads. The builder saves a
    /// password only once this succeeds, so a wrong one is never kept.
    pub fn verify_and_list(src: &PostgresSource, password: Option<&str>) -> Result<(Verified, Vec<TableInfo>), String> {
        let mut client = match password {
            Some(pw) => connect_with(src, pw.to_string())?,
            None => connect(src)?,
        };
        let verified = prove(&mut client, None)?;
        Ok((verified, list_tables(&mut client)?))
    }

    /// A table or view's oid, qualified name, and whether row-level
    /// security is on, from `schema.table` or a name unique in the database.
    fn resolve_table(client: &mut Client, name: &str) -> Result<(u32, String, bool), String> {
        let name = name.trim();
        let (schema, table) = match name.split_once('.') {
            Some((s, t)) => (Some(s.trim().trim_matches('"')), t.trim().trim_matches('"')),
            None => (None, name.trim_matches('"')),
        };
        let rows = client
            .query(
                "SELECT c.oid, c.oid::regclass::text, c.relrowsecurity
                   FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
                  WHERE c.relkind IN ('r', 'p', 'v', 'm', 'f') AND c.relname = $1
                    AND ($2::text IS NULL OR n.nspname = $2)
                    AND n.nspname NOT IN ('pg_catalog', 'information_schema') AND n.nspname NOT LIKE 'pg\\_%'",
                &[&table, &schema],
            )
            .map_err(|e| format!("finding table {name}: {e}"))?;
        match rows.as_slice() {
            [one] => Ok((one.get(0), one.get(1), one.get(2))),
            [] => Err(format!("the database has no table or view {name}")),
            many => Err(format!(
                "{name} names {} tables ({}); write it as schema.table",
                many.len(),
                many.iter().map(|r| r.get::<_, String>(1)).collect::<Vec<_>>().join(", ")
            )),
        }
    }

    /// The tables and views the role can read, for choosing one.
    pub fn tables(src: &PostgresSource) -> Result<Vec<TableInfo>, String> {
        list_tables(&mut connect(src)?)
    }

    fn list_tables(client: &mut Client) -> Result<Vec<TableInfo>, String> {
        let rows = client
            .query(
                "SELECT c.oid::regclass::text, c.relkind::text
                   FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
                  WHERE c.relkind IN ('r', 'p', 'v', 'm', 'f')
                    AND n.nspname NOT IN ('pg_catalog', 'information_schema') AND n.nspname NOT LIKE 'pg\\_%'
                    AND has_table_privilege(c.oid, 'SELECT')
                  ORDER BY n.nspname <> 'public', 1",
                &[],
            )
            .map_err(|e| format!("listing tables: {e}"))?;
        Ok(rows
            .iter()
            .map(|r| TableInfo {
                name: r.get(0),
                kind: match r.get::<_, String>(1).as_str() {
                    "v" => "view",
                    "m" => "materialized view",
                    "f" => "foreign table",
                    _ => "table",
                },
            })
            .collect())
    }

    /// Whether a policy lets the role (or PUBLIC) select from a table with
    /// row-level security on.
    fn rls_hides_all(tx: &mut Transaction, oid: u32) -> bool {
        tx.query_one(
            "SELECT NOT EXISTS (
               SELECT 1 FROM pg_policy p
                WHERE p.polrelid = $1 AND p.polcmd IN ('r', '*')
                  AND (0::oid = ANY (p.polroles)
                       OR EXISTS (SELECT 1 FROM unnest(p.polroles) r(oid) WHERE pg_has_role(current_user, r.oid, 'USAGE'))))",
            &[&oid],
        )
        .map(|r| r.get(0))
        .unwrap_or(false)
    }

    /// Read the source: connect, prove read-only, then one read. The rows
    /// and their column types are the snapshot.
    pub fn fetch(src: &PostgresSource, max_rows: usize, max_cols: usize) -> Result<(Snapshot, Verified), String> {
        check_settings(src)?;
        let mut client = connect(src)?;
        let table = if src.query.trim().is_empty() { Some(resolve_table(&mut client, &src.table)?) } else { None };
        let verified = prove(&mut client, table.as_ref().map(|t| t.0))?;

        let mut tx = client.build_transaction().read_only(true).start().map_err(|e| format!("starting the read: {e}"))?;
        // ISO dates and the shortest exact floats, whatever the role's settings
        tx.batch_execute("SET LOCAL datestyle = 'ISO, YMD'; SET LOCAL intervalstyle = 'postgres'; SET LOCAL extra_float_digits = 1")
            .map_err(|e| e.to_string())?;
        let timeout: String = tx.query_one("SELECT current_setting('statement_timeout')", &[]).map(|r| r.get(0)).unwrap_or_default();
        if timeout == "0" {
            tx.batch_execute("SET LOCAL statement_timeout = '5min'").map_err(|e| e.to_string())?;
        }
        let base = match &table {
            Some((_, qualified, _)) => format!("SELECT * FROM {qualified}"),
            None => src.query.trim().trim_end_matches(';').trim_end().to_string(),
        };
        // Preparing (the extended protocol) takes exactly one statement
        let statement = tx.prepare(&base).map_err(|e| query_error("the query", &e))?;
        let cols = statement.columns();
        if cols.is_empty() {
            return Err("the query returns no columns; it must be a SELECT".into());
        }
        if cols.len() > max_cols {
            return Err(format!("the source has {} columns; a sheet holds {max_cols}", cols.len()));
        }
        // Every value as text, by position (so repeated names work); a data-
        // changing statement can't sit in FROM. Newlines end a -- comment
        let list = cols
            .iter()
            .enumerate()
            .map(|(i, c)| match c.type_().name() {
                "timestamptz" => format!("(q.c{i} AT TIME ZONE 'UTC')::text"),
                _ => format!("q.c{i}::text"),
            })
            .collect::<Vec<_>>()
            .join(", ");
        let aliases = (0..cols.len()).map(|i| format!("c{i}")).collect::<Vec<_>>().join(", ");
        let sql = format!("SELECT {list} FROM (\n{base}\n) AS q({aliases}) LIMIT {}", max_rows + 1);
        let columns: Vec<ReadColumn> = cols.iter().map(|c| ReadColumn { name: c.name().to_string(), pg_type: c.type_().name().to_string() }).collect();
        let found = tx.query(&sql, &[]).map_err(|e| query_error("reading", &e))?;
        if found.len() > max_rows {
            return Err(format!("the source has more than {max_rows} rows, the most a sheet holds; filter it with a query"));
        }
        let rows: Vec<Vec<Option<String>>> = found.iter().map(|r| (0..columns.len()).map(|i| r.get(i)).collect()).collect();
        let mut warnings = Vec::new();
        if let Some((oid, qualified, true)) = &table {
            if rows.is_empty() && rls_hides_all(&mut tx, *oid) {
                warnings.push(format!(
                    "0 rows: row-level security on {qualified} hides every row from {role}. To let it read them, run as the table's owner: CREATE POLICY visigrid_read ON {qualified} FOR SELECT TO {role} USING (true);",
                    role = verified.role
                ));
            }
        }
        let _ = tx.rollback();
        let bytes = serde_json::to_vec(&Read { columns, rows, warnings }).map_err(|e| e.to_string())?;
        if bytes.len() > MAX_SNAPSHOT_BYTES {
            return Err(format!("the source is more than {} MB as text; filter it with a query", MAX_SNAPSHOT_BYTES / (1024 * 1024)));
        }
        Ok((Snapshot::from_bytes(std::path::Path::new(&src.identity()), bytes), verified))
    }

    fn query_error(doing: &str, e: &::postgres::Error) -> String {
        match e.as_db_error() {
            Some(db) if db.code().code() == "42601" && db.message().contains("multiple commands") => {
                "the query must be a single SELECT".to_string()
            }
            Some(db) if db.code().code() == "57014" => "the read took longer than the server's statement timeout".to_string(),
            Some(db) if db.code().code() == "25006" => format!("{doing}: the query tries to change data; a source only reads"),
            Some(db) => format!("{doing}: {}", db.message()),
            None => format!("{doing}: {e}"),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn src() -> PostgresSource {
        let mut s = PostgresSource::new("db.example.com".into(), "reader".into());
        s.table = "public.orders".into();
        s
    }

    #[test]
    fn settings_refuse_a_url_a_missing_table_and_tls_off_remotely() {
        assert!(check_settings(&src()).is_ok());
        let mut s = src();
        s.host = "postgresql://reader:pw@db.example.com/postgres".into();
        assert!(check_settings(&s).unwrap_err().contains("keychain"));
        let mut s = src();
        s.table.clear();
        assert!(check_settings(&s).unwrap_err().contains("no table or query"));
        s.query = "select 1".into();
        s.table = "t".into();
        assert!(check_settings(&s).unwrap_err().contains("both"));
        let mut s = src();
        s.tls = false;
        assert!(check_settings(&s).unwrap_err().contains("this computer"));
        s.host = "localhost".into();
        assert!(check_settings(&s).is_ok());
    }

    #[test]
    fn identity_changes_with_the_query_and_names_no_password() {
        let a = src();
        assert_eq!(a.identity(), "postgres://reader@db.example.com:5432/postgres/public.orders");
        let mut q = src();
        q.table.clear();
        q.query = "select 1".into();
        let mut q2 = q.clone();
        q2.query = "select 2".into();
        assert_ne!(q.identity(), q2.identity(), "a changed query is approved again");
        assert_eq!(a.keychain_account(), "postgres/reader@db.example.com:5432");
    }

    #[test]
    fn a_password_in_the_recipe_is_refused() {
        let toml = "host = \"h\"\nuser = \"u\"\ntable = \"t\"\npassword = \"x\"\n";
        assert!(toml::from_str::<PostgresSource>(toml).is_err());
    }

    fn frame(read: serde_json::Value) -> (Frame, Vec<String>) {
        read_frame(&Snapshot::from_bytes(std::path::Path::new("x"), serde_json::to_vec(&read).unwrap())).unwrap()
    }

    #[test]
    fn types_become_columns_and_exactness_decides_numbers() {
        let (f, _) = frame(serde_json::json!({
            "columns": [
                {"name": "id", "type": "int8"}, {"name": "ref", "type": "int8"}, {"name": "total", "type": "numeric"},
                {"name": "wide", "type": "numeric"}, {"name": "gift", "type": "bool"}, {"name": "on", "type": "date"},
                {"name": "at", "type": "timestamptz"}, {"name": "tags", "type": "_text"}, {"name": "f", "type": "float8"}
            ],
            "rows": [
                ["1", "9007199254740993", "129.50", "1234567890.1234567", "true", "2026-09-03", "2026-09-01 10:00:00", "{a,b}", "NaN"],
                ["2", "12", "0.10", "1", "false", null, "2026-09-02 11:30:00.25", "{}", "1.5"]
            ]
        }));
        let rules: Vec<_> = f.columns.iter().map(|c| (c.rule, c.kind)).collect();
        assert_eq!(rules[0], (ColumnRule::Number, ValueKind::Number));
        assert_eq!(rules[1].0, ColumnRule::Text, "an int8 past 2^53 keeps its digits as text");
        assert_eq!(rules[2], (ColumnRule::Number, ValueKind::Number));
        assert_eq!(rules[3].0, ColumnRule::Text, "17 significant digits stay text");
        assert_eq!(rules[4].0, ColumnRule::Text);
        assert_eq!(rules[5].0, ColumnRule::Date(DateOrder::Ymd));
        assert_eq!(rules[6], (ColumnRule::Text, ValueKind::DateTime));
        assert_eq!(rules[7].0, ColumnRule::Text);
        assert_eq!(rules[8].0, ColumnRule::Text, "NaN keeps a float column text");
        assert_eq!(f.rows[0][1], "9007199254740993");
        assert_eq!(f.rows[0][4], "TRUE");
        assert_eq!(f.rows[1][4], "FALSE");
        assert_eq!(f.rows[1][5], "", "NULL is empty");
        assert_eq!(f.rows[0][2], "129.50", "numeric text kept exactly");
    }

    #[test]
    fn infinity_and_bc_dates_keep_a_date_column_text() {
        let (f, _) = frame(serde_json::json!({
            "columns": [{"name": "d", "type": "date"}, {"name": "t", "type": "timestamp"}],
            "rows": [["2026-01-01", "2026-01-01 00:00:00"], ["infinity", "0044-03-15 00:00:00 BC"]]
        }));
        assert_eq!(f.columns[0].rule, ColumnRule::Text);
        assert_eq!(f.columns[1].kind, ValueKind::Plain);
    }

    #[test]
    fn warnings_in_the_snapshot_reach_the_report() {
        let (_, w) = frame(serde_json::json!({"columns": [{"name": "a", "type": "text"}], "rows": [], "warnings": ["0 rows: row-level security"]}));
        assert_eq!(w, ["0 rows: row-level security"]);
    }
}
