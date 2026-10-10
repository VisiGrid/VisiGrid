//! `vgrid postgres`: the password a PostgreSQL recipe source connects with,
//! the read-only proof on its own, and the tables a role can read.

use visigrid_io::recipe::postgres::{self as pg, PostgresSource};

use crate::visibooks_cmd::read_secret;
use crate::CliError;

/// The server and role every subcommand names.
#[derive(clap::Args)]
pub(crate) struct Server {
    /// Server host name (on Supabase, the pooler host from Connect)
    #[arg(long)]
    host: String,

    /// The read-only role (on Supabase's pooler, role.project_ref)
    #[arg(long)]
    user: String,

    #[arg(long, default_value_t = pg::DEFAULT_PORT)]
    port: u16,

    #[arg(long, default_value = pg::DEFAULT_DATABASE)]
    database: String,

    /// Connect without TLS (a server on this computer only)
    #[arg(long)]
    no_tls: bool,
}

impl Server {
    fn source(&self) -> PostgresSource {
        let mut src = PostgresSource::new(self.host.clone(), self.user.clone());
        src.port = self.port;
        src.database = self.database.clone();
        src.tls = !self.no_tls;
        src
    }
}

/// Save (or with `delete`, remove) the role's password in the keychain,
/// read from stdin, unechoed at a terminal.
pub(crate) fn cmd_password(server: Server, delete: bool) -> Result<(), CliError> {
    let src = server.source();
    let account = src.keychain_account();
    let who = format!("{}@{}:{}", src.user.trim(), src.host.trim(), src.port);
    if delete {
        visigrid_config::secrets::delete(&account).map_err(CliError::io)?;
        eprintln!("Removed the password for {who}");
        return Ok(());
    }
    let pw = read_secret(&format!("Password for {who}: "))?;
    if pw.is_empty() {
        return Err(CliError::args("no password given; nothing saved"));
    }
    visigrid_config::secrets::set(&account, &pw).map_err(CliError::io)?;
    eprintln!("Saved in the system keychain. Recipes reading {who} will use it.");
    Ok(())
}

/// Connect and run the read-only proof, without reading any rows.
pub(crate) fn cmd_check(server: Server) -> Result<(), CliError> {
    let src = server.source();
    let verified = pg::verify(&src).map_err(CliError::io)?;
    println!("{} on PostgreSQL {}", verified.summary(), verified.server_version);
    Ok(())
}

/// List the tables and views the role can read.
pub(crate) fn cmd_tables(server: Server) -> Result<(), CliError> {
    let src = server.source();
    let tables = pg::tables(&src).map_err(CliError::io)?;
    if tables.is_empty() {
        eprintln!("The role can read no tables; GRANT SELECT on the ones a recipe should read.");
    }
    for t in tables {
        println!("{}\t{}", t.name, t.kind);
    }
    Ok(())
}
