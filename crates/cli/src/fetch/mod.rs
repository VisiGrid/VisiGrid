//! `vgrid fetch` — pull data from external sources into canonical CSV.

mod brex;
mod common;
mod gusto;
mod mercury;
mod qbo;
mod sftp;
mod stripe;

use std::path::PathBuf;

use clap::Subcommand;

use crate::CliError;

#[derive(Subcommand)]
pub enum FetchCommands {
    /// Fetch balance transactions from Stripe
    #[command(after_help = "\
Examples:
  vgrid fetch stripe --from 2026-01-01 --to 2026-01-31
  vgrid fetch stripe --from 2026-01-01 --to 2026-01-31 --out stripe.csv
  vgrid fetch stripe --from 2026-01-01 --to 2026-01-31 --api-key sk_live_...
  STRIPE_API_KEY=sk_live_... vgrid fetch stripe --from 2026-01-01 --to 2026-01-31")]
    Stripe {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Stripe secret key (default: STRIPE_API_KEY env)
        #[arg(long)]
        api_key: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Stripe Connect account ID (acct_...)
        #[arg(long)]
        account: Option<String>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,
    },

    /// Fetch payroll data from Gusto
    #[command(after_help = "\
Examples:
  vgrid fetch gusto --credentials ~/.config/vgrid/gusto.json --from 2026-01-01 --to 2026-01-31
  vgrid fetch gusto --credentials ~/gusto.json --from 2026-01-01 --to 2026-01-31 --out gusto.csv
  vgrid fetch gusto --access-token gp_... --company-uuid abc-123 --from 2026-01-01 --to 2026-01-31")]
    Gusto {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Path to Gusto OAuth2 credentials JSON file
        #[arg(long)]
        credentials: Option<PathBuf>,

        /// Gusto access token (skips refresh, requires --company-uuid)
        #[arg(long)]
        access_token: Option<String>,

        /// Gusto company UUID (required with --access-token)
        #[arg(long)]
        company_uuid: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,
    },

    /// Fetch posted ledger transactions from QuickBooks Online
    #[command(after_help = "\
Examples:
  # Fetch posted transactions for a bank account
  vgrid fetch qbo --credentials ~/.config/vgrid/qbo.json --from 2026-01-01 --to 2026-01-31 --account \"Checking\"
  vgrid fetch qbo --credentials ~/qbo.json --from 2026-01-01 --to 2026-01-31 --account \"Checking\" --out qbo.csv

  # Only deposits (Stripe payouts)
  vgrid fetch qbo --credentials ~/qbo.json --from 2026-01-01 --to 2026-01-31 --account \"Checking\" --include deposit

  # Direct token (skips refresh)
  vgrid fetch qbo --access-token eyJ... --realm-id 123456789 --from 2026-01-01 --to 2026-01-31 --account-id 35")]
    Qbo {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Path to QBO OAuth2 credentials JSON file
        #[arg(long)]
        credentials: Option<PathBuf>,

        /// QBO access token (skips refresh, requires --realm-id)
        #[arg(long)]
        access_token: Option<String>,

        /// QBO company realm ID (required with --access-token)
        #[arg(long)]
        realm_id: Option<String>,

        /// Bank account name in QBO (resolved to ID via Account query)
        #[arg(long)]
        account: Option<String>,

        /// Bank account ID directly (skip name resolution)
        #[arg(long)]
        account_id: Option<String>,

        /// Entity types to query, comma-separated (default: deposit,purchase,transfer)
        #[arg(long)]
        include: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,

        /// Use QBO sandbox API base URL
        #[arg(long)]
        sandbox: bool,
    },

    /// Fetch card transactions from Brex
    #[command(name = "brex-card", after_help = "\
Examples:
  vgrid fetch brex-card --from 2026-01-01 --to 2026-01-31
  vgrid fetch brex-card --from 2026-01-01 --to 2026-01-31 --out brex-card.csv
  vgrid fetch brex-card --from 2026-01-01 --to 2026-01-31 --api-key brex_token_...
  BREX_API_KEY=brex_token_... vgrid fetch brex-card --from 2026-01-01 --to 2026-01-31")]
    BrexCard {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Brex API token (default: BREX_API_KEY env)
        #[arg(long)]
        api_key: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,
    },

    /// Fetch bank (cash account) transactions from Brex
    #[command(name = "brex-bank", after_help = "\
Examples:
  vgrid fetch brex-bank --from 2026-01-01 --to 2026-01-31
  vgrid fetch brex-bank --from 2026-01-01 --to 2026-01-31 --out brex-bank.csv
  vgrid fetch brex-bank --from 2026-01-01 --to 2026-01-31 --api-key brex_token_...
  vgrid fetch brex-bank --from 2026-01-01 --to 2026-01-31 --account cash_abc123
  BREX_API_KEY=brex_token_... vgrid fetch brex-bank --from 2026-01-01 --to 2026-01-31")]
    BrexBank {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Brex API token (default: BREX_API_KEY env)
        #[arg(long)]
        api_key: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Brex cash account ID (auto-detected if only one)
        #[arg(long)]
        account: Option<String>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,
    },

    /// Fetch bank transactions from Mercury
    #[command(after_help = "\
Examples:
  vgrid fetch mercury --from 2026-01-01 --to 2026-01-31
  vgrid fetch mercury --from 2026-01-01 --to 2026-01-31 --out mercury.csv
  vgrid fetch mercury --from 2026-01-01 --to 2026-01-31 --api-key secret-token:mercury_...
  MERCURY_API_KEY=secret-token:mercury_... vgrid fetch mercury --from 2026-01-01 --to 2026-01-31")]
    Mercury {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Mercury API token (default: MERCURY_API_KEY env)
        #[arg(long)]
        api_key: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Mercury account ID (acc_...)
        #[arg(long)]
        account: Option<String>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,
    },

    /// Download files from an SFTP server
    #[command(after_help = "\
Examples:
  # List files (dry run, newest-first)
  vgrid fetch sftp --host sftp.partner.com --user client123 --key ~/.ssh/id_ed25519 \\
    --remote-dir /reports/ --list-only

  # Download with TOFU on first connect
  vgrid fetch sftp --host sftp.partner.com --user client123 --key ~/.ssh/id_ed25519 \\
    --remote-dir /reports/ --glob 'settlement_*.csv' --out ./downloads/ --trust-on-first-use

  # Pipe single file to stdout, provenance on stderr
  vgrid fetch sftp --host sftp.partner.com --user client123 --key ~/.ssh/id_ed25519 \\
    --remote-dir /reports/ --glob 'settlement_20260215.csv' --stdout

  # Download with state tracking (skip already-seen files)
  vgrid fetch sftp --host sftp.partner.com --user client123 --key ~/.ssh/id_ed25519 \\
    --remote-dir /reports/ --out ./downloads/ --state-dir ~/.visigrid/sftp/partner/

  # Provenance as sidecar JSON (default in --out mode), or suppress it
  vgrid fetch sftp ... --out ./downloads/ --provenance sidecar
  vgrid fetch sftp ... --stdout --provenance none")]
    Sftp {
        /// SFTP hostname
        #[arg(long)]
        host: String,

        /// SFTP port
        #[arg(long, default_value = "22")]
        port: u16,

        /// SSH username
        #[arg(long, visible_alias = "user")]
        username: String,

        /// Path to SSH private key file
        #[arg(long, visible_alias = "key")]
        private_key: Option<PathBuf>,

        /// Passphrase for private key (or SFTP_KEY_PASSPHRASE env)
        #[arg(long, env = "SFTP_KEY_PASSPHRASE")]
        passphrase: Option<String>,

        /// Password for password auth (or SFTP_PASSWORD env)
        #[arg(long, env = "SFTP_PASSWORD")]
        password: Option<String>,

        /// Remote directory to scan
        #[arg(long)]
        remote_dir: String,

        /// Filename filter pattern (basename only)
        #[arg(long, default_value = "*.csv")]
        glob: String,

        /// Minimum age in seconds since mtime before ingesting
        #[arg(long, default_value = "120")]
        min_age: u64,

        /// Output directory for downloaded files
        #[arg(long, required_unless_present = "stdout", conflicts_with = "stdout")]
        out: Option<PathBuf>,

        /// Write single file to stdout instead of --out
        #[arg(long)]
        stdout: bool,

        /// Path to known_hosts file
        #[arg(long, default_value = "~/.ssh/known_hosts")]
        known_hosts: String,

        /// Where to write TOFU'd keys (default: same as --known-hosts)
        #[arg(long)]
        known_hosts_out: Option<String>,

        /// Accept unknown host key and write it to --known-hosts-out
        #[arg(long)]
        trust_on_first_use: bool,

        /// Overwrite even if content hash matches existing file
        #[arg(long)]
        overwrite: bool,

        /// Limit number of files to download
        #[arg(long)]
        max_files: Option<usize>,

        /// Only files with mtime >= this ISO date (YYYY-MM-DD)
        #[arg(long)]
        since: Option<String>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,

        /// Dry run: print matching files (newest-first) and exit
        #[arg(long)]
        list_only: bool,

        /// Where to write provenance records [default: sidecar with --out, stderr with --stdout]
        #[arg(long, value_enum)]
        provenance: Option<sftp::ProvenanceMode>,

        /// Directory for seen.jsonl state file (skip already-processed files)
        #[arg(long)]
        state_dir: Option<PathBuf>,

        /// Re-download files even if already recorded in --state-dir
        #[arg(long)]
        reprocess: bool,
    },
}

pub fn cmd_fetch(command: FetchCommands) -> Result<(), CliError> {
    match command {
        FetchCommands::Gusto {
            from,
            to,
            credentials,
            access_token,
            company_uuid,
            out,
            quiet,
        } => gusto::cmd_fetch_gusto(
            from,
            to,
            credentials,
            access_token,
            company_uuid,
            out,
            quiet,
        ),
        FetchCommands::Qbo {
            from,
            to,
            credentials,
            access_token,
            realm_id,
            account,
            account_id,
            include,
            out,
            quiet,
            sandbox,
        } => qbo::cmd_fetch_qbo(
            from,
            to,
            credentials,
            access_token,
            realm_id,
            account,
            account_id,
            include,
            out,
            quiet,
            sandbox,
        ),
        FetchCommands::BrexCard {
            from,
            to,
            api_key,
            out,
            quiet,
        } => brex::card::cmd_fetch_brex_card(from, to, api_key, out, quiet),
        FetchCommands::BrexBank {
            from,
            to,
            api_key,
            out,
            account,
            quiet,
        } => brex::bank::cmd_fetch_brex_bank(from, to, api_key, out, account, quiet),
        FetchCommands::Stripe {
            from,
            to,
            api_key,
            out,
            account,
            quiet,
        } => stripe::cmd_fetch_stripe(from, to, api_key, out, account, quiet),
        FetchCommands::Mercury {
            from,
            to,
            api_key,
            out,
            account,
            quiet,
        } => mercury::cmd_fetch_mercury(from, to, api_key, out, account, quiet),
        FetchCommands::Sftp {
            host,
            port,
            username,
            private_key,
            passphrase,
            password,
            remote_dir,
            glob,
            min_age,
            out,
            stdout,
            known_hosts,
            known_hosts_out,
            trust_on_first_use,
            overwrite,
            max_files,
            since,
            quiet,
            list_only,
            provenance,
            state_dir,
            reprocess,
        } => sftp::cmd_fetch_sftp(sftp::SftpArgs {
            host,
            port,
            username,
            private_key,
            passphrase,
            password,
            remote_dir,
            glob_pattern: glob,
            min_age,
            out,
            stdout_mode: stdout,
            known_hosts_path: known_hosts,
            known_hosts_out,
            trust_on_first_use,
            overwrite,
            max_files,
            since,
            quiet,
            list_only,
            provenance,
            state_dir,
            reprocess,
        }),
    }
}
