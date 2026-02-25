//! `vgrid fetch` — pull data from external sources into canonical CSV.

mod authorizenet;
mod brex;
mod common;
mod digits;
mod fiserv;
mod gusto;
mod netsuite;
pub(crate) mod http;
mod mercury;
mod qbo;
mod ramp;
mod sftp;
mod stripe;
mod xero;

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

    /// Fetch settled transactions from Authorize.net
    #[command(after_help = "\
Examples:
  vgrid fetch authorizenet --from 2026-01-01 --to 2026-01-31
  vgrid fetch authorizenet --from 2026-01-01 --to 2026-01-31 --out authorizenet.csv
  vgrid fetch authorizenet --from 2026-01-01 --to 2026-01-31 --sandbox
  AUTHORIZENET_API_LOGIN_ID=xxx AUTHORIZENET_TRANSACTION_KEY=yyy vgrid fetch authorizenet --from 2026-01-01 --to 2026-01-31")]
    Authorizenet {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// API Login ID (default: AUTHORIZENET_API_LOGIN_ID env)
        #[arg(long)]
        api_login_id: Option<String>,

        /// Transaction Key (default: AUTHORIZENET_TRANSACTION_KEY env)
        #[arg(long)]
        transaction_key: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,

        /// Use sandbox API endpoint
        #[arg(long)]
        sandbox: bool,
    },

    /// Fetch settled transactions from Fiserv/CardPointe
    #[command(after_help = "\
Examples:
  vgrid fetch fiserv --from 2026-01-01 --to 2026-01-31
  vgrid fetch fiserv --from 2026-01-01 --to 2026-01-31 --out fiserv.csv
  vgrid fetch fiserv --from 2026-01-01 --to 2026-01-31 --funding --save-raw /tmp/fiserv_raw
  FISERV_MERCHANT_ID=xxx FISERV_API_USERNAME=user FISERV_API_PASSWORD=pass vgrid fetch fiserv --from 2026-01-01 --to 2026-01-31")]
    Fiserv {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// CardPointe API base URL (default: FISERV_API_URL env or UAT)
        #[arg(long)]
        api_url: Option<String>,

        /// Merchant ID (default: FISERV_MERCHANT_ID env)
        #[arg(long)]
        merchant_id: Option<String>,

        /// API username (default: FISERV_API_USERNAME env)
        #[arg(long)]
        api_username: Option<String>,

        /// API password (default: FISERV_API_PASSWORD env)
        #[arg(long)]
        api_password: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Fetch funding/deposit data instead of settled transactions
        #[arg(long)]
        funding: bool,

        /// Directory to save raw JSON responses (one file per day)
        #[arg(long)]
        save_raw: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,
    },

    /// Fetch general ledger transactions from NetSuite
    #[command(after_help = "\
Examples:
  vgrid fetch netsuite --from 2026-01-01 --to 2026-01-31 --account-id 1234567 --consumer-key abc --consumer-secret def --token-id ghi --token-secret jkl
  NETSUITE_ACCOUNT_ID=1234567 NETSUITE_CONSUMER_KEY=abc NETSUITE_CONSUMER_SECRET=def NETSUITE_TOKEN_ID=ghi NETSUITE_TOKEN_SECRET=jkl vgrid fetch netsuite --from 2026-01-01 --to 2026-01-31")]
    Netsuite {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// NetSuite account ID (default: NETSUITE_ACCOUNT_ID env)
        #[arg(long)]
        account_id: Option<String>,

        /// OAuth consumer key (default: NETSUITE_CONSUMER_KEY env)
        #[arg(long)]
        consumer_key: Option<String>,

        /// OAuth consumer secret (default: NETSUITE_CONSUMER_SECRET env)
        #[arg(long)]
        consumer_secret: Option<String>,

        /// OAuth token ID (default: NETSUITE_TOKEN_ID env)
        #[arg(long)]
        token_id: Option<String>,

        /// OAuth token secret (default: NETSUITE_TOKEN_SECRET env)
        #[arg(long)]
        token_secret: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

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

    /// Fetch bank transactions from Xero
    #[command(after_help = "\
Examples:
  # Fetch bank transactions for a bank account
  vgrid fetch xero --credentials ~/.config/vgrid/xero.json --from 2026-01-01 --to 2026-01-31 --account \"Checking\"
  vgrid fetch xero --credentials ~/xero.json --from 2026-01-01 --to 2026-01-31 --account \"Checking\" --out xero.csv

  # Only transactions (no transfers)
  vgrid fetch xero --credentials ~/xero.json --from 2026-01-01 --to 2026-01-31 --account \"Checking\" --include transaction

  # Direct token (skips refresh)
  vgrid fetch xero --access-token eyJ... --tenant-id abc-def-123 --from 2026-01-01 --to 2026-01-31 --account-id guid-here")]
    Xero {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Path to Xero OAuth2 credentials JSON file
        #[arg(long)]
        credentials: Option<PathBuf>,

        /// Xero access token (skips refresh, requires --tenant-id)
        #[arg(long)]
        access_token: Option<String>,

        /// Xero tenant ID (required with --access-token)
        #[arg(long)]
        tenant_id: Option<String>,

        /// Bank account name in Xero (resolved to ID via Account query)
        #[arg(long)]
        account: Option<String>,

        /// Bank account ID directly (skip name resolution)
        #[arg(long)]
        account_id: Option<String>,

        /// Entity types to query, comma-separated (default: transaction,transfer)
        #[arg(long)]
        include: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,
    },

    /// Fetch ledger entries from Digits
    #[command(after_help = "\
Examples:
  vgrid fetch digits --credentials ~/.config/vgrid/digits.json --from 2026-01-01 --to 2026-01-31
  vgrid fetch digits --access-token eyJ... --legal-entity-id le_123 --from 2026-01-01 --to 2026-01-31 --out digits.csv")]
    Digits {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Path to Digits OAuth2 credentials JSON file
        #[arg(long)]
        credentials: Option<PathBuf>,

        /// Digits access token (skips refresh, requires --legal-entity-id)
        #[arg(long)]
        access_token: Option<String>,

        /// Digits legal entity ID (required with --access-token)
        #[arg(long)]
        legal_entity_id: Option<String>,

        /// Account name filter
        #[arg(long)]
        account: Option<String>,

        /// Account ID filter
        #[arg(long)]
        account_id: Option<String>,

        /// Entry types to include, comma-separated (default: credit,debit)
        #[arg(long)]
        include: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,
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

    /// Fetch card transactions from Ramp
    #[command(name = "ramp-card", after_help = "\
Examples:
  vgrid fetch ramp-card --from 2026-01-01 --to 2026-01-31
  vgrid fetch ramp-card --from 2026-01-01 --to 2026-01-31 --out ramp-card.csv
  vgrid fetch ramp-card --from 2026-01-01 --to 2026-01-31 --api-key ramp_token_...
  vgrid fetch ramp-card --from 2026-01-01 --to 2026-01-31 --card card_id_123
  RAMP_ACCESS_TOKEN=ramp_token_... vgrid fetch ramp-card --from 2026-01-01 --to 2026-01-31")]
    RampCard {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Ramp access token (default: RAMP_ACCESS_TOKEN env)
        #[arg(long)]
        api_key: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,

        /// Transaction state filter (default: CLEARED)
        #[arg(long)]
        state: Option<String>,

        /// Filter by card ID
        #[arg(long)]
        card: Option<String>,

        /// Filter by entity ID
        #[arg(long)]
        entity: Option<String>,
    },

    /// Fetch business account transactions from Ramp
    #[command(name = "ramp-bank", after_help = "\
Examples:
  vgrid fetch ramp-bank --from 2026-01-01 --to 2026-01-31
  vgrid fetch ramp-bank --from 2026-01-01 --to 2026-01-31 --out ramp-bank.csv
  vgrid fetch ramp-bank --from 2026-01-01 --to 2026-01-31 --api-key ramp_token_...
  vgrid fetch ramp-bank --from 2026-01-01 --to 2026-01-31 --entity entity_id_123
  RAMP_ACCESS_TOKEN=ramp_token_... vgrid fetch ramp-bank --from 2026-01-01 --to 2026-01-31")]
    RampBank {
        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Ramp access token (default: RAMP_ACCESS_TOKEN env)
        #[arg(long)]
        api_key: Option<String>,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,

        /// Transaction state filter (default: CLEARED)
        #[arg(long)]
        state: Option<String>,

        /// Filter by entity ID
        #[arg(long)]
        entity: Option<String>,
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

    /// Fetch data from any HTTP API using a mapping file
    #[command(after_help = "\
Examples:
  # Fetch with bearer token from env var
  vgrid fetch http --url https://api.vendor.com/v1/payments \\
    --auth bearer-env:VENDOR_API_TOKEN --from 2026-01-01 --to 2026-01-31 \\
    --map mapping.json --out payments.csv

  # Preview raw API response (no mapping applied)
  vgrid fetch http --url https://api.vendor.com/v1/payments \\
    --auth bearer-env:VENDOR_API_TOKEN --from 2026-01-01 --to 2026-01-31 \\
    --map mapping.json --sample

  # Save raw response for audit trail
  vgrid fetch http --url https://api.vendor.com/v1/payments \\
    --auth bearer-env:VENDOR_API_TOKEN --from 2026-01-01 --to 2026-01-31 \\
    --map mapping.json --out payments.csv --save-raw raw.json

  # API key in custom header
  vgrid fetch http --url https://api.vendor.com/v1/data \\
    --auth header-env:X-API-Key:MY_API_KEY --from 2026-01-01 --to 2026-01-31 \\
    --map mapping.json --out data.csv

  # No auth (public API)
  vgrid fetch http --url https://api.example.com/rates \\
    --auth none --from 2026-01-01 --to 2026-01-31 \\
    --map mapping.json --out rates.csv")]
    Http {
        /// HTTPS URL of the API endpoint
        #[arg(long)]
        url: String,

        /// Auth method (env-var indirection for safety)
        ///
        /// bearer-env:VAR — Bearer token from environment variable
        /// header-env:NAME:VAR — Custom header from environment variable
        /// basic-env:USER_VAR:PASS_VAR — Basic auth from environment variables
        /// none — No authentication
        #[arg(long)]
        auth: String,

        /// Start date inclusive (YYYY-MM-DD)
        #[arg(long)]
        from: String,

        /// End date exclusive (YYYY-MM-DD)
        #[arg(long)]
        to: String,

        /// Path to mapping JSON file
        #[arg(long, visible_alias = "map")]
        mapping: PathBuf,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Save raw API response to file (for audit trail)
        #[arg(long)]
        save_raw: Option<PathBuf>,

        /// Print raw response JSON and exit (no mapping)
        #[arg(long)]
        sample: bool,

        /// Request timeout in seconds (default: 15)
        #[arg(long)]
        timeout: Option<u64>,

        /// Maximum items to process (default: 10000)
        #[arg(long, default_value = "10000")]
        max_items: Option<usize>,

        /// Maximum pages to fetch when paginating (default: 100)
        #[arg(long)]
        max_pages: Option<u32>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,

        /// Write a signed request fingerprint (JSON) to this path
        #[arg(long)]
        fingerprint: Option<PathBuf>,
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
        FetchCommands::Authorizenet {
            from,
            to,
            api_login_id,
            transaction_key,
            out,
            quiet,
            sandbox,
        } => authorizenet::cmd_fetch_authorizenet(from, to, api_login_id, transaction_key, out, quiet, sandbox),
        FetchCommands::Fiserv {
            from,
            to,
            api_url,
            merchant_id,
            api_username,
            api_password,
            out,
            funding,
            save_raw,
            quiet,
        } => fiserv::cmd_fetch_fiserv(from, to, api_url, merchant_id, api_username, api_password, out, funding, save_raw, quiet),
        FetchCommands::Netsuite {
            from,
            to,
            account_id,
            consumer_key,
            consumer_secret,
            token_id,
            token_secret,
            out,
            quiet,
        } => netsuite::cmd_fetch_netsuite(from, to, account_id, consumer_key, consumer_secret, token_id, token_secret, out, quiet),
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
        FetchCommands::Xero {
            from,
            to,
            credentials,
            access_token,
            tenant_id,
            account,
            account_id,
            include,
            out,
            quiet,
        } => xero::cmd_fetch_xero(
            from,
            to,
            credentials,
            access_token,
            tenant_id,
            account,
            account_id,
            include,
            out,
            quiet,
        ),
        FetchCommands::Digits {
            from,
            to,
            credentials,
            access_token,
            legal_entity_id,
            account,
            account_id,
            include,
            out,
            quiet,
        } => digits::cmd_fetch_digits(
            from,
            to,
            credentials,
            access_token,
            legal_entity_id,
            account,
            account_id,
            include,
            out,
            quiet,
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
        FetchCommands::RampCard {
            from,
            to,
            api_key,
            out,
            quiet,
            state,
            card,
            entity,
        } => ramp::card::cmd_fetch_ramp_card(from, to, api_key, out, quiet, state, card, entity),
        FetchCommands::RampBank {
            from,
            to,
            api_key,
            out,
            quiet,
            state,
            entity,
        } => ramp::bank::cmd_fetch_ramp_bank(from, to, api_key, out, quiet, state, entity),
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
        FetchCommands::Http {
            url,
            auth,
            from,
            to,
            mapping,
            out,
            save_raw,
            sample,
            timeout,
            max_items,
            max_pages,
            quiet,
            fingerprint,
        } => http::cmd_fetch_http(url, auth, from, to, mapping, out, save_raw, sample, timeout, max_items, max_pages, quiet, fingerprint),
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
