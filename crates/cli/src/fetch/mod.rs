//! `vgrid fetch` — pull data from external sources into canonical CSV.

mod brex;
mod common;
mod gusto;
mod mercury;
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
    }
}
