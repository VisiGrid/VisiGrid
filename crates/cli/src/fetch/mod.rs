//! `vgrid fetch` — pull data from external sources into canonical CSV.

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
}

pub fn cmd_fetch(command: FetchCommands) -> Result<(), CliError> {
    match command {
        FetchCommands::Stripe {
            from,
            to,
            api_key,
            out,
            account,
            quiet,
        } => stripe::cmd_fetch_stripe(from, to, api_key, out, account, quiet),
    }
}
