//! `vgrid parse` — transform artifacts into canonical CSV.

mod fiserv_cardpointe_v1;
mod statement_pdf;

use std::path::PathBuf;

use clap::Subcommand;

use crate::CliError;

#[derive(Subcommand)]
pub enum ParseCommands {
    /// Parse a processor statement PDF into canonical CSV
    #[command(name = "statement-pdf", after_help = "\
Examples:
  vgrid parse statement-pdf --template fiserv_cardpointe_v1 --file statement.pdf
  vgrid parse statement-pdf --template fiserv_cardpointe_v1 --file statement.pdf --out out.csv
  vgrid parse statement-pdf --template fiserv_cardpointe_v1 --file statement.pdf --save-raw /tmp/raw")]
    StatementPdf {
        /// Template ID (available: fiserv_cardpointe_v1)
        #[arg(long)]
        template: String,

        /// Path to PDF file
        #[arg(long)]
        file: PathBuf,

        /// Output CSV file path (default: stdout)
        #[arg(long)]
        out: Option<PathBuf>,

        /// Directory to save pdftotext.txt + engine_meta.json
        #[arg(long)]
        save_raw: Option<PathBuf>,

        /// Suppress progress on stderr
        #[arg(long, short = 'q')]
        quiet: bool,
    },
}

pub fn cmd_parse(command: ParseCommands) -> Result<(), CliError> {
    match command {
        ParseCommands::StatementPdf { template, file, out, save_raw, quiet } => {
            statement_pdf::cmd_parse_statement_pdf(&template, &file, &out, &save_raw, quiet)
        }
    }
}
