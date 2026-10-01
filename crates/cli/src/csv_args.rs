//! CSV import flags shared by the commands that read CSV (convert, calc).
//!
//! These are the options of VisiGrid's import settings dialog; the dialog
//! prints the same flags as "Same import from the command line".

use visigrid_io::csv::{ColumnRule, CsvOptions, DateOrder, Encoding};

use crate::CliError;

#[derive(clap::Args, Debug, Clone, Default)]
pub struct CsvImportArgs {
    /// CSV: keep these columns as text (names or letters, comma-separated)
    #[arg(long, value_name = "COLS", value_delimiter = ',')]
    pub text: Vec<String>,

    /// CSV: read these columns as numbers, even 007 (names or letters, comma-separated)
    #[arg(long, value_name = "COLS", value_delimiter = ',')]
    pub number: Vec<String>,

    /// CSV: read a column as dates, COL=ymd|dmy|mdy (repeatable). Nothing else becomes a date.
    #[arg(long, value_name = "COL=ORDER")]
    pub date: Vec<String>,

    /// CSV: leave these columns out (names or letters, comma-separated)
    #[arg(long, value_name = "COLS", value_delimiter = ',')]
    pub skip: Vec<String>,

    /// CSV: numbers use a decimal comma (1.234,56)
    #[arg(long)]
    pub decimal_comma: bool,

    /// CSV: text encoding (utf-8, utf-16, windows-1252); detected when omitted
    #[arg(long, value_name = "ENCODING")]
    pub encoding: Option<String>,

    /// CSV: the first row is data, not column names (column flags then use letters)
    #[arg(long)]
    pub no_header: bool,

    /// CSV: evaluate cells that start with = (left as text by default)
    #[arg(long)]
    pub formulas: bool,
}

impl CsvImportArgs {
    /// Options for reading CSV. A delimiter other than the default comma is
    /// used as given; the default lets the importer sniff it.
    pub fn to_options(&self, delimiter: char) -> Result<CsvOptions, CliError> {
        let mut columns: Vec<(String, ColumnRule)> = Vec::new();
        let mut add = |names: &[String], rule: ColumnRule| {
            for n in names.iter().map(|n| n.trim()).filter(|n| !n.is_empty()) {
                columns.push((n.to_string(), rule));
            }
        };
        add(&self.text, ColumnRule::Text);
        add(&self.number, ColumnRule::Number);
        add(&self.skip, ColumnRule::Skip);
        for spec in &self.date {
            let (col, order) = spec
                .rsplit_once('=')
                .ok_or_else(|| CliError::args(format!("--date expects COL=ORDER, got {spec}")).with_hint("e.g. --date order_date=dmy"))?;
            let order = DateOrder::parse(order.trim()).ok_or_else(|| {
                CliError::args(format!("unknown date order {order:?}")).with_hint("use ymd, dmy or mdy")
            })?;
            columns.push((col.trim().to_string(), ColumnRule::Date(order)));
        }
        let encoding = match &self.encoding {
            None => None,
            Some(e) => Some(Encoding::parse(e).ok_or_else(|| {
                CliError::args(format!("unknown encoding {e:?}")).with_hint("use utf-8, utf-16 or windows-1252")
            })?),
        };
        Ok(CsvOptions {
            delimiter: (delimiter != ',').then_some(delimiter as u8),
            encoding,
            no_header: self.no_header,
            decimal_comma: self.decimal_comma,
            evaluate_formulas: self.formulas,
            columns,
            origin: (0, 0),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn flags_become_options_and_back() {
        let args = CsvImportArgs {
            text: vec!["zip".into()],
            number: vec!["amount".into()],
            date: vec!["order date=dmy".into()],
            skip: vec!["notes".into()],
            decimal_comma: true,
            formulas: true,
            ..Default::default()
        };
        let options = args.to_options(';').unwrap();
        assert_eq!(
            options.cli_flags().join(" "),
            "--text zip --number amount --date 'order date'=dmy --skip notes --delimiter ';' --decimal-comma --formulas"
        );
        assert!(CsvImportArgs { date: vec!["x".into()], ..Default::default() }.to_options(',').is_err());
        assert!(CsvImportArgs { encoding: Some("ebcdic".into()), ..Default::default() }.to_options(',').is_err());
        assert_eq!(CsvImportArgs::default().to_options(',').unwrap().delimiter, None, "comma is the default: sniff");
    }
}
