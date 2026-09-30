//! `vgrid pivot` — summarize a table from the terminal.
//!
//! With a file, the pivot is computed in memory and printed; the file is
//! never written. With `--session`, the same request is sent to a running
//! VisiGrid (GUI or `vgrid serve`) as a `create_pivot` structure op, and the
//! pivot lands on a new sheet there. Both paths resolve fields through
//! `visigrid_session_host::resolve_create_pivot`, so they accept and refuse
//! exactly the same requests.

use std::path::Path;

use visigrid_engine::formula::eval::Value;
use visigrid_engine::workbook::Workbook;
use visigrid_protocol::{PivotValueSpec, StructureOp};

use crate::util::{display_width, pad_right};
use crate::CliError;

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum PivotFormat {
    Table,
    Csv,
    Json,
}

pub struct PivotArgs {
    pub rows: Vec<String>,
    pub column: Option<String>,
    pub values: Vec<String>,
    pub sheet: Option<String>,
    pub range: Option<String>,
}

/// Split repeated and comma-separated flags into one list.
fn flatten(items: &[String]) -> Vec<String> {
    items
        .iter()
        .flat_map(|s| s.split(','))
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .collect()
}

/// `sum:Amount` → (sum, Amount). A bare field name takes the desktop's
/// default: sum for a numeric column, count otherwise.
fn value_specs(values: &[String]) -> Vec<PivotValueSpec> {
    flatten(values)
        .into_iter()
        .map(|v| match v.split_once(':') {
            Some((agg, field)) => PivotValueSpec { field: field.trim().to_string(), aggregation: Some(agg.trim().to_string()) },
            None => PivotValueSpec { field: v, aggregation: None },
        })
        .collect()
}

pub(crate) fn create_op(args: &PivotArgs, sheet: Option<usize>) -> StructureOp {
    StructureOp::CreatePivot {
        sheet,
        source: args.range.clone(),
        rows: flatten(&args.rows),
        column: args.column.clone(),
        values: value_specs(&args.values),
    }
}

/// Load any file `vgrid` reads into a workbook, computed.
fn load(path: &Path, delimiter: Option<&str>) -> Result<Workbook, CliError> {
    if !path.exists() {
        return Err(CliError::io(format!("{} not found", path.display())));
    }
    let ext = path.extension().and_then(|e| e.to_str()).map(|e| e.to_lowercase());
    let single = |sheet| Workbook::from_sheets(vec![sheet], 0);
    let mut wb = match ext.as_deref() {
        Some("sheet") | Some("vgrid") => visigrid_io::native::load_workbook(path).map_err(CliError::io)?,
        Some("json") => {
            let content = std::fs::read_to_string(path).map_err(|e| CliError::io(e.to_string()))?;
            visigrid_io::json::import_any(&content).map_err(CliError::io)?.0
        }
        Some("xlsx") | Some("xlsm") | Some("xls") | Some("xlsb") | Some("ods") => {
            visigrid_io::xlsx::import(path).map_err(CliError::parse)?.0
        }
        Some("parquet") => {
            let imported = visigrid_io::parquet::import(path).map_err(CliError::parse)?;
            // A pivot over the rows that fit is not the file's pivot.
            if let Some(note) = imported.truncation_message() {
                return Err(CliError::parse(format!("{}: {}", path.display(), note))
                    .with_hint("a pivot over part of the file would be wrong; split or filter it first"));
            }
            single(imported.sheet)
        }
        Some("tsv") | Some("tab") if delimiter.is_none() => single(visigrid_io::csv::import_tsv(path).map_err(CliError::parse)?),
        _ => match delimiter {
            Some(d) => {
                let d = crate::util::parse_delimiter(d)?;
                single(visigrid_io::csv::import_with_delimiter(path, d).map_err(CliError::parse)?)
            }
            None => single(visigrid_io::csv::import(path).map_err(CliError::parse)?),
        },
    };
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    Ok(wb)
}

pub fn cmd_pivot_file(path: &Path, args: &PivotArgs, delimiter: Option<&str>, format: PivotFormat) -> Result<(), CliError> {
    let mut wb = load(path, delimiter)?;
    let sheet_idx = match args.sheet.as_deref() {
        None => 0,
        Some(s) => crate::sheet_ops::resolve_sheet_by_arg(&wb, s)?,
    };
    let op = create_op(args, Some(sheet_idx));
    let (source, definition) = visigrid_session_host::resolve_create_pivot(&op, &wb)
        .map_err(|(_, msg)| CliError::args(msg))?;
    let (id, out_idx) = wb.create_pivot(source, definition).map_err(CliError::args)?;
    let (h, w) = wb.find_pivot(id).and_then(|(_, t)| t.extent).unwrap_or((0, 0));
    let out = wb.sheet(out_idx).ok_or_else(|| CliError::io("pivot output sheet missing"))?;
    let (h, w) = (h as usize, w as usize);

    let mut text = String::new();
    match format {
        PivotFormat::Json => {
            let header_rows = if matches!(op, StructureOp::CreatePivot { column: Some(_), .. }) { 2 } else { 1 };
            let json_value = |r: usize, c: usize| match out.get_computed_value(r, c) {
                Value::Number(n) if n.fract() == 0.0 && n.abs() < 9e15 => serde_json::json!(n as i64),
                Value::Number(n) => serde_json::json!(n),
                Value::Boolean(b) => serde_json::json!(b),
                Value::Empty => serde_json::Value::Null,
                _ => serde_json::json!(out.get_display(r, c)),
            };
            let headers: Vec<Vec<String>> =
                (0..header_rows.min(h)).map(|r| (0..w).map(|c| out.get_display(r, c)).collect()).collect();
            let rows: Vec<Vec<serde_json::Value>> =
                (header_rows.min(h)..h).map(|r| (0..w).map(|c| json_value(r, c)).collect()).collect();
            let doc = serde_json::json!({ "headers": headers, "rows": rows, "source_rows": source.data_rows() });
            text = serde_json::to_string_pretty(&doc).map_err(|e| CliError::io(e.to_string()))? + "\n";
        }
        PivotFormat::Csv => {
            let mut writer = csv::Writer::from_writer(Vec::new());
            for r in 0..h {
                let record: Vec<String> = (0..w).map(|c| raw_text(&out.get_computed_value(r, c), &out.get_display(r, c))).collect();
                writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
            }
            let bytes = writer.into_inner().map_err(|e| CliError::io(e.to_string()))?;
            text = String::from_utf8_lossy(&bytes).into_owned();
        }
        PivotFormat::Table => {
            let label_cols = match &op {
                StructureOp::CreatePivot { rows, .. } => rows.len().max(1),
                _ => 1,
            };
            let cells: Vec<Vec<String>> =
                (0..h).map(|r| (0..w).map(|c| out.get_formatted_display(r, c)).collect()).collect();
            let widths: Vec<usize> =
                (0..w).map(|c| cells.iter().map(|row| display_width(&row[c])).max().unwrap_or(0)).collect();
            for row in &cells {
                let line: Vec<String> = row
                    .iter()
                    .enumerate()
                    .map(|(c, v)| {
                        // Numbers right-align in value columns; row labels
                        // (a year, say) stay left.
                        if c >= label_cols && v.chars().next().is_some_and(|ch| ch.is_ascii_digit() || ch == '-' || ch == '(') {
                            format!("{}{}", " ".repeat(widths[c] - display_width(v)), v)
                        } else {
                            pad_right(v, widths[c])
                        }
                    })
                    .collect();
                text.push_str(line.join("  ").trim_end());
                text.push('\n');
            }
        }
    }
    // `vgrid pivot … | head` closes the pipe early; that is not an error.
    use std::io::Write;
    match std::io::stdout().lock().write_all(text.as_bytes()) {
        Err(e) if e.kind() != std::io::ErrorKind::BrokenPipe => return Err(CliError::io(e.to_string())),
        _ => {}
    }
    if format == PivotFormat::Table {
        eprintln!("{} source rows → {} × {}", source.data_rows(), h, w);
    }
    Ok(())
}

/// CSV wants unformatted numbers (no thousands separators).
fn raw_text(v: &Value, display: &str) -> String {
    match v {
        Value::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        _ => display.to_string(),
    }
}

pub fn cmd_pivot_session(session_id: Option<&str>, args: &PivotArgs, refresh: Option<Option<String>>) -> Result<(), CliError> {
    let discovery = crate::resolve_session(session_id)?;
    let token = crate::get_session_token()?;
    let mut client = crate::session::SessionClient::connect(&discovery, &token).map_err(CliError::session)?;
    let op = match refresh {
        Some(pivot) => StructureOp::RefreshPivot { pivot },
        None => {
            let sheet = match args.sheet.as_deref() {
                None => None,
                Some(s) => Some(s.parse::<usize>().map_err(|_| {
                    CliError::args("with --session, --sheet is a 0-based sheet index")
                })?),
            };
            create_op(args, sheet)
        }
    };
    let r = client.structure(op).map_err(|e| {
        let text = e.to_string();
        let err = CliError::session(e);
        if text.contains("malformed_message") {
            err.with_hint("that VisiGrid predates pivot support; update the app or `vgrid serve` to match this vgrid")
        } else {
            err
        }
    })?;
    println!("{} (revision {})", r.description, r.revision);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn value_specs_accept_bare_names_and_lists() {
        let specs = value_specs(&["sum:Amount, count:Order".into(), "Qty".into()]);
        assert_eq!(
            specs,
            vec![
                PivotValueSpec { field: "Amount".into(), aggregation: Some("sum".into()) },
                PivotValueSpec { field: "Order".into(), aggregation: Some("count".into()) },
                PivotValueSpec { field: "Qty".into(), aggregation: None },
            ]
        );
        assert_eq!(flatten(&["Region,Rep".into(), " Month ".into()]), vec!["Region", "Rep", "Month"]);
    }
}
