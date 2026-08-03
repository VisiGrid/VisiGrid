//! `vgrid convert` and its readers/writers/filter helpers.
//! Extracted from main.rs 2026-07-29 (pure move; visibility widened to
//! pub(crate) so main.rs and calc keep calling these).

use std::collections::HashMap;
use std::io::{self, Read, Write};
use std::path::PathBuf;

use crate::sheet_ops;
use crate::{CliError, Format, InspectFormat, WhereClause, parse_where, resolve_where_columns, filter_row_indices, resolve_sheet};

// ============================================================================
// --select helpers
// ============================================================================

/// Find the first non-empty row in the sheet. Returns 0 if all rows are empty.
pub(crate) fn find_header_row(sheet: &visigrid_engine::sheet::Sheet, rows: usize, cols: usize) -> usize {
    for row in 0..rows {
        for col in 0..cols {
            if !sheet.get_display(row, col).trim().is_empty() {
                return row;
            }
        }
    }
    0
}

pub(crate) fn check_ambiguous_headers(canonical_headers: &[String]) -> Result<(), CliError> {
    let mut seen: HashMap<String, Vec<String>> = HashMap::new();
    for h in canonical_headers {
        let canon = h.trim();
        if canon.is_empty() { continue; }
        seen.entry(canon.to_lowercase())
            .or_default()
            .push(canon.to_string());
    }

    for (key, names) in &seen {
        if names.len() > 1 {
            return Err(CliError::args(format!(
                "ambiguous column name \"{}\" (matches: {})",
                key,
                names.join(", ")
            )));
        }
    }
    Ok(())
}

pub(crate) fn parse_rename_specs(spec: &str) -> Result<Vec<(String, String)>, CliError> {
    let mut result = Vec::new();
    for pair in spec.split(',') {
        let pair = pair.trim();
        if pair.is_empty() {
            continue;
        }
        let colon_pos = pair.find(':').ok_or_else(|| {
            CliError::args(format!("invalid --rename spec {:?}: expected OLD:NEW", pair))
                .with_hint("example: --rename 'order_number:Invoice,amount:Amount'")
        })?;
        let old_name = pair[..colon_pos].trim().to_string();
        let new_name = pair[colon_pos + 1..].trim().to_string();
        if old_name.is_empty() || new_name.is_empty() {
            return Err(CliError::args(format!("invalid --rename spec {:?}: both names required", pair))
                .with_hint("example: --rename 'order_number:Invoice'"));
        }
        result.push((old_name, new_name));
    }
    Ok(result)
}

pub(crate) fn parse_select_args(select_args: &[String]) -> Vec<String> {
    select_args
        .iter()
        .flat_map(|arg| arg.split(','))
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .collect()
}

pub(crate) fn resolve_select_columns(
    select_names: &[String],
    canonical_headers: &[String],
) -> Result<Vec<(usize, String)>, CliError> {
    // Build O(1) lookup: lowercased → (index, canonical name)
    let mut map: HashMap<String, (usize, String)> = HashMap::new();
    for (i, h) in canonical_headers.iter().enumerate() {
        let key = h.trim().to_lowercase();
        if key.is_empty() { continue; }
        map.insert(key, (i, h.clone()));
    }

    let mut result = Vec::with_capacity(select_names.len());
    let mut seen_indices = std::collections::HashSet::new();

    for name in select_names {
        let needle = name.trim().to_lowercase();
        match map.get(&needle) {
            Some((idx, canonical)) => {
                if !seen_indices.insert(*idx) {
                    return Err(CliError::args(
                        format!("duplicate column in --select: \"{}\"", name)
                    ));
                }
                result.push((*idx, canonical.clone()));
            }
            None => {
                let non_empty_count = canonical_headers.iter().filter(|h| !h.trim().is_empty()).count();
                let available: Vec<&str> = canonical_headers
                    .iter()
                    .map(|h| h.as_str())
                    .filter(|h| !h.trim().is_empty())
                    .take(25)
                    .collect();
                let suffix = if non_empty_count > 25 {
                    format!(" (+{} more)", non_empty_count - 25)
                } else {
                    String::new()
                };
                return Err(CliError::args(
                    format!("unknown column in --select: \"{}\"", name)
                ).with_hint(format!("available columns: {}{}", available.join(", "), suffix)));
            }
        }
    }

    Ok(result)
}

// ============================================================================
// convert
// ============================================================================

pub(crate) fn cmd_convert(
    input: Option<PathBuf>,
    from: Option<Format>,
    to: Format,
    output: Option<PathBuf>,
    sheet_arg: Option<String>,
    delimiter: char,
    headers: bool,
    where_clauses: Vec<String>,
    select_args: Vec<String>,
    rename: Option<String>,
    quiet: bool,
) -> Result<(), CliError> {

    // Validate --where requires --headers
    if !where_clauses.is_empty() && !headers {
        return Err(CliError::args("--where requires --headers")
            .with_hint("add --headers so column names can be resolved"));
    }

    // Validate --select requires --headers
    if !select_args.is_empty() && !headers {
        return Err(CliError::args("--select requires --headers")
            .with_hint("add --headers so column names can be resolved"));
    }

    // Validate --rename requires --headers
    if rename.is_some() && !headers {
        return Err(CliError::args("--rename requires --headers")
            .with_hint("add --headers so column names can be resolved"));
    }

    // Parse rename specs early (fail fast)
    let rename_specs = match &rename {
        Some(spec) => parse_rename_specs(spec)?,
        None => vec![],
    };

    // Determine input format
    let input_format = match (&input, from) {
        (None, None) => return Err(CliError::args("stdin requires --from to specify the input format")
            .with_hint("vgrid convert --from csv -t json")),
        (None, Some(f)) => f,
        (Some(path), None) => infer_format(path)?,
        (Some(_), Some(f)) => f, // --from overrides extension
    };

    // Validate --sheet is only used with multi-sheet formats
    if sheet_arg.is_some() && !matches!(input_format, Format::Xlsx | Format::Sheet | Format::JsonFull) {
        return Err(CliError::args("--sheet is not supported for single-sheet formats")
            .with_hint("--sheet works with .sheet, .xlsx, and json-full inputs"));
    }

    // ── Full-fidelity path: keep the whole workbook, never collapse ──
    //
    // Everything below this block funnels through `read_file`, which returns a
    // single `Sheet` and drops the workbook, the per-sheet layout and the
    // active-sheet index on the floor. That is fine for csv/json/lines, whose
    // shape is one table anyway, and wrong for the formats that carry a whole
    // workbook: it silently discarded every sheet but one, plus every column
    // width, row height and frozen pane.
    //
    // So workbook-in/workbook-out conversions are handled here instead, before
    // the collapse — but only when no reshaping flag is present, since
    // --sheet/--where/--select/--rename all mean "give me one reshaped table"
    // and must keep their existing single-sheet semantics.
    let reshaping = sheet_arg.is_some()
        || !where_clauses.is_empty()
        || !select_args.is_empty()
        || rename.is_some();

    let workbook_in = matches!(
        input_format,
        Format::JsonFull | Format::Xlsx | Format::Sheet
    );

    if !reshaping && workbook_in && matches!(to, Format::JsonFull | Format::Xlsx) {
        let (wb, layouts, active) = read_workbook_full(input.as_deref(), input_format)?;

        match to {
            // json-full always emits v2 workbook form — see docs/cli/convert.md.
            // A shape that varied with the input is what silently dropped
            // sheet 2 for anyone converting a multi-sheet file.
            Format::JsonFull => {
                let out = visigrid_io::json::export_workbook(&wb, &layouts, active)
                    .map_err(CliError::io)?;
                match &output {
                    Some(path) => std::fs::write(path, out.as_bytes())
                        .map_err(|e| CliError::io(e.to_string()))?,
                    None => {
                        let mut stdout = io::stdout();
                        stdout.write_all(out.as_bytes()).and_then(|_| stdout.write_all(b"\n"))
                            .map_err(|e| CliError::io(e.to_string()))?;
                    }
                }
            }
            // Layout has to be threaded here too, or a round trip through
            // json-full loses the panes and widths on the way back OUT —
            // import parsing them is only half the trip.
            Format::Xlsx => {
                let export_layouts: Vec<visigrid_io::xlsx::ExportLayout> =
                    layouts.iter().map(sheet_layout_to_export).collect();
                match &output {
                    Some(path) => {
                        visigrid_io::xlsx::export(&wb, path, Some(&export_layouts))
                            .map_err(CliError::io)?;
                    }
                    None => {
                        let (bytes, _) =
                            visigrid_io::xlsx::export_to_buffer(&wb, Some(&export_layouts))
                                .map_err(CliError::io)?;
                        io::stdout()
                            .write_all(&bytes)
                            .map_err(|e| CliError::io(e.to_string()))?;
                    }
                }
            }
            _ => unreachable!("guarded by the matches! above"),
        }

        if !quiet {
            if let Some(path) = &output {
                eprintln!("Wrote {} ({} sheets)", path.display(), wb.sheets().len());
            }
        }
        return Ok(());
    }

    // Read input into sheet (convert always starts at A1)
    let mut sheet = match &input {
        Some(path) => read_file(path, input_format, delimiter, sheet_arg.as_deref())?,
        None => read_stdin(input_format, delimiter, 0, 0)?,
    };

    let (bounds_rows, bounds_cols) = get_data_bounds(&sheet);

    // Find the actual header row (first non-empty row)
    let header_row = if headers && bounds_rows > 0 && bounds_cols > 0 {
        find_header_row(&sheet, bounds_rows, bounds_cols)
    } else {
        0
    };

    // Apply --rename to header cells (before canonical_headers, so renames flow through)
    if !rename_specs.is_empty() && headers && bounds_cols > 0 {
        for (old_name, new_name) in &rename_specs {
            let old_lower = old_name.to_lowercase();
            let mut found = false;
            for c in 0..bounds_cols {
                let cell_val = sheet.get_display(header_row, c);
                if cell_val.trim().to_lowercase() == old_lower {
                    sheet.set_value(header_row, c, new_name);
                    found = true;
                    break;
                }
            }
            if !found {
                let available: Vec<String> = (0..bounds_cols)
                    .map(|c| sheet.get_display(header_row, c).trim().to_string())
                    .collect();
                return Err(CliError::args(format!("rename: column {:?} not found", old_name))
                    .with_hint(format!("available columns: {}", available.join(", "))));
            }
        }
    }

    // Build canonical headers list once
    let canonical_headers: Vec<String> = if headers && bounds_cols > 0 {
        (0..bounds_cols).map(|c| sheet.get_display(header_row, c).trim().to_string()).collect()
    } else {
        vec![]
    };

    // Ambiguous header check (once, before --where or --select resolution)
    if (!where_clauses.is_empty() || !select_args.is_empty()) && headers {
        check_ambiguous_headers(&canonical_headers)?;
    }

    // Resolve and apply --where filters
    let row_filter = if !where_clauses.is_empty() {
        let parsed: Vec<WhereClause> = where_clauses
            .iter()
            .map(|e| parse_where(e))
            .collect::<Result<Vec<_>, _>>()?;
        let resolved = resolve_where_columns(&parsed, &canonical_headers)?;
        let (indices, skip_counts) = filter_row_indices(&sheet, &resolved, header_row);

        // Report unparseable cells to stderr (suppressed by --quiet)
        if !quiet {
            for (i, &count) in skip_counts.iter().enumerate() {
                if count > 0 {
                    eprintln!("note: {} rows skipped ({} not numeric)", count, parsed[i].column);
                }
            }
        }

        Some(indices)
    } else {
        None
    };

    // Resolve column selection (after --where, before write)
    let col_filter = if !select_args.is_empty() {
        let select_names = parse_select_args(&select_args);
        if select_names.is_empty() {
            return Err(CliError::args("empty --select list"));
        }
        let resolved = resolve_select_columns(&select_names, &canonical_headers)?;
        Some(resolved)
    } else {
        None
    };

    // json-full is a full-fidelity format: row/column filters would strip
    // formulas, which contradicts its purpose — refuse rather than surprise
    if matches!(to, Format::JsonFull) && (row_filter.is_some() || col_filter.is_some()) {
        return Err(CliError::args("--where/--select are not supported with -t json-full")
            .with_hint("filter to csv/json first, or export the full sheet"));
    }

    // Binary spreadsheet outputs: full fidelity (formulas, formats) when
    // unfiltered; --where/--select materialize display values first (same
    // semantics as the CSV writer — filtered rows can't keep live formulas).
    match to {
        Format::Xlsx => {
            let out_sheet = if row_filter.is_some() || col_filter.is_some() {
                materialize_filtered(&sheet, header_row, headers, row_filter.as_deref(), col_filter.as_deref())
            } else {
                sheet.clone()
            };
            let wb = visigrid_engine::workbook::Workbook::from_sheets(vec![out_sheet], 0);
            match output {
                Some(path) => {
                    visigrid_io::xlsx::export(&wb, &path, None)
                        .map_err(|e| CliError::io(e))?;
                }
                None => {
                    let (bytes, _) = visigrid_io::xlsx::export_to_buffer(&wb, None)
                        .map_err(|e| CliError::io(e))?;
                    io::stdout()
                        .write_all(&bytes)
                        .map_err(|e| CliError::io(e.to_string()))?;
                }
            }
            return Ok(());
        }
        Format::Sheet => {
            let Some(path) = output else {
                return Err(CliError::format("sheet format cannot be written to stdout")
                    .with_hint("use -o output.sheet to write to a file"));
            };
            let out_sheet = if row_filter.is_some() || col_filter.is_some() {
                materialize_filtered(&sheet, header_row, headers, row_filter.as_deref(), col_filter.as_deref())
            } else {
                sheet.clone()
            };
            visigrid_io::native::save(&out_sheet, &path)
                .map_err(|e| CliError::io(e))?;
            return Ok(());
        }
        _ => {}
    }

    // Write output
    let output_bytes = write_format(
        &sheet, to, delimiter, headers, header_row,
        row_filter.as_deref(),
        col_filter.as_deref(),
    )?;

    match output {
        Some(path) => {
            std::fs::write(&path, &output_bytes)
                .map_err(|e| CliError::io(format!("{}: {}", path.display(), e)))?;
        }
        None => {
            io::stdout()
                .write_all(&output_bytes)
                .map_err(|e| CliError::io(e.to_string()))?;
        }
    }

    Ok(())
}

/// Build a values-only sheet from filtered rows/columns (display values,
/// like the CSV writer — formulas can't survive row removal).
pub(crate) fn materialize_filtered(
    sheet: &visigrid_engine::sheet::Sheet,
    header_row: usize,
    headers: bool,
    row_filter: Option<&[usize]>,
    col_filter: Option<&[(usize, String)]>,
) -> visigrid_engine::sheet::Sheet {
    use visigrid_engine::sheet::{Sheet, SheetId};
    let (bounds_rows, bounds_cols) = get_data_bounds(sheet);
    let rows: Vec<usize> = match row_filter {
        Some(f) => {
            let mut v: Vec<usize> = Vec::new();
            if headers {
                v.push(header_row);
            }
            v.extend_from_slice(f);
            v
        }
        None => (0..bounds_rows).collect(),
    };
    let cols: Vec<usize> = match col_filter {
        Some(f) => f.iter().map(|(idx, _)| *idx).collect(),
        None => (0..bounds_cols).collect(),
    };
    let mut out = Sheet::new(SheetId(1), rows.len().max(1), cols.len().max(1));
    for (out_r, &src_r) in rows.iter().enumerate() {
        for (out_c, &src_c) in cols.iter().enumerate() {
            let v = sheet.get_display(src_r, src_c);
            if !v.is_empty() {
                out.set_value(out_r, out_c, &v);
            }
        }
    }
    out
}

pub(crate) fn infer_inspect_format(path: &PathBuf) -> Result<InspectFormat, CliError> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase());

    match ext.as_deref() {
        Some("csv") => Ok(InspectFormat::Csv),
        Some("tsv") => Ok(InspectFormat::Tsv),
        Some("xlsx") | Some("xls") | Some("xlsb") | Some("ods") => Ok(InspectFormat::Xlsx),
        Some("sheet") => Ok(InspectFormat::Sheet),
        _ => Err(CliError::args(format!(
            "cannot infer inspect format from extension {:?}",
            ext.as_deref().unwrap_or("(none)")
        )).with_hint("supported: .sheet, .xlsx, .xls, .xlsb, .ods, .csv, .tsv (or use --format)")),
    }
}

pub(crate) fn infer_format(path: &PathBuf) -> Result<Format, CliError> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase());

    match ext.as_deref() {
        Some("csv") => Ok(Format::Csv),
        Some("tsv") => Ok(Format::Tsv),
        Some("json") => Ok(Format::Json),
        Some("xlsx") | Some("xls") | Some("xlsb") | Some("ods") => Ok(Format::Xlsx),
        Some("sheet") => Ok(Format::Sheet),
        _ => Err(CliError::args(format!(
            "cannot infer format from extension {:?}",
            ext.as_deref().unwrap_or("(none)")
        )).with_hint("use --from with one of: csv, tsv, json, xlsx, sheet")),
    }
}

pub(crate) fn read_file(path: &PathBuf, format: Format, _delimiter: char, sheet_arg: Option<&str>) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    // TODO: Use custom delimiter when io crate supports it
    match format {
        Format::Csv => {
            visigrid_io::csv::import(path)
                .map_err(|e| CliError::parse(e))
        }
        Format::Tsv => {
            visigrid_io::csv::import_tsv(path)
                .map_err(|e| CliError::parse(e))
        }
        Format::Xlsx => {
            let (workbook, _stats) = visigrid_io::xlsx::import(path)
                .map_err(|e| CliError::parse(e))?;
            let (_, sheet) = resolve_sheet(&workbook, sheet_arg)?;
            Ok(sheet.clone())
        }
        Format::Sheet => {
            let workbook = visigrid_io::native::load_workbook(path)
                .map_err(|e| CliError::io(e))?;
            let (_, sheet) = resolve_sheet(&workbook, sheet_arg)?;
            Ok(sheet.clone())
        }
        Format::JsonFull => {
            let content = std::fs::read_to_string(path)
                .map_err(|e| CliError::io(e.to_string()))?;
            let (workbook, _, active) = visigrid_io::json::import_any(&content).map_err(CliError::io)?;
            return match sheet_arg {
                Some(_) => {
                    let (_, sheet) = resolve_sheet(&workbook, sheet_arg)?;
                    Ok(sheet.clone())
                }
                None => Ok(workbook.sheets()[active].clone()),
            };
        }
        Format::Json => {
            let content = std::fs::read_to_string(path)
                .map_err(|e| CliError::io(e.to_string()))?;
            parse_json(&content, 0, 0)
        }
        Format::Lines => {
            let content = std::fs::read_to_string(path)
                .map_err(|e| CliError::io(e.to_string()))?;
            parse_lines(&content, 0, 0)
        }
    }
}

/// Read an input as a whole workbook: every sheet, per-sheet layout in pixels,
/// and the active-sheet index. The counterpart to `read_file`, which collapses
/// to one sheet — see the full-fidelity block in `cmd_convert`.
fn read_workbook_full(
    input: Option<&std::path::Path>,
    format: Format,
) -> Result<
    (
        visigrid_engine::workbook::Workbook,
        Vec<visigrid_io::json::SheetLayout>,
        usize,
    ),
    CliError,
> {
    match format {
        Format::JsonFull => {
            let content = match input {
                Some(path) => std::fs::read_to_string(path)
                    .map_err(|e| CliError::io(e.to_string()))?,
                None => {
                    let mut buf = String::new();
                    io::stdin()
                        .read_to_string(&mut buf)
                        .map_err(|e| CliError::io(e.to_string()))?;
                    buf
                }
            };
            visigrid_io::json::import_any(&content).map_err(CliError::io)
        }
        Format::Xlsx => {
            let path = input.ok_or_else(|| {
                CliError::args("xlsx and sheet formats require file input")
            })?;
            let (wb, stats) = visigrid_io::xlsx::import(path).map_err(CliError::parse)?;
            // Raw Excel units (character widths, points) → pixels.
            let layouts: Vec<_> = stats
                .imported_layouts
                .iter()
                .map(|l| l.to_sheet_layout())
                .collect();
            let active = wb.active_sheet_index();
            let sheet_count = wb.sheets().len();
            Ok((wb, pad_layouts(layouts, sheet_count), active))
        }
        Format::Sheet => {
            let path = input.ok_or_else(|| {
                CliError::args("xlsx and sheet formats require file input")
            })?;
            let wb = visigrid_io::native::load_workbook(path).map_err(CliError::io)?;
            // .sheet stores layout workbook-wide, keyed by sheet index, and
            // already in pixels. It has no frozen-pane column.
            let native_layout = visigrid_io::native::load_layout(path);
            let layouts: Vec<_> = (0..wb.sheets().len())
                .map(|i| visigrid_io::json::SheetLayout {
                    col_widths: native_layout
                        .col_widths
                        .get(&i)
                        .map(|m| m.iter().map(|(&c, &w)| (c, w)).collect())
                        .unwrap_or_default(),
                    row_heights: native_layout
                        .row_heights
                        .get(&i)
                        .map(|m| m.iter().map(|(&r, &h)| (r, h)).collect())
                        .unwrap_or_default(),
                    ..Default::default()
                })
                .collect();
            let active = wb.active_sheet_index();
            Ok((wb, layouts, active))
        }
        _ => Err(CliError::format(
            "input format does not carry a workbook".to_string(),
        )),
    }
}

/// `export_workbook` indexes layouts positionally, so a short vector would
/// silently give later sheets a default layout. Pad to the sheet count.
fn pad_layouts(
    mut layouts: Vec<visigrid_io::json::SheetLayout>,
    sheet_count: usize,
) -> Vec<visigrid_io::json::SheetLayout> {
    while layouts.len() < sheet_count {
        layouts.push(Default::default());
    }
    layouts
}

fn sheet_layout_to_export(
    layout: &visigrid_io::json::SheetLayout,
) -> visigrid_io::xlsx::ExportLayout {
    visigrid_io::xlsx::ExportLayout {
        col_widths: layout.col_widths.iter().map(|(&c, &w)| (c, w)).collect(),
        row_heights: layout.row_heights.iter().map(|(&r, &h)| (r, h)).collect(),
        frozen_rows: layout.frozen_rows,
        frozen_cols: layout.frozen_cols,
        ..Default::default()
    }
}

pub(crate) fn read_stdin(format: Format, delimiter: char, into_row: usize, into_col: usize) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    let mut input = String::new();
    io::stdin()
        .read_to_string(&mut input)
        .map_err(|e| CliError::io(e.to_string()))?;

    if input.is_empty() {
        return Err(CliError::parse("no input received on stdin")
            .with_hint("cat file.csv | vgrid calc '=SUM(A:A)' --from csv"));
    }

    match format {
        Format::Csv => parse_csv(&input, delimiter as u8, into_row, into_col),
        Format::Tsv => parse_csv(&input, b'\t', into_row, into_col),
        Format::Json => parse_json(&input, into_row, into_col),
        Format::JsonFull => visigrid_io::json::import_full(&input).map_err(CliError::io),
        Format::Lines => parse_lines(&input, into_row, into_col),
        Format::Xlsx | Format::Sheet => {
            Err(CliError::args("xlsx and sheet formats require file input"))
        }
    }
}

pub(crate) fn parse_csv(content: &str, delimiter: u8, into_row: usize, into_col: usize) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    use visigrid_engine::sheet::{Sheet, SheetId};

    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(false)
        .from_reader(content.as_bytes());

    let mut sheet = Sheet::new(SheetId(1), 1000, 26);

    for (row_idx, result) in reader.records().enumerate() {
        let record = result.map_err(|e| CliError::parse(format!("line {}: {}", row_idx + 1, e)))?;
        for (col_idx, field) in record.iter().enumerate() {
            if !field.is_empty() {
                sheet.set_value(into_row + row_idx, into_col + col_idx, field);
            }
        }
    }

    Ok(sheet)
}

pub(crate) fn parse_json(content: &str, into_row: usize, into_col: usize) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    use visigrid_engine::sheet::{Sheet, SheetId};

    let value: serde_json::Value = serde_json::from_str(content)
        .map_err(|e| CliError::parse(format!("JSON parse error: {}", e)))?;

    let mut sheet = Sheet::new(SheetId(1), 1000, 26);

    match value {
        serde_json::Value::Array(rows) => {
            if rows.is_empty() {
                return Err(CliError::parse("empty input"));
            }

            // Check if array of arrays or array of objects
            if let Some(serde_json::Value::Object(_)) = rows.first() {
                // Array of objects - collect all keys lexicographically
                let mut all_keys: std::collections::BTreeSet<String> = std::collections::BTreeSet::new();
                for row in &rows {
                    if let serde_json::Value::Object(obj) = row {
                        for key in obj.keys() {
                            all_keys.insert(key.clone());
                        }
                    }
                }
                let keys: Vec<String> = all_keys.into_iter().collect();

                // Write header row
                for (col, key) in keys.iter().enumerate() {
                    sheet.set_value(into_row, into_col + col, key);
                }

                // Write data rows
                for (row_idx, row) in rows.iter().enumerate() {
                    if let serde_json::Value::Object(obj) = row {
                        for (col, key) in keys.iter().enumerate() {
                            if let Some(val) = obj.get(key) {
                                let cell_value = json_value_to_string(val, row_idx + 1, key)?;
                                if !cell_value.is_empty() {
                                    sheet.set_value(into_row + row_idx + 1, into_col + col, &cell_value);
                                }
                            }
                        }
                    }
                }
            } else {
                // Array of arrays
                for (row_idx, row) in rows.iter().enumerate() {
                    if let serde_json::Value::Array(cols) = row {
                        for (col_idx, val) in cols.iter().enumerate() {
                            let cell_value = json_value_to_string(val, row_idx, &col_idx.to_string())?;
                            if !cell_value.is_empty() {
                                sheet.set_value(into_row + row_idx, into_col + col_idx, &cell_value);
                            }
                        }
                    } else {
                        return Err(CliError::parse(format!("row {}: expected array", row_idx)));
                    }
                }
            }
        }
        _ => return Err(CliError::parse("JSON must be array of arrays or array of objects")),
    }

    Ok(sheet)
}

pub(crate) fn json_value_to_string(val: &serde_json::Value, row: usize, key: &str) -> Result<String, CliError> {
    match val {
        serde_json::Value::Null => Ok(String::new()),
        serde_json::Value::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
        serde_json::Value::Number(n) => Ok(n.to_string()),
        serde_json::Value::String(s) => Ok(s.clone()),
        serde_json::Value::Array(_) | serde_json::Value::Object(_) => {
            Err(CliError::parse(format!("non-scalar value at row {}, key \"{}\"", row, key)))
        }
    }
}

pub(crate) fn parse_lines(content: &str, into_row: usize, into_col: usize) -> Result<visigrid_engine::sheet::Sheet, CliError> {
    use visigrid_engine::sheet::{Sheet, SheetId};

    let lines: Vec<&str> = content.lines().collect();
    if lines.is_empty() {
        return Err(CliError::parse("empty input"));
    }

    let mut sheet = Sheet::new(SheetId(1), 1000, 26);
    for (row, line) in lines.iter().enumerate() {
        if !line.is_empty() {
            sheet.set_value(into_row + row, into_col, line);
        }
    }

    Ok(sheet)
}

pub(crate) fn write_format(
    sheet: &visigrid_engine::sheet::Sheet,
    format: Format,
    delimiter: char,
    headers: bool,
    header_row: usize,
    row_filter: Option<&[usize]>,
    col_filter: Option<&[(usize, String)]>,
) -> Result<Vec<u8>, CliError> {
    match format {
        Format::Csv => write_csv(sheet, delimiter as u8, headers, header_row, row_filter, col_filter),
        Format::Tsv => write_csv(sheet, b'\t', headers, header_row, row_filter, col_filter),
        Format::Json => write_json(sheet, headers, header_row, row_filter, col_filter),
        Format::JsonFull => visigrid_io::json::export_full(sheet)
            .map(|s| s.into_bytes())
            .map_err(CliError::io),
        Format::Lines => write_lines(sheet, header_row, row_filter, col_filter),
        Format::Xlsx => Err(CliError::format("xlsx output is handled before write_format")
            .with_hint("this is a bug — please report it")),
        Format::Sheet => Err(CliError::format("sheet format cannot be written to stdout")
            .with_hint("use -o output.sheet to write to a file")),
    }
}

pub(crate) fn write_csv(
    sheet: &visigrid_engine::sheet::Sheet,
    delimiter: u8,
    headers: bool,
    header_row: usize,
    row_filter: Option<&[usize]>,
    col_filter: Option<&[(usize, String)]>,
) -> Result<Vec<u8>, CliError> {
    let mut writer = csv::WriterBuilder::new()
        .delimiter(delimiter)
        .from_writer(Vec::new());

    let (rows, cols) = get_data_bounds(sheet);

    // Helper: push columns for a given row into the record
    let push_row = |record: &mut Vec<String>, row: usize| {
        match col_filter {
            Some(selected) => {
                for (idx, _) in selected {
                    record.push(sheet.get_display(row, *idx));
                }
            }
            None => {
                for col in 0..cols {
                    record.push(sheet.get_display(row, col));
                }
            }
        }
    };

    match row_filter {
        Some(indices) => {
            // Write header row + filtered data rows
            if rows > 0 {
                let mut record: Vec<String> = Vec::new();
                push_row(&mut record, header_row);
                writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
            }
            for &row in indices {
                let mut record: Vec<String> = Vec::new();
                push_row(&mut record, row);
                writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
            }
        }
        None => {
            if headers {
                // Write header row, then data rows starting after header
                if rows > 0 {
                    let mut record: Vec<String> = Vec::new();
                    push_row(&mut record, header_row);
                    writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
                }
                for row in (header_row + 1)..rows {
                    let mut record: Vec<String> = Vec::new();
                    push_row(&mut record, row);
                    writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
                }
            } else {
                for row in 0..rows {
                    let mut record: Vec<String> = Vec::new();
                    push_row(&mut record, row);
                    writer.write_record(&record).map_err(|e| CliError::io(e.to_string()))?;
                }
            }
        }
    }

    writer.into_inner().map_err(|e| CliError::io(e.to_string()))
}

pub(crate) fn write_json(
    sheet: &visigrid_engine::sheet::Sheet,
    headers: bool,
    header_row: usize,
    row_filter: Option<&[usize]>,
    col_filter: Option<&[(usize, String)]>,
) -> Result<Vec<u8>, CliError> {
    let (rows, cols) = get_data_bounds(sheet);

    if headers && rows > 0 {
        let data_rows: Vec<usize> = match row_filter {
            Some(indices) => indices.to_vec(),
            None => ((header_row + 1)..rows).collect(),
        };

        if let Some(selected) = col_filter {
            // --select path: build JSON manually to preserve key order
            let json_keys: Vec<(usize, String)> = selected.iter().map(|(idx, name)| {
                let sanitized: String = name
                    .to_lowercase()
                    .chars()
                    .map(|c| if c.is_alphanumeric() { c } else { '_' })
                    .collect();
                let key = if sanitized.is_empty() {
                    format!("col{}", idx)
                } else {
                    sanitized
                };
                (*idx, key)
            }).collect();

            let mut rows_json: Vec<Vec<(String, serde_json::Value)>> = Vec::new();
            for row in data_rows {
                let mut pairs = Vec::new();
                for (col_idx, key) in &json_keys {
                    let value = sheet.get_display(row, *col_idx);
                    pairs.push((key.clone(), string_to_json_value(&value)));
                }
                rows_json.push(pairs);
            }

            // Format manually to preserve key order
            let mut bytes = Vec::new();
            bytes.extend_from_slice(b"[\n");
            for (i, pairs) in rows_json.iter().enumerate() {
                bytes.extend_from_slice(b"  {\n");
                for (j, (key, value)) in pairs.iter().enumerate() {
                    let val_str = serde_json::to_string(value).map_err(|e| CliError::io(e.to_string()))?;
                    bytes.extend_from_slice(b"    ");
                    bytes.extend_from_slice(serde_json::to_string(key).map_err(|e| CliError::io(e.to_string()))?.as_bytes());
                    bytes.extend_from_slice(b": ");
                    bytes.extend_from_slice(val_str.as_bytes());
                    if j < pairs.len() - 1 {
                        bytes.push(b',');
                    }
                    bytes.push(b'\n');
                }
                bytes.extend_from_slice(b"  }");
                if i < rows_json.len() - 1 {
                    bytes.push(b',');
                }
                bytes.push(b'\n');
            }
            bytes.extend_from_slice(b"]\n");
            Ok(bytes)
        } else {
            // Standard path: array of objects with BTreeMap key ordering
            let mut header_names: Vec<String> = Vec::new();
            for col in 0..cols {
                let name = sheet.get_display(header_row, col);
                let sanitized: String = name
                    .to_lowercase()
                    .chars()
                    .map(|c| if c.is_alphanumeric() { c } else { '_' })
                    .collect();
                header_names.push(if sanitized.is_empty() {
                    format!("col{}", col)
                } else {
                    sanitized
                });
            }

            let mut objects: Vec<serde_json::Map<String, serde_json::Value>> = Vec::new();
            for row in data_rows {
                let mut obj = serde_json::Map::new();
                for (col, key) in header_names.iter().enumerate() {
                    let value = sheet.get_display(row, col);
                    obj.insert(key.clone(), string_to_json_value(&value));
                }
                objects.push(obj);
            }

            let mut bytes = serde_json::to_vec_pretty(&objects).map_err(|e| CliError::io(e.to_string()))?;
            bytes.push(b'\n');
            Ok(bytes)
        }
    } else {
        // Array of arrays (no col_filter since --select requires --headers)
        let mut rows_vec: Vec<Vec<serde_json::Value>> = Vec::new();
        let all_rows: Vec<usize> = match row_filter {
            Some(indices) => {
                let mut v = vec![header_row];
                v.extend_from_slice(indices);
                v
            }
            None => (0..rows).collect(),
        };
        for row in all_rows {
            let mut row_vec: Vec<serde_json::Value> = Vec::new();
            for col in 0..cols {
                let value = sheet.get_display(row, col);
                row_vec.push(string_to_json_value(&value));
            }
            rows_vec.push(row_vec);
        }

        let mut bytes = serde_json::to_vec_pretty(&rows_vec).map_err(|e| CliError::io(e.to_string()))?;
        bytes.push(b'\n');
        Ok(bytes)
    }
}

/// Convert a display string to a typed JSON value
/// Numbers become JSON numbers, booleans become JSON booleans, rest are strings
pub(crate) fn string_to_json_value(s: &str) -> serde_json::Value {
    if s.is_empty() {
        return serde_json::Value::String(String::new());
    }

    // Try to parse as number first
    if let Ok(n) = s.parse::<f64>() {
        // Check if it's an integer
        if n.fract() == 0.0 && n.abs() < i64::MAX as f64 {
            serde_json::json!(n as i64)
        } else {
            serde_json::json!(n)
        }
    } else if s == "TRUE" || s == "true" {
        serde_json::json!(true)
    } else if s == "FALSE" || s == "false" {
        serde_json::json!(false)
    } else {
        serde_json::json!(s)
    }
}

pub(crate) fn write_lines(
    sheet: &visigrid_engine::sheet::Sheet,
    header_row: usize,
    row_filter: Option<&[usize]>,
    col_filter: Option<&[(usize, String)]>,
) -> Result<Vec<u8>, CliError> {
    let mut output = Vec::new();
    let (rows, _) = get_data_bounds(sheet);

    // With --select: output the first selected column; without: column 0
    let output_col = match col_filter {
        Some(selected) => selected[0].0,
        None => 0,
    };

    let all_rows: Vec<usize> = match row_filter {
        Some(indices) => {
            let mut v = vec![header_row];
            v.extend_from_slice(indices);
            v
        }
        None => (0..rows).collect(),
    };

    for row in all_rows {
        let value = sheet.get_display(row, output_col);
        output.extend_from_slice(value.as_bytes());
        output.push(b'\n');
    }

    Ok(output)
}

pub(crate) fn get_data_bounds(sheet: &visigrid_engine::sheet::Sheet) -> (usize, usize) {
    sheet_ops::get_data_bounds(sheet)
}

