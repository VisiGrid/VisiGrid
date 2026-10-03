//! Values-only Parquet export. Analysis and writing share the same typed cells;
//! display formatting never supplies numeric data. This is not a source-file
//! round trip: the spreadsheet has already normalised imported values.
use std::{collections::HashSet, io::Write, path::Path, sync::Arc};

use parquet::{
    basic::{Compression, LogicalType, Repetition, TimeUnit, Type as PhysicalType},
    data_type::{BoolType, ByteArray, ByteArrayType, DoubleType, Int32Type, Int64Type},
    file::{properties::WriterProperties, writer::SerializedFileWriter},
    schema::types::Type,
};
use serde::Serialize;
use visigrid_engine::{
    cell::NumberFormat,
    formula::eval::Value,
    sheet::{custom_code_temporal_format, Sheet, TemporalFormat},
};

const DAY_MICROS: i64 = 86_400_000_000;
const ROW_GROUP_ROWS: usize = 2048;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum ColumnType {
    String,
    Int64,
    Double,
    Boolean,
    Date,
    Time,
    Timestamp,
}

#[derive(Clone, Debug)]
pub struct Column {
    /// Zero-based source column.
    pub index: usize,
    pub name: String,
    /// Explicit conversion of mixed values to text (errors still block export).
    pub as_text: bool,
}

#[derive(Debug, Serialize)]
pub struct Issue {
    pub cell: String,
    pub message: String,
}

#[derive(Debug, Serialize)]
pub struct ColumnReport {
    pub name: String,
    pub source_column: usize,
    pub data_type: Option<ColumnType>,
    pub null_count: usize,
    pub issue_count: usize,
    /// Bounded examples, even for a very large invalid column.
    pub issues: Vec<Issue>,
    pub warnings: Vec<String>,
}

#[derive(Debug, Serialize)]
pub struct Report {
    pub ready: bool,
    pub row_count: usize,
    pub columns: Vec<ColumnReport>,
}

/// Holds an immutable sheet reference: the cells written are the cells checked.
pub struct Plan<'a> {
    sheet: &'a Sheet,
    rows: Vec<usize>,
    columns: Vec<Column>,
    report: Report,
}

pub fn column_name(mut index: usize) -> String {
    let mut bytes = Vec::new();
    loop {
        bytes.push(b'A' + (index % 26) as u8);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    bytes.reverse();
    String::from_utf8(bytes).unwrap()
}

#[derive(Debug)]
enum Cell {
    Null,
    Text(String),
    Int(i64),
    Float(f64),
    Bool(bool),
    Date(i32),
    Time(i64),
    Timestamp(i64),
}

impl Cell {
    fn kind(&self) -> Option<ColumnType> {
        Some(match self {
            Self::Null => return None,
            Self::Text(_) => ColumnType::String,
            Self::Int(_) => ColumnType::Int64,
            Self::Float(_) => ColumnType::Double,
            Self::Bool(_) => ColumnType::Boolean,
            Self::Date(_) => ColumnType::Date,
            Self::Time(_) => ColumnType::Time,
            Self::Timestamp(_) => ColumnType::Timestamp,
        })
    }
}

// Use arithmetic, not the display date converter (which clamps serial <= 0).
// Limit dates to 0001..9999 for interoperable calendar values.
fn date_days(n: f64) -> Result<i32, String> {
    let day = n.floor();
    if !(-693_594.0..=2_958_465.0).contains(&day) {
        return Err("date is outside years 0001..9999".into());
    }
    if day == 60.0 {
        return Err("1900-02-29 is not a real calendar date".into());
    }
    Ok(day as i32 - if day < 60.0 { 25_568 } else { 25_569 })
}

fn read_cell(sheet: &Sheet, row: usize, col: usize, as_text: bool) -> Result<Cell, String> {
    if sheet
        .get_cell_opt(row, col)
        .is_some_and(|c| c.value().is_formula())
        && sheet.get_cached_value(row, col).is_none()
    {
        return Err("formula has no computed result; recalculate before exporting".into());
    }
    let value = sheet.get_computed_value(row, col);
    if let Value::Error(error) = &value {
        return Err(format!("formula error: {error}"));
    }
    if let Value::Number(n) = &value {
        if !n.is_finite() {
            return Err("non-finite number".into());
        }
    }
    if as_text {
        return Ok(match value {
            Value::Empty => Cell::Null,
            Value::Text(s) => Cell::Text(s),
            // Stable raw numbers, not rounded/formatted display text.
            Value::Number(n) => Cell::Text(n.to_string()),
            Value::Boolean(b) => Cell::Text(if b { "TRUE" } else { "FALSE" }.into()),
            Value::Error(_) => unreachable!(),
        });
    }
    let format = if matches!(value, Value::Number(_)) {
        sheet.get_format(row, col).number_format
    } else {
        NumberFormat::General
    };
    let number_format = match format {
        NumberFormat::Custom(code) => {
            let temporal = custom_code_temporal_format(&code)?;
            match temporal {
                Some(TemporalFormat::Date) => NumberFormat::Date {
                    style: Default::default(),
                },
                Some(TemporalFormat::Time) => NumberFormat::Time,
                Some(TemporalFormat::DateTime) => NumberFormat::DateTime,
                None => NumberFormat::Custom(code),
            }
        }
        other => other,
    };
    Ok(match value {
        Value::Empty => Cell::Null,
        Value::Text(s) => Cell::Text(s),
        Value::Boolean(b) => Cell::Bool(b),
        Value::Error(_) => unreachable!(),
        Value::Number(n) => match number_format {
            NumberFormat::Date { .. } => {
                if n.fract() != 0.0 {
                    return Err(
                        "date contains a time component; use DateTime format or export as text"
                            .into(),
                    );
                }
                Cell::Date(date_days(n)?)
            }
            NumberFormat::DateTime => {
                let days = date_days(n)? as i64;
                let fraction = ((n - n.floor()) * DAY_MICROS as f64).round() as i64;
                Cell::Timestamp(days * DAY_MICROS + fraction)
            }
            NumberFormat::Time => {
                if !(0.0..1.0).contains(&n) {
                    return Err(
                        "time must be within one day; durations need a numeric format or text"
                            .into(),
                    );
                }
                let micros = (n * DAY_MICROS as f64).round() as i64;
                if micros >= DAY_MICROS {
                    return Err("time rounds past midnight at microsecond precision".into());
                }
                Cell::Time(micros)
            }
            _ if n.fract() == 0.0 && n.abs() <= 9_007_199_254_740_991.0 => Cell::Int(n as i64),
            _ => Cell::Float(n),
        },
    })
}

fn merge(a: ColumnType, b: ColumnType) -> Option<ColumnType> {
    use ColumnType::*;
    if a == b {
        return Some(a);
    }
    match (a, b) {
        (Int64, Double) | (Double, Int64) => Some(Double),
        (Date, Timestamp) | (Timestamp, Date) => Some(Timestamp),
        _ => None,
    }
}

pub fn analyze(sheet: &Sheet, rows: Vec<usize>, columns: Vec<Column>) -> Result<Plan<'_>, String> {
    if columns.is_empty() {
        return Err("no columns to export".into());
    }
    let mut names = HashSet::new();
    for col in &columns {
        if col.name.trim().is_empty() {
            return Err(format!(
                "column {} has a blank header",
                column_name(col.index)
            ));
        }
        if !names.insert(col.name.trim().to_lowercase()) {
            return Err(format!("duplicate header: {:?}", col.name));
        }
    }
    let mut reports = Vec::with_capacity(columns.len());
    for col in &columns {
        let mut report = ColumnReport {
            name: col.name.clone(),
            source_column: col.index,
            data_type: if col.as_text {
                Some(ColumnType::String)
            } else {
                None
            },
            null_count: 0,
            issue_count: 0,
            issues: Vec::new(),
            warnings: Vec::new(),
        };
        for &row in &rows {
            let result = read_cell(sheet, row, col.index, col.as_text).and_then(|cell| {
                match (report.data_type, cell.kind()) {
                    (_, None) => report.null_count += 1,
                    (None, Some(kind)) => report.data_type = Some(kind),
                    (Some(a), Some(b)) => {
                        report.data_type = Some(merge(a, b).ok_or_else(|| {
                            format!(
                                "mixed column types: {a:?} and {b:?}; export this column as text"
                            )
                        })?);
                    }
                }
                Ok(())
            });
            if let Err(message) = result {
                report.issue_count += 1;
                if report.issues.len() < 8 {
                    report.issues.push(Issue {
                        cell: format!("{}{}", column_name(col.index), row + 1),
                        message,
                    });
                }
            }
        }
        if report.data_type.is_none() {
            report.data_type = Some(ColumnType::String);
            report
                .warnings
                .push("No non-null values: exported as nullable string".into());
        }
        if matches!(
            report.data_type,
            Some(ColumnType::Time | ColumnType::Timestamp)
        ) {
            report
                .warnings
                .push("Microsecond precision; local time with no timezone assigned".into());
        }
        reports.push(report);
    }
    let ready = reports.iter().all(|c| c.issue_count == 0);
    Ok(Plan {
        sheet,
        report: Report {
            ready,
            row_count: rows.len(),
            columns: reports,
        },
        rows,
        columns,
    })
}

impl Plan<'_> {
    pub fn report(&self) -> &Report {
        &self.report
    }

    pub fn validation_error(&self) -> Option<String> {
        if self.report.ready {
            return None;
        }
        Some(
            self.report
                .columns
                .iter()
                .flat_map(|c| {
                    c.issues
                        .iter()
                        .map(move |i| format!("{} ({:?}): {}", i.cell, c.name, i.message))
                })
                .collect::<Vec<_>>()
                .join("\n"),
        )
    }

    /// Validate before opening/replacing the destination. Persist only after a
    /// complete footer is written; failures preserve an existing output file.
    pub fn write_path(&self, path: &Path) -> Result<(), String> {
        if let Some(error) = self.validation_error() {
            return Err(error);
        }
        let parent = path
            .parent()
            .filter(|p| !p.as_os_str().is_empty())
            .unwrap_or(Path::new("."));
        let mut temp = tempfile::NamedTempFile::new_in(parent).map_err(|e| e.to_string())?;
        self.write(temp.as_file_mut())?;
        temp.persist(path).map_err(|e| e.to_string())?;
        Ok(())
    }

    pub fn write<W: Write + Send>(&self, output: W) -> Result<(), String> {
        if let Some(error) = self.validation_error() {
            return Err(error);
        }
        let fields = self
            .report
            .columns
            .iter()
            .map(|c| {
                use ColumnType::*;
                let (physical, logical) = match c.data_type.unwrap() {
                    String => (PhysicalType::BYTE_ARRAY, Some(LogicalType::String)),
                    Int64 => (PhysicalType::INT64, None),
                    Double => (PhysicalType::DOUBLE, None),
                    Boolean => (PhysicalType::BOOLEAN, None),
                    Date => (PhysicalType::INT32, Some(LogicalType::Date)),
                    Time => (
                        PhysicalType::INT64,
                        Some(LogicalType::time(false, TimeUnit::MICROS)),
                    ),
                    Timestamp => (
                        PhysicalType::INT64,
                        Some(LogicalType::timestamp(false, TimeUnit::MICROS)),
                    ),
                };
                Type::primitive_type_builder(&c.name, physical)
                    .with_repetition(Repetition::OPTIONAL)
                    .with_logical_type(logical)
                    .build()
                    .map(Arc::new)
            })
            .collect::<Result<Vec<_>, _>>()
            .map_err(|e| e.to_string())?;
        let schema = Arc::new(
            Type::group_type_builder("sheet")
                .with_fields(fields)
                .build()
                .map_err(|e| e.to_string())?,
        );
        let props = Arc::new(
            WriterProperties::builder()
                .set_compression(Compression::SNAPPY)
                .build(),
        );
        let mut writer =
            SerializedFileWriter::new(output, schema, props).map_err(|e| e.to_string())?;
        for rows in self.rows.chunks(ROW_GROUP_ROWS) {
            let mut group = writer.next_row_group().map_err(|e| e.to_string())?;
            for (column, report) in self.columns.iter().zip(&self.report.columns) {
                let cells = rows
                    .iter()
                    .map(|&row| read_cell(self.sheet, row, column.index, column.as_text))
                    .collect::<Result<Vec<_>, _>>()?;
                let defs: Vec<i16> = cells
                    .iter()
                    .map(|v| if matches!(v, Cell::Null) { 0 } else { 1 })
                    .collect();
                let mut out = group
                    .next_column()
                    .map_err(|e| e.to_string())?
                    .ok_or("missing output column")?;
                macro_rules! write_values {
                    ($ty:ty, $pat:pat => $value:expr) => {{
                        let values = cells
                            .iter()
                            .filter_map(|c| match c {
                                $pat => Some($value),
                                Cell::Null => None,
                                _ => unreachable!("validated type"),
                            })
                            .collect::<Vec<_>>();
                        out.typed::<$ty>()
                            .write_batch(&values, Some(&defs), None)
                            .map_err(|e| e.to_string())?;
                    }};
                }
                match report.data_type.unwrap() {
                    ColumnType::String => {
                        write_values!(ByteArrayType, Cell::Text(s) => ByteArray::from(s.as_str()))
                    }
                    ColumnType::Int64 => write_values!(Int64Type, Cell::Int(n) => *n),
                    ColumnType::Boolean => write_values!(BoolType, Cell::Bool(b) => *b),
                    ColumnType::Date => write_values!(Int32Type, Cell::Date(d) => *d),
                    ColumnType::Time => write_values!(Int64Type, Cell::Time(t) => *t),
                    ColumnType::Double => {
                        let values: Vec<f64> = cells
                            .iter()
                            .filter_map(|c| match c {
                                Cell::Int(n) => Some(*n as f64),
                                Cell::Float(n) => Some(*n),
                                Cell::Null => None,
                                _ => unreachable!(),
                            })
                            .collect();
                        out.typed::<DoubleType>()
                            .write_batch(&values, Some(&defs), None)
                            .map_err(|e| e.to_string())?;
                    }
                    ColumnType::Timestamp => {
                        let values: Vec<i64> = cells
                            .iter()
                            .filter_map(|c| match c {
                                Cell::Date(d) => Some(*d as i64 * DAY_MICROS),
                                Cell::Timestamp(t) => Some(*t),
                                Cell::Null => None,
                                _ => unreachable!(),
                            })
                            .collect();
                        out.typed::<Int64Type>()
                            .write_batch(&values, Some(&defs), None)
                            .map_err(|e| e.to_string())?;
                    }
                }
                out.close().map_err(|e| e.to_string())?;
            }
            group.close().map_err(|e| e.to_string())?;
        }
        writer.close().map_err(|e| e.to_string())?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use parquet::{
        file::reader::{FileReader, SerializedFileReader},
        record::Field,
    };
    use visigrid_engine::{cell::DateStyle, sheet::SheetId};

    fn column(index: usize) -> Column {
        Column {
            index,
            name: column_name(index),
            as_text: false,
        }
    }
    fn sheet() -> Sheet {
        Sheet::new(SheetId(1), 10_000, 20)
    }
    fn fields(plan: &Plan<'_>) -> Vec<Vec<Field>> {
        // A path in a temp folder, not an open NamedTempFile: Windows
        // refuses to replace a file that is still open
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("out.parquet");
        plan.write_path(&path).unwrap();
        let reader = SerializedFileReader::new(std::fs::File::open(&path).unwrap()).unwrap();
        reader
            .get_row_iter(None)
            .unwrap()
            .map(|r| {
                r.unwrap()
                    .get_column_iter()
                    .map(|(_, v)| v.clone())
                    .collect()
            })
            .collect()
    }

    #[test]
    fn exact_values_nulls_empty_strings_and_formula_types() {
        let mut s = sheet();
        s.set_value(0, 0, "1.123456789012345");
        s.set_value(1, 0, "2");
        s.set_text(0, 1, "007");
        s.set_value(1, 1, "=\"\"");
        s.set_value(0, 2, "=1=1");
        s.set_value(1, 2, "=1=2");
        s.set_text(0, 3, "TRUE");
        let p = analyze(&s, vec![0, 1, 2], (0..4).map(column).collect()).unwrap();
        assert!(p.report.ready, "{:?}", p.report);
        let rows = fields(&p);
        assert_eq!(
            rows[0],
            vec![
                Field::Double(1.123456789012345),
                Field::Str("007".into()),
                Field::Bool(true),
                Field::Str("TRUE".into())
            ]
        );
        assert_eq!(rows[1][0], Field::Double(2.0));
        assert_eq!(rows[1][1], Field::Str("".into()));
        assert!(rows[2].iter().all(|c| matches!(c, Field::Null)));
    }

    #[test]
    fn mixed_types_fail_without_touching_destination_and_text_is_explicit() {
        let mut s = sheet();
        s.set_value(0, 0, "1.123456789");
        s.set_text(1, 0, "N/A");
        let p = analyze(&s, vec![0, 1], vec![column(0)]).unwrap();
        assert!(!p.report.ready);
        assert_eq!(p.report.columns[0].issues[0].cell, "A2");
        let file = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(file.path(), "keep me").unwrap();
        assert!(p.write_path(file.path()).is_err());
        assert_eq!(std::fs::read(file.path()).unwrap(), b"keep me");
        let mut c = column(0);
        c.as_text = true;
        let p = analyze(&s, vec![0, 1], vec![c]).unwrap();
        assert_eq!(fields(&p)[0][0], Field::Str("1.123456789".into()));
    }

    #[test]
    fn errors_and_uncalculated_formulas_cannot_be_hidden_by_text_override() {
        let mut s = sheet();
        s.set_value(0, 0, "=1/0");
        s.set_value_deferred(1, 0, "=1+1");
        let mut c = column(0);
        c.as_text = true;
        let p = analyze(&s, vec![0, 1], vec![c]).unwrap();
        assert_eq!(p.report.columns[0].issue_count, 2);
        assert!(p.validation_error().unwrap().contains("no computed result"));
    }

    #[test]
    fn temporal_schema_and_calendar_validation() {
        let mut s = sheet();
        s.set_value(0, 0, "25569");
        s.set_value(0, 1, "25569.5");
        s.set_value(0, 2, "0.5");
        s.set_number_format(
            0,
            0,
            NumberFormat::Date {
                style: DateStyle::Iso,
            },
        );
        s.set_number_format(0, 1, NumberFormat::DateTime);
        s.set_number_format(0, 2, NumberFormat::Time);
        let p = analyze(&s, vec![0], (0..3).map(column).collect()).unwrap();
        assert!(p.report.ready, "{:?}", p.report);
        let row = &fields(&p)[0];
        assert_eq!(row[0], Field::Date(0));
        assert_eq!(row[1], Field::TimestampMicros(DAY_MICROS / 2));
        assert_eq!(row[2], Field::TimeMicros(DAY_MICROS / 2));
        assert_eq!(date_days(1.0).unwrap(), -25567);
        assert_eq!(date_days(0.0).unwrap(), -25568);
        assert!(date_days(60.0).is_err());
        s.set_value(0, 0, "61.5");
        assert!(!analyze(&s, vec![0], vec![column(0)]).unwrap().report.ready);
    }

    #[test]
    fn scans_past_first_row_group_and_bounds_issue_samples() {
        let mut s = sheet();
        s.set_value(0, 0, "1");
        for row in 2048..2100 {
            s.set_text(row, 0, "x");
        }
        let p = analyze(&s, (0..2100).collect(), vec![column(0)]).unwrap();
        assert_eq!(p.report.columns[0].issue_count, 52);
        assert_eq!(p.report.columns[0].issues.len(), 8);
    }

    #[test]
    fn writes_multiple_groups_and_empty_schema_only_table() {
        let mut s = sheet();
        s.set_value(0, 0, "9007199254740992");
        s.set_text(0, 1, "9223372036854775807");
        let p = analyze(&s, (0..2050).collect(), vec![column(0), column(1)]).unwrap();
        let rows = fields(&p);
        assert_eq!(rows.len(), 2050);
        assert_eq!(rows[0][0], Field::Double(9007199254740992.0));
        assert_eq!(rows[0][1], Field::Str("9223372036854775807".into()));
        let p = analyze(&s, vec![], vec![column(0)]).unwrap();
        assert_eq!(p.report.columns[0].data_type, Some(ColumnType::String));
        assert!(fields(&p).is_empty());
    }

    #[test]
    fn custom_web_date_formats_are_typed_and_conditional_ambiguity_is_refused() {
        let mut s = sheet();
        s.set_value(0, 0, "25569");
        s.set_number_format(0, 0, NumberFormat::Custom("yyyy-mm-dd".into()));
        let p = analyze(&s, vec![0], vec![column(0)]).unwrap();
        assert_eq!(p.report.columns[0].data_type, Some(ColumnType::Date));
        assert_eq!(fields(&p)[0][0], Field::Date(0));
        s.set_number_format(0, 0, NumberFormat::Custom("0;yyyy-mm-dd".into()));
        assert!(!analyze(&s, vec![0], vec![column(0)]).unwrap().report.ready);
        s.set_number_format(0, 0, NumberFormat::Custom("[h]:mm:ss".into()));
        assert_eq!(
            analyze(&s, vec![0], vec![column(0)])
                .unwrap()
                .report
                .columns[0]
                .data_type,
            Some(ColumnType::Int64)
        );
    }

    #[test]
    fn duplicate_or_blank_headers_are_refused() {
        let s = sheet();
        assert!(analyze(&s, vec![], vec![column(0), column(0)]).is_err());
        let mut c = column(0);
        c.name.clear();
        assert!(analyze(&s, vec![], vec![c]).is_err());
    }
}
