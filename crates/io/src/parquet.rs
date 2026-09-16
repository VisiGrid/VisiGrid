// Apache Parquet import (read-only)
//
// A Parquet file is one table: column names in row 1, one record per row
// below. Values keep the type the file declared for them. Nothing is inferred
// from text, the way CSV import has to — a string column holding "007" or
// "=1+1" stays exactly those characters.

use std::fs::File;
use std::panic::{catch_unwind, AssertUnwindSafe};
use std::path::Path;

use parquet::basic::{LogicalType, TimeUnit};
use parquet::file::reader::{FileReader, SerializedFileReader};
use parquet::record::Field;
use parquet::schema::types::Type;
use visigrid_engine::cell::{DateStyle, NumberFormat};
use visigrid_engine::sheet::{Sheet, SheetId};

/// Sheet capacity, header row included. Same limits as xlsx import.
pub const MAX_ROWS: usize = visigrid_engine::sheet::NUM_ROWS;
pub const MAX_COLS: usize = visigrid_engine::sheet::NUM_COLS;

/// Days from the spreadsheet epoch (serial 0, 1899-12-30) to 1970-01-01.
const UNIX_EPOCH_SERIAL: f64 = 25569.0;
/// First serial past the fictional 1900-02-29. Earlier serials are off by a
/// day in the 1900 date system, so dates before it are stored as text.
const FIRST_SAFE_SERIAL: f64 = 61.0;
/// Largest integer an f64 holds exactly. Beyond it an ID column would round.
const MAX_EXACT_INT: i128 = 1 << 53;

const MICROS_PER_DAY: f64 = 86_400_000_000.0;

pub struct ParquetImport {
    pub sheet: Sheet,
    /// Records in the file.
    pub total_rows: u64,
    /// Records placed on the sheet (below the header row).
    pub rows_loaded: usize,
    /// Top-level columns in the file.
    pub total_cols: usize,
    /// Columns placed on the sheet.
    pub cols_loaded: usize,
}

impl ParquetImport {
    pub fn truncated(&self) -> bool {
        (self.rows_loaded as u64) < self.total_rows || self.cols_loaded < self.total_cols
    }

    /// One line for a status bar or CLI error, or None when nothing was cut.
    pub fn truncation_message(&self) -> Option<String> {
        if !self.truncated() {
            return None;
        }
        let mut parts = Vec::new();
        if (self.rows_loaded as u64) < self.total_rows {
            parts.push(format!("first {} of {} rows", self.rows_loaded, self.total_rows));
        }
        if self.cols_loaded < self.total_cols {
            parts.push(format!("first {} of {} columns", self.cols_loaded, self.total_cols));
        }
        Some(format!(
            "Loaded {} (sheet limit is {} rows including the header, {} columns)",
            parts.join(" and "),
            MAX_ROWS,
            MAX_COLS
        ))
    }
}

/// How a column's raw values need reinterpreting.
///
/// The record reader only knows the legacy converted types. Nanosecond times
/// and timestamps have no converted type, so it hands them over as bare
/// integers — and pyarrow writes nanoseconds for nanosecond data, which is
/// what pandas held by default before 3.0.
/// UUIDs arrive as their 16 raw bytes.
#[derive(Clone, Copy)]
enum ColumnKind {
    Plain,
    TimestampNanos,
    TimeNanos,
    Uuid,
}

fn column_kind(ty: &Type) -> ColumnKind {
    if !ty.is_primitive() {
        return ColumnKind::Plain;
    }
    match ty.get_basic_info().logical_type_ref() {
        Some(LogicalType::Timestamp(t)) if t.unit == TimeUnit::NANOS => ColumnKind::TimestampNanos,
        Some(LogicalType::Time(t)) if t.unit == TimeUnit::NANOS => ColumnKind::TimeNanos,
        Some(LogicalType::Uuid) => ColumnKind::Uuid,
        _ => ColumnKind::Plain,
    }
}

pub fn import(path: &Path) -> Result<ParquetImport, String> {
    // The record reader panics (unimplemented!) on the few legacy column types
    // it has no conversion for, INTERVAL among them. Opening a file must not
    // take the app down with it.
    match catch_unwind(AssertUnwindSafe(|| import_inner(path))) {
        Ok(result) => result,
        Err(payload) => {
            let detail = payload
                .downcast_ref::<String>()
                .map(String::as_str)
                .or_else(|| payload.downcast_ref::<&str>().copied())
                .unwrap_or("unsupported column type");
            Err(format!("This Parquet file uses a column type VisiGrid can't read yet: {}", detail))
        }
    }
}

fn import_inner(path: &Path) -> Result<ParquetImport, String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let reader = SerializedFileReader::new(file)
        .map_err(|e| format!("Not a readable Parquet file: {}", e))?;

    let metadata = reader.metadata().file_metadata();
    let total_rows = metadata.num_rows().max(0) as u64;
    let schema = metadata.schema();
    let fields = schema.get_fields();
    let total_cols = fields.len();
    let cols_loaded = total_cols.min(MAX_COLS);

    // Past the column limit, project the schema down so the columns that
    // won't be shown are never decoded.
    let projection = if cols_loaded < total_cols {
        Some(
            Type::group_type_builder(schema.name())
                .with_fields(fields[..cols_loaded].to_vec())
                .build()
                .map_err(|e| e.to_string())?,
        )
    } else {
        None
    };

    let kinds: Vec<ColumnKind> = fields[..cols_loaded].iter().map(|f| column_kind(f)).collect();

    let mut sheet = Sheet::new(SheetId(1), MAX_ROWS, MAX_COLS);
    for (col, field) in fields[..cols_loaded].iter().enumerate() {
        sheet.set_text(0, col, field.name());
    }

    let mut rows_loaded = 0usize;
    let rows = reader
        .get_row_iter(projection)
        .map_err(|e| format!("Failed to read Parquet rows: {}", e))?;
    for record in rows.take(MAX_ROWS - 1) {
        let record = record.map_err(|e| format!("Failed to read Parquet row {}: {}", rows_loaded + 1, e))?;
        let row = rows_loaded + 1;
        for (col, (_, value)) in record.get_column_iter().enumerate().take(cols_loaded) {
            write_field(&mut sheet, row, col, value, kinds[col]);
        }
        rows_loaded += 1;
    }

    sheet.rows = (rows_loaded + 1).max(1000);
    sheet.cols = cols_loaded.max(26);

    Ok(ParquetImport { sheet, total_rows, rows_loaded, total_cols, cols_loaded })
}

fn write_field(sheet: &mut Sheet, row: usize, col: usize, value: &Field, kind: ColumnKind) {
    match value {
        Field::Null => {}
        Field::Bool(b) => sheet.set_value_deferred(row, col, if *b { "TRUE" } else { "FALSE" }),

        Field::Byte(n) => set_number(sheet, row, col, *n as f64),
        Field::Short(n) => set_number(sheet, row, col, *n as f64),
        Field::Int(n) => set_number(sheet, row, col, *n as f64),
        Field::UByte(n) => set_number(sheet, row, col, *n as f64),
        Field::UShort(n) => set_number(sheet, row, col, *n as f64),
        Field::UInt(n) => set_number(sheet, row, col, *n as f64),
        Field::Long(n) => match kind {
            ColumnKind::TimestampNanos => {
                let micros = n.div_euclid(1000);
                set_timestamp_micros(sheet, row, col, micros);
            }
            ColumnKind::TimeNanos => set_time(sheet, row, col, (*n / 1000) as f64 / MICROS_PER_DAY),
            ColumnKind::Plain | ColumnKind::Uuid => set_integer(sheet, row, col, *n as i128),
        },
        Field::ULong(n) => set_integer(sheet, row, col, *n as i128),

        // Through the shortest decimal that round-trips the f32, so 0.1 lands
        // as 0.1 and not 0.10000000149011612.
        Field::Float16(n) => set_float(sheet, row, col, f32::from(*n).to_string().parse().unwrap_or(f64::NAN)),
        Field::Float(n) => set_float(sheet, row, col, n.to_string().parse().unwrap_or(f64::NAN)),
        Field::Double(n) => set_float(sheet, row, col, *n),

        Field::Decimal(_) => {
            // Display renders the exact decimal ("1234.50"). More significant
            // digits than a double keeps would silently change the amount, so
            // those stay text.
            let text = value.to_string();
            let digits = text.chars().filter(|c| c.is_ascii_digit()).collect::<String>();
            if digits.trim_start_matches('0').len() <= 15 {
                sheet.set_value_deferred(row, col, &text);
            } else {
                sheet.set_text(row, col, &text);
            }
        }

        Field::Str(s) => sheet.set_text(row, col, s),
        Field::Bytes(b) => {
            let bytes = b.data();
            if let (ColumnKind::Uuid, Ok(id)) = (kind, <[u8; 16]>::try_from(bytes)) {
                sheet.set_text(row, col, &format_uuid(&id));
            } else {
                // Older writers leave strings unannotated, so readable text is
                // shown as text. Anything else is binary and shown as hex
                // rather than as a run of control characters.
                match std::str::from_utf8(bytes) {
                    Ok(s) if !s.chars().any(|c| c.is_control() && !matches!(c, '\t' | '\n' | '\r')) => {
                        sheet.set_text(row, col, s)
                    }
                    _ => sheet.set_text(row, col, &format!("0x{}", hex(bytes))),
                }
            }
        }

        Field::Date(days) => {
            let serial = *days as f64 + UNIX_EPOCH_SERIAL;
            if serial >= FIRST_SAFE_SERIAL {
                set_number(sheet, row, col, serial);
                set_number_format(sheet, row, col, NumberFormat::Date { style: DateStyle::Short });
            } else {
                sheet.set_text(row, col, &value.to_string());
            }
        }
        Field::TimeMillis(ms) => set_time(sheet, row, col, *ms as f64 * 1000.0 / MICROS_PER_DAY),
        Field::TimeMicros(us) => set_time(sheet, row, col, *us as f64 / MICROS_PER_DAY),
        Field::TimestampMillis(ms) => set_timestamp_micros(sheet, row, col, ms.saturating_mul(1000)),
        Field::TimestampMicros(us) => set_timestamp_micros(sheet, row, col, *us),

        // Structs, lists and maps don't fit in a cell. JSON keeps them
        // readable and parseable.
        Field::Group(_) | Field::ListInternal(_) | Field::MapInternal(_) => {
            let json = serde_json::to_string(&value.to_json_value()).unwrap_or_else(|_| value.to_string());
            sheet.set_text(row, col, &json);
        }
    }
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| format!("{:02x}", byte)).collect()
}

fn format_uuid(id: &[u8; 16]) -> String {
    format!("{}-{}-{}-{}-{}", hex(&id[..4]), hex(&id[4..6]), hex(&id[6..8]), hex(&id[8..10]), hex(&id[10..]))
}

fn set_number(sheet: &mut Sheet, row: usize, col: usize, n: f64) {
    // Debug formatting is the shortest round-trip form and switches to an
    // exponent for very large or small values, which from_input parses.
    sheet.set_value_deferred(row, col, &format!("{:?}", n));
}

fn set_integer(sheet: &mut Sheet, row: usize, col: usize, n: i128) {
    if n.abs() <= MAX_EXACT_INT {
        set_number(sheet, row, col, n as f64);
    } else {
        sheet.set_text(row, col, &n.to_string());
    }
}

fn set_float(sheet: &mut Sheet, row: usize, col: usize, n: f64) {
    if n.is_finite() {
        set_number(sheet, row, col, n);
    } else {
        sheet.set_text(row, col, &n.to_string());
    }
}

fn set_time(sheet: &mut Sheet, row: usize, col: usize, fraction_of_day: f64) {
    set_number(sheet, row, col, fraction_of_day);
    set_number_format(sheet, row, col, NumberFormat::Time);
}

/// Timestamps are placed as UTC wall-clock time. A spreadsheet serial has no
/// zone, and converting to the viewer's zone would make the same file show
/// different values on different machines.
fn set_timestamp_micros(sheet: &mut Sheet, row: usize, col: usize, micros: i64) {
    let serial = micros as f64 / MICROS_PER_DAY + UNIX_EPOCH_SERIAL;
    if serial >= FIRST_SAFE_SERIAL {
        set_number(sheet, row, col, serial);
        set_number_format(sheet, row, col, NumberFormat::DateTime);
    } else {
        let text = chrono::DateTime::from_timestamp_micros(micros)
            .map(|dt| dt.format("%Y-%m-%d %H:%M:%S%.f").to_string())
            .unwrap_or_else(|| micros.to_string());
        sheet.set_text(row, col, &text);
    }
}

fn set_number_format(sheet: &mut Sheet, row: usize, col: usize, number_format: NumberFormat) {
    let mut format = sheet.get_format(row, col);
    format.number_format = number_format;
    sheet.set_format(row, col, format);
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;

    use parquet::data_type::{BoolType, ByteArray, ByteArrayType, DoubleType, FixedLenByteArray, FixedLenByteArrayType, Int32Type, Int64Type};
    use parquet::file::properties::WriterProperties;
    use parquet::file::writer::SerializedFileWriter;
    use parquet::schema::parser::parse_message_type;
    use visigrid_engine::cell::CellValue;

    /// One column's values for a single row group. `None` is a null.
    enum Col {
        Bool(Vec<Option<bool>>),
        I32(Vec<Option<i32>>),
        I64(Vec<Option<i64>>),
        F64(Vec<Option<f64>>),
        Bytes(Vec<Option<&'static [u8]>>),
        Fixed(Vec<Option<Vec<u8>>>),
    }

    fn split<T: Clone>(values: &[Option<T>]) -> (Vec<T>, Vec<i16>) {
        let defs = values.iter().map(|v| v.is_some() as i16).collect();
        (values.iter().flatten().cloned().collect(), defs)
    }

    /// Write a flat Parquet file whose columns are all OPTIONAL.
    fn write(schema: &str, cols: Vec<Col>) -> tempfile::NamedTempFile {
        let file = tempfile::Builder::new().suffix(".parquet").tempfile().unwrap();
        let schema = Arc::new(parse_message_type(schema).unwrap());
        let mut writer = SerializedFileWriter::new(file.reopen().unwrap(), schema, Arc::new(WriterProperties::builder().build())).unwrap();
        let mut group = writer.next_row_group().unwrap();
        for col in cols {
            let mut w = group.next_column().unwrap().unwrap();
            match col {
                Col::Bool(v) => { let (vals, defs) = split(&v); w.typed::<BoolType>().write_batch(&vals, Some(&defs), None).unwrap(); }
                Col::I32(v) => { let (vals, defs) = split(&v); w.typed::<Int32Type>().write_batch(&vals, Some(&defs), None).unwrap(); }
                Col::I64(v) => { let (vals, defs) = split(&v); w.typed::<Int64Type>().write_batch(&vals, Some(&defs), None).unwrap(); }
                Col::F64(v) => { let (vals, defs) = split(&v); w.typed::<DoubleType>().write_batch(&vals, Some(&defs), None).unwrap(); }
                Col::Bytes(v) => {
                    let (vals, defs) = split(&v);
                    let vals: Vec<ByteArray> = vals.into_iter().map(ByteArray::from).collect();
                    w.typed::<ByteArrayType>().write_batch(&vals, Some(&defs), None).unwrap();
                }
                Col::Fixed(v) => {
                    let (vals, defs) = split(&v);
                    let vals: Vec<FixedLenByteArray> = vals.into_iter().map(|b| ByteArray::from(b).into()).collect();
                    w.typed::<FixedLenByteArrayType>().write_batch(&vals, Some(&defs), None).unwrap();
                }
            }
            w.close().unwrap();
        }
        group.close().unwrap();
        writer.close().unwrap();
        file
    }

    fn number(sheet: &Sheet, row: usize, col: usize) -> f64 {
        match sheet.get_cell(row, col).value {
            CellValue::Number(n) => n,
            other => panic!("({row},{col}) is {other:?}, expected a number"),
        }
    }

    fn text(sheet: &Sheet, row: usize, col: usize) -> String {
        match sheet.get_cell(row, col).value {
            CellValue::Text(s) => s,
            other => panic!("({row},{col}) is {other:?}, expected text"),
        }
    }

    #[test]
    fn column_names_become_the_header_row_and_values_keep_their_types() {
        let file = write(
            "message m {
                OPTIONAL BYTE_ARRAY zip (UTF8);
                OPTIONAL INT32 qty;
                OPTIONAL DOUBLE price;
                OPTIONAL BOOLEAN paid;
            }",
            vec![
                Col::Bytes(vec![Some(b"007"), Some(b"=1+1")]),
                Col::I32(vec![Some(3), None]),
                Col::F64(vec![Some(0.1), Some(-2.5)]),
                Col::Bool(vec![Some(true), Some(false)]),
            ],
        );
        let result = import(file.path()).unwrap();
        let sheet = &result.sheet;

        assert_eq!((result.rows_loaded, result.total_rows, result.cols_loaded), (2, 2, 4));
        assert!(!result.truncated());
        assert_eq!(
            (0..4).map(|c| text(sheet, 0, c)).collect::<Vec<_>>(),
            ["zip", "qty", "price", "paid"]
        );

        // A declared string is never reinterpreted: no lost zeros, no formula.
        assert_eq!(text(sheet, 1, 0), "007");
        assert_eq!(text(sheet, 2, 0), "=1+1");

        assert_eq!(number(sheet, 1, 1), 3.0);
        assert!(matches!(sheet.get_cell(2, 1).value, CellValue::Empty), "null stays empty");
        assert_eq!(number(sheet, 1, 2), 0.1);
        assert_eq!(number(sheet, 2, 2), -2.5);
        assert_eq!(sheet.get_display(1, 3), "TRUE");
    }

    #[test]
    fn dates_and_timestamps_become_serials_in_every_unit() {
        // 2024-03-15 = 19797 days after the Unix epoch = serial 45366.
        // 2024-03-15T12:00:00Z is half a day on top of that.
        let noon_micros: i64 = 19797 * 86_400_000_000 + 43_200_000_000;
        let file = write(
            "message m {
                OPTIONAL INT32 d (DATE);
                OPTIONAL INT64 ms (TIMESTAMP(MILLIS,true));
                OPTIONAL INT64 us (TIMESTAMP(MICROS,true));
                OPTIONAL INT64 ns (TIMESTAMP(NANOS,false));
                OPTIONAL INT64 t (TIME(NANOS,false));
            }",
            vec![
                Col::I32(vec![Some(19797)]),
                Col::I64(vec![Some(noon_micros / 1000)]),
                Col::I64(vec![Some(noon_micros)]),
                Col::I64(vec![Some(noon_micros * 1000)]),
                Col::I64(vec![Some(6 * 3_600_000_000_000)]),
            ],
        );
        let sheet = import(file.path()).unwrap().sheet;

        assert_eq!(number(&sheet, 1, 0), 45366.0);
        assert!(matches!(sheet.get_format(1, 0).number_format, NumberFormat::Date { .. }));
        for col in 1..=3 {
            assert_eq!(number(&sheet, 1, col), 45366.5, "column {col}");
            assert!(matches!(sheet.get_format(1, col).number_format, NumberFormat::DateTime));
        }
        // What the grid shows.
        assert_eq!(sheet.get_formatted_display(1, 0), "3/15/2024");
        assert_eq!(sheet.get_formatted_display(1, 1), "3/15/2024 12:00:00");
        assert_eq!(sheet.get_formatted_display(1, 4), "06:00:00");
        assert_eq!(number(&sheet, 1, 4), 0.25);
        assert!(matches!(sheet.get_format(1, 4).number_format, NumberFormat::Time));
    }

    #[test]
    fn values_a_double_would_change_are_kept_as_text() {
        let file = write(
            "message m {
                OPTIONAL INT64 id;
                OPTIONAL INT64 small;
                OPTIONAL INT64 amount (DECIMAL(10,2));
                OPTIONAL BYTE_ARRAY wide (DECIMAL(30,2));
                OPTIONAL DOUBLE nan;
            }",
            vec![
                Col::I64(vec![Some(9_007_199_254_740_993)]),
                Col::I64(vec![Some(42)]),
                Col::I64(vec![Some(123_450)]),
                // 12345678901234567890123.45 as big-endian two's complement.
                Col::Bytes(vec![Some(&[0x01, 0x05, 0x6e, 0x0f, 0x36, 0xa6, 0x44, 0x3d, 0xe2, 0xdf, 0x79])]),
                Col::F64(vec![Some(f64::NAN)]),
            ],
        );
        let sheet = import(file.path()).unwrap().sheet;

        assert_eq!(text(&sheet, 1, 0), "9007199254740993");
        assert_eq!(number(&sheet, 1, 1), 42.0);
        assert_eq!(number(&sheet, 1, 2), 1234.5);
        assert_eq!(text(&sheet, 1, 3), "12345678901234567890123.45");
        assert_eq!(text(&sheet, 1, 4), "NaN");
    }

    #[test]
    fn rows_past_the_sheet_limit_are_reported_not_silently_dropped() {
        let n = MAX_ROWS + 10;
        let file = write(
            "message m { OPTIONAL INT32 n; }",
            vec![Col::I32((0..n as i32).map(Some).collect())],
        );
        let result = import(file.path()).unwrap();

        assert_eq!(result.total_rows, n as u64);
        assert_eq!(result.rows_loaded, MAX_ROWS - 1);
        assert!(result.truncated());
        assert_eq!(number(&result.sheet, MAX_ROWS - 1, 0), (MAX_ROWS - 2) as f64);
        let message = result.truncation_message().unwrap();
        assert!(message.contains(&format!("first {} of {} rows", MAX_ROWS - 1, n)), "{message}");
    }

    #[test]
    fn columns_past_the_sheet_limit_are_reported() {
        let count = MAX_COLS + 4;
        let schema = format!(
            "message m {{ {} }}",
            (0..count).map(|i| format!("OPTIONAL INT32 c{i};")).collect::<String>()
        );
        let file = write(&schema, (0..count as i32).map(|i| Col::I32(vec![Some(i)])).collect());
        let result = import(file.path()).unwrap();

        assert_eq!((result.cols_loaded, result.total_cols), (MAX_COLS, count));
        assert_eq!(text(&result.sheet, 0, MAX_COLS - 1), format!("c{}", MAX_COLS - 1));
        assert_eq!(number(&result.sheet, 1, MAX_COLS - 1), (MAX_COLS - 1) as f64);
        assert!(result.truncation_message().unwrap().contains(&format!("first {MAX_COLS} of {count} columns")));
    }

    #[test]
    fn nested_values_are_shown_as_json() {
        let file = write(
            // Fields in alphabetical order, so the JSON reads the same whether
            // or not serde_json's preserve_order is unified in by another crate.
            "message m { REQUIRED group point { OPTIONAL BYTE_ARRAY label (UTF8); OPTIONAL INT32 x; } }",
            vec![Col::Bytes(vec![Some(b"a")]), Col::I32(vec![Some(1)])],
        );
        let result = import(file.path()).unwrap();

        assert_eq!(result.cols_loaded, 1, "one top-level column, not one per leaf");
        assert_eq!(text(&result.sheet, 0, 0), "point");
        assert_eq!(text(&result.sheet, 1, 0), r#"{"label":"a","x":1}"#);
    }

    #[test]
    fn uuids_are_formatted_and_other_binary_is_hex() {
        let file = write(
            "message m {
                OPTIONAL FIXED_LEN_BYTE_ARRAY (16) id (UUID);
                OPTIONAL BYTE_ARRAY raw;
                OPTIONAL BYTE_ARRAY legacy_text;
            }",
            vec![
                Col::Fixed(vec![Some((0u8..16).collect())]),
                Col::Bytes(vec![Some(&[0x00, 0x07, 0xff])]),
                Col::Bytes(vec![Some(b"plain")]),
            ],
        );
        let sheet = import(file.path()).unwrap().sheet;

        assert_eq!(text(&sheet, 1, 0), "00010203-0405-0607-0809-0a0b0c0d0e0f");
        assert_eq!(text(&sheet, 1, 1), "0x0007ff");
        assert_eq!(text(&sheet, 1, 2), "plain");
    }

    #[test]
    fn a_column_type_the_reader_cannot_convert_is_an_error_not_a_crash() {
        let file = write(
            "message m { OPTIONAL FIXED_LEN_BYTE_ARRAY (12) span (INTERVAL); }",
            vec![Col::Fixed(vec![Some(vec![1u8; 12])])],
        );
        let err = import(file.path()).err().expect("INTERVAL columns are not readable yet");
        assert!(err.contains("can't read yet"), "{err}");
    }

    #[test]
    fn a_file_that_is_not_parquet_is_an_error() {
        let file = tempfile::Builder::new().suffix(".parquet").tempfile().unwrap();
        std::fs::write(file.path(), "a,b\n1,2\n").unwrap();
        let err = import(file.path()).err().unwrap();
        assert!(err.starts_with("Not a readable Parquet file"), "{err}");
    }
}
