//! Saved Table sort intent. Physical export ordering and reference checks live
//! in xlsx_sorted_export; this module only reads/writes the saved criterion.
use super::{
    xlsx_table_filters::boolean,
    xlsx_tables::{attr, range},
};
use quick_xml::{
    events::{BytesEnd, BytesStart, Event},
    Reader, Writer,
};
use visigrid_engine::{
    filter::SortDirection,
    sheet::Sheet,
    table::{DataTable, TableRange},
    table_view::TableSort,
};

fn body(table: &DataTable) -> Result<TableRange, String> {
    if table.range.data_rows() == 0 {
        return Err("Saved sorting requires Table body rows".into());
    }
    Ok(TableRange {
        start_row: table.range.start_row + 1,
        ..table.range
    })
}

pub(super) fn parse(xml: &str, table: &DataTable) -> Result<Option<TableSort>, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().expand_empty_elements = true;
    let mut path = Vec::<String>::new();
    let mut state = None;
    let mut sort = None;
    let mut sort_depth = None;
    loop {
        match reader.read_event().map_err(|e| e.to_string())? {
            Event::Start(e) => {
                let name = String::from_utf8_lossy(e.local_name().as_ref()).into_owned();
                let parent_is_table = path.len() == 1 && path[0] == "table";
                let parent_is_auto =
                    path.len() == 2 && path[0] == "table" && path[1] == "autoFilter";
                if name == "sortState" && (parent_is_table || parent_is_auto) {
                    if state.is_some() {
                        return Err("Multiple saved sort states are unsupported".into());
                    }
                    if boolean(&e, b"columnSort", false)? {
                        return Err("Left-to-right sorting is unsupported".into());
                    }
                    if boolean(&e, b"caseSensitive", false)? {
                        return Err("Case-sensitive sorting is unsupported".into());
                    }
                    if attr(&e, b"sortMethod")?
                        .as_deref()
                        .is_some_and(|s| s != "none")
                    {
                        return Err("Language-specific sort methods are unsupported".into());
                    }
                    let r = range(&attr(&e, b"ref")?.ok_or("Missing sort range")?)?;
                    if r != body(table)? && r != table.range {
                        return Err("Saved sort must cover the complete Table body".into());
                    }
                    state = Some(r);
                    sort_depth = Some(path.len() + 1);
                } else if sort_depth.is_some() {
                    if name != "sortCondition" || sort_depth != Some(path.len()) {
                        return Err(format!("Unsupported saved sort element: {name}"));
                    }
                    if sort.is_some() {
                        return Err("Multi-column sorting is unsupported".into());
                    }
                    if attr(&e, b"sortBy")?
                        .as_deref()
                        .is_some_and(|s| s != "value")
                    {
                        return Err("Color/icon sorting is unsupported".into());
                    }
                    for key in [b"customList".as_slice(), b"dxfId", b"iconSet", b"iconId"] {
                        if attr(&e, key)?.is_some() {
                            return Err("Custom-list or style-based sorting is unsupported".into());
                        }
                    }
                    let r = range(&attr(&e, b"ref")?.ok_or("Missing sort column range")?)?;
                    let body = body(table)?;
                    let state = state.unwrap();
                    if r.start_col != r.end_col
                        || r.start_col < body.start_col
                        || r.end_col > body.end_col
                        || r.end_row != body.end_row
                        || (r.start_row != body.start_row && r.start_row != state.start_row)
                    {
                        return Err("Saved sort column must cover one complete Table field".into());
                    }
                    sort = Some(TableSort {
                        column: table.columns[r.start_col - body.start_col].id,
                        direction: if boolean(&e, b"descending", false)? {
                            SortDirection::Descending
                        } else {
                            SortDirection::Ascending
                        },
                    });
                }
                path.push(name);
            }
            Event::End(e) => {
                if e.local_name().as_ref() == b"sortState" && sort_depth == Some(path.len()) {
                    sort_depth = None;
                }
                path.pop();
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if state.is_some() && sort.is_none() {
        return Err("Saved sort has no sort condition".into());
    }
    Ok(sort)
}

fn reference(r: TableRange) -> String {
    format!(
        "{}{}:{}{}",
        super::xlsx::col_to_letter(r.start_col),
        r.start_row + 1,
        super::xlsx::col_to_letter(r.end_col),
        r.end_row + 1
    )
}

pub(super) fn write(
    sheet: &Sheet,
    table: &DataTable,
    writer: &mut Writer<Vec<u8>>,
) -> Result<(), String> {
    let Some(sort) = sheet
        .table_view_spec()
        .filter(|s| s.table == table.id)
        .and_then(|s| s.sort.as_ref())
    else {
        return Ok(());
    };
    let body = body(table)?;
    let offset = table
        .columns
        .iter()
        .position(|c| c.id == sort.column)
        .ok_or("Missing saved sort column")?;
    let all = reference(body);
    let col = body.start_col + offset;
    let column = reference(TableRange {
        start_col: col,
        end_col: col,
        ..body
    });
    let mut state = BytesStart::new("sortState");
    state.push_attribute(("ref", all.as_str()));
    writer
        .write_event(Event::Start(state))
        .map_err(|e| e.to_string())?;
    let mut condition = BytesStart::new("sortCondition");
    condition.push_attribute(("ref", column.as_str()));
    condition.push_attribute((
        "descending",
        if sort.direction == SortDirection::Descending {
            "1"
        } else {
            "0"
        },
    ));
    writer
        .write_event(Event::Empty(condition))
        .map_err(|e| e.to_string())?;
    writer
        .write_event(Event::End(BytesEnd::new("sortState")))
        .map_err(|e| e.to_string())
}
