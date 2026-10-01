//! Capture a resolved sheet view. No workbook references survive capture.
use crate::{AxisItem, LayoutInput, Merge, TextCell, MAX_CELL_POSITIONS};
use std::collections::{HashMap, HashSet};
use std::ops::Range;
use visigrid_engine::cell::{Alignment, CellFormat, CellStyle};
use visigrid_engine::sheet::Sheet;

/// Axes in display order, already excluding hidden and filtered positions.
#[derive(Clone, Debug)]
pub struct SheetView {
    pub rows: Vec<AxisItem>,
    pub columns: Vec<AxisItem>,
}

#[derive(Clone, Debug)]
pub struct Scope {
    /// Positions in SheetView, not source coordinates. None uses displayed text bounds.
    pub rows: Range<usize>,
    pub columns: Range<usize>,
}

#[derive(Clone, Debug)]
pub struct PrintCell {
    pub rows: Range<usize>,
    pub columns: Range<usize>,
    pub source: (usize, usize),
    pub text: String,
    pub format: CellFormat,
}

#[derive(Clone, Debug)]
pub struct Snapshot {
    pub name: String,
    pub layout: LayoutInput,
    pub cells: Vec<PrintCell>,
    pub default_font: String,
    pub default_size: f32,
}

/// Bounds ignore distant formatting, empty formula results, and hidden values.
/// Selected ranges include blank styled cells. Merges intersecting the scope
/// expand it; merges made discontinuous by sorting are rejected explicitly.
pub fn capture(
    sheet: &Sheet,
    view: &SheetView,
    selection: Option<Scope>,
    default_font: &str,
    default_size: f32,
) -> Result<Snapshot, String> {
    if !default_size.is_finite() || default_size <= 0.0 {
        return Err("Invalid default font size".into());
    }
    let rows: HashMap<_, _> = view
        .rows
        .iter()
        .enumerate()
        .map(|(i, a)| (a.source_index, i))
        .collect();
    let cols: HashMap<_, _> = view
        .columns
        .iter()
        .enumerate()
        .map(|(i, a)| (a.source_index, i))
        .collect();
    let mut scope = if let Some(s) = selection {
        s
    } else {
        let mut lo = (usize::MAX, usize::MAX);
        let mut hi = (0, 0);
        for (r, c) in sheet
            .cells_iter()
            .map(|(p, _)| p)
            .chain(sheet.spill_receiver_coords())
        {
            let (Some(&vr), Some(&vc)) = (rows.get(&r), cols.get(&c)) else {
                continue;
            };
            if sheet.is_merge_hidden(r, c) || sheet.get_formatted_display(r, c).is_empty() {
                continue;
            }
            lo = (lo.0.min(vr), lo.1.min(vc));
            hi = (hi.0.max(vr + 1), hi.1.max(vc + 1));
            // A Center Across Selection title spans the empty Center Across
            // cells to its right; print them so the title isn't cut off.
            if sheet.get_format(r, c).alignment == Alignment::CenterAcrossSelection {
                let mut next = c + 1;
                while next < sheet.cols
                    && sheet.get_format(r, next).alignment == Alignment::CenterAcrossSelection
                    && sheet.get_formatted_display(r, next).is_empty()
                    && !sheet.is_merge_hidden(r, next)
                {
                    if let Some(&vn) = cols.get(&next) {
                        hi.1 = hi.1.max(vn + 1);
                    }
                    next += 1;
                }
            }
        }
        if lo.0 == usize::MAX {
            return Err(
                "Nothing visible to export. Select a range to include blank formatted cells."
                    .into(),
            );
        }
        Scope {
            rows: lo.0..hi.0,
            columns: lo.1..hi.1,
        }
    };
    if scope.rows.is_empty()
        || scope.columns.is_empty()
        || scope.rows.end > view.rows.len()
        || scope.columns.end > view.columns.len()
    {
        return Err("The selected range has no visible cells.".into());
    }
    // Discover merge geometry from its visible members, including when the
    // origin is hidden. Never promote hidden origin text to a visible cell.
    let mut merges = Vec::new();
    for m in &sheet.merged_regions {
        let rr: Vec<_> = (m.start.0..=m.end.0)
            .filter_map(|r| rows.get(&r).copied())
            .collect();
        let cc: Vec<_> = (m.start.1..=m.end.1)
            .filter_map(|c| cols.get(&c).copied())
            .collect();
        if rr.is_empty() || cc.is_empty() {
            continue;
        }
        let mr = *rr.iter().min().unwrap()..rr.iter().max().unwrap() + 1;
        let mc = *cc.iter().min().unwrap()..cc.iter().max().unwrap() + 1;
        let ordered =
            rr.windows(2).all(|w| w[1] == w[0] + 1) && cc.windows(2).all(|w| w[1] == w[0] + 1);
        merges.push((mr, mc, (m.start.0, m.start.1), ordered));
    }
    loop {
        let old = (scope.rows.clone(), scope.columns.clone());
        for (r, c, _, ordered) in &merges {
            if intersects(r, &scope.rows) && intersects(c, &scope.columns) {
                if !ordered {
                    return Err("A merged cell is split by the current sort. Clear the sort before exporting this range.".into());
                }
                scope.rows = scope.rows.start.min(r.start)..scope.rows.end.max(r.end);
                scope.columns = scope.columns.start.min(c.start)..scope.columns.end.max(c.end);
            }
        }
        if old == (scope.rows.clone(), scope.columns.clone()) {
            break;
        }
    }
    if scope
        .rows
        .len()
        .checked_mul(scope.columns.len())
        .is_none_or(|n| n > MAX_CELL_POSITIONS)
    {
        return Err("Export exceeds one million visible cells. Select a smaller range.".into());
    }
    let mut layout = LayoutInput {
        rows: view.rows[scope.rows.clone()].to_vec(),
        columns: view.columns[scope.columns.clone()].to_vec(),
        ..Default::default()
    };
    let mut covered = HashSet::new();
    let mut origins = HashMap::new();
    for (r, c, origin, _) in merges {
        if !intersects(&r, &scope.rows) || !intersects(&c, &scope.columns) {
            continue;
        }
        let merge = Merge {
            rows: r.start - scope.rows.start..r.end - scope.rows.start,
            columns: c.start - scope.columns.start..c.end - scope.columns.start,
        };
        for rr in merge.rows.clone() {
            for cc in merge.columns.clone() {
                covered.insert((rr, cc));
            }
        }
        covered.remove(&(merge.rows.start, merge.columns.start));
        origins.insert(
            (merge.rows.start, merge.columns.start),
            (merge.clone(), origin),
        );
        layout.merges.push(merge);
    }
    let mut cells = Vec::new();
    for (r, row) in layout.rows.iter().enumerate() {
        for (c, col) in layout.columns.iter().enumerate() {
            if covered.contains(&(r, c)) {
                continue;
            }
            let source = (row.source_index, col.source_index);
            let mut format = sheet.get_effective_format(source.0, source.1);
            let mut text = sheet.get_formatted_display(source.0, source.1);
            let (rr, cc) = if let Some((merge, origin)) = origins.get(&(r, c)) {
                if source != *origin {
                    text.clear();
                } else if let Some(m) = sheet.get_merge(source.0, source.1) {
                    let (t, ri, b, l) = sheet.resolve_merge_borders(m);
                    format.border_top = t;
                    format.border_right = ri;
                    format.border_bottom = b;
                    format.border_left = l;
                }
                (merge.rows.clone(), merge.columns.clone())
            } else {
                (r..r + 1, c..c + 1)
            };
            resolve_semantic_style(&mut format);
            if format.alignment == Alignment::General {
                format.alignment = if matches!(
                    sheet.get_computed_value(source.0, source.1),
                    visigrid_engine::formula::eval::Value::Number(_)
                ) {
                    Alignment::Right
                } else {
                    Alignment::Left
                };
            }
            if !text.is_empty() {
                layout.text.push(TextCell {
                    row: r,
                    column: c,
                    font_size_pt: f64::from(format.font_size.unwrap_or(default_size)),
                });
            }
            cells.push(PrintCell {
                rows: rr,
                columns: cc,
                source,
                text,
                format,
            });
        }
    }
    Ok(Snapshot {
        name: sheet.name.clone(),
        layout,
        cells,
        default_font: default_font.into(),
        default_size,
    })
}

fn intersects(a: &Range<usize>, b: &Range<usize>) -> bool {
    a.start < b.end && b.start < a.end
}

// Print uses a light paper palette, independent of the desktop theme.
fn resolve_semantic_style(f: &mut CellFormat) {
    let (bg, fg) = match f.cell_style {
        CellStyle::None => return,
        CellStyle::Error => ([254, 226, 226, 255], [153, 27, 27, 255]),
        CellStyle::Warning => ([254, 243, 199, 255], [146, 64, 14, 255]),
        CellStyle::Success => ([220, 252, 231, 255], [22, 101, 52, 255]),
        CellStyle::Input => ([219, 234, 254, 255], [30, 64, 175, 255]),
        CellStyle::Total => {
            f.bold = true;
            ([241, 245, 249, 255], [15, 23, 42, 255])
        }
        CellStyle::Note => {
            f.italic = true;
            ([248, 250, 252, 255], [71, 85, 105, 255])
        }
    };
    f.background_color.get_or_insert(bg);
    f.font_color.get_or_insert(fg);
}
