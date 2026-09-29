//! Device-independent pagination of an already resolved, immutable sheet view.
//!
//! The caller captures calculation results, scope, visibility and display order.
//! This crate does not read the workbook, query a printer, or mutate sheet sizes.
//! All dimensions are physical points. Convert unzoomed grid units at the input
//! boundary with [`GRID_UNIT_PT`], never with screen DPI or current app zoom.

use std::collections::HashSet;
use std::ops::Range;

pub const GRID_UNIT_PT: f64 = 72.0 / 96.0;
pub const MAX_PAGES: usize = 1_000;
pub const MAX_CELL_POSITIONS: usize = 1_000_000;
pub const READABILITY_THRESHOLD_PT: f64 = 8.0;
// Physical tolerance for floating point sums, not a fraction of the page size.
const EPSILON_PT: f64 = 1e-7;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Paper {
    A4,
    Letter,
    Legal,
}

impl Paper {
    pub fn size(self) -> (f64, f64) {
        match self {
            Self::A4 => (210.0 * 72.0 / 25.4, 297.0 * 72.0 / 25.4),
            Self::Letter => (612.0, 792.0),
            Self::Legal => (612.0, 1008.0),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Margins {
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
    pub left: f64,
}

impl Default for Margins {
    fn default() -> Self {
        Self {
            top: 36.0,
            right: 36.0,
            bottom: 36.0,
            left: 36.0,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Scale {
    Actual,
    FitColumns,
    FitSheet,
    /// A factor, not a percentage: 0.6 means 60%.
    Custom(f64),
}

#[derive(Debug, Clone, PartialEq)]
pub struct PageSettings {
    pub paper: Paper,
    pub landscape: bool,
    pub margins: Margins,
    pub scale: Scale,
    /// Fixed footer reservation, in physical points (not scaled).
    pub footer: bool,
    /// Counts of leading *visible* rows/columns to repeat, not sheet coordinates.
    pub repeat_rows: usize,
    pub repeat_columns: usize,
}

impl Default for PageSettings {
    fn default() -> Self {
        Self {
            paper: Paper::A4,
            landscape: false,
            margins: Margins::default(),
            scale: Scale::FitColumns,
            footer: false,
            repeat_rows: 0,
            repeat_columns: 0,
        }
    }
}

/// One visible row/column. Source coordinates survive filtering and sorting.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AxisItem {
    pub source_index: usize,
    pub size_pt: f64,
}

/// Half-open positions in the visible axes, already resolved by the caller.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Merge {
    pub rows: Range<usize>,
    pub columns: Range<usize>,
}

/// One nonempty displayed cell's typography, used for readability diagnostics.
/// Covered merge cells and hidden values must not appear here.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TextCell {
    pub row: usize,
    pub column: usize,
    pub font_size_pt: f64,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct LayoutInput {
    pub rows: Vec<AxisItem>,
    pub columns: Vec<AxisItem>,
    pub merges: Vec<Merge>,
    pub text: Vec<TextCell>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Axis {
    Row,
    Column,
}

#[derive(Debug, Clone, PartialEq)]
pub enum LayoutError {
    NothingVisible,
    InvalidMargins,
    InvalidScale,
    InvalidAxis {
        axis: Axis,
        position: usize,
    },
    DuplicateSourceIndex {
        axis: Axis,
        source_index: usize,
    },
    InvalidMerge {
        index: usize,
    },
    OverlappingMerges {
        first: usize,
        second: usize,
    },
    InvalidTextCell {
        index: usize,
    },
    InvalidTitles {
        axis: Axis,
    },
    TitlesCutMerge {
        axis: Axis,
        merge: usize,
    },
    FitTooSmall {
        scale: f64,
    },
    /// Positions in the visible axis; caller can identify source row/column(s).
    CannotFit {
        axis: Axis,
        positions: Range<usize>,
    },
    TooManyCells,
    TooManyPages,
}

impl std::fmt::Display for LayoutError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NothingVisible => write!(f, "Nothing visible to print"),
            Self::InvalidMargins => write!(f, "Margins must leave a positive page body"),
            Self::InvalidScale => write!(f, "Scale must be between 10% and 400%"),
            Self::InvalidAxis { axis, position } => write!(f, "Invalid {axis:?} size at visible position {position}"),
            Self::DuplicateSourceIndex { axis, source_index } => write!(f, "Duplicate {axis:?} source index {source_index}"),
            Self::InvalidMerge { index } => write!(f, "Invalid visible merge {index}"),
            Self::OverlappingMerges { first, second } => write!(f, "Visible merges {first} and {second} overlap"),
            Self::InvalidTextCell { index } => write!(f, "Invalid visible text cell {index}"),
            Self::InvalidTitles { axis } => write!(f, "Repeated {axis:?} titles must leave space for body content"),
            Self::TitlesCutMerge { axis, merge } => write!(f, "Repeated {axis:?} titles cut merge {merge}"),
            Self::FitTooSmall { scale } => write!(f, "Fit requires {:.1}% scale; narrow the area or use another scale mode", scale * 100.0),
            Self::CannotFit { axis, positions } => write!(f, "Visible {axis:?} positions {positions:?} cannot fit; reduce scale or change paper/orientation"),
            Self::TooManyCells => write!(f, "Print area exceeds {MAX_CELL_POSITIONS} visible cell positions; narrow the area"),
            Self::TooManyPages => write!(f, "Print area exceeds {MAX_PAGES} pages; narrow the area"),
        }
    }
}

impl std::error::Error for LayoutError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScaleConstraint {
    None,
    Width,
    Height,
    Both,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Rect {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

/// A page references bands rather than copying all cells or axis metrics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Page {
    /// Non-title rows and columns. Leading title bands are implicit in the plan.
    pub rows: Range<usize>,
    pub columns: Range<usize>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Readability {
    pub smallest_text_pt: Option<f64>,
    /// Unique source cells, not multiplied by repeated headers or page count.
    pub cells_below_threshold: usize,
}

/// Immutable result; private metrics keep placement consistent with pagination.
#[derive(Debug, Clone, PartialEq)]
pub struct PagePlan {
    pages: Vec<Page>,
    rows: Vec<AxisItem>,
    columns: Vec<AxisItem>,
    row_offsets: Vec<f64>,
    column_offsets: Vec<f64>,
    repeat_rows: usize,
    repeat_columns: usize,
    page_size: (f64, f64),
    body: Rect,
    scale: f64,
    constraint: ScaleConstraint,
    readability: Readability,
}

impl PagePlan {
    pub fn pages(&self) -> &[Page] {
        &self.pages
    }
    pub fn page_size(&self) -> (f64, f64) {
        self.page_size
    }
    pub fn body(&self) -> Rect {
        self.body
    }
    pub fn scale(&self) -> f64 {
        self.scale
    }
    pub fn scale_constraint(&self) -> ScaleConstraint {
        self.constraint
    }
    pub fn readability(&self) -> &Readability {
        &self.readability
    }

    pub fn source_cell(&self, row: usize, column: usize) -> Option<(usize, usize)> {
        Some((
            self.rows.get(row)?.source_index,
            self.columns.get(column)?.source_index,
        ))
    }

    /// Physical output rectangle for a cell or merge on a particular page.
    /// Returns None if any part lies outside that page (never silently clips).
    pub fn rect(&self, page: usize, rows: Range<usize>, columns: Range<usize>) -> Option<Rect> {
        let page = self.pages.get(page)?;
        let (y, height) = placement(&rows, &page.rows, self.repeat_rows, &self.row_offsets)?;
        let (x, width) = placement(
            &columns,
            &page.columns,
            self.repeat_columns,
            &self.column_offsets,
        )?;
        Some(Rect {
            x: self.body.x + x * self.scale,
            y: self.body.y + y * self.scale,
            width: width * self.scale,
            height: height * self.scale,
        })
    }
}

fn placement(
    range: &Range<usize>,
    band: &Range<usize>,
    titles: usize,
    offsets: &[f64],
) -> Option<(f64, f64)> {
    if range.start >= range.end || range.end >= offsets.len() {
        return None;
    }
    let start = if range.end <= titles {
        offsets[range.start]
    } else if range.start >= band.start && range.end <= band.end {
        offsets[titles] + offsets[range.start] - offsets[band.start]
    } else {
        return None;
    };
    Some((start, offsets[range.end] - offsets[range.start]))
}

/// Validate and paginate in down-then-over order. Fit modes only shrink.
pub fn paginate(input: &LayoutInput, settings: &PageSettings) -> Result<PagePlan, LayoutError> {
    if input.rows.is_empty() || input.columns.is_empty() {
        return Err(LayoutError::NothingVisible);
    }
    if input
        .rows
        .len()
        .checked_mul(input.columns.len())
        .is_none_or(|n| n > MAX_CELL_POSITIONS)
    {
        return Err(LayoutError::TooManyCells);
    }
    let row_offsets = offsets(&input.rows, Axis::Row)?;
    let column_offsets = offsets(&input.columns, Axis::Column)?;
    let (mut w, mut h) = settings.paper.size();
    if settings.landscape {
        std::mem::swap(&mut w, &mut h);
    }
    let m = settings.margins;
    let body = Rect {
        x: m.left,
        y: m.top,
        width: w - m.left - m.right,
        height: h - m.top - m.bottom - if settings.footer { 18.0 } else { 0.0 },
    };
    if [m.top, m.right, m.bottom, m.left]
        .iter()
        .any(|v| !v.is_finite() || *v < 0.0)
        || body.width <= 0.0
        || body.height <= 0.0
    {
        return Err(LayoutError::InvalidMargins);
    }
    for (i, merge) in input.merges.iter().enumerate() {
        if !valid_range(&merge.rows, input.rows.len())
            || !valid_range(&merge.columns, input.columns.len())
        {
            return Err(LayoutError::InvalidMerge { index: i });
        }
    }
    // Marking occupied merge positions is bounded by the visible-cell limit.
    // It also avoids quadratic comparisons for sheets with many small merges.
    let mut occupied = std::collections::HashMap::new();
    for (i, merge) in input.merges.iter().enumerate() {
        for row in merge.rows.clone() {
            for column in merge.columns.clone() {
                if let Some(first) = occupied.insert((row, column), i) {
                    return Err(LayoutError::OverlappingMerges { first, second: i });
                }
            }
        }
    }
    validate_titles(input, settings.repeat_rows, Axis::Row)?;
    validate_titles(input, settings.repeat_columns, Axis::Column)?;
    let fit_w = body.width / column_offsets.last().unwrap();
    let fit_h = body.height / row_offsets.last().unwrap();
    let (scale, constraint) = match settings.scale {
        Scale::Actual => (1.0, ScaleConstraint::None),
        Scale::Custom(s) if s.is_finite() && (0.1..=4.0).contains(&s) => (s, ScaleConstraint::None),
        Scale::Custom(_) => return Err(LayoutError::InvalidScale),
        Scale::FitColumns => (
            fit_w.min(1.0),
            if fit_w < 1.0 {
                ScaleConstraint::Width
            } else {
                ScaleConstraint::None
            },
        ),
        Scale::FitSheet => {
            let s = fit_w.min(fit_h).min(1.0);
            let c = if s == 1.0 {
                ScaleConstraint::None
            } else if (fit_w - fit_h).abs() < 1e-12 {
                ScaleConstraint::Both
            } else if fit_w < fit_h {
                ScaleConstraint::Width
            } else {
                ScaleConstraint::Height
            };
            (s, c)
        }
    };
    if scale < 0.1 {
        return Err(LayoutError::FitTooSmall { scale });
    }
    let rows = bands(
        input,
        Axis::Row,
        &row_offsets,
        settings.repeat_rows,
        body.height / scale,
    )?;
    let columns = bands(
        input,
        Axis::Column,
        &column_offsets,
        settings.repeat_columns,
        body.width / scale,
    )?;
    if rows
        .len()
        .checked_mul(columns.len())
        .is_none_or(|n| n > MAX_PAGES)
    {
        return Err(LayoutError::TooManyPages);
    }
    let mut readability = Readability {
        smallest_text_pt: None,
        cells_below_threshold: 0,
    };
    let mut seen = HashSet::new();
    for (index, cell) in input.text.iter().enumerate() {
        if cell.row >= input.rows.len()
            || cell.column >= input.columns.len()
            || !cell.font_size_pt.is_finite()
            || cell.font_size_pt <= 0.0
            || !seen.insert((cell.row, cell.column))
        {
            return Err(LayoutError::InvalidTextCell { index });
        }
        if let Some(&merge_index) = occupied.get(&(cell.row, cell.column)) {
            let merge = &input.merges[merge_index];
            if cell.row != merge.rows.start || cell.column != merge.columns.start {
                return Err(LayoutError::InvalidTextCell { index });
            }
        }
        let size = cell.font_size_pt * scale;
        if !size.is_finite() {
            return Err(LayoutError::InvalidTextCell { index });
        }
        readability.smallest_text_pt =
            Some(readability.smallest_text_pt.map_or(size, |s| s.min(size)));
        if size < READABILITY_THRESHOLD_PT {
            readability.cells_below_threshold += 1;
        }
    }
    let pages = columns
        .iter()
        .flat_map(|columns| {
            rows.iter().map(|rows| Page {
                rows: rows.clone(),
                columns: columns.clone(),
            })
        })
        .collect();
    Ok(PagePlan {
        pages,
        rows: input.rows.clone(),
        columns: input.columns.clone(),
        row_offsets,
        column_offsets,
        repeat_rows: settings.repeat_rows,
        repeat_columns: settings.repeat_columns,
        page_size: (w, h),
        body,
        scale,
        constraint,
        readability,
    })
}

fn valid_range(range: &Range<usize>, len: usize) -> bool {
    range.start < range.end && range.end <= len
}

fn offsets(items: &[AxisItem], axis: Axis) -> Result<Vec<f64>, LayoutError> {
    let mut result = vec![0.0];
    let mut seen = HashSet::new();
    for (position, item) in items.iter().enumerate() {
        let sum = result.last().unwrap() + item.size_pt;
        if !item.size_pt.is_finite()
            || item.size_pt <= 0.0
            || !sum.is_finite()
            || sum <= *result.last().unwrap()
        {
            return Err(LayoutError::InvalidAxis { axis, position });
        }
        if !seen.insert(item.source_index) {
            return Err(LayoutError::DuplicateSourceIndex {
                axis,
                source_index: item.source_index,
            });
        }
        result.push(sum);
    }
    Ok(result)
}

fn validate_titles(input: &LayoutInput, titles: usize, axis: Axis) -> Result<(), LayoutError> {
    let len = match axis {
        Axis::Row => input.rows.len(),
        Axis::Column => input.columns.len(),
    };
    if titles >= len {
        return Err(LayoutError::InvalidTitles { axis });
    }
    for (merge, m) in input.merges.iter().enumerate() {
        let r = match axis {
            Axis::Row => &m.rows,
            Axis::Column => &m.columns,
        };
        if r.start < titles && r.end > titles {
            return Err(LayoutError::TitlesCutMerge { axis, merge });
        }
    }
    Ok(())
}

fn bands(
    input: &LayoutInput,
    axis: Axis,
    offsets: &[f64],
    titles: usize,
    capacity: f64,
) -> Result<Vec<Range<usize>>, LayoutError> {
    let available = capacity - offsets[titles];
    if available <= 0.0 {
        return Err(LayoutError::InvalidTitles { axis });
    }
    let len = offsets.len() - 1;
    // A difference array marks grid lines crossed by merges. Overlapping spans
    // on one axis combine transitively, even when their cells don't overlap.
    let mut crossings = vec![0i64; len + 1];
    for m in &input.merges {
        let r = match axis {
            Axis::Row => &m.rows,
            Axis::Column => &m.columns,
        };
        if r.end - r.start > 1 {
            crossings[r.start + 1] += 1;
            crossings[r.end] -= 1;
        }
    }
    let mut forbidden = vec![false; len + 1];
    let mut count = 0i64;
    for (i, delta) in crossings.into_iter().enumerate() {
        count += delta;
        forbidden[i] = count > 0;
    }
    let mut result = Vec::new();
    let mut start = titles;
    while start < len {
        let mut end = start;
        let mut candidate = start + 1;
        while candidate <= len {
            if offsets[candidate] - offsets[start] > available + EPSILON_PT {
                break;
            }
            if !forbidden[candidate] {
                end = candidate;
            }
            candidate += 1;
        }
        if end == start {
            let mut group_end = start + 1;
            while group_end < len && forbidden[group_end] {
                group_end += 1;
            }
            return Err(LayoutError::CannotFit {
                axis,
                positions: start..group_end,
            });
        }
        result.push(start..end);
        if result.len() > MAX_PAGES {
            return Err(LayoutError::TooManyPages);
        }
        start = end;
    }
    Ok(result)
}
