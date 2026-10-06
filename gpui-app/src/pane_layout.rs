//! Shared frozen/scrolling geometry. Scroll and freeze boundaries are view slots,
//! never ranks in the filtered row list. A zero extent denotes a hidden slot.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct Slot {
    pub index: usize,
    pub offset: f32,
    pub size: f32,
}
#[derive(Debug)]
pub(crate) struct AxisLayout {
    pub frozen: Vec<Slot>,
    pub body: Vec<Slot>,
    pub body_offset: f32,
}
impl AxisLayout {
    pub fn hit(&self, position: f32) -> Option<usize> {
        if position < 0.0 {
            return None;
        }
        self.frozen
            .iter()
            .chain(&self.body)
            .find(|s| position >= s.offset && position < s.offset + s.size)
            .map(|s| s.index)
    }
}
pub(crate) fn layout(
    limit: usize,
    boundary: usize,
    scroll: usize,
    viewport: f32,
    size: impl Fn(usize) -> f32,
) -> AxisLayout {
    from_slots(
        boundary,
        viewport,
        (0..boundary.min(limit)).map(|i| (i, size(i))),
        (scroll.max(boundary)..limit).map(|i| (i, size(i))),
    )
}

fn from_slots(
    boundary: usize,
    viewport: f32,
    frozen_slots: impl Iterator<Item = (usize, f32)>,
    body_slots: impl Iterator<Item = (usize, f32)>,
) -> AxisLayout {
    let mut frozen = Vec::new();
    let mut body = Vec::new();
    let mut offset = 0.0;
    for (index, extent) in frozen_slots {
        if offset >= viewport {
            break;
        }
        if extent > 0.0 {
            frozen.push(Slot {
                index,
                offset,
                size: extent,
            });
            offset += extent;
        }
    }
    if boundary > 0 {
        offset += 1.0;
    }
    let body_offset = offset;
    for (index, extent) in body_slots {
        if offset >= viewport {
            break;
        }
        if extent > 0.0 {
            body.push(Slot {
                index,
                offset,
                size: extent,
            });
            offset += extent;
        }
    }
    AxisLayout {
        frozen,
        body,
        body_offset,
    }
}

pub(crate) fn offset(
    target: usize,
    boundary: usize,
    scroll: usize,
    size: impl Fn(usize) -> f32,
) -> f32 {
    if target < boundary {
        return (0..target).map(&size).sum();
    }
    let frozen: f32 = (0..boundary).map(&size).sum();
    let start = scroll.max(boundary);
    let delta: f32 = if target >= start {
        (start..target).map(&size).sum()
    } else {
        -(target..start).map(&size).sum::<f32>()
    };
    frozen + if boundary > 0 { 1.0 } else { 0.0 } + delta
}

/// Keep the target fully visible, walking backwards from it only when needed.
pub(crate) fn ensure_visible(
    limit: usize,
    boundary: usize,
    scroll: usize,
    target: usize,
    viewport: f32,
    size: impl Fn(usize) -> f32,
) -> usize {
    if boundary >= limit {
        return boundary;
    }
    let current = scroll.max(boundary).min(limit.saturating_sub(1));
    if target < boundary || target >= limit || size(target) <= 0.0 {
        return current;
    }
    let frozen: f32 = (0..boundary.min(limit)).map(&size).sum();
    let space = viewport - frozen - if boundary > 0 { 1.0 } else { 0.0 };
    if space <= 0.0 {
        return current;
    }
    if target < current {
        return target;
    }
    let used: f32 = (current..=target).map(&size).sum();
    if used <= space {
        return current;
    }
    let mut first = target;
    let mut used = size(target);
    for index in (boundary..target).rev() {
        let extent = size(index);
        if extent <= 0.0 {
            continue;
        }
        if used + extent > space {
            break;
        }
        first = index;
        used += extent;
    }
    first
}

pub(crate) fn scroll(
    limit: usize,
    boundary: usize,
    current: usize,
    delta: i32,
    visible: impl Fn(usize) -> bool,
) -> usize {
    if boundary >= limit {
        return boundary;
    }
    let position = current.max(boundary).min(limit.saturating_sub(1));
    let start = (position..limit).find(|&i| visible(i)).unwrap_or(position);
    let mut next = start;
    let mut remaining = delta.unsigned_abs();
    if delta > 0 {
        for index in start.saturating_add(1)..limit {
            if visible(index) {
                next = index;
                remaining -= 1;
                if remaining == 0 {
                    break;
                }
            }
        }
    } else if delta < 0 {
        for index in (boundary..start).rev() {
            if visible(index) {
                next = index;
                remaining -= 1;
                if remaining == 0 {
                    break;
                }
            }
        }
    }
    next
}

/// Adjacent displayed data rows for a view slot. Filtering has an indexed
/// inverse; only immediately adjacent manually hidden rows need to be skipped.
pub(crate) fn row_neighbors(
    rows: &visigrid_engine::filter::RowView,
    hidden: Option<&std::collections::BTreeSet<usize>>,
    slot: usize,
) -> (Option<usize>, Option<usize>) {
    (
        crate::formatting::plan::row_neighbor(rows, hidden, slot, false),
        crate::formatting::plan::row_neighbor(rows, hidden, slot, true),
    )
}

use crate::{app::Spreadsheet, workbook_view::WorkbookViewState};
impl Spreadsheet {
    pub(crate) fn displayed_row_height(&self, row: usize) -> f32 {
        if !self.row_view.is_view_row_visible(row)
            || self.is_row_hidden(self.row_view.view_to_data(row))
        {
            return 0.0;
        }
        self.metrics
            .row_height(self.row_height(self.row_view.view_to_data(row)))
    }
    pub(crate) fn displayed_col_width(&self, col: usize) -> f32 {
        if self.is_col_hidden(col) {
            0.0
        } else {
            self.metrics.col_width(self.col_width(col))
        }
    }
    pub(crate) fn pane_rows(&self, view: &WorkbookViewState) -> AxisLayout {
        let rows = &self.row_view;
        let hidden = self.display_hidden_rows();
        let height = |(slot, data)| (slot, self.metrics.row_height(self.row_height(data)));
        from_slots(
            view.frozen_rows,
            self.grid_layout.viewport_size.1,
            crate::formatting::plan::visible_rows(
                rows,
                hidden,
                0,
                view.frozen_rows.saturating_sub(1),
            )
            .take_while(|(slot, _)| *slot < view.frozen_rows)
            .map(height),
            crate::formatting::plan::visible_rows(
                rows,
                hidden,
                view.scroll_row.max(view.frozen_rows),
                rows.row_count().saturating_sub(1),
            )
            .map(height),
        )
    }
    pub(crate) fn pane_cols(&self, view: &WorkbookViewState) -> AxisLayout {
        layout(
            crate::app::NUM_COLS,
            view.frozen_cols,
            view.scroll_col,
            self.grid_layout.viewport_size.0,
            |c| self.displayed_col_width(c),
        )
    }
}

#[cfg(test)]
#[path = "pane_layout_tests.rs"]
mod tests;

#[cfg(test)]
mod displayed_neighbor_tests {
    use super::row_neighbors;
    use visigrid_engine::filter::RowView;

    #[test]
    fn borders_use_displayed_neighbors_through_sort_filter_and_manual_hiding() {
        let mut rows = RowView::new(8);
        rows.apply_sort(vec![0, 4, 2, 3, 1, 5, 6, 7]);
        rows.apply_filter((0..8).map(|data| data != 2).collect());
        let hidden = [4, 6].into();
        assert_eq!(row_neighbors(&rows, Some(&hidden), 3), (Some(0), Some(1)));
        assert_eq!(row_neighbors(&rows, Some(&hidden), 5), (Some(1), Some(7)));
        assert_eq!(row_neighbors(&rows, Some(&hidden), 0), (None, Some(3)));
        assert_eq!(row_neighbors(&rows, Some(&hidden), 7), (Some(5), None));
    }
}
