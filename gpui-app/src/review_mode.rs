//! Window-local view state for an immutable operation-plan preview.
//!
//! The prepared plan remains the single owner of materialized changes and
//! workbook snapshots. This index stores only change offsets and row markers,
//! so entering Review Mode does not duplicate a potentially large plan.

use std::collections::{BTreeMap, HashMap, HashSet};

use visigrid_engine::operation_plan::{
    ChangeKind, GroupId, MaterializedChange, OperationPlan, PlanId, ReviewRowState,
};
use visigrid_engine::sheet::SheetId;

use crate::app::Spreadsheet;
use crate::terminal::state::PendingResult;

pub const OVERVIEW_BUCKET_COUNT: usize = 64;

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ReviewEndpoint {
    Before,
    #[default]
    After,
}

#[derive(Debug)]
pub struct ReviewModeState {
    pub plan_id: PlanId,
    pub source_sheet_id: SheetId,
    pub endpoint: ReviewEndpoint,
    by_before_cell: HashMap<(usize, usize), usize>,
    by_before_row: BTreeMap<usize, Vec<usize>>,
    deleted_rows: HashSet<usize>,
    group_counts: HashMap<GroupId, usize>,
    navigable_cells: Vec<(usize, usize)>,
    navigable_groups: Vec<(usize, usize)>,
    overview_buckets: Vec<usize>,
    source_row_count: usize,
}

impl ReviewModeState {
    pub fn from_plan(plan: &OperationPlan) -> Self {
        let mut by_before_cell = HashMap::new();
        let mut by_before_row: BTreeMap<usize, Vec<usize>> = BTreeMap::new();

        for (index, change) in plan.changes.iter().enumerate() {
            let Some(coordinate) = change.before_coordinate else {
                continue;
            };
            by_before_row.entry(coordinate.row).or_default().push(index);
            // RowDeleted is a row-level marker at column zero. Do not let it
            // shadow the actual cleared-cell change at the same coordinate.
            if change.kind != ChangeKind::RowDeleted {
                by_before_cell.insert((coordinate.row, coordinate.col), index);
            }
        }

        let deleted_rows = plan
            .row_lineage
            .iter()
            .filter(|row| row.state == ReviewRowState::Deleted)
            .filter_map(|row| row.before_data_row)
            .collect();
        let mut group_counts = HashMap::new();
        for change in &plan.changes {
            if let Some(group_id) = &change.group_id {
                *group_counts.entry(group_id.clone()).or_default() += 1;
            }
        }
        let mut navigable_cells: Vec<_> = plan
            .changes
            .iter()
            .filter_map(|change| change.before_coordinate)
            .map(|coordinate| (coordinate.row, coordinate.col))
            .collect();
        navigable_cells.sort_unstable();
        navigable_cells.dedup();
        let mut navigable_groups: Vec<_> = plan
            .groups
            .iter()
            .filter_map(|group| {
                plan.changes
                    .iter()
                    .filter(|change| change.group_id.as_ref() == Some(&group.id))
                    .filter_map(|change| change.before_coordinate)
                    .map(|coordinate| (coordinate.row, coordinate.col))
                    .min()
            })
            .collect();
        navigable_groups.sort_unstable();
        navigable_groups.dedup();
        let source_row_count = plan
            .row_lineage
            .iter()
            .filter_map(|row| row.before_data_row)
            .max()
            .map_or(1, |row| row + 1);
        let mut overview_buckets = vec![0; OVERVIEW_BUCKET_COUNT];
        for change in &plan.changes {
            if change.kind == ChangeKind::RowDeleted {
                continue;
            }
            if let Some(coordinate) = change.before_coordinate {
                let bucket = (coordinate.row * OVERVIEW_BUCKET_COUNT / source_row_count)
                    .min(OVERVIEW_BUCKET_COUNT - 1);
                overview_buckets[bucket] += 1;
            }
        }

        Self {
            plan_id: plan.id.clone(),
            source_sheet_id: plan.source_sheet_id,
            endpoint: ReviewEndpoint::After,
            by_before_cell,
            by_before_row,
            deleted_rows,
            group_counts,
            navigable_cells,
            navigable_groups,
            overview_buckets,
            source_row_count,
        }
    }

    pub fn change_index_at_source(&self, row: usize, col: usize) -> Option<usize> {
        self.by_before_cell.get(&(row, col)).copied()
    }

    pub fn is_deleted_source_row(&self, row: usize) -> bool {
        self.deleted_rows.contains(&row)
    }

    pub fn change_indices_at_source_row(&self, row: usize) -> &[usize] {
        self.by_before_row
            .get(&row)
            .map(Vec::as_slice)
            .unwrap_or(&[])
    }

    pub fn group_change_count(&self, group_id: &GroupId) -> usize {
        self.group_counts.get(group_id).copied().unwrap_or(0)
    }

    pub fn adjacent_source_change(
        &self,
        row: usize,
        col: usize,
        forward: bool,
    ) -> Option<(usize, usize)> {
        adjacent_coordinate(&self.navigable_cells, row, col, forward)
    }

    pub fn adjacent_source_group(
        &self,
        row: usize,
        col: usize,
        forward: bool,
    ) -> Option<(usize, usize)> {
        adjacent_coordinate(&self.navigable_groups, row, col, forward)
    }

    pub fn overview_buckets(&self) -> &[usize] {
        &self.overview_buckets
    }

    pub fn source_row_for_bucket(&self, bucket: usize) -> usize {
        bucket.min(OVERVIEW_BUCKET_COUNT - 1) * self.source_row_count / OVERVIEW_BUCKET_COUNT
    }

    pub fn bucket_for_source_row(&self, row: usize) -> usize {
        (row * OVERVIEW_BUCKET_COUNT / self.source_row_count).min(OVERVIEW_BUCKET_COUNT - 1)
    }

    pub fn set_endpoint(&mut self, endpoint: ReviewEndpoint) {
        self.endpoint = endpoint;
    }
}

fn adjacent_coordinate(
    coordinates: &[(usize, usize)],
    row: usize,
    col: usize,
    forward: bool,
) -> Option<(usize, usize)> {
    if coordinates.is_empty() {
        return None;
    }

    let current = (row, col);
    let index = if forward {
        let next = coordinates.partition_point(|coordinate| *coordinate <= current);
        if next == coordinates.len() {
            0
        } else {
            next
        }
    } else {
        let previous = coordinates.partition_point(|coordinate| *coordinate < current);
        if previous == 0 {
            coordinates.len() - 1
        } else {
            previous - 1
        }
    };
    coordinates.get(index).copied()
}

impl Spreadsheet {
    pub fn navigate_review_change(
        &mut self,
        forward: bool,
        by_group: bool,
        cx: &mut gpui::Context<Self>,
    ) {
        let source_sheet_id = match self.review_mode.as_ref() {
            Some(state) => state.source_sheet_id,
            None => return,
        };
        let current = if self.wb(cx).active_sheet_id() == source_sheet_id {
            (
                self.view_to_data(self.view_state.selected.0, cx),
                self.view_state.selected.1,
            )
        } else if forward {
            (usize::MAX, usize::MAX)
        } else {
            (0, 0)
        };
        let Some((sheet_id, row, col)) = self.review_mode.as_ref().and_then(|state| {
            let target = if by_group {
                state.adjacent_source_group(current.0, current.1, forward)
            } else {
                state.adjacent_source_change(current.0, current.1, forward)
            };
            target.map(|(row, col)| (state.source_sheet_id, row, col))
        }) else {
            return;
        };
        let Some(sheet_index) = self.wb(cx).sheet_index_by_id(sheet_id) else {
            return;
        };
        let view_row = self.data_to_view(row, cx).unwrap_or(row);
        self.reveal_cell(sheet_index, view_row, col, cx);
    }

    /// Resolve a visible source-grid coordinate in O(1). The returned change
    /// remains owned by the immutable prepared plan in the terminal preview.
    pub fn review_change_at_source(
        &self,
        sheet_id: SheetId,
        row: usize,
        col: usize,
    ) -> Option<&MaterializedChange> {
        let state = self.review_mode.as_ref()?;
        if state.source_sheet_id != sheet_id {
            return None;
        }
        let index = state.change_index_at_source(row, col)?;
        let PendingResult::LuaPreview(preview) = self.terminal.pending_result.as_ref()? else {
            return None;
        };
        let prepared = preview.prepared_plan.as_ref()?;
        if prepared.plan().id != state.plan_id {
            return None;
        }
        prepared.plan().changes.get(index)
    }

    pub fn review_change_for_source_selection(
        &self,
        sheet_id: SheetId,
        row: usize,
        col: usize,
    ) -> Option<&MaterializedChange> {
        let state = self.review_mode.as_ref()?;
        if state.source_sheet_id != sheet_id {
            return None;
        }
        let PendingResult::LuaPreview(preview) = self.terminal.pending_result.as_ref()? else {
            return None;
        };
        let prepared = preview.prepared_plan.as_ref()?;
        if prepared.plan().id != state.plan_id {
            return None;
        }
        if state.is_deleted_source_row(row) {
            return state
                .change_indices_at_source_row(row)
                .iter()
                .filter_map(|index| prepared.plan().changes.get(*index))
                .find(|change| change.kind == ChangeKind::RowDeleted);
        }
        let index = state
            .change_index_at_source(row, col)
            .or_else(|| state.change_indices_at_source_row(row).first().copied())?;
        prepared.plan().changes.get(index)
    }

    pub fn review_row_is_deleted(&self, sheet_id: SheetId, row: usize) -> bool {
        self.review_mode.as_ref().is_some_and(|state| {
            state.source_sheet_id == sheet_id && state.is_deleted_source_row(row)
        })
    }
}
