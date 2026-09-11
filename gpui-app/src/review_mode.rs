//! Window-local view state for an immutable operation-plan preview.
//!
//! The prepared plan remains the single owner of materialized changes and
//! workbook snapshots. This index stores only change offsets and row markers,
//! so entering Review Mode does not duplicate a potentially large plan.

use std::cell::RefCell;
use std::collections::{BTreeMap, HashMap, HashSet};

use visigrid_engine::operation_plan::{
    ChangeKind, DeterminismClass, ExecutionContextFingerprint, GroupId, MaterializedChange,
    OperationPlan, PlanId, PreparedOperationPlan, ProblemSeverity, ReviewRowState,
};
use visigrid_engine::sheet::{Sheet, SheetId};
use visigrid_engine::workbook::Workbook;

use crate::app::Spreadsheet;
use crate::terminal::state::PendingResult;

pub const OVERVIEW_BUCKET_COUNT: usize = 64;

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ReviewOverviewBucket {
    pub change_count: usize,
    pub has_deleted_row: bool,
    pub has_new_formula_error: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReviewApplyStatus {
    pub stale: bool,
    pub blocking: bool,
    pub determinism: DeterminismClass,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ReviewEligibilityKey {
    workbook_revision: u64,
    execution_context: crate::scripting::ExecutionContextGenerationKey,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ReviewEligibilityCache {
    key: ReviewEligibilityKey,
    status: ReviewApplyStatus,
}

impl ReviewApplyStatus {
    pub fn evaluate(
        prepared: &PreparedOperationPlan,
        workbook: &Workbook,
        context: &ExecutionContextFingerprint,
    ) -> Self {
        Self {
            stale: prepared.is_stale(workbook, context),
            blocking: prepared
                .plan()
                .problems
                .iter()
                .any(|problem| problem.severity == ProblemSeverity::Blocking),
            determinism: prepared.plan().determinism,
        }
    }

    pub fn can_apply(self) -> bool {
        !self.stale && !self.blocking && self.determinism == DeterminismClass::Full
    }

    pub fn disabled_reason(self) -> Option<&'static str> {
        if self.stale {
            Some("Plan is stale · re-preview required")
        } else if self.blocking {
            Some("Resolve blocking review problems")
        } else {
            match self.determinism {
                DeterminismClass::Full => None,
                DeterminismClass::Conditional => {
                    Some("Conditional plan · explicit override is not available yet")
                }
                DeterminismClass::Unresolved => Some("Plan determinism is unresolved"),
            }
        }
    }

    pub fn determinism_label(self) -> &'static str {
        match self.determinism {
            DeterminismClass::Full => "Deterministic",
            DeterminismClass::Conditional => "Conditional",
            DeterminismClass::Unresolved => "Unresolved",
        }
    }
}

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
    overview_buckets: Vec<ReviewOverviewBucket>,
    overview_bucket_targets: Vec<Vec<(usize, usize)>>,
    after_row_by_before: Vec<Option<usize>>,
    source_row_count: usize,
    eligibility_cache: RefCell<Option<ReviewEligibilityCache>>,
}

impl ReviewModeState {
    pub fn from_prepared(prepared: &PreparedOperationPlan, workbook: &Workbook) -> Self {
        let mut state = Self::from_plan(prepared.plan());
        let key = Self::eligibility_key(workbook);
        let context = crate::scripting::execution_context_fingerprint(
            workbook,
            &prepared.plan().operations,
        );
        let status = ReviewApplyStatus::evaluate(prepared, workbook, &context);
        *state.eligibility_cache.get_mut() = Some(ReviewEligibilityCache { key, status });
        state
    }

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
        let mut overview_buckets = vec![ReviewOverviewBucket::default(); OVERVIEW_BUCKET_COUNT];
        let mut overview_bucket_targets = vec![Vec::new(); OVERVIEW_BUCKET_COUNT];
        for change in &plan.changes {
            if let Some(coordinate) = change.before_coordinate {
                let bucket = (coordinate.row * OVERVIEW_BUCKET_COUNT / source_row_count)
                    .min(OVERVIEW_BUCKET_COUNT - 1);
                let overview = &mut overview_buckets[bucket];
                overview.change_count += 1;
                overview.has_deleted_row |= change.kind == ChangeKind::RowDeleted;
                overview.has_new_formula_error |= change.after.display.starts_with('#')
                    && !change.before.display.starts_with('#');
                overview_bucket_targets[bucket].push((coordinate.row, coordinate.col));
            }
        }
        for targets in &mut overview_bucket_targets {
            targets.sort_unstable();
            targets.dedup();
        }
        let mut after_row_by_before = vec![None; source_row_count];
        for row in &plan.row_lineage {
            if let Some(before_row) = row.before_data_row {
                if before_row < after_row_by_before.len() {
                    after_row_by_before[before_row] = row.after_data_row;
                }
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
            overview_bucket_targets,
            after_row_by_before,
            source_row_count,
            eligibility_cache: RefCell::new(None),
        }
    }

    fn eligibility_key(workbook: &Workbook) -> ReviewEligibilityKey {
        ReviewEligibilityKey {
            workbook_revision: workbook.revision(),
            execution_context: crate::scripting::execution_context_generation_key(workbook),
        }
    }

    pub fn apply_status(
        &self,
        prepared: &PreparedOperationPlan,
        workbook: &Workbook,
    ) -> ReviewApplyStatus {
        let key = Self::eligibility_key(workbook);
        if let Some(cache) = self.eligibility_cache.borrow().as_ref() {
            if cache.key == key {
                return cache.status;
            }
        }
        let context = crate::scripting::execution_context_fingerprint(
            workbook,
            &prepared.plan().operations,
        );
        let status = ReviewApplyStatus::evaluate(prepared, workbook, &context);
        *self.eligibility_cache.borrow_mut() = Some(ReviewEligibilityCache { key, status });
        status
    }

    fn after_row_for_source(&self, source_row: usize) -> Option<usize> {
        self.after_row_by_before.get(source_row).copied().flatten()
    }

    pub fn endpoint_sheet_row<'a>(
        &self,
        prepared: &'a PreparedOperationPlan,
        sheet_id: SheetId,
        source_row: usize,
    ) -> Option<(&'a Sheet, usize)> {
        if self.source_sheet_id != sheet_id || prepared.plan().id != self.plan_id {
            return None;
        }
        match self.endpoint {
            ReviewEndpoint::Before => prepared
                .source_workbook()
                .sheet_by_id(sheet_id)
                .map(|sheet| (sheet, source_row)),
            ReviewEndpoint::After => match self.after_row_for_source(source_row) {
                Some(after_row) => prepared
                    .preview_workbook()
                    .sheet_by_id(sheet_id)
                    .map(|sheet| (sheet, after_row)),
                None => prepared
                    .source_workbook()
                    .sheet_by_id(sheet_id)
                    .map(|sheet| (sheet, source_row)),
            },
        }
    }

    pub fn source_sheet<'a>(
        &self,
        prepared: &'a PreparedOperationPlan,
        sheet_id: SheetId,
    ) -> Option<&'a Sheet> {
        if self.source_sheet_id != sheet_id || prepared.plan().id != self.plan_id {
            return None;
        }
        prepared.source_workbook().sheet_by_id(sheet_id)
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

    pub fn adjacent_visible_source_change(
        &self,
        row: usize,
        col: usize,
        forward: bool,
        by_group: bool,
        mut is_visible: impl FnMut(usize) -> bool,
    ) -> Option<(usize, usize)> {
        let mut cursor = (row, col);
        for _ in 0..self.navigation_len(by_group) {
            let target = if by_group {
                self.adjacent_source_group(cursor.0, cursor.1, forward)
            } else {
                self.adjacent_source_change(cursor.0, cursor.1, forward)
            }?;
            cursor = target;
            if is_visible(target.0) {
                return Some(target);
            }
        }
        None
    }

    pub fn navigation_len(&self, by_group: bool) -> usize {
        if by_group {
            self.navigable_groups.len()
        } else {
            self.navigable_cells.len()
        }
    }

    pub fn overview_buckets(&self) -> &[ReviewOverviewBucket] {
        &self.overview_buckets
    }

    pub fn overview_bucket_targets(&self, bucket: usize) -> &[(usize, usize)] {
        self.overview_bucket_targets
            .get(bucket)
            .map(Vec::as_slice)
            .unwrap_or(&[])
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
    pub fn block_review_sheet_switch(&mut self, cx: &mut gpui::Context<Self>) -> bool {
        if self.review_mode.is_none() {
            return false;
        }
        self.status_message =
            Some("Apply or dismiss Review Mode before switching sheets.".into());
        cx.notify();
        true
    }

    pub fn review_apply_status(
        &self,
        prepared: &PreparedOperationPlan,
        cx: &gpui::App,
    ) -> Option<ReviewApplyStatus> {
        let state = self.review_mode.as_ref()?;
        if state.plan_id != prepared.plan().id {
            return None;
        }
        Some(state.apply_status(prepared, self.workbook.read(cx)))
    }

    /// Resolve a source-geometry row against the immutable Before/After
    /// workbook snapshots. Deleted rows deliberately retain their frozen
    /// source content so the deletion tint and strikethrough remain legible.
    pub fn review_endpoint_sheet_row(
        &self,
        sheet_id: SheetId,
        source_row: usize,
    ) -> Option<(&Sheet, usize)> {
        let state = self.review_mode.as_ref()?;
        let PendingResult::LuaPreview(preview) = self.terminal.pending_result.as_ref()? else {
            return None;
        };
        let prepared = preview.prepared_plan.as_ref()?;
        state.endpoint_sheet_row(prepared, sheet_id, source_row)
    }

    pub fn review_source_sheet(&self, sheet_id: SheetId) -> Option<&Sheet> {
        let state = self.review_mode.as_ref()?;
        let PendingResult::LuaPreview(preview) = self.terminal.pending_result.as_ref()? else {
            return None;
        };
        let prepared = preview.prepared_plan.as_ref()?;
        state.source_sheet(prepared, sheet_id)
    }

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
        let Some(sheet_index) = self.wb(cx).sheet_index_by_id(source_sheet_id) else {
            return;
        };
        let target = {
            let row_view = &self.row_view;
            self.review_mode.as_ref().and_then(|state| {
                state.adjacent_visible_source_change(
                    current.0,
                    current.1,
                    forward,
                    by_group,
                    |row| row_view.data_to_view(row).is_some(),
                )
            })
        };
        if let Some((row, col)) = target {
            let view_row = self
                .row_view
                .data_to_view(row)
                .expect("review target was filtered for visibility");
            self.reveal_cell(sheet_index, view_row, col, cx);
            return;
        }
        self.status_message = Some("Review changes are hidden by the active filter.".into());
        cx.notify();
    }

    pub fn navigate_review_bucket(&mut self, bucket: usize, cx: &mut gpui::Context<Self>) {
        let Some(source_sheet_id) = self
            .review_mode
            .as_ref()
            .map(|state| state.source_sheet_id)
        else {
            return;
        };
        let Some(sheet_index) = self.wb(cx).sheet_index_by_id(source_sheet_id) else {
            return;
        };
        let target_count = self
            .review_mode
            .as_ref()
            .map(|state| state.overview_bucket_targets(bucket).len())
            .unwrap_or(0);
        for index in 0..target_count {
            let target = self
                .review_mode
                .as_ref()
                .and_then(|state| state.overview_bucket_targets(bucket).get(index))
                .copied();
            let Some((row, col)) = target else {
                continue;
            };
            if let Some(view_row) = self.data_to_view(row, cx) {
                self.reveal_cell(sheet_index, view_row, col, cx);
                return;
            }
        }
        self.status_message = Some("Review changes in this region are hidden by the active filter.".into());
        cx.notify();
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
