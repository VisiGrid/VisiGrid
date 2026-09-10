//! Producer-neutral, immutable operation plans for Review Mode.
//!
//! A producer supplies primitive operations. The engine normalizes them,
//! applies them to a cloned workbook, materializes the visible differences,
//! and freezes the hashes required to commit that exact preview later.

use std::collections::{BTreeMap, BTreeSet, HashSet};
use std::time::SystemTime;

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::cell::{CellFormat, CellStyle};
use crate::cell_id::CellId;
use crate::sheet::{Sheet, SheetId};
use crate::structural::Axis;
use crate::workbook::Workbook;

pub const OPERATION_PLAN_CONTRACT_VERSION: u32 = 1;
pub const MAX_PLAN_OPERATIONS: usize = 100_000;
pub const MAX_MATERIALIZED_CHANGES: usize = 250_000;

#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct PlanId(pub String);

#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct GroupId(pub String);

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct PlanProducer {
    pub kind: String,
    pub name: String,
    pub source_path: Option<String>,
    pub source_hash: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ExecutionContextFingerprint {
    pub engine_version: String,
    pub functions_generation: u64,
    pub functions_source_hash: String,
    pub locale: String,
    pub timezone: String,
    pub auto_recalc: bool,
    pub iterative_calculation_enabled: bool,
    pub iterative_max_iterations: u32,
    /// Exact IEEE-754 representation, used instead of lossy text or `f64`
    /// equality in the wire-safe fingerprint.
    pub iterative_tolerance_bits: u64,
    #[serde(default)]
    pub volatile_inputs: Vec<String>,
}

impl ExecutionContextFingerprint {
    pub fn hash(&self) -> String {
        hash_serializable(self)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DeterminismClass {
    Full,
    Conditional,
    Unresolved,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
pub struct CellCoordinate {
    pub row: usize,
    pub col: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct CellRange {
    pub start: CellCoordinate,
    pub end: CellCoordinate,
}

impl CellRange {
    pub fn single(row: usize, col: usize) -> Self {
        let coordinate = CellCoordinate { row, col };
        Self {
            start: coordinate,
            end: coordinate,
        }
    }

    fn normalized(self) -> Self {
        Self {
            start: CellCoordinate {
                row: self.start.row.min(self.end.row),
                col: self.start.col.min(self.end.col),
            },
            end: CellCoordinate {
                row: self.start.row.max(self.end.row),
                col: self.start.col.max(self.end.col),
            },
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "snake_case")]
pub enum PlannedCellValue {
    Text(String),
    Number(f64),
    Boolean(bool),
    Error(String),
}

impl PlannedCellValue {
    fn as_input(&self) -> String {
        match self {
            Self::Text(value) => value.clone(),
            Self::Number(value) if value.fract() == 0.0 && value.abs() < 1e15 => {
                format!("{value:.0}")
            }
            Self::Number(value) => value.to_string(),
            Self::Boolean(value) => {
                if *value {
                    "TRUE".into()
                } else {
                    "FALSE".into()
                }
            }
            Self::Error(value) => format!("#ERROR: {value}"),
        }
    }
}

/// Primitive operation vocabulary. Operations always address the one source
/// sheet identified by their containing plan.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum PlannedOp {
    SetCellValue {
        coordinate: CellCoordinate,
        value: PlannedCellValue,
    },
    SetCellFormula {
        coordinate: CellCoordinate,
        formula: String,
    },
    ClearCell {
        coordinate: CellCoordinate,
    },
    ClearRange {
        range: CellRange,
    },
    SetCellStyle {
        range: CellRange,
        style: CellStyle,
    },
    DeleteRows {
        at: usize,
        count: usize,
    },
    // Reserved in v1 so the wire vocabulary remains primitive. Materializing
    // these is Phase 1.1 and returns `unsupported_operation` for now.
    InsertRows {
        at: usize,
        count: usize,
    },
    InsertColumns {
        at: usize,
        count: usize,
    },
    DeleteColumns {
        at: usize,
        count: usize,
    },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OperationGroup {
    pub id: GroupId,
    pub title: String,
    pub description: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ProducerClaim(pub String);

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum VerificationDefinition {
    NoNewFormulaErrors { id: String, label: Option<String> },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum VerificationStatus {
    Passed,
    Failed,
    Unknown,
    NotRun,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct VerificationResult {
    pub id: String,
    pub label: Option<String>,
    pub status: VerificationStatus,
    pub evidence: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ProblemSeverity {
    Warning,
    Blocking,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct PlanProblem {
    pub code: String,
    pub message: String,
    pub severity: ProblemSeverity,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct ReviewRowId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ReviewRowState {
    Unchanged,
    Changed,
    Inserted,
    Deleted,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ReviewRowLineage {
    pub id: ReviewRowId,
    pub before_data_row: Option<usize>,
    pub after_data_row: Option<usize>,
    pub display_position: usize,
    pub state: ReviewRowState,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellSnapshot {
    pub raw: String,
    pub display: String,
    pub format: CellFormat,
}

impl CellSnapshot {
    fn from_sheet(sheet: &Sheet, row: usize, col: usize) -> Self {
        Self {
            raw: sheet.get_raw(row, col),
            display: sheet.get_display(row, col),
            format: sheet.get_format(row, col),
        }
    }

    fn empty() -> Self {
        Self {
            raw: String::new(),
            display: String::new(),
            format: CellFormat::default(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ChangeKind {
    Value,
    Formula,
    Cleared,
    Format,
    RowDeleted,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ChangeCause {
    Direct,
    Recalculated,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct MaterializedChange {
    pub sheet_id: SheetId,
    pub row_id: ReviewRowId,
    pub before_coordinate: Option<CellCoordinate>,
    pub after_coordinate: Option<CellCoordinate>,
    pub kind: ChangeKind,
    pub cause: ChangeCause,
    pub before: CellSnapshot,
    pub after: CellSnapshot,
    pub group_id: Option<GroupId>,
    pub reason: Option<ProducerClaim>,
    pub sources: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct AffectedRange {
    pub sheet_id: SheetId,
    pub range: CellRange,
}

#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct PlanSummary {
    pub cells_changed: usize,
    pub cells_cleared: usize,
    pub formulas_changed: usize,
    pub formatting_changes: usize,
    pub rows_deleted: usize,
    pub recalculated_cells: usize,
}

impl PlanSummary {
    pub fn total_changes(&self) -> usize {
        self.cells_changed
            + self.cells_cleared
            + self.formulas_changed
            + self.formatting_changes
            + self.rows_deleted
            + self.recalculated_cells
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct OperationPlan {
    pub id: PlanId,
    pub contract_version: u32,
    pub workbook_session_id: String,
    pub source_sheet_id: SheetId,
    pub source_revision: u64,
    pub source_fingerprint: String,
    pub execution_context: ExecutionContextFingerprint,
    pub determinism: DeterminismClass,
    pub plan_hash: String,
    pub producer: PlanProducer,
    pub title: String,
    pub description: Option<String>,
    pub operations: Vec<PlannedOp>,
    pub groups: Vec<OperationGroup>,
    pub changes: Vec<MaterializedChange>,
    pub affected_ranges: Vec<AffectedRange>,
    pub verification: Vec<VerificationResult>,
    pub problems: Vec<PlanProblem>,
    pub summary: PlanSummary,
    pub row_lineage: Vec<ReviewRowLineage>,
    pub preview_fingerprint: String,
    pub created_at: SystemTime,
}

#[derive(Debug, Clone)]
pub struct OperationPlanRequest {
    pub id: PlanId,
    pub workbook_session_id: String,
    pub source_sheet_id: SheetId,
    pub expected_revision: u64,
    pub execution_context: ExecutionContextFingerprint,
    pub producer: PlanProducer,
    pub title: String,
    pub description: Option<String>,
    pub operations: Vec<PlannedOp>,
    pub groups: Vec<OperationGroup>,
    pub verification: Vec<VerificationDefinition>,
}

/// Runtime-only snapshots paired with the persisted/wire-safe plan.
#[derive(Debug, Clone)]
pub struct PreparedOperationPlan {
    plan: OperationPlan,
    source_workbook: Workbook,
    preview_workbook: Workbook,
}

#[derive(Debug, Clone)]
pub struct PlanCommit {
    pub plan_id: PlanId,
    pub source: Workbook,
    pub applied: Workbook,
    pub affected_cells: Vec<CellId>,
}

impl PlanCommit {
    pub fn undo_into(&self, workbook: &mut Workbook) {
        let revision = workbook.revision().saturating_add(1);
        *workbook = self.source.clone();
        workbook.set_revision_after_atomic_commit(revision);
    }

    pub fn redo_into(&self, workbook: &mut Workbook) {
        let revision = workbook.revision().saturating_add(1);
        *workbook = self.applied.clone();
        workbook.set_revision_after_atomic_commit(revision);
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PlanError {
    RevisionMismatch { expected: u64, actual: u64 },
    SourceSheetMissing,
    OperationLimitExceeded { limit: usize },
    MaterializedChangeLimitExceeded { limit: usize },
    MultiSheetUnsupported,
    InvalidOperation(String),
    UnsupportedOperation(String),
    ContextChanged,
    PlanHashMismatch,
    PreviewFingerprintMismatch,
    Unresolved,
    BlockingProblems,
    ConditionalRequiresOverride,
}

impl std::fmt::Display for PlanError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RevisionMismatch { expected, actual } => {
                write!(
                    f,
                    "source revision changed (expected {expected}, found {actual})"
                )
            }
            Self::SourceSheetMissing => write!(f, "source sheet no longer exists"),
            Self::OperationLimitExceeded { limit } => {
                write!(f, "normalized operation limit exceeded ({limit})")
            }
            Self::MaterializedChangeLimitExceeded { limit } => {
                write!(f, "materialized change limit exceeded ({limit})")
            }
            Self::MultiSheetUnsupported => {
                write!(
                    f,
                    "multi_sheet_unsupported: the preview affects another sheet"
                )
            }
            Self::InvalidOperation(message) => write!(f, "invalid operation: {message}"),
            Self::UnsupportedOperation(message) => write!(f, "unsupported operation: {message}"),
            Self::ContextChanged => write!(f, "calculation context changed"),
            Self::PlanHashMismatch => write!(f, "operation plan hash does not match its contents"),
            Self::PreviewFingerprintMismatch => {
                write!(f, "candidate does not match reviewed preview")
            }
            Self::Unresolved => write!(f, "plan calculation did not settle"),
            Self::BlockingProblems => write!(f, "plan has blocking problems"),
            Self::ConditionalRequiresOverride => {
                write!(
                    f,
                    "conditional plan requires an explicit determinism override"
                )
            }
        }
    }
}

impl std::error::Error for PlanError {}

impl PreparedOperationPlan {
    pub fn materialize(
        source: &Workbook,
        request: OperationPlanRequest,
    ) -> Result<Self, PlanError> {
        if source.revision() != request.expected_revision {
            return Err(PlanError::RevisionMismatch {
                expected: request.expected_revision,
                actual: source.revision(),
            });
        }
        let sheet_index = source
            .sheet_index_by_id(request.source_sheet_id)
            .ok_or(PlanError::SourceSheetMissing)?;
        let source_sheet = source
            .sheet(sheet_index)
            .ok_or(PlanError::SourceSheetMissing)?;
        let operations = normalize_operations(source_sheet, request.operations)?;
        let source_fingerprint = workbook_fingerprint(source);
        let plan_hash = compute_plan_hash(
            OPERATION_PLAN_CONTRACT_VERSION,
            &request.workbook_session_id,
            request.source_sheet_id,
            request.expected_revision,
            &request.execution_context,
            &operations,
        );

        let mut preview = source.clone();
        let report = apply_operations(&mut preview, request.source_sheet_id, &operations)?;
        let mut row_lineage = build_row_lineage(source_sheet.rows, &operations);
        let changes = materialize_changes(
            source,
            &preview,
            request.source_sheet_id,
            &operations,
            &row_lineage,
        )?;
        let changed_row_ids: HashSet<_> = changes
            .iter()
            .filter(|change| change.sheet_id == request.source_sheet_id)
            .map(|change| change.row_id.0)
            .collect();
        for row in &mut row_lineage {
            if row.state == ReviewRowState::Unchanged && changed_row_ids.contains(&row.id.0) {
                row.state = ReviewRowState::Changed;
            }
        }
        let affected_ranges = affected_ranges(&changes);
        let summary = summarize(&changes);
        let new_errors = changes
            .iter()
            .filter(|change| {
                is_formula_error(&change.after.display) && !is_formula_error(&change.before.display)
            })
            .count();
        let verification = request
            .verification
            .into_iter()
            .map(|definition| match definition {
                VerificationDefinition::NoNewFormulaErrors { id, label } => VerificationResult {
                    id,
                    label,
                    status: if new_errors == 0 {
                        VerificationStatus::Passed
                    } else {
                        VerificationStatus::Failed
                    },
                    evidence: format!("{new_errors} new formula error(s)"),
                },
            })
            .collect();

        let mut problems = Vec::new();
        if new_errors > 0 {
            problems.push(PlanProblem {
                code: "new_formula_errors".into(),
                message: format!("The preview introduces {new_errors} new formula error(s)."),
                severity: ProblemSeverity::Blocking,
            });
        }
        if !report.errors.is_empty() || (report.scc_count > 0 && !report.converged) {
            problems.push(PlanProblem {
                code: "recalculation_unresolved".into(),
                message: "Preview recalculation did not settle.".into(),
                severity: ProblemSeverity::Blocking,
            });
        }

        let determinism = if problems
            .iter()
            .any(|problem| problem.code == "recalculation_unresolved")
        {
            DeterminismClass::Unresolved
        } else if request.execution_context.volatile_inputs.is_empty() {
            DeterminismClass::Full
        } else {
            DeterminismClass::Conditional
        };
        let preview_fingerprint = workbook_fingerprint(&preview);
        let plan = OperationPlan {
            id: request.id,
            contract_version: OPERATION_PLAN_CONTRACT_VERSION,
            workbook_session_id: request.workbook_session_id,
            source_sheet_id: request.source_sheet_id,
            source_revision: request.expected_revision,
            source_fingerprint,
            execution_context: request.execution_context,
            determinism,
            plan_hash,
            producer: request.producer,
            title: request.title,
            description: request.description,
            operations,
            groups: request.groups,
            changes,
            affected_ranges,
            verification,
            problems,
            summary,
            row_lineage,
            preview_fingerprint,
            created_at: SystemTime::now(),
        };

        Ok(Self {
            plan,
            source_workbook: source.clone(),
            preview_workbook: preview,
        })
    }

    pub fn source_workbook(&self) -> &Workbook {
        &self.source_workbook
    }

    /// The frozen, wire-safe plan. Mutating operations requires constructing
    /// and materializing a new `PreparedOperationPlan` with a new plan ID.
    pub fn plan(&self) -> &OperationPlan {
        &self.plan
    }

    pub fn preview_workbook(&self) -> &Workbook {
        &self.preview_workbook
    }

    pub fn is_stale(&self, workbook: &Workbook, context: &ExecutionContextFingerprint) -> bool {
        workbook.revision() != self.plan.source_revision
            || workbook_fingerprint(workbook) != self.plan.source_fingerprint
            || context != &self.plan.execution_context
    }

    pub fn verify_candidate(
        &self,
        current: &Workbook,
        context: &ExecutionContextFingerprint,
    ) -> Result<PlanCommit, PlanError> {
        self.verify_candidate_inner(current, context, false)
    }

    /// Verify a conditional plan after the human has explicitly acknowledged
    /// its volatile/external inputs. Unresolved and blocking plans remain
    /// impossible to commit.
    pub fn verify_candidate_with_conditional_override(
        &self,
        current: &Workbook,
        context: &ExecutionContextFingerprint,
    ) -> Result<PlanCommit, PlanError> {
        self.verify_candidate_inner(current, context, true)
    }

    fn verify_candidate_inner(
        &self,
        current: &Workbook,
        context: &ExecutionContextFingerprint,
        allow_conditional: bool,
    ) -> Result<PlanCommit, PlanError> {
        if current.revision() != self.plan.source_revision {
            return Err(PlanError::RevisionMismatch {
                expected: self.plan.source_revision,
                actual: current.revision(),
            });
        }
        if workbook_fingerprint(current) != self.plan.source_fingerprint {
            return Err(PlanError::PreviewFingerprintMismatch);
        }
        if context != &self.plan.execution_context {
            return Err(PlanError::ContextChanged);
        }
        if self
            .plan
            .problems
            .iter()
            .any(|problem| problem.severity == ProblemSeverity::Blocking)
        {
            return Err(PlanError::BlockingProblems);
        }
        match self.plan.determinism {
            DeterminismClass::Unresolved => return Err(PlanError::Unresolved),
            DeterminismClass::Conditional if !allow_conditional => {
                return Err(PlanError::ConditionalRequiresOverride)
            }
            DeterminismClass::Conditional => {}
            DeterminismClass::Full => {}
        }
        let expected_hash = compute_plan_hash(
            self.plan.contract_version,
            &self.plan.workbook_session_id,
            self.plan.source_sheet_id,
            self.plan.source_revision,
            &self.plan.execution_context,
            &self.plan.operations,
        );
        if expected_hash != self.plan.plan_hash {
            return Err(PlanError::PlanHashMismatch);
        }

        let mut candidate = current.clone();
        let _ = apply_operations(
            &mut candidate,
            self.plan.source_sheet_id,
            &self.plan.operations,
        )?;
        if self.plan.determinism == DeterminismClass::Full
            && workbook_fingerprint(&candidate) != self.plan.preview_fingerprint
        {
            return Err(PlanError::PreviewFingerprintMismatch);
        }
        candidate.set_revision_after_atomic_commit(self.plan.source_revision.saturating_add(1));

        let affected_cells = self
            .plan
            .changes
            .iter()
            .filter_map(|change| {
                change
                    .after_coordinate
                    .or(change.before_coordinate)
                    .map(|coordinate| CellId::new(change.sheet_id, coordinate.row, coordinate.col))
            })
            .collect();
        Ok(PlanCommit {
            plan_id: self.plan.id.clone(),
            source: current.clone(),
            applied: candidate,
            affected_cells,
        })
    }
}

fn normalize_operations(
    sheet: &Sheet,
    operations: Vec<PlannedOp>,
) -> Result<Vec<PlannedOp>, PlanError> {
    if operations.len() > MAX_PLAN_OPERATIONS {
        return Err(PlanError::OperationLimitExceeded {
            limit: MAX_PLAN_OPERATIONS,
        });
    }

    let mut cell_writes: BTreeMap<CellCoordinate, PlannedOp> = BTreeMap::new();
    let mut styles = Vec::new();
    let mut deletes = Vec::new();
    let mut expanded_cell_touches = 0usize;
    for operation in operations {
        match operation {
            PlannedOp::SetCellValue { coordinate, value } => {
                expanded_cell_touches = expanded_cell_touches.saturating_add(1);
                validate_coordinate(sheet, coordinate)?;
                if matches!(value, PlannedCellValue::Number(number) if !number.is_finite()) {
                    return Err(PlanError::InvalidOperation(
                        "cell numbers must be finite".into(),
                    ));
                }
                cell_writes.insert(coordinate, PlannedOp::SetCellValue { coordinate, value });
            }
            PlannedOp::SetCellFormula {
                coordinate,
                formula,
            } => {
                expanded_cell_touches = expanded_cell_touches.saturating_add(1);
                validate_coordinate(sheet, coordinate)?;
                if !formula.starts_with('=') {
                    return Err(PlanError::InvalidOperation(
                        "formulas must start with '='".into(),
                    ));
                }
                cell_writes.insert(
                    coordinate,
                    PlannedOp::SetCellFormula {
                        coordinate,
                        formula,
                    },
                );
            }
            PlannedOp::ClearCell { coordinate } => {
                expanded_cell_touches = expanded_cell_touches.saturating_add(1);
                validate_coordinate(sheet, coordinate)?;
                cell_writes.insert(coordinate, PlannedOp::ClearCell { coordinate });
            }
            PlannedOp::ClearRange { range } => {
                let range = range.normalized();
                validate_coordinate(sheet, range.start)?;
                validate_coordinate(sheet, range.end)?;
                let count = (range.end.row - range.start.row + 1)
                    .saturating_mul(range.end.col - range.start.col + 1);
                expanded_cell_touches = expanded_cell_touches.saturating_add(count);
                if expanded_cell_touches > MAX_PLAN_OPERATIONS {
                    return Err(PlanError::OperationLimitExceeded {
                        limit: MAX_PLAN_OPERATIONS,
                    });
                }
                for row in range.start.row..=range.end.row {
                    for col in range.start.col..=range.end.col {
                        let coordinate = CellCoordinate { row, col };
                        cell_writes.insert(coordinate, PlannedOp::ClearCell { coordinate });
                    }
                }
            }
            PlannedOp::SetCellStyle { range, style } => {
                let range = range.normalized();
                validate_coordinate(sheet, range.start)?;
                validate_coordinate(sheet, range.end)?;
                let count = (range.end.row - range.start.row + 1)
                    .saturating_mul(range.end.col - range.start.col + 1);
                expanded_cell_touches = expanded_cell_touches.saturating_add(count);
                if expanded_cell_touches > MAX_PLAN_OPERATIONS {
                    return Err(PlanError::OperationLimitExceeded {
                        limit: MAX_PLAN_OPERATIONS,
                    });
                }
                let changes_any_cell = (range.start.row..=range.end.row).any(|row| {
                    (range.start.col..=range.end.col)
                        .any(|col| sheet.get_format(row, col).cell_style != style)
                });
                if changes_any_cell {
                    styles.push(PlannedOp::SetCellStyle { range, style });
                }
            }
            PlannedOp::DeleteRows { at, count } => {
                if count == 0 || at >= sheet.rows || at.saturating_add(count) > sheet.rows {
                    return Err(PlanError::InvalidOperation(
                        "row deletion is outside the sheet".into(),
                    ));
                }
                deletes.push((at, count));
            }
            PlannedOp::InsertRows { .. } => {
                return Err(PlanError::UnsupportedOperation(
                    "row insertion is Phase 1.1".into(),
                ))
            }
            PlannedOp::InsertColumns { .. } | PlannedOp::DeleteColumns { .. } => {
                return Err(PlanError::UnsupportedOperation(
                    "column structural review is Phase 1.1".into(),
                ));
            }
        }
        if expanded_cell_touches > MAX_PLAN_OPERATIONS {
            return Err(PlanError::OperationLimitExceeded {
                limit: MAX_PLAN_OPERATIONS,
            });
        }
    }

    deletes.sort_unstable_by_key(|(at, _)| *at);
    for pair in deletes.windows(2) {
        if pair[0].0 + pair[0].1 > pair[1].0 {
            return Err(PlanError::InvalidOperation("row deletions overlap".into()));
        }
    }
    let deleted = |row: usize| {
        deletes
            .iter()
            .any(|(at, count)| row >= *at && row < at + count)
    };
    cell_writes.retain(|coordinate, operation| {
        if deleted(coordinate.row) {
            return false;
        }
        let before = CellSnapshot::from_sheet(sheet, coordinate.row, coordinate.col);
        match operation {
            PlannedOp::SetCellValue { value, .. } => before.raw != value.as_input(),
            PlannedOp::SetCellFormula { formula, .. } => before.raw != *formula,
            PlannedOp::ClearCell { .. } => before != CellSnapshot::empty(),
            _ => true,
        }
    });

    let mut normalized: Vec<_> = cell_writes.into_values().collect();
    normalized.extend(styles);
    deletes.sort_unstable_by(|left, right| right.0.cmp(&left.0));
    normalized.extend(
        deletes
            .into_iter()
            .map(|(at, count)| PlannedOp::DeleteRows { at, count }),
    );
    Ok(normalized)
}

fn validate_coordinate(sheet: &Sheet, coordinate: CellCoordinate) -> Result<(), PlanError> {
    if coordinate.row >= sheet.rows || coordinate.col >= sheet.cols {
        Err(PlanError::InvalidOperation(format!(
            "cell ({}, {}) is outside the sheet",
            coordinate.row, coordinate.col
        )))
    } else {
        Ok(())
    }
}

fn apply_operations(
    workbook: &mut Workbook,
    sheet_id: SheetId,
    operations: &[PlannedOp],
) -> Result<crate::recalc::RecalcReport, PlanError> {
    let sheet_index = workbook
        .sheet_index_by_id(sheet_id)
        .ok_or(PlanError::SourceSheetMissing)?;
    workbook.begin_batch();
    for operation in operations {
        match operation {
            PlannedOp::SetCellValue { coordinate, value } => {
                workbook.set_cell_value_tracked(
                    sheet_index,
                    coordinate.row,
                    coordinate.col,
                    &value.as_input(),
                );
            }
            PlannedOp::SetCellFormula {
                coordinate,
                formula,
            } => {
                workbook.set_cell_value_tracked(
                    sheet_index,
                    coordinate.row,
                    coordinate.col,
                    formula,
                );
            }
            PlannedOp::ClearCell { coordinate } => {
                workbook.clear_cell_tracked(sheet_index, coordinate.row, coordinate.col);
            }
            PlannedOp::SetCellStyle { range, style } => {
                let mut changed = Vec::new();
                if let Some(sheet) = workbook.sheet_mut(sheet_index) {
                    for row in range.start.row..=range.end.row {
                        for col in range.start.col..=range.end.col {
                            if sheet.get_format(row, col).cell_style != *style {
                                sheet.set_cell_style(row, col, *style);
                                changed.push(CellId::new(sheet_id, row, col));
                            }
                        }
                    }
                }
                for cell in changed {
                    workbook.note_format_changed(cell);
                }
            }
            PlannedOp::DeleteRows { .. } => {}
            PlannedOp::ClearRange { .. }
            | PlannedOp::InsertRows { .. }
            | PlannedOp::InsertColumns { .. }
            | PlannedOp::DeleteColumns { .. } => {
                workbook.end_batch();
                return Err(PlanError::InvalidOperation(
                    "plan was not normalized".into(),
                ));
            }
        }
    }
    workbook.end_batch();

    for operation in operations {
        if let PlannedOp::DeleteRows { at, count } = operation {
            workbook
                .structural_edit(sheet_index, Axis::Row, *at, *count, true)
                .map_err(PlanError::InvalidOperation)?;
        }
    }
    Ok(workbook.recompute_full_ordered())
}

fn build_row_lineage(row_count: usize, operations: &[PlannedOp]) -> Vec<ReviewRowLineage> {
    let deletes: Vec<(usize, usize)> = operations
        .iter()
        .filter_map(|operation| match operation {
            PlannedOp::DeleteRows { at, count } => Some((*at, *count)),
            _ => None,
        })
        .collect();
    let changed_rows: HashSet<usize> = operations
        .iter()
        .filter_map(|operation| match operation {
            PlannedOp::SetCellValue { coordinate, .. }
            | PlannedOp::SetCellFormula { coordinate, .. }
            | PlannedOp::ClearCell { coordinate } => Some(coordinate.row),
            PlannedOp::SetCellStyle { range, .. } => Some(range.start.row),
            _ => None,
        })
        .collect();

    (0..row_count)
        .map(|row| {
            let is_deleted = deletes
                .iter()
                .any(|(at, count)| row >= *at && row < at + count);
            let deleted_before = deletes
                .iter()
                .map(|(at, count)| if row >= at + count { *count } else { 0 })
                .sum::<usize>();
            ReviewRowLineage {
                id: ReviewRowId(row as u64 + 1),
                before_data_row: Some(row),
                after_data_row: (!is_deleted).then_some(row - deleted_before),
                display_position: row,
                state: if is_deleted {
                    ReviewRowState::Deleted
                } else if changed_rows.contains(&row) {
                    ReviewRowState::Changed
                } else {
                    ReviewRowState::Unchanged
                },
            }
        })
        .collect()
}

fn materialize_changes(
    before: &Workbook,
    after: &Workbook,
    source_sheet_id: SheetId,
    operations: &[PlannedOp],
    lineage: &[ReviewRowLineage],
) -> Result<Vec<MaterializedChange>, PlanError> {
    let direct_cells: BTreeSet<CellCoordinate> = operations
        .iter()
        .flat_map(|operation| match operation {
            PlannedOp::SetCellValue { coordinate, .. }
            | PlannedOp::SetCellFormula { coordinate, .. }
            | PlannedOp::ClearCell { coordinate } => vec![*coordinate],
            PlannedOp::SetCellStyle { range, .. } => {
                let mut cells = Vec::new();
                for row in range.start.row..=range.end.row {
                    for col in range.start.col..=range.end.col {
                        cells.push(CellCoordinate { row, col });
                    }
                }
                cells
            }
            _ => Vec::new(),
        })
        .collect();
    let mut changes = Vec::new();
    for before_sheet in before.sheets() {
        let sheet_id = before_sheet.id;
        let Some(after_sheet) = after.sheet_by_id(sheet_id) else {
            continue;
        };
        if sheet_id == source_sheet_id {
            for row in lineage {
                let before_row = row
                    .before_data_row
                    .expect("source lineage always has before row");
                if row.state == ReviewRowState::Deleted {
                    push_materialized_change(
                        &mut changes,
                        MaterializedChange {
                            sheet_id,
                            row_id: row.id,
                            before_coordinate: Some(CellCoordinate {
                                row: before_row,
                                col: 0,
                            }),
                            after_coordinate: None,
                            kind: ChangeKind::RowDeleted,
                            cause: ChangeCause::Direct,
                            before: CellSnapshot::empty(),
                            after: CellSnapshot::empty(),
                            group_id: None,
                            reason: None,
                            sources: Vec::new(),
                        },
                    )?;
                    let mut coordinates: Vec<_> = before_sheet
                        .cells_iter()
                        .filter(|((cell_row, _), _)| *cell_row == before_row)
                        .map(|(&(cell_row, col), _)| CellCoordinate { row: cell_row, col })
                        .collect();
                    coordinates.sort_unstable();
                    for coordinate in coordinates {
                        let old =
                            CellSnapshot::from_sheet(before_sheet, coordinate.row, coordinate.col);
                        if old != CellSnapshot::empty() {
                            push_materialized_change(
                                &mut changes,
                                MaterializedChange {
                                    sheet_id,
                                    row_id: row.id,
                                    before_coordinate: Some(coordinate),
                                    after_coordinate: None,
                                    kind: ChangeKind::Cleared,
                                    cause: ChangeCause::Direct,
                                    before: old,
                                    after: CellSnapshot::empty(),
                                    group_id: None,
                                    reason: None,
                                    sources: Vec::new(),
                                },
                            )?;
                        }
                    }
                    continue;
                }
                let Some(after_row) = row.after_data_row else {
                    continue;
                };
                let mut cols = BTreeSet::new();
                cols.extend(
                    before_sheet
                        .cells_iter()
                        .filter(|((r, _), _)| *r == before_row)
                        .map(|((_, c), _)| *c),
                );
                cols.extend(
                    after_sheet
                        .cells_iter()
                        .filter(|((r, _), _)| *r == after_row)
                        .map(|((_, c), _)| *c),
                );
                for col in cols {
                    let before_coordinate = CellCoordinate {
                        row: before_row,
                        col,
                    };
                    let after_coordinate = CellCoordinate {
                        row: after_row,
                        col,
                    };
                    push_cell_change(
                        &mut changes,
                        sheet_id,
                        row.id,
                        before_coordinate,
                        after_coordinate,
                        CellSnapshot::from_sheet(before_sheet, before_row, col),
                        CellSnapshot::from_sheet(after_sheet, after_row, col),
                        direct_cells.contains(&before_coordinate),
                    )?;
                }
            }
        } else {
            let mut coordinates = BTreeSet::new();
            coordinates.extend(
                before_sheet
                    .cells_iter()
                    .map(|(&(row, col), _)| CellCoordinate { row, col }),
            );
            coordinates.extend(
                after_sheet
                    .cells_iter()
                    .map(|(&(row, col), _)| CellCoordinate { row, col }),
            );
            for coordinate in coordinates {
                let before = CellSnapshot::from_sheet(before_sheet, coordinate.row, coordinate.col);
                let after = CellSnapshot::from_sheet(after_sheet, coordinate.row, coordinate.col);
                if before != after {
                    return Err(PlanError::MultiSheetUnsupported);
                }
            }
        }
    }
    changes.sort_by_key(|change| {
        let coordinate = change
            .after_coordinate
            .or(change.before_coordinate)
            .unwrap_or(CellCoordinate { row: 0, col: 0 });
        (
            change.sheet_id.raw(),
            coordinate.row,
            coordinate.col,
            change.kind as u8,
        )
    });
    Ok(changes)
}

fn push_materialized_change(
    changes: &mut Vec<MaterializedChange>,
    change: MaterializedChange,
) -> Result<(), PlanError> {
    ensure_materialized_change_capacity(changes.len())?;
    changes.push(change);
    Ok(())
}

fn ensure_materialized_change_capacity(current_len: usize) -> Result<(), PlanError> {
    if current_len >= MAX_MATERIALIZED_CHANGES {
        return Err(PlanError::MaterializedChangeLimitExceeded {
            limit: MAX_MATERIALIZED_CHANGES,
        });
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_cell_change(
    changes: &mut Vec<MaterializedChange>,
    sheet_id: SheetId,
    row_id: ReviewRowId,
    before_coordinate: CellCoordinate,
    after_coordinate: CellCoordinate,
    before: CellSnapshot,
    after: CellSnapshot,
    direct: bool,
) -> Result<(), PlanError> {
    if before == after {
        return Ok(());
    }
    let kind = if before.raw != after.raw {
        if after.raw.is_empty() {
            ChangeKind::Cleared
        } else if after.raw.starts_with('=') {
            ChangeKind::Formula
        } else {
            ChangeKind::Value
        }
    } else if before.format != after.format {
        ChangeKind::Format
    } else {
        ChangeKind::Value
    };
    push_materialized_change(
        changes,
        MaterializedChange {
            sheet_id,
            row_id,
            before_coordinate: Some(before_coordinate),
            after_coordinate: Some(after_coordinate),
            kind,
            cause: if direct {
                ChangeCause::Direct
            } else {
                ChangeCause::Recalculated
            },
            before,
            after,
            group_id: None,
            reason: None,
            sources: Vec::new(),
        },
    )
}

fn affected_ranges(changes: &[MaterializedChange]) -> Vec<AffectedRange> {
    let mut bounds: BTreeMap<u64, (SheetId, usize, usize, usize, usize)> = BTreeMap::new();
    for change in changes {
        let Some(coordinate) = change.after_coordinate.or(change.before_coordinate) else {
            continue;
        };
        bounds
            .entry(change.sheet_id.raw())
            .and_modify(|(_, min_row, min_col, max_row, max_col)| {
                *min_row = (*min_row).min(coordinate.row);
                *min_col = (*min_col).min(coordinate.col);
                *max_row = (*max_row).max(coordinate.row);
                *max_col = (*max_col).max(coordinate.col);
            })
            .or_insert((
                change.sheet_id,
                coordinate.row,
                coordinate.col,
                coordinate.row,
                coordinate.col,
            ));
    }
    bounds
        .into_values()
        .map(
            |(sheet_id, min_row, min_col, max_row, max_col)| AffectedRange {
                sheet_id,
                range: CellRange {
                    start: CellCoordinate {
                        row: min_row,
                        col: min_col,
                    },
                    end: CellCoordinate {
                        row: max_row,
                        col: max_col,
                    },
                },
            },
        )
        .collect()
}

fn summarize(changes: &[MaterializedChange]) -> PlanSummary {
    let mut summary = PlanSummary::default();
    for change in changes {
        match (change.cause, change.kind) {
            (_, ChangeKind::RowDeleted) => summary.rows_deleted += 1,
            (ChangeCause::Recalculated, _) => summary.recalculated_cells += 1,
            (_, ChangeKind::Cleared) => summary.cells_cleared += 1,
            (_, ChangeKind::Formula) => summary.formulas_changed += 1,
            (_, ChangeKind::Format) => summary.formatting_changes += 1,
            (_, ChangeKind::Value) => summary.cells_changed += 1,
        }
    }
    summary
}

fn is_formula_error(display: &str) -> bool {
    display.starts_with('#')
}

#[derive(Serialize)]
struct PlanHashInput<'a> {
    contract_version: u32,
    workbook_session_id: &'a str,
    source_sheet_id: SheetId,
    source_revision: u64,
    execution_context: &'a ExecutionContextFingerprint,
    operations: &'a [PlannedOp],
}

fn compute_plan_hash(
    contract_version: u32,
    workbook_session_id: &str,
    source_sheet_id: SheetId,
    source_revision: u64,
    execution_context: &ExecutionContextFingerprint,
    operations: &[PlannedOp],
) -> String {
    hash_serializable(&PlanHashInput {
        contract_version,
        workbook_session_id,
        source_sheet_id,
        source_revision,
        execution_context,
        operations,
    })
}

pub fn workbook_fingerprint(workbook: &Workbook) -> String {
    let mut hasher = Sha256::new();
    let mut named_ranges = workbook.named_ranges().list();
    named_ranges.sort_by(|left, right| {
        left.name
            .to_ascii_lowercase()
            .cmp(&right.name.to_ascii_lowercase())
    });
    let named_ranges = serde_json::to_vec(&named_ranges).expect("named ranges serialize");
    hasher.update((named_ranges.len() as u64).to_le_bytes());
    hasher.update(named_ranges);
    for sheet in workbook.sheets() {
        hasher.update(sheet.id.raw().to_le_bytes());
        hasher.update((sheet.name.len() as u64).to_le_bytes());
        hasher.update(sheet.name.as_bytes());
        hasher.update((sheet.rows as u64).to_le_bytes());
        hasher.update((sheet.cols as u64).to_le_bytes());
        let mut coordinates: Vec<_> = sheet
            .cells_iter()
            .map(|(&(row, col), _)| (row, col))
            .collect();
        coordinates.sort_unstable();
        for (row, col) in coordinates {
            hasher.update((row as u64).to_le_bytes());
            hasher.update((col as u64).to_le_bytes());
            let snapshot = CellSnapshot::from_sheet(sheet, row, col);
            let encoded = serde_json::to_vec(&snapshot).expect("cell snapshots serialize");
            hasher.update((encoded.len() as u64).to_le_bytes());
            hasher.update(encoded);
        }
    }
    format!("v1:{:x}", hasher.finalize())
}

fn hash_serializable(value: &impl Serialize) -> String {
    let encoded = serde_json::to_vec(value).expect("plan hash inputs serialize");
    format!("v1:{:x}", Sha256::digest(encoded))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn context() -> ExecutionContextFingerprint {
        ExecutionContextFingerprint {
            engine_version: env!("CARGO_PKG_VERSION").into(),
            functions_generation: 1,
            functions_source_hash: "none".into(),
            locale: "en-US".into(),
            timezone: "UTC".into(),
            auto_recalc: true,
            iterative_calculation_enabled: false,
            iterative_max_iterations: 100,
            iterative_tolerance_bits: 0.001_f64.to_bits(),
            volatile_inputs: Vec::new(),
        }
    }

    fn request(workbook: &Workbook, operations: Vec<PlannedOp>) -> OperationPlanRequest {
        OperationPlanRequest {
            id: PlanId("pv_test".into()),
            workbook_session_id: "session-test".into(),
            source_sheet_id: workbook.active_sheet_id(),
            expected_revision: workbook.revision(),
            execution_context: context(),
            producer: PlanProducer {
                kind: "test".into(),
                name: "Test".into(),
                source_path: None,
                source_hash: None,
            },
            title: "Test plan".into(),
            description: None,
            operations,
            groups: Vec::new(),
            verification: vec![VerificationDefinition::NoNewFormulaErrors {
                id: "errors".into(),
                label: None,
            }],
        }
    }

    #[test]
    fn normalizes_repeated_writes_and_removes_noops() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 0, 0, "old");
        let coordinate = CellCoordinate { row: 0, col: 0 };
        let prepared = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![
                    PlannedOp::SetCellValue {
                        coordinate,
                        value: PlannedCellValue::Text("first".into()),
                    },
                    PlannedOp::SetCellValue {
                        coordinate,
                        value: PlannedCellValue::Text("old".into()),
                    },
                ],
            ),
        )
        .unwrap();
        assert!(prepared.plan.operations.is_empty());
        assert!(prepared.plan.changes.is_empty());
    }

    #[test]
    fn materializes_direct_and_recalculated_changes_on_a_clone() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 0, 0, "10");
        workbook.set_cell_value_tracked(0, 0, 1, "=A1*2");
        let source_revision = workbook.revision();
        let prepared = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![PlannedOp::SetCellValue {
                    coordinate: CellCoordinate { row: 0, col: 0 },
                    value: PlannedCellValue::Number(20.0),
                }],
            ),
        )
        .unwrap();
        assert_eq!(
            workbook.active_sheet().get_display(0, 0),
            "10",
            "source stays inert"
        );
        assert_eq!(workbook.revision(), source_revision);
        assert_eq!(
            prepared.preview_workbook().active_sheet().get_display(0, 1),
            "40"
        );
        assert!(prepared.plan.changes.iter().any(|change| {
            change.after_coordinate == Some(CellCoordinate { row: 0, col: 0 })
                && change.cause == ChangeCause::Direct
        }));
        assert!(prepared.plan.changes.iter().any(|change| {
            change.after_coordinate == Some(CellCoordinate { row: 0, col: 1 })
                && change.cause == ChangeCause::Recalculated
        }));
    }

    #[test]
    fn row_deletion_keeps_tombstone_lineage() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 0, 0, "header");
        workbook.set_cell_value_tracked(0, 1, 0, "delete me");
        workbook.set_cell_value_tracked(0, 2, 0, "keep me");
        let prepared = PreparedOperationPlan::materialize(
            &workbook,
            request(&workbook, vec![PlannedOp::DeleteRows { at: 1, count: 1 }]),
        )
        .unwrap();
        let deleted = &prepared.plan.row_lineage[1];
        assert_eq!(deleted.state, ReviewRowState::Deleted);
        assert_eq!(deleted.before_data_row, Some(1));
        assert_eq!(deleted.after_data_row, None);
        assert_eq!(prepared.plan.row_lineage[2].after_data_row, Some(1));
        assert_eq!(
            prepared.preview_workbook().active_sheet().get_display(1, 0),
            "keep me"
        );
        assert!(prepared
            .plan
            .changes
            .iter()
            .any(|change| change.kind == ChangeKind::RowDeleted));
    }

    #[test]
    fn candidate_commit_is_atomic_and_snapshot_undoable() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 0, 0, "before");
        let prepared = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![PlannedOp::SetCellValue {
                    coordinate: CellCoordinate { row: 0, col: 0 },
                    value: PlannedCellValue::Text("after".into()),
                }],
            ),
        )
        .unwrap();
        let commit = prepared.verify_candidate(&workbook, &context()).unwrap();
        assert_eq!(workbook.active_sheet().get_display(0, 0), "before");
        assert_eq!(commit.applied.revision(), workbook.revision() + 1);
        commit.redo_into(&mut workbook);
        assert_eq!(workbook.active_sheet().get_display(0, 0), "after");
        commit.undo_into(&mut workbook);
        assert_eq!(workbook.active_sheet().get_display(0, 0), "before");
    }

    #[test]
    fn stale_revision_and_context_are_rejected() {
        let mut workbook = Workbook::new();
        let prepared = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![PlannedOp::ClearCell {
                    coordinate: CellCoordinate { row: 0, col: 0 },
                }],
            ),
        )
        .unwrap();
        workbook.set_cell_value_tracked(0, 0, 0, "drift");
        assert!(matches!(
            prepared.verify_candidate(&workbook, &context()),
            Err(PlanError::RevisionMismatch { .. })
        ));

        let workbook = prepared.source_workbook().clone();
        let mut changed_context = context();
        changed_context.iterative_tolerance_bits = 0.01_f64.to_bits();
        assert_eq!(
            prepared
                .verify_candidate(&workbook, &changed_context)
                .unwrap_err(),
            PlanError::ContextChanged
        );
    }

    #[test]
    fn wire_plan_is_versioned_and_round_trips() {
        let workbook = Workbook::new();
        let prepared = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![PlannedOp::SetCellValue {
                    coordinate: CellCoordinate { row: 0, col: 0 },
                    value: PlannedCellValue::Text("reviewed".into()),
                }],
            ),
        )
        .unwrap();

        let encoded = serde_json::to_string(prepared.plan()).unwrap();
        let decoded: OperationPlan = serde_json::from_str(&encoded).unwrap();
        assert_eq!(decoded.contract_version, OPERATION_PLAN_CONTRACT_VERSION);
        assert_eq!(decoded, *prepared.plan());
    }

    #[test]
    fn altered_normalized_operations_fail_hash_verification() {
        let workbook = Workbook::new();
        let mut prepared = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![PlannedOp::SetCellValue {
                    coordinate: CellCoordinate { row: 0, col: 0 },
                    value: PlannedCellValue::Text("reviewed".into()),
                }],
            ),
        )
        .unwrap();
        prepared.plan.operations[0] = PlannedOp::SetCellValue {
            coordinate: CellCoordinate { row: 0, col: 0 },
            value: PlannedCellValue::Text("tampered".into()),
        };

        assert_eq!(
            prepared
                .verify_candidate(&workbook, &context())
                .unwrap_err(),
            PlanError::PlanHashMismatch
        );
    }

    #[test]
    fn formatting_plans_drop_noops_and_bound_materialization() {
        let workbook = Workbook::new();
        let no_op = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![PlannedOp::SetCellStyle {
                    range: CellRange::single(0, 0),
                    style: CellStyle::default(),
                }],
            ),
        )
        .unwrap();
        assert!(no_op.plan().operations.is_empty());

        let oversized = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![PlannedOp::SetCellStyle {
                    range: CellRange {
                        start: CellCoordinate { row: 0, col: 0 },
                        end: CellCoordinate { row: 400, col: 255 },
                    },
                    style: CellStyle::from_int(1),
                }],
            ),
        );
        assert!(matches!(
            oversized,
            Err(PlanError::OperationLimitExceeded {
                limit: MAX_PLAN_OPERATIONS
            })
        ));
    }

    #[test]
    fn v1_plan_and_materialized_change_limits_are_enforced() {
        let workbook = Workbook::new();
        let repeated_clear = PlannedOp::ClearCell {
            coordinate: CellCoordinate { row: 0, col: 0 },
        };
        let too_many_operations = PreparedOperationPlan::materialize(
            &workbook,
            request(&workbook, vec![repeated_clear; MAX_PLAN_OPERATIONS + 1]),
        );
        assert!(matches!(
            too_many_operations,
            Err(PlanError::OperationLimitExceeded {
                limit: MAX_PLAN_OPERATIONS
            })
        ));
        assert!(ensure_materialized_change_capacity(MAX_MATERIALIZED_CHANGES - 1).is_ok());
        assert_eq!(
            ensure_materialized_change_capacity(MAX_MATERIALIZED_CHANGES).unwrap_err(),
            PlanError::MaterializedChangeLimitExceeded {
                limit: MAX_MATERIALIZED_CHANGES
            }
        );
    }

    #[test]
    fn v1_rejects_cross_sheet_recalculation_effects() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 0, 0, "10");
        let dependent_sheet = workbook.add_sheet();
        workbook.set_cell_value_tracked(dependent_sheet, 0, 0, "=Sheet1!A1*2");
        let _ = workbook.set_active_sheet(0);

        let result = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![PlannedOp::SetCellValue {
                    coordinate: CellCoordinate { row: 0, col: 0 },
                    value: PlannedCellValue::Number(20.0),
                }],
            ),
        );
        assert!(matches!(result, Err(PlanError::MultiSheetUnsupported)));
    }
}
