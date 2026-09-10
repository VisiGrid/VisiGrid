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
pub const MAX_DELETE_ROWS_PER_OPERATION: usize = 1_000;
pub const MAX_ROWS_DELETED: usize = 65_536;
pub const MAX_VERIFICATION_DEFINITIONS: usize = 8;
pub const MAX_PLAN_GROUPS: usize = 256;
pub const MAX_SOURCES_PER_OPERATION: usize = 16;
pub const MAX_METADATA_TEXT_BYTES: usize = 500;
pub const MAX_ID_BYTES: usize = 64;
pub const MAX_LABEL_BYTES: usize = 120;
pub const MAX_SOURCE_REFERENCE_BYTES: usize = 64;

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

#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct OperationMetadata {
    pub group_id: Option<GroupId>,
    pub reason: Option<ProducerClaim>,
    #[serde(default)]
    pub sources: Vec<String>,
}

/// A primitive mutation paired with optional, untrusted review metadata.
/// Only `operation` contributes to execution authority; metadata is carried
/// into materialized changes for explanation and cannot alter the mutation.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PlannedOperation {
    pub operation: PlannedOp,
    #[serde(default)]
    pub metadata: OperationMetadata,
}

impl PlannedOperation {
    pub fn plain(operation: PlannedOp) -> Self {
        Self {
            operation,
            metadata: OperationMetadata::default(),
        }
    }
}

impl From<PlannedOp> for PlannedOperation {
    fn from(operation: PlannedOp) -> Self {
        Self::plain(operation)
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum VerificationDefinition {
    NoNewFormulaErrors {
        id: String,
        label: Option<String>,
    },
    #[serde(rename = "gross_minus_group_equals_preview")]
    RetainedTotal {
        id: String,
        label: Option<String>,
        source_range: CellRange,
        amount_column: usize,
        excluded_group: GroupId,
        tolerance: f64,
        currency: String,
    },
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
    #[serde(flatten)]
    pub evidence: VerificationEvidence,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum VerificationEvidence {
    NoNewFormulaErrors {
        new_errors: usize,
    },
    #[serde(rename = "gross_minus_group_equals_preview")]
    RetainedTotal {
        expected: String,
        actual: String,
        tolerance: String,
        currency: String,
        gross_source: String,
        classified_exclusions: String,
        difference: String,
    },
    Unavailable {
        message: String,
    },
}

impl VerificationEvidence {
    fn summary(&self) -> String {
        match self {
            Self::NoNewFormulaErrors { new_errors } => {
                format!("{new_errors} new formula error(s)")
            }
            Self::RetainedTotal {
                expected,
                actual,
                tolerance,
                currency,
                ..
            } => format!(
                "expected={expected} {currency}; actual={actual} {currency}; tolerance={tolerance}"
            ),
            Self::Unavailable { message } => message.clone(),
        }
    }
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
    pub operations: Vec<PlannedOperation>,
    pub groups: Vec<OperationGroup>,
    pub changes: Vec<MaterializedChange>,
    pub affected_ranges: Vec<AffectedRange>,
    pub verification_definitions: Vec<VerificationDefinition>,
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
    pub operations: Vec<PlannedOperation>,
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
    pub verification: Vec<VerificationResult>,
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
    CandidateVerificationFailed,
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
            Self::CandidateVerificationFailed => {
                write!(f, "final candidate failed verification")
            }
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
        validate_plan_metadata(source_sheet, &request.operations, &request.groups)?;
        validate_verification_definitions(source_sheet, &request.verification, &request.groups)?;
        let normalized = normalize_operations(source_sheet, request.operations)?;
        let operations = normalized.operations;
        let source_fingerprint = workbook_fingerprint(source);
        let verification_definitions = request.verification.clone();
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
        let mut problems = normalized.problems;
        let (verification, verification_problems) = evaluate_verifications(
            source,
            &preview,
            request.source_sheet_id,
            &row_lineage,
            &operations,
            &verification_definitions,
            &changes,
        )?;
        problems.extend(verification_problems);
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
            verification_definitions,
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
        let report = apply_operations(
            &mut candidate,
            self.plan.source_sheet_id,
            &self.plan.operations,
        )?;
        if !report.errors.is_empty() || (report.scc_count > 0 && !report.converged) {
            return Err(PlanError::Unresolved);
        }
        let source_sheet = current
            .sheet_by_id(self.plan.source_sheet_id)
            .ok_or(PlanError::SourceSheetMissing)?;
        let row_lineage = build_row_lineage(source_sheet.rows, &self.plan.operations);
        let changes = materialize_changes(
            current,
            &candidate,
            self.plan.source_sheet_id,
            &self.plan.operations,
            &row_lineage,
        )?;
        let (verification, verification_problems) = evaluate_verifications(
            current,
            &candidate,
            self.plan.source_sheet_id,
            &row_lineage,
            &self.plan.operations,
            &self.plan.verification_definitions,
            &changes,
        )?;
        if verification_problems
            .iter()
            .any(|problem| problem.severity == ProblemSeverity::Blocking)
        {
            return Err(PlanError::CandidateVerificationFailed);
        }
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
            verification,
        })
    }
}

struct NormalizationResult {
    operations: Vec<PlannedOperation>,
    problems: Vec<PlanProblem>,
}

fn normalize_operations(
    sheet: &Sheet,
    operations: Vec<PlannedOperation>,
) -> Result<NormalizationResult, PlanError> {
    if operations.len() > MAX_PLAN_OPERATIONS {
        return Err(PlanError::OperationLimitExceeded {
            limit: MAX_PLAN_OPERATIONS,
        });
    }

    let mut cell_writes: BTreeMap<CellCoordinate, PlannedOperation> = BTreeMap::new();
    let mut styles = Vec::new();
    let mut deletes = Vec::new();
    let mut rows_deleted = 0usize;
    let mut expanded_cell_touches = 0usize;
    for planned in operations {
        let PlannedOperation {
            operation,
            metadata,
        } = planned;
        match operation {
            PlannedOp::SetCellValue { coordinate, value } => {
                expanded_cell_touches = expanded_cell_touches.saturating_add(1);
                validate_coordinate(sheet, coordinate)?;
                if matches!(value, PlannedCellValue::Number(number) if !number.is_finite()) {
                    return Err(PlanError::InvalidOperation(
                        "cell numbers must be finite".into(),
                    ));
                }
                cell_writes.insert(
                    coordinate,
                    PlannedOperation {
                        operation: PlannedOp::SetCellValue { coordinate, value },
                        metadata,
                    },
                );
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
                    PlannedOperation {
                        operation: PlannedOp::SetCellFormula {
                            coordinate,
                            formula,
                        },
                        metadata,
                    },
                );
            }
            PlannedOp::ClearCell { coordinate } => {
                expanded_cell_touches = expanded_cell_touches.saturating_add(1);
                validate_coordinate(sheet, coordinate)?;
                cell_writes.insert(
                    coordinate,
                    PlannedOperation {
                        operation: PlannedOp::ClearCell { coordinate },
                        metadata,
                    },
                );
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
                        cell_writes.insert(
                            coordinate,
                            PlannedOperation {
                                operation: PlannedOp::ClearCell { coordinate },
                                metadata: metadata.clone(),
                            },
                        );
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
                    styles.push(PlannedOperation {
                        operation: PlannedOp::SetCellStyle { range, style },
                        metadata,
                    });
                }
            }
            PlannedOp::DeleteRows { at, count } => {
                if count == 0 || at >= sheet.rows || at.saturating_add(count) > sheet.rows {
                    return Err(PlanError::InvalidOperation(
                        "row deletion is outside the sheet".into(),
                    ));
                }
                if count > MAX_DELETE_ROWS_PER_OPERATION {
                    return Err(PlanError::InvalidOperation(format!(
                        "row deletion exceeds {MAX_DELETE_ROWS_PER_OPERATION} rows per operation"
                    )));
                }
                rows_deleted = rows_deleted.saturating_add(count);
                if rows_deleted > MAX_ROWS_DELETED {
                    return Err(PlanError::InvalidOperation(format!(
                        "plan deletes more than {MAX_ROWS_DELETED} rows"
                    )));
                }
                deletes.push((at, count, metadata));
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

    deletes.sort_by_key(|(at, _, _)| *at);
    for pair in deletes.windows(2) {
        if pair[0].0 + pair[0].1 > pair[1].0 {
            return Err(PlanError::InvalidOperation("row deletions overlap".into()));
        }
    }
    let deleted = |row: usize| {
        deletes
            .iter()
            .any(|(at, count, _)| row >= *at && row < at + count)
    };
    let dropped_writes = cell_writes
        .keys()
        .filter(|coordinate| deleted(coordinate.row))
        .count();
    cell_writes.retain(|coordinate, operation| {
        if deleted(coordinate.row) {
            return false;
        }
        let before = CellSnapshot::from_sheet(sheet, coordinate.row, coordinate.col);
        match &operation.operation {
            PlannedOp::SetCellValue { value, .. } => before.raw != value.as_input(),
            PlannedOp::SetCellFormula { formula, .. } => before.raw != *formula,
            PlannedOp::ClearCell { .. } => before != CellSnapshot::empty(),
            _ => true,
        }
    });

    let mut normalized: Vec<_> = cell_writes.into_values().collect();
    normalized.extend(styles);
    deletes.sort_unstable_by_key(|delete| std::cmp::Reverse(delete.0));
    normalized.extend(
        deletes
            .into_iter()
            .map(|(at, count, metadata)| PlannedOperation {
                operation: PlannedOp::DeleteRows { at, count },
                metadata,
            }),
    );
    let problems = if dropped_writes == 0 {
        Vec::new()
    } else {
        vec![PlanProblem {
            code: "write_to_deleted_row_dropped".into(),
            message: format!(
                "Normalization dropped {dropped_writes} write(s) targeting rows deleted by this plan."
            ),
            severity: ProblemSeverity::Warning,
        }]
    };
    Ok(NormalizationResult {
        operations: normalized,
        problems,
    })
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

fn validate_plan_metadata(
    sheet: &Sheet,
    operations: &[PlannedOperation],
    groups: &[OperationGroup],
) -> Result<(), PlanError> {
    if groups.len() > MAX_PLAN_GROUPS {
        return Err(PlanError::InvalidOperation(format!(
            "plan has more than {MAX_PLAN_GROUPS} review groups"
        )));
    }

    let mut group_ids = HashSet::new();
    for group in groups {
        validate_text_bytes("group id", &group.id.0, MAX_ID_BYTES)?;
        validate_text_bytes("group title", &group.title, MAX_LABEL_BYTES)?;
        if group.id.0.trim().is_empty() || group.title.trim().is_empty() {
            return Err(PlanError::InvalidOperation(
                "review group ids and titles must not be empty".into(),
            ));
        }
        if !group_ids.insert(group.id.0.as_str()) {
            return Err(PlanError::InvalidOperation(format!(
                "duplicate review group id '{}'",
                group.id.0
            )));
        }
        if let Some(description) = &group.description {
            validate_text_bytes("group description", description, MAX_METADATA_TEXT_BYTES)?;
        }
    }

    for planned in operations {
        if let Some(group_id) = &planned.metadata.group_id {
            if !group_ids.contains(group_id.0.as_str()) {
                return Err(PlanError::InvalidOperation(format!(
                    "operation references unknown review group '{}'",
                    group_id.0
                )));
            }
        }
        if let Some(reason) = &planned.metadata.reason {
            validate_text_bytes("operation reason", &reason.0, MAX_METADATA_TEXT_BYTES)?;
        }
        if planned.metadata.sources.len() > MAX_SOURCES_PER_OPERATION {
            return Err(PlanError::InvalidOperation(format!(
                "operation has more than {MAX_SOURCES_PER_OPERATION} source references"
            )));
        }
        for source in &planned.metadata.sources {
            validate_text_bytes("source reference", source, MAX_SOURCE_REFERENCE_BYTES)?;
            validate_source_reference(sheet, source)?;
        }
    }
    Ok(())
}

fn validate_verification_definitions(
    sheet: &Sheet,
    definitions: &[VerificationDefinition],
    groups: &[OperationGroup],
) -> Result<(), PlanError> {
    if definitions.len() > MAX_VERIFICATION_DEFINITIONS {
        return Err(PlanError::InvalidOperation(format!(
            "plan has more than {MAX_VERIFICATION_DEFINITIONS} verification definitions"
        )));
    }
    let group_ids: HashSet<_> = groups.iter().map(|group| &group.id).collect();
    let mut ids = HashSet::new();
    for definition in definitions {
        let id = match definition {
            VerificationDefinition::NoNewFormulaErrors { id, label } => {
                if let Some(label) = label {
                    validate_text_bytes("verification label", label, MAX_LABEL_BYTES)?;
                }
                id
            }
            VerificationDefinition::RetainedTotal {
                id,
                label,
                source_range,
                amount_column,
                excluded_group,
                tolerance,
                currency,
            } => {
                if let Some(label) = label {
                    validate_text_bytes("verification label", label, MAX_LABEL_BYTES)?;
                }
                let range = source_range.normalized();
                validate_coordinate(sheet, range.start)?;
                validate_coordinate(sheet, range.end)?;
                if *amount_column >= sheet.cols
                    || *amount_column < range.start.col
                    || *amount_column > range.end.col
                {
                    return Err(PlanError::InvalidOperation(
                        "retained-total amount column is outside the source range".into(),
                    ));
                }
                if !group_ids.contains(excluded_group) {
                    return Err(PlanError::InvalidOperation(format!(
                        "retained-total verification references unknown group '{}'",
                        excluded_group.0
                    )));
                }
                if !tolerance.is_finite() || *tolerance < 0.0 {
                    return Err(PlanError::InvalidOperation(
                        "retained-total tolerance must be finite and nonnegative".into(),
                    ));
                }
                if currency.len() != 3
                    || !currency.bytes().all(|byte| byte.is_ascii_alphabetic())
                {
                    return Err(PlanError::InvalidOperation(
                        "verification currency must be a three-letter code".into(),
                    ));
                }
                id
            }
        };
        validate_text_bytes("verification id", id, MAX_ID_BYTES)?;
        if id.trim().is_empty() || !ids.insert(id.as_str()) {
            return Err(PlanError::InvalidOperation(
                "verification ids must be nonempty and unique".into(),
            ));
        }
    }
    Ok(())
}

fn evaluate_verifications(
    source: &Workbook,
    candidate: &Workbook,
    source_sheet_id: SheetId,
    lineage: &[ReviewRowLineage],
    operations: &[PlannedOperation],
    definitions: &[VerificationDefinition],
    changes: &[MaterializedChange],
) -> Result<(Vec<VerificationResult>, Vec<PlanProblem>), PlanError> {
    let source_sheet = source
        .sheet_by_id(source_sheet_id)
        .ok_or(PlanError::SourceSheetMissing)?;
    let candidate_sheet = candidate
        .sheet_by_id(source_sheet_id)
        .ok_or(PlanError::SourceSheetMissing)?;
    let new_errors = changes
        .iter()
        .filter(|change| {
            is_formula_error(&change.after.display) && !is_formula_error(&change.before.display)
        })
        .count();
    let mut results = Vec::with_capacity(definitions.len());
    let mut problems = if new_errors == 0 {
        Vec::new()
    } else {
        vec![PlanProblem {
            code: "new_formula_errors".into(),
            message: format!("The candidate introduces {new_errors} new formula error(s)."),
            severity: ProblemSeverity::Blocking,
        }]
    };

    for definition in definitions {
        let (result, problem_code) = match definition {
            VerificationDefinition::NoNewFormulaErrors { id, label } => (
                VerificationResult {
                    id: id.clone(),
                    label: label.clone(),
                    status: if new_errors == 0 {
                        VerificationStatus::Passed
                    } else {
                        VerificationStatus::Failed
                    },
                    evidence: VerificationEvidence::NoNewFormulaErrors { new_errors },
                },
                None,
            ),
            VerificationDefinition::RetainedTotal {
                id,
                label,
                source_range,
                amount_column,
                excluded_group,
                tolerance,
                currency,
            } => (
                evaluate_retained_total(
                    source_sheet,
                    candidate_sheet,
                    lineage,
                    operations,
                    id.clone(),
                    label.clone(),
                    *source_range,
                    *amount_column,
                    excluded_group,
                    *tolerance,
                    currency,
                ),
                Some("retained_total_verification_failed"),
            ),
        };
        if let Some(problem_code) = problem_code {
            if matches!(
                result.status,
                VerificationStatus::Failed | VerificationStatus::Unknown
            ) {
                problems.push(PlanProblem {
                    code: problem_code.into(),
                    message: format!(
                        "Verification '{}': {}",
                        result.id,
                        result.evidence.summary()
                    ),
                    severity: ProblemSeverity::Blocking,
                });
            }
        }
        results.push(result);
    }

    Ok((results, problems))
}

#[allow(clippy::too_many_arguments)]
fn evaluate_retained_total(
    source: &Sheet,
    preview: &Sheet,
    lineage: &[ReviewRowLineage],
    operations: &[PlannedOperation],
    id: String,
    label: Option<String>,
    source_range: CellRange,
    amount_column: usize,
    excluded_group: &GroupId,
    tolerance: f64,
    currency: &str,
) -> VerificationResult {
    let range = source_range.normalized();
    let mut gross_source = 0.0;
    let mut classified_exclusions = 0.0;
    let mut preview_retained = 0.0;

    for row in lineage.iter().filter(|row| {
        row.before_data_row
            .is_some_and(|before| before >= range.start.row && before <= range.end.row)
    }) {
        let before_row = row.before_data_row.expect("filtered to source rows");
        let amount = match strict_numeric_amount(source, before_row, amount_column) {
            Ok(amount) => amount.unwrap_or(0.0),
            Err(reason) => {
                return VerificationResult {
                    id,
                    label,
                    status: VerificationStatus::Unknown,
                    evidence: VerificationEvidence::Unavailable {
                        message: format!("source row {}: {reason}", before_row + 1),
                    },
                };
            }
        };
        gross_source += amount;

        if row.state == ReviewRowState::Deleted {
            let is_classified = operations.iter().any(|planned| {
                matches!(
                    planned.operation,
                    PlannedOp::DeleteRows { at, count }
                        if before_row >= at && before_row < at + count
                ) && planned.metadata.group_id.as_ref() == Some(excluded_group)
            });
            if is_classified {
                classified_exclusions += amount;
            }
        } else if let Some(after_row) = row.after_data_row {
            match strict_numeric_amount(preview, after_row, amount_column) {
                Ok(Some(amount)) => preview_retained += amount,
                Ok(None) => {}
                Err(reason) => {
                    return VerificationResult {
                        id,
                        label,
                        status: VerificationStatus::Unknown,
                        evidence: VerificationEvidence::Unavailable {
                            message: format!("preview row {}: {reason}", after_row + 1),
                        },
                    };
                }
            }
        }
    }

    let expected_retained = gross_source - classified_exclusions;
    let difference = (expected_retained - preview_retained).abs();
    VerificationResult {
        id,
        label,
        status: if difference <= tolerance {
            VerificationStatus::Passed
        } else {
            VerificationStatus::Failed
        },
        evidence: VerificationEvidence::RetainedTotal {
            expected: currency_evidence(expected_retained),
            actual: currency_evidence(preview_retained),
            tolerance: decimal_evidence(tolerance),
            currency: currency.to_ascii_uppercase(),
            gross_source: currency_evidence(gross_source),
            classified_exclusions: currency_evidence(classified_exclusions),
            difference: currency_evidence(difference),
        },
    }
}

fn currency_evidence(value: f64) -> String {
    format!("{value:.2}")
}

fn decimal_evidence(value: f64) -> String {
    if value == 0.0 {
        "0".into()
    } else {
        value.to_string()
    }
}

fn strict_numeric_amount(sheet: &Sheet, row: usize, col: usize) -> Result<Option<f64>, String> {
    use crate::formula::eval::Value;

    match sheet.get_computed_value(row, col) {
        Value::Empty => Ok(None),
        Value::Number(value) if value.is_finite() => Ok(Some(value)),
        Value::Number(_) => Err("amount is not finite".into()),
        Value::Text(value) if value.is_empty() => Ok(None),
        Value::Text(_) => Err("amount is nonnumeric text".into()),
        Value::Boolean(_) => Err("amount is boolean".into()),
        Value::Error(error) => Err(format!("amount contains {error}")),
    }
}

fn validate_text_bytes(label: &str, value: &str, limit: usize) -> Result<(), PlanError> {
    if value.len() > limit {
        return Err(PlanError::InvalidOperation(format!(
            "{label} exceeds {limit} bytes"
        )));
    }
    Ok(())
}

fn validate_source_reference(sheet: &Sheet, source: &str) -> Result<(), PlanError> {
    let mut endpoints = source.split(':');
    let Some(start) = endpoints.next().and_then(parse_a1_coordinate) else {
        return Err(PlanError::InvalidOperation(format!(
            "invalid same-sheet source reference '{source}'"
        )));
    };
    let end = match endpoints.next() {
        Some(value) => parse_a1_coordinate(value).ok_or_else(|| {
            PlanError::InvalidOperation(format!("invalid same-sheet source reference '{source}'"))
        })?,
        None => start,
    };
    if endpoints.next().is_some() {
        return Err(PlanError::InvalidOperation(format!(
            "invalid same-sheet source reference '{source}'"
        )));
    }
    validate_coordinate(sheet, start)?;
    validate_coordinate(sheet, end)
}

fn parse_a1_coordinate(value: &str) -> Option<CellCoordinate> {
    if value.is_empty() || !value.is_ascii() {
        return None;
    }
    let letters = value.bytes().take_while(u8::is_ascii_alphabetic).count();
    if letters == 0 || letters == value.len() {
        return None;
    }
    let mut col = 0usize;
    for byte in value.bytes().take(letters) {
        col = col
            .checked_mul(26)?
            .checked_add((byte.to_ascii_uppercase() - b'A' + 1) as usize)?;
    }
    let row = value[letters..].parse::<usize>().ok()?;
    if row == 0 {
        return None;
    }
    Some(CellCoordinate {
        row: row - 1,
        col: col - 1,
    })
}

fn apply_operations(
    workbook: &mut Workbook,
    sheet_id: SheetId,
    operations: &[PlannedOperation],
) -> Result<crate::recalc::RecalcReport, PlanError> {
    let sheet_index = workbook
        .sheet_index_by_id(sheet_id)
        .ok_or(PlanError::SourceSheetMissing)?;
    workbook.begin_batch();
    for planned in operations {
        match &planned.operation {
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

    for planned in operations {
        if let PlannedOp::DeleteRows { at, count } = &planned.operation {
            workbook
                .structural_edit(sheet_index, Axis::Row, *at, *count, true)
                .map_err(PlanError::InvalidOperation)?;
        }
    }
    Ok(workbook.recompute_full_ordered())
}

fn build_row_lineage(row_count: usize, operations: &[PlannedOperation]) -> Vec<ReviewRowLineage> {
    let deletes: Vec<(usize, usize)> = operations
        .iter()
        .filter_map(|planned| match &planned.operation {
            PlannedOp::DeleteRows { at, count } => Some((*at, *count)),
            _ => None,
        })
        .collect();
    let changed_rows: HashSet<usize> = operations
        .iter()
        .filter_map(|planned| match &planned.operation {
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
    operations: &[PlannedOperation],
    lineage: &[ReviewRowLineage],
) -> Result<Vec<MaterializedChange>, PlanError> {
    let direct_metadata: BTreeMap<CellCoordinate, OperationMetadata> = operations
        .iter()
        .flat_map(|planned| match &planned.operation {
            PlannedOp::SetCellValue { coordinate, .. }
            | PlannedOp::SetCellFormula { coordinate, .. }
            | PlannedOp::ClearCell { coordinate } => {
                vec![(*coordinate, planned.metadata.clone())]
            }
            PlannedOp::SetCellStyle { range, .. } => {
                let mut cells = Vec::new();
                for row in range.start.row..=range.end.row {
                    for col in range.start.col..=range.end.col {
                        cells.push((CellCoordinate { row, col }, planned.metadata.clone()));
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
                    let metadata = operations
                        .iter()
                        .find_map(|planned| match &planned.operation {
                            PlannedOp::DeleteRows { at, count }
                                if before_row >= *at && before_row < at + count =>
                            {
                                Some(&planned.metadata)
                            }
                            _ => None,
                        })
                        .expect("deleted row has normalized deletion metadata");
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
                            group_id: metadata.group_id.clone(),
                            reason: metadata.reason.clone(),
                            sources: metadata.sources.clone(),
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
                                    group_id: metadata.group_id.clone(),
                                    reason: metadata.reason.clone(),
                                    sources: metadata.sources.clone(),
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
                        direct_metadata.get(&before_coordinate),
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
    direct_metadata: Option<&OperationMetadata>,
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
            cause: if direct_metadata.is_some() {
                ChangeCause::Direct
            } else {
                ChangeCause::Recalculated
            },
            before,
            after,
            group_id: direct_metadata.and_then(|metadata| metadata.group_id.clone()),
            reason: direct_metadata.and_then(|metadata| metadata.reason.clone()),
            sources: direct_metadata
                .map(|metadata| metadata.sources.clone())
                .unwrap_or_default(),
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
    operations: Vec<&'a PlannedOp>,
}

fn compute_plan_hash(
    contract_version: u32,
    workbook_session_id: &str,
    source_sheet_id: SheetId,
    source_revision: u64,
    execution_context: &ExecutionContextFingerprint,
    operations: &[PlannedOperation],
) -> String {
    let operations = operations
        .iter()
        .map(|planned| &planned.operation)
        .collect();
    hash_serializable(&PlanHashInput {
        contract_version,
        workbook_session_id,
        source_sheet_id,
        source_revision,
        execution_context,
        operations,
    })
}

/// Hash all workbook state used by plan materialization. This is deliberately
/// complete and therefore O(populated cells); long-lived callers such as MCP
/// should retain prepared plans instead of polling by recomputing this hash.
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
            operations: operations.into_iter().map(Into::into).collect(),
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
    fn conditional_apply_reevaluates_assertions_instead_of_trusting_preview_results() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 1, 2, "100");
        let mut request = request(&workbook, Vec::new());
        request.execution_context.volatile_inputs = vec!["test-clock".into()];
        request.groups = vec![OperationGroup {
            id: GroupId("duplicates".into()),
            title: "Duplicates".into(),
            description: None,
        }];
        request.operations = vec![PlannedOperation::plain(PlannedOp::DeleteRows {
            at: 1,
            count: 1,
        })];
        request.verification = vec![VerificationDefinition::RetainedTotal {
            id: "retained".into(),
            label: None,
            source_range: CellRange {
                start: CellCoordinate { row: 1, col: 0 },
                end: CellCoordinate { row: 1, col: 2 },
            },
            amount_column: 2,
            excluded_group: GroupId("duplicates".into()),
            tolerance: 0.01,
            currency: "USD".into(),
        }];
        let context = request.execution_context.clone();
        let mut prepared = PreparedOperationPlan::materialize(&workbook, request).unwrap();

        // Simulate a stale green result reaching Apply. The final candidate
        // must be evaluated from the retained definition, not this snapshot.
        prepared.plan.problems.clear();
        prepared.plan.verification[0].status = VerificationStatus::Passed;
        assert_eq!(
            prepared
                .verify_candidate_with_conditional_override(&workbook, &context)
                .unwrap_err(),
            PlanError::CandidateVerificationFailed
        );
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
        prepared.plan.operations[0].operation = PlannedOp::SetCellValue {
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
    fn v1_delete_primitive_limit_is_enforced() {
        let workbook = Workbook::new();
        let result = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![PlannedOp::DeleteRows {
                    at: 0,
                    count: MAX_DELETE_ROWS_PER_OPERATION + 1,
                }],
            ),
        );
        assert!(
            matches!(result, Err(PlanError::InvalidOperation(message)) if message.contains("1000"))
        );
    }

    #[test]
    fn v1_total_deleted_row_limit_is_enforced() {
        let workbook = Workbook::from_sheets(
            vec![Sheet::new(SheetId::from_raw(1), 70_000, 10)],
            0,
        );
        let operations = (0..66)
            .map(|block| {
                PlannedOperation::plain(PlannedOp::DeleteRows {
                    at: block * 1_000,
                    count: 1_000,
                })
            })
            .collect();
        let result = normalize_operations(workbook.active_sheet(), operations);
        assert!(
            matches!(result, Err(PlanError::InvalidOperation(message)) if message.contains("65536"))
        );
    }

    #[test]
    fn writes_to_rows_deleted_by_the_same_plan_are_reported() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 1, 0, "before");
        let prepared = PreparedOperationPlan::materialize(
            &workbook,
            request(
                &workbook,
                vec![
                    PlannedOp::SetCellValue {
                        coordinate: CellCoordinate { row: 1, col: 0 },
                        value: PlannedCellValue::Text("never visible".into()),
                    },
                    PlannedOp::DeleteRows { at: 1, count: 1 },
                ],
            ),
        )
        .unwrap();

        assert_eq!(prepared.plan().operations.len(), 1);
        assert!(prepared.plan().problems.iter().any(|problem| {
            problem.code == "write_to_deleted_row_dropped"
                && problem.severity == ProblemSeverity::Warning
        }));
    }

    #[test]
    fn review_metadata_reaches_direct_changes_but_not_the_execution_hash() {
        let workbook = Workbook::new();
        let mut first = request(&workbook, Vec::new());
        first.groups = vec![OperationGroup {
            id: GroupId("cleanup".into()),
            title: "Cleanup".into(),
            description: None,
        }];
        first.operations = vec![PlannedOperation {
            operation: PlannedOp::SetCellValue {
                coordinate: CellCoordinate { row: 0, col: 0 },
                value: PlannedCellValue::Text("after".into()),
            },
            metadata: OperationMetadata {
                group_id: Some(GroupId("cleanup".into())),
                reason: Some(ProducerClaim("First explanation".into())),
                sources: vec!["A2".into()],
            },
        }];
        let mut second = first.clone();
        second.operations[0].metadata.reason = Some(ProducerClaim("Reworded".into()));

        let first = PreparedOperationPlan::materialize(&workbook, first).unwrap();
        let second = PreparedOperationPlan::materialize(&workbook, second).unwrap();
        assert_eq!(first.plan().plan_hash, second.plan().plan_hash);
        let change = first
            .plan()
            .changes
            .iter()
            .find(|change| change.after_coordinate == Some(CellCoordinate { row: 0, col: 0 }))
            .unwrap();
        assert_eq!(change.group_id, Some(GroupId("cleanup".into())));
        assert_eq!(
            change.reason,
            Some(ProducerClaim("First explanation".into()))
        );
        assert_eq!(change.sources, vec!["A2"]);
    }

    #[test]
    fn retained_total_is_computed_from_typed_source_and_preview_values() {
        let mut workbook = Workbook::new();
        for (row, value) in [(1, "100"), (2, "200"), (3, "100"), (5, "50")] {
            workbook.set_cell_value_tracked(0, row, 2, value);
        }
        let duplicate_group = OperationGroup {
            id: GroupId("exact_duplicates".into()),
            title: "Exact duplicates".into(),
            description: None,
        };
        let empty_group = OperationGroup {
            id: GroupId("empty_rows".into()),
            title: "Empty rows".into(),
            description: None,
        };
        let mut request = request(&workbook, Vec::new());
        request.groups = vec![duplicate_group.clone(), empty_group.clone()];
        request.operations = vec![
            PlannedOperation {
                operation: PlannedOp::DeleteRows { at: 3, count: 1 },
                metadata: OperationMetadata {
                    group_id: Some(duplicate_group.id.clone()),
                    reason: Some(ProducerClaim("Duplicate transaction".into())),
                    sources: vec!["A2:C2".into(), "A4:C4".into()],
                },
            },
            PlannedOperation {
                operation: PlannedOp::DeleteRows { at: 4, count: 1 },
                metadata: OperationMetadata {
                    group_id: Some(empty_group.id.clone()),
                    reason: Some(ProducerClaim("Empty transaction row".into())),
                    sources: vec!["A5:C5".into()],
                },
            },
        ];
        request.verification = vec![VerificationDefinition::RetainedTotal {
            id: "retained_payments".into(),
            label: Some("Retained payments".into()),
            source_range: CellRange {
                start: CellCoordinate { row: 1, col: 0 },
                end: CellCoordinate { row: 5, col: 2 },
            },
            amount_column: 2,
            excluded_group: duplicate_group.id,
            tolerance: 0.01,
            currency: "USD".into(),
        }];

        let prepared = PreparedOperationPlan::materialize(&workbook, request).unwrap();
        let verification = &prepared.plan().verification[0];
        assert_eq!(verification.status, VerificationStatus::Passed);
        assert!(matches!(
            &verification.evidence,
            VerificationEvidence::RetainedTotal {
                expected,
                actual,
                currency,
                ..
            } if expected == "350.00" && actual == "350.00" && currency == "USD"
        ));
        let wire = serde_json::to_value(verification).unwrap();
        assert_eq!(wire["kind"], "gross_minus_group_equals_preview");
        assert_eq!(wire["expected"], "350.00");
        assert_eq!(wire["actual"], "350.00");
        assert_eq!(wire["tolerance"], "0.01");
        assert!(wire.get("evidence").is_none());
        let deletion = prepared
            .plan()
            .changes
            .iter()
            .find(|change| {
                change.kind == ChangeKind::RowDeleted
                    && change.group_id == Some(GroupId("exact_duplicates".into()))
            })
            .unwrap();
        assert_eq!(deletion.sources, vec!["A2:C2", "A4:C4"]);
        assert!(!prepared
            .plan()
            .problems
            .iter()
            .any(|problem| problem.severity == ProblemSeverity::Blocking));
    }

    #[test]
    fn nonnumeric_retained_amount_is_unknown_and_blocks_apply() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 1, 2, "not an amount");
        let mut request = request(&workbook, Vec::new());
        request.groups = vec![OperationGroup {
            id: GroupId("duplicates".into()),
            title: "Duplicates".into(),
            description: None,
        }];
        request.verification = vec![VerificationDefinition::RetainedTotal {
            id: "retained".into(),
            label: None,
            source_range: CellRange {
                start: CellCoordinate { row: 1, col: 0 },
                end: CellCoordinate { row: 1, col: 2 },
            },
            amount_column: 2,
            excluded_group: GroupId("duplicates".into()),
            tolerance: 0.01,
            currency: "USD".into(),
        }];

        let prepared = PreparedOperationPlan::materialize(&workbook, request).unwrap();
        assert_eq!(
            prepared.plan().verification[0].status,
            VerificationStatus::Unknown
        );
        assert!(matches!(
            &prepared.plan().verification[0].evidence,
            VerificationEvidence::Unavailable { message }
                if message.contains("nonnumeric text")
        ));
        assert_eq!(
            prepared.verify_candidate(&workbook, &context()).unwrap_err(),
            PlanError::BlockingProblems
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
