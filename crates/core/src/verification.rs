use serde::{Deserialize, Serialize};

/// The governing profile for a verification run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum VerificationProfile {
    /// Strict mode for compliance. Fails on any pollutant or ambiguity.
    Audit,
    /// Pragmatic mode for operational reconciliation.
    Reconcile,
    /// Detailed mode capturing all changes including formatting.
    Forensics,
}

/// Status of a verification run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum VerificationStatus {
    /// Verification succeeded and hashes are valid.
    Verified,
    /// File contains pollutants or ambiguity and cannot be audited.
    NotAuditable,
    /// Verification failed (e.g., hash mismatch).
    Failed,
}

/// The primary artifact of a VisiGrid verification run.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ProofObject {
    /// Contract version (e.g., "1.0")
    pub v: String,
    /// Engine version that generated the proof
    pub engine: String,
    /// Profile used for verification
    pub profile: VerificationProfile,
    /// RFC 3339 timestamp
    pub timestamp: String,
    /// Source file information
    pub source: SourceInfo,
    /// Scope of the verification
    pub scope: VerificationScope,
    /// Cryptographic fingerprints
    pub fingerprints: Fingerprints,
    /// Final status of the verification
    pub status: VerificationStatus,
    /// List of detected pollutants (if any)
    pub pollutants: Vec<PollutantInfo>,
    /// Non-fatal warnings
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SourceInfo {
    pub filename: String,
    /// SHA-256 of the raw source file
    pub hash: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct VerificationScope {
    pub sheets: Vec<String>,
    pub input_regions: Vec<String>,
    pub excluded_ranges: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Fingerprints {
    /// Hash of Structure and Logic (Formulas, Topology, Named Ranges)
    pub model: String,
    /// Hash of literal Data values
    pub data: String,
    /// Optional hash of the computed results
    #[serde(skip_serializing_if = "Option::is_none")]
    pub result: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PollutantInfo {
    pub code: String,
    pub pollutant_type: String,
    pub description: String,
    pub location: Option<String>,
}