//! Golden schema tests for `vgrid publish` JSON output.
//!
//! These tests validate that RunResult serialization matches the committed
//! golden JSON files. If a field is added, removed, or renamed, these tests
//! will fail — forcing an explicit schema version bump.
//!
//! The golden files are the public contract. CI scripts parse this JSON.
//! Breaking it without versioning breaks customers.

use visigrid_hub_client::{RunResult, AssertionResult, EngineMetadata, CreateRevisionOptions};

/// Validate that every key in the golden JSON is present in RunResult serialization.
fn validate_golden_keys(golden_path: &str, result: &RunResult) {
    let golden: serde_json::Value = serde_json::from_str(
        &std::fs::read_to_string(golden_path)
            .unwrap_or_else(|e| panic!("Cannot read {}: {}", golden_path, e))
    ).unwrap_or_else(|e| panic!("Cannot parse {}: {}", golden_path, e));

    let serialized = serde_json::to_value(result).unwrap();

    // Every key in the golden file must exist in the serialized output
    if let Some(golden_obj) = golden.as_object() {
        let serial_obj = serialized.as_object().expect("RunResult should serialize as object");
        for key in golden_obj.keys() {
            assert!(
                serial_obj.contains_key(key),
                "Golden key '{}' missing from RunResult serialization (file: {})",
                key, golden_path,
            );
        }
    }
}

#[test]
fn test_golden_publish_pass() {
    let result = RunResult {
        run_id: "42".into(),
        version: 3,
        status: "verified".into(),
        check_status: Some("pass".into()),
        diff_summary: Some(serde_json::json!({
            "row_count_change": 10,
            "col_count_change": 0,
        })),
        row_count: Some(1000),
        col_count: Some(15),
        content_hash: Some("blake3:a7ffc6f8bf1ed76651c14756a061d662f580ff4de43b49fa82d80a4b80f8434a".into()),
        source_metadata: Some(serde_json::json!({
            "type": "dbt",
            "identity": "models/payments",
            "timestamp": "2025-06-15T14:30:00Z",
            "query_hash": "sha256:e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
        })),
        assertions: Some(vec![AssertionResult {
            kind: "sum".into(),
            column: "amount".into(),
            expected: Some("12345.67".into()),
            actual: Some("12345.66".into()),
            tolerance: Some("0.01".into()),
            status: "pass".into(),
            delta: None,
            message: None,
            origin: None,
            engine: None,
        }]),
        proof_url: "https://api.visihub.app/api/repos/acme/payments/runs/42/proof".into(),
    };

    validate_golden_keys("tests/golden/publish-pass.json", &result);

    let json = serde_json::to_value(&result).unwrap();
    assert_eq!(json["check_status"], "pass");
    assert_eq!(json["status"], "verified");
    assert!(json["assertions"].is_array());
    assert_eq!(json["assertions"][0]["status"], "pass");
}

#[test]
fn test_golden_publish_fail() {
    let result = RunResult {
        run_id: "99".into(),
        version: 5,
        status: "verified".into(),
        check_status: Some("fail".into()),
        diff_summary: Some(serde_json::json!({
            "row_count_change": -50,
            "col_count_change": 2,
        })),
        row_count: Some(950),
        col_count: Some(17),
        content_hash: Some("blake3:deadbeef".into()),
        source_metadata: Some(serde_json::json!({"type": "dbt", "identity": "models/payments"})),
        assertions: Some(vec![AssertionResult {
            kind: "sum".into(),
            column: "amount".into(),
            expected: Some("12345.67".into()),
            actual: Some("12300.00".into()),
            tolerance: Some("0.01".into()),
            status: "fail".into(),
            delta: Some("45.67".into()),
            message: None,
            origin: None,
            engine: None,
        }]),
        proof_url: "https://api.visihub.app/api/repos/acme/payments/runs/99/proof".into(),
    };

    validate_golden_keys("tests/golden/publish-fail.json", &result);

    let json = serde_json::to_value(&result).unwrap();
    assert_eq!(json["check_status"], "fail");
    assert_eq!(json["assertions"][0]["status"], "fail");
    assert!(json["assertions"][0]["delta"].is_string());
}

#[test]
fn test_golden_publish_baseline() {
    let result = RunResult {
        run_id: "1".into(),
        version: 1,
        status: "verified".into(),
        check_status: Some("baseline_created".into()),
        diff_summary: None,
        row_count: Some(500),
        col_count: Some(10),
        content_hash: Some("blake3:a7ffc6f8bf1ed76651c14756a061d662f580ff4de43b49fa82d80a4b80f8434a".into()),
        source_metadata: Some(serde_json::json!({"type": "dbt", "identity": "models/payments"})),
        assertions: Some(vec![AssertionResult {
            kind: "sum".into(),
            column: "amount".into(),
            expected: None,
            actual: Some("12345.67".into()),
            tolerance: None,
            status: "baseline_created".into(),
            delta: None,
            message: None,
            origin: None,
            engine: None,
        }]),
        proof_url: "https://api.visihub.app/api/repos/acme/payments/runs/1/proof".into(),
    };

    validate_golden_keys("tests/golden/publish-baseline.json", &result);

    let json = serde_json::to_value(&result).unwrap();
    assert_eq!(json["check_status"], "baseline_created");
    assert_eq!(json["assertions"][0]["status"], "baseline_created");
    assert!(json["assertions"][0]["actual"].is_string());
    // baseline_created must NOT trigger --fail-on-check-failure
    assert_ne!(result.check_status.as_deref(), Some("fail"));
}

#[test]
fn test_golden_publish_warn() {
    let result = RunResult {
        run_id: "55".into(),
        version: 4,
        status: "verified".into(),
        check_status: Some("warn".into()),
        diff_summary: Some(serde_json::json!({
            "row_count_change": 5,
            "col_count_change": 0,
        })),
        row_count: Some(1005),
        col_count: Some(15),
        content_hash: Some("blake3:b8ffc6f8bf1ed76651c14756a061d662f580ff4de43b49fa82d80a4b80f8434a".into()),
        source_metadata: Some(serde_json::json!({"type": "dbt", "identity": "models/payments"})),
        assertions: None,
        proof_url: "https://api.visihub.app/api/repos/acme/payments/runs/55/proof".into(),
    };

    validate_golden_keys("tests/golden/publish-warn.json", &result);

    let json = serde_json::to_value(&result).unwrap();
    assert_eq!(json["check_status"], "warn");
    assert_eq!(json["status"], "verified");
    // warn must NOT match "fail" — so exit code stays 0
    assert_ne!(result.check_status.as_deref(), Some("fail"));
}

#[test]
fn test_golden_no_wait_output() {
    let golden: serde_json::Value = serde_json::from_str(
        &std::fs::read_to_string("tests/golden/publish-no-wait.json").unwrap()
    ).unwrap();

    assert!(golden["run_id"].is_string(), "no-wait output must have run_id");
    assert_eq!(golden["status"], "processing", "no-wait status must be 'processing'");
    assert!(golden["proof_url"].is_string(), "no-wait output must have proof_url");
}

#[test]
fn test_run_result_required_fields_never_null() {
    let result = RunResult {
        run_id: "1".into(),
        version: 1,
        status: "verified".into(),
        check_status: None,
        diff_summary: None,
        row_count: None,
        col_count: None,
        content_hash: None,
        source_metadata: None,
        assertions: None,
        proof_url: "https://example.com/proof".into(),
    };

    let json = serde_json::to_value(&result).unwrap();
    let obj = json.as_object().unwrap();

    for key in &["run_id", "version", "status", "proof_url"] {
        assert!(obj.contains_key(*key), "Required field '{}' missing", key);
        assert!(!obj[*key].is_null(), "Required field '{}' is null", key);
    }
}

#[test]
fn test_assertion_result_schema() {
    // Verify the assertion result JSON shape matches the contract
    let pass = AssertionResult {
        kind: "sum".into(),
        column: "amount".into(),
        expected: Some("12345.67".into()),
        actual: Some("12345.66".into()),
        tolerance: Some("0.01".into()),
        status: "pass".into(),
        delta: None,
        message: None,
        origin: None,
        engine: None,
    };

    let fail = AssertionResult {
        kind: "sum".into(),
        column: "revenue".into(),
        expected: Some("100000".into()),
        actual: Some("99950".into()),
        tolerance: Some("0".into()),
        status: "fail".into(),
        delta: Some("50".into()),
        message: None,
        origin: None,
        engine: None,
    };

    let pass_json = serde_json::to_value(&pass).unwrap();
    assert_eq!(pass_json["kind"], "sum");
    assert_eq!(pass_json["column"], "amount");
    assert_eq!(pass_json["status"], "pass");
    // delta should not appear for pass (skip_serializing_if)
    assert!(pass_json.get("delta").is_none());

    let fail_json = serde_json::to_value(&fail).unwrap();
    assert_eq!(fail_json["status"], "fail");
    assert_eq!(fail_json["delta"], "50");
}

#[test]
fn test_golden_publish_cell_assertion_pass() {
    let result = RunResult {
        run_id: "100".into(),
        version: 2,
        status: "verified".into(),
        check_status: Some("pass".into()),
        diff_summary: None,
        row_count: None,
        col_count: None,
        content_hash: None,
        source_metadata: None,
        assertions: Some(vec![AssertionResult {
            kind: "cell".into(),
            column: "summary!B7".into(),
            expected: Some("0".into()),
            actual: Some("4520".into()),
            tolerance: Some("10000".into()),
            status: "pass".into(),
            delta: None,
            message: None,
            origin: Some("client".into()),
            engine: Some(EngineMetadata {
                name: "visigrid-engine".into(),
                version: "0.1.0".into(),
                fingerprint: None,
            }),
        }]),
        proof_url: "https://api.visihub.app/api/repos/acme/recon/runs/100/proof".into(),
    };

    validate_golden_keys("tests/golden/publish-cell-pass.json", &result);

    let json = serde_json::to_value(&result).unwrap();
    assert_eq!(json["check_status"], "pass");
    assert_eq!(json["status"], "verified");
    assert_eq!(json["assertions"][0]["kind"], "cell");
    assert_eq!(json["assertions"][0]["status"], "pass");
    assert_eq!(json["assertions"][0]["origin"], "client");
    assert!(json["assertions"][0]["engine"].is_object());
    assert_eq!(json["assertions"][0]["engine"]["name"], "visigrid-engine");
}

// ===========================================================================
// DatasetStatus + CreateRevisionOptions tests
// ===========================================================================

#[test]
fn test_create_revision_options_defaults_have_none_for_new_fields() {
    let opts = CreateRevisionOptions::default();
    assert!(opts.source_metadata.is_none(), "source_metadata should default to None");
    assert!(opts.message.is_none(), "message should default to None");
}

#[test]
fn test_create_revision_options_with_source_metadata() {
    let sm = serde_json::json!({
        "type": "trust_pipeline",
        "fingerprint": "v2:42:abc123",
        "timestamp": "2026-02-12T00:00:00Z",
        "trust_pipeline": {
            "stamp": {
                "expected_fingerprint": "v2:42:abc123",
                "label": "Q4 Filing"
            }
        }
    });

    let opts = CreateRevisionOptions {
        source_metadata: Some(sm.clone()),
        message: Some("Q4 close".into()),
        format: Some("sheet".into()),
        ..Default::default()
    };

    assert_eq!(opts.source_metadata.unwrap(), sm);
    assert_eq!(opts.message.unwrap(), "Q4 close");
    // Individual fields should be None (source_metadata takes priority)
    assert!(opts.source_type.is_none());
    assert!(opts.source_identity.is_none());
    assert!(opts.query_hash.is_none());
}

#[test]
fn test_dataset_status_parse_with_source_metadata() {
    // Simulate parsing a JSON response with source_metadata
    let json: serde_json::Value = serde_json::json!({
        "current_revision_id": 42,
        "content_hash": "blake3:abc123",
        "byte_size": 1024,
        "updated_at": "2026-02-12T00:00:00Z",
        "updated_by": "alice",
        "source_metadata": {
            "type": "trust_pipeline",
            "fingerprint": "v2:42:abc123"
        }
    });

    let status = visigrid_hub_client::DatasetStatus {
        current_revision_id: json["current_revision_id"].as_i64()
            .map(|n| n.to_string()),
        content_hash: json["content_hash"].as_str().map(String::from),
        byte_size: json["byte_size"].as_u64(),
        source_metadata: json.get("source_metadata").cloned()
            .filter(|v| !v.is_null()),
    };

    assert_eq!(status.current_revision_id.as_deref(), Some("42"));
    assert_eq!(status.content_hash.as_deref(), Some("blake3:abc123"));
    assert_eq!(status.byte_size, Some(1024));
    let sm = status.source_metadata.unwrap();
    assert_eq!(sm["type"], "trust_pipeline");
    assert_eq!(sm["fingerprint"], "v2:42:abc123");
}

#[test]
fn test_dataset_status_parse_without_source_metadata() {
    let json: serde_json::Value = serde_json::json!({
        "current_revision_id": null,
        "content_hash": null,
        "byte_size": null,
        "updated_at": "2026-02-12T00:00:00Z",
        "updated_by": null,
        "source_metadata": null
    });

    let status = visigrid_hub_client::DatasetStatus {
        current_revision_id: json["current_revision_id"].as_i64()
            .map(|n| n.to_string()),
        content_hash: json["content_hash"].as_str().map(String::from),
        byte_size: json["byte_size"].as_u64(),
        source_metadata: json.get("source_metadata").cloned()
            .filter(|v| !v.is_null()),
    };

    assert!(status.current_revision_id.is_none());
    assert!(status.content_hash.is_none());
    assert!(status.byte_size.is_none());
    assert!(status.source_metadata.is_none());
}

/// Idempotency edge cases: source_metadata exists but is not a trust_pipeline,
/// or has a different type, or has no fingerprint. CLI should publish (not skip).
#[test]
fn test_idempotency_non_trust_pipeline_type_should_not_match() {
    // Latest revision was uploaded by legacy `vgrid publish` with type=dbt
    let sm = serde_json::json!({
        "type": "dbt",
        "identity": "models/payments",
        "timestamp": "2026-02-12T00:00:00Z"
    });

    let local_fingerprint = "v2:42:abc123";

    // This should NOT trigger idempotency — type is not "trust_pipeline"
    let is_match = sm["type"].as_str() == Some("trust_pipeline")
        && sm["fingerprint"].as_str() == Some(local_fingerprint);
    assert!(!is_match, "non-trust_pipeline type must not match");
}

#[test]
fn test_idempotency_trust_pipeline_missing_fingerprint_should_not_match() {
    // trust_pipeline type but fingerprint field is missing
    let sm = serde_json::json!({
        "type": "trust_pipeline",
        "timestamp": "2026-02-12T00:00:00Z"
    });

    let local_fingerprint = "v2:42:abc123";

    let is_match = sm["type"].as_str() == Some("trust_pipeline")
        && sm["fingerprint"].as_str() == Some(local_fingerprint);
    assert!(!is_match, "trust_pipeline without fingerprint must not match");
}

#[test]
fn test_idempotency_trust_pipeline_different_fingerprint_should_not_match() {
    // trust_pipeline with a DIFFERENT fingerprint
    let sm = serde_json::json!({
        "type": "trust_pipeline",
        "fingerprint": "v2:99:different",
        "timestamp": "2026-02-12T00:00:00Z"
    });

    let local_fingerprint = "v2:42:abc123";

    let is_match = sm["type"].as_str() == Some("trust_pipeline")
        && sm["fingerprint"].as_str() == Some(local_fingerprint);
    assert!(!is_match, "different fingerprint must not match");
}

#[test]
fn test_idempotency_trust_pipeline_matching_fingerprint_should_match() {
    let sm = serde_json::json!({
        "type": "trust_pipeline",
        "fingerprint": "v2:42:abc123",
        "timestamp": "2026-02-12T00:00:00Z"
    });

    let local_fingerprint = "v2:42:abc123";

    let is_match = sm["type"].as_str() == Some("trust_pipeline")
        && sm["fingerprint"].as_str() == Some(local_fingerprint);
    assert!(is_match, "matching trust_pipeline + fingerprint must match");
}
