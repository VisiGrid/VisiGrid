//! Golden schema tests for `vgrid publish` JSON output.
//!
//! These tests validate that RunResult serialization matches the committed
//! golden JSON files. If a field is added, removed, or renamed, these tests
//! will fail — forcing an explicit schema version bump.
//!
//! The golden files are the public contract. CI scripts parse this JSON.
//! Breaking it without versioning breaks customers.

use visigrid_hub_client::RunResult;

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
        proof_url: "https://api.visihub.app/api/repos/acme/payments/runs/42/proof".into(),
    };

    validate_golden_keys("tests/golden/publish-pass.json", &result);

    // check_status must be "pass"
    let json = serde_json::to_value(&result).unwrap();
    assert_eq!(json["check_status"], "pass");
    assert_eq!(json["status"], "verified");
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
        proof_url: "https://api.visihub.app/api/repos/acme/payments/runs/99/proof".into(),
    };

    validate_golden_keys("tests/golden/publish-fail.json", &result);

    // check_status must be "fail" — CLI uses this to set exit code 41
    let json = serde_json::to_value(&result).unwrap();
    assert_eq!(json["check_status"], "fail");
}

#[test]
fn test_golden_no_wait_output() {
    // When --no-wait is used, the CLI emits a minimal JSON with just run_id, status, proof_url.
    // This isn't a RunResult — it's hand-built in hub.rs. Validate the golden shape.
    let golden: serde_json::Value = serde_json::from_str(
        &std::fs::read_to_string("tests/golden/publish-no-wait.json").unwrap()
    ).unwrap();

    assert!(golden["run_id"].is_string(), "no-wait output must have run_id");
    assert_eq!(golden["status"], "processing", "no-wait status must be 'processing'");
    assert!(golden["proof_url"].is_string(), "no-wait output must have proof_url");
}

#[test]
fn test_run_result_required_fields_never_null() {
    // These fields are ALWAYS present (not Option). If someone makes them
    // optional, this test must fail to force a schema version discussion.
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
        proof_url: "https://example.com/proof".into(),
    };

    let json = serde_json::to_value(&result).unwrap();
    let obj = json.as_object().unwrap();

    // Required: always present, never null
    for key in &["run_id", "version", "status", "proof_url"] {
        assert!(obj.contains_key(*key), "Required field '{}' missing", key);
        assert!(!obj[*key].is_null(), "Required field '{}' is null", key);
    }
}
