//! Golden vector compatibility test for v1 protocol.
//!
//! This test ensures the visigrid-protocol crate types can deserialize the frozen v1 golden vectors.
//! If this test fails, the protocol types have drifted from the canonical wire format.
//!
//! Golden vectors live in: gpui-app/src/session_server/protocol_golden/*.jsonl
//! These files are the source of truth for wire format compatibility.
//!
//! **Rule**: The golden vectors MUST NOT change. If the test fails, fix the types, not the vectors.

use std::fs;
use std::path::PathBuf;

use serde_json::Value;

// Use the shared protocol types
use visigrid_protocol::{
    ClientMessage, ServerMessage, EventPayload, InspectResult, InspectTarget,
};

/// Find the protocol_golden directory relative to workspace root.
fn golden_dir() -> PathBuf {
    let manifest_dir = env!("CARGO_MANIFEST_DIR");
    PathBuf::from(manifest_dir)
        .parent() // crates/
        .unwrap()
        .parent() // workspace root
        .unwrap()
        .join("gpui-app/src/session_server/protocol_golden")
}

/// Load all lines from a golden vector file.
fn load_golden_lines(filename: &str) -> Vec<String> {
    let path = golden_dir().join(filename);
    let contents = fs::read_to_string(&path)
        .unwrap_or_else(|e| panic!("Failed to read {}: {}", path.display(), e));
    contents
        .lines()
        .filter(|line| !line.trim().is_empty())
        .map(String::from)
        .collect()
}

/// Test that a line can be deserialized as generic JSON (sanity check).
fn assert_valid_json(line: &str, context: &str) {
    serde_json::from_str::<Value>(line)
        .unwrap_or_else(|e| panic!("Invalid JSON in {}: {} - line: {}", context, e, line));
}

// =============================================================================
// Golden Vector Tests
// =============================================================================

#[test]
fn test_hello_ok() {
    let lines = load_golden_lines("hello_ok.jsonl");
    assert_eq!(lines.len(), 2, "hello_ok.jsonl should have 2 lines");

    // Line 1: Client hello
    let client_msg: ClientMessage = serde_json::from_str(&lines[0])
        .expect("Failed to deserialize hello message");
    match client_msg {
        ClientMessage::Hello(hello) => {
            assert_eq!(hello.client, "test-agent");
            assert_eq!(hello.protocol_version, 1);
        }
        _ => panic!("Expected Hello message"),
    }

    // Line 2: Server welcome
    let server_msg: ServerMessage = serde_json::from_str(&lines[1])
        .expect("Failed to deserialize welcome message");
    match server_msg {
        ServerMessage::Welcome(welcome) => {
            assert_eq!(welcome.protocol_version, 1);
            assert!(welcome.capabilities.contains(&"apply".to_string()));
            assert!(welcome.capabilities.contains(&"inspect".to_string()));
        }
        _ => panic!("Expected Welcome message"),
    }
}

#[test]
fn test_apply_ops_ok() {
    let lines = load_golden_lines("apply_ops_ok.jsonl");
    assert_eq!(lines.len(), 2, "apply_ops_ok.jsonl should have 2 lines");

    // Line 1: Client apply_ops
    let client_msg: ClientMessage = serde_json::from_str(&lines[0])
        .expect("Failed to deserialize apply_ops message");
    match client_msg {
        ClientMessage::ApplyOps(apply) => {
            assert_eq!(apply.ops.len(), 3);
            assert!(apply.atomic);
            assert_eq!(apply.expected_revision, Some(42));
        }
        _ => panic!("Expected ApplyOps message"),
    }

    // Line 2: Server result
    let server_msg: ServerMessage = serde_json::from_str(&lines[1])
        .expect("Failed to deserialize apply_ops_result message");
    match server_msg {
        ServerMessage::ApplyOpsResult(result) => {
            assert_eq!(result.applied, 3);
            assert_eq!(result.total, 3);
            assert_eq!(result.revision, 43);
            assert!(result.error.is_none());
        }
        _ => panic!("Expected ApplyOpsResult message"),
    }
}

#[test]
fn test_inspect() {
    let lines = load_golden_lines("inspect.jsonl");
    assert!(lines.len() >= 6, "inspect.jsonl should have at least 6 lines");

    // Line 1: Inspect cell
    let client_msg: ClientMessage = serde_json::from_str(&lines[0])
        .expect("Failed to deserialize inspect message");
    match client_msg {
        ClientMessage::Inspect(inspect) => {
            match inspect.target {
                InspectTarget::Cell { sheet, row, col } => {
                    assert_eq!(sheet, 0);
                    assert_eq!(row, 0);
                    assert_eq!(col, 0);
                }
                _ => panic!("Expected Cell target"),
            }
        }
        _ => panic!("Expected Inspect message"),
    }

    // Line 2: Inspect result (cell)
    let server_msg: ServerMessage = serde_json::from_str(&lines[1])
        .expect("Failed to deserialize inspect_result message");
    match server_msg {
        ServerMessage::InspectResult(result) => {
            match result.result {
                InspectResult::Cell(info) => {
                    assert_eq!(info.raw, "Hello");
                    assert_eq!(info.display, "Hello");
                    assert!(info.formula.is_none());
                }
                _ => panic!("Expected Cell result"),
            }
        }
        _ => panic!("Expected InspectResult message"),
    }

    // Line 5: Inspect workbook
    let client_msg: ClientMessage = serde_json::from_str(&lines[4])
        .expect("Failed to deserialize workbook inspect message");
    match client_msg {
        ClientMessage::Inspect(inspect) => {
            assert!(matches!(inspect.target, InspectTarget::Workbook));
        }
        _ => panic!("Expected Inspect message"),
    }

    // Line 6: Workbook result
    let server_msg: ServerMessage = serde_json::from_str(&lines[5])
        .expect("Failed to deserialize workbook result message");
    match server_msg {
        ServerMessage::InspectResult(result) => {
            match result.result {
                InspectResult::Workbook(info) => {
                    assert_eq!(info.sheet_count, 1);
                    assert_eq!(info.title, "Untitled");
                }
                _ => panic!("Expected Workbook result"),
            }
        }
        _ => panic!("Expected InspectResult message"),
    }
}

#[test]
fn test_errors() {
    let lines = load_golden_lines("errors.jsonl");
    assert!(lines.len() >= 10, "errors.jsonl should have at least 10 error types");

    for (i, line) in lines.iter().enumerate() {
        let server_msg: ServerMessage = serde_json::from_str(line)
            .unwrap_or_else(|e| panic!("Failed to deserialize error line {}: {}", i + 1, e));

        match server_msg {
            ServerMessage::Error(err) => {
                assert!(!err.code.is_empty(), "Error code should not be empty");
                assert!(!err.message.is_empty(), "Error message should not be empty");

                // Specific checks for errors with retry_after_ms
                if err.code == "rate_limited" {
                    assert!(err.retry_after_ms.is_some(), "rate_limited should have retry_after_ms");
                }
                if err.code == "writer_conflict" {
                    assert!(err.retry_after_ms.is_some(), "writer_conflict should have retry_after_ms");
                }
            }
            _ => panic!("Expected Error message on line {}", i + 1),
        }
    }
}

#[test]
fn test_subscribe_events() {
    let lines = load_golden_lines("subscribe_events.jsonl");
    assert!(lines.len() >= 6, "subscribe_events.jsonl should have at least 6 lines");

    // Line 1: Subscribe
    let client_msg: ClientMessage = serde_json::from_str(&lines[0])
        .expect("Failed to deserialize subscribe message");
    match client_msg {
        ClientMessage::Subscribe(sub) => {
            assert!(sub.topics.contains(&"cells".to_string()));
        }
        _ => panic!("Expected Subscribe message"),
    }

    // Line 2: Subscribed
    let server_msg: ServerMessage = serde_json::from_str(&lines[1])
        .expect("Failed to deserialize subscribed message");
    match server_msg {
        ServerMessage::Subscribed(sub) => {
            assert!(sub.topics.contains(&"cells".to_string()));
        }
        _ => panic!("Expected Subscribed message"),
    }

    // Line 3: Event
    let server_msg: ServerMessage = serde_json::from_str(&lines[2])
        .expect("Failed to deserialize event message");
    match server_msg {
        ServerMessage::Event(event) => {
            assert_eq!(event.topic, "cells");
            match event.payload {
                EventPayload::CellsChanged { ranges } => {
                    assert!(!ranges.is_empty());
                }
                _ => panic!("Expected CellsChanged payload"),
            }
        }
        _ => panic!("Expected Event message"),
    }
}

#[test]
fn test_ping_pong() {
    let lines = load_golden_lines("ping_pong.jsonl");
    assert_eq!(lines.len(), 2, "ping_pong.jsonl should have 2 lines");

    // Line 1: Ping
    let client_msg: ClientMessage = serde_json::from_str(&lines[0])
        .expect("Failed to deserialize ping message");
    assert!(matches!(client_msg, ClientMessage::Ping(_)));

    // Line 2: Pong
    let server_msg: ServerMessage = serde_json::from_str(&lines[1])
        .expect("Failed to deserialize pong message");
    assert!(matches!(server_msg, ServerMessage::Pong(_)));
}

#[test]
fn test_stats() {
    let lines = load_golden_lines("stats.jsonl");
    assert_eq!(lines.len(), 2, "stats.jsonl should have 2 lines");

    // Line 1: Stats request
    let client_msg: ClientMessage = serde_json::from_str(&lines[0])
        .expect("Failed to deserialize stats message");
    assert!(matches!(client_msg, ClientMessage::Stats(_)));

    // Line 2: Stats result
    let server_msg: ServerMessage = serde_json::from_str(&lines[1])
        .expect("Failed to deserialize stats_result message");
    match server_msg {
        ServerMessage::StatsResult(stats) => {
            assert!(stats.active_connections > 0);
        }
        _ => panic!("Expected StatsResult message"),
    }
}

#[test]
fn test_writer_conflict() {
    let lines = load_golden_lines("writer_conflict.jsonl");
    assert!(lines.len() >= 2, "writer_conflict.jsonl should have at least 2 lines");

    // Should contain an error with writer_conflict code
    for line in &lines {
        if let Ok(ServerMessage::Error(err)) = serde_json::from_str::<ServerMessage>(line) {
            if err.code == "writer_conflict" {
                assert!(err.retry_after_ms.is_some());
                return;
            }
        }
    }
    panic!("writer_conflict.jsonl should contain a writer_conflict error");
}

#[test]
fn test_events_dropped() {
    let lines = load_golden_lines("events_dropped.jsonl");
    assert!(!lines.is_empty(), "events_dropped.jsonl should not be empty");

    // Should contain an EventsDropped event
    for line in &lines {
        if let Ok(ServerMessage::Event(event)) = serde_json::from_str::<ServerMessage>(line) {
            if matches!(event.payload, EventPayload::EventsDropped { .. }) {
                return;
            }
        }
    }
    panic!("events_dropped.jsonl should contain an EventsDropped event");
}

/// Meta-test: Ensure all golden vector files are valid JSON.
#[test]
fn test_all_golden_files_valid_json() {
    let golden_files = [
        "hello_ok.jsonl",
        "hello_protocol_mismatch.jsonl",
        "apply_ops_ok.jsonl",
        "errors.jsonl",
        "inspect.jsonl",
        "ping_pong.jsonl",
        "stats.jsonl",
        "subscribe_events.jsonl",
        "subscribe_events_large_paste.jsonl",
        "writer_conflict.jsonl",
        "events_dropped.jsonl",
    ];

    for filename in &golden_files {
        let lines = load_golden_lines(filename);
        for (i, line) in lines.iter().enumerate() {
            assert_valid_json(line, &format!("{}:{}", filename, i + 1));
        }
    }
}

/// Verify round-trip: deserialize then serialize back should produce equivalent JSON.
/// This catches any field name mismatches or missing fields.
#[test]
fn test_round_trip_hello() {
    let lines = load_golden_lines("hello_ok.jsonl");

    // Client message round-trip
    let original: Value = serde_json::from_str(&lines[0]).unwrap();
    let typed: ClientMessage = serde_json::from_str(&lines[0]).unwrap();
    let reserialized: Value = serde_json::to_value(&typed).unwrap();

    // Compare specific fields (not full equality, as serialization order may differ)
    assert_eq!(original["type"], reserialized["type"]);
    assert_eq!(original["id"], reserialized["id"]);
    assert_eq!(original["client"], reserialized["client"]);
    assert_eq!(original["token"], reserialized["token"]);
    assert_eq!(original["protocol_version"], reserialized["protocol_version"]);

    // Server message round-trip
    let original: Value = serde_json::from_str(&lines[1]).unwrap();
    let typed: ServerMessage = serde_json::from_str(&lines[1]).unwrap();
    let reserialized: Value = serde_json::to_value(&typed).unwrap();

    assert_eq!(original["type"], reserialized["type"]);
    assert_eq!(original["id"], reserialized["id"]);
    assert_eq!(original["session_id"], reserialized["session_id"]);
    assert_eq!(original["protocol_version"], reserialized["protocol_version"]);
    assert_eq!(original["revision"], reserialized["revision"]);
    assert_eq!(original["capabilities"], reserialized["capabilities"]);
}

#[test]
fn test_round_trip_apply_ops() {
    let lines = load_golden_lines("apply_ops_ok.jsonl");

    // Client message round-trip
    let original: Value = serde_json::from_str(&lines[0]).unwrap();
    let typed: ClientMessage = serde_json::from_str(&lines[0]).unwrap();
    let reserialized: Value = serde_json::to_value(&typed).unwrap();

    assert_eq!(original["type"], reserialized["type"]);
    assert_eq!(original["id"], reserialized["id"]);
    assert_eq!(original["atomic"], reserialized["atomic"]);
    assert_eq!(original["expected_revision"], reserialized["expected_revision"]);
    assert_eq!(original["ops"].as_array().unwrap().len(), reserialized["ops"].as_array().unwrap().len());

    // Server message round-trip
    let original: Value = serde_json::from_str(&lines[1]).unwrap();
    let typed: ServerMessage = serde_json::from_str(&lines[1]).unwrap();
    let reserialized: Value = serde_json::to_value(&typed).unwrap();

    assert_eq!(original["type"], reserialized["type"]);
    assert_eq!(original["id"], reserialized["id"]);
    assert_eq!(original["applied"], reserialized["applied"]);
    assert_eq!(original["total"], reserialized["total"]);
    assert_eq!(original["current_revision"], reserialized["current_revision"]);
}

// =============================================================================
// Byte-Exact Serialization Tests (Tripwire for wire format drift)
// =============================================================================
//
// These tests verify that our serialization produces EXACTLY the same bytes
// as the golden vectors. This catches:
// - Key ordering changes in serde_json
// - Accidental field renames
// - Float formatting drift
// - Missing/extra fields
//
// Rule: Do NOT use HashMap in protocol types. Use structs or BTreeMap only.
// Rule: Keep #[serde(rename = "...")] and field order stable.

use visigrid_protocol::{
    HelloMessage, ApplyOpsMessage, InspectMessage, Op, ErrorMessage,
};

/// Byte-exact test for ClientMessage::Hello serialization.
/// Verifies CLI can produce exactly what the server expects.
#[test]
fn test_client_hello_byte_exact() {
    let golden = load_golden_lines("hello_ok.jsonl")[0].clone();

    // Construct the exact message from golden
    let msg = ClientMessage::Hello(HelloMessage {
        id: "req-1".to_string(),
        client: "test-agent".to_string(),
        version: "1.0.0".to_string(),
        token: "dGVzdC10b2tlbi1mb3ItZ29sZGVuLXZlY3RvcnM=".to_string(),
        protocol_version: 1,
    });

    let serialized = serde_json::to_string(&msg).expect("serialization failed");

    assert_eq!(
        serialized, golden,
        "\nByte-exact serialization mismatch for Hello!\n\
         Expected (golden): {}\n\
         Got (serialized):  {}\n\
         This indicates wire format drift. Check field order in HelloMessage.",
        golden, serialized
    );
}

/// Byte-exact test for ClientMessage::ApplyOps serialization.
#[test]
fn test_client_apply_ops_byte_exact() {
    let golden = load_golden_lines("apply_ops_ok.jsonl")[0].clone();

    let msg = ClientMessage::ApplyOps(ApplyOpsMessage {
        id: "req-2".to_string(),
        ops: vec![
            Op::SetCellValue {
                sheet: 0,
                row: 0,
                col: 0,
                value: "Hello".to_string(),
            },
            Op::SetCellFormula {
                sheet: 0,
                row: 1,
                col: 0,
                formula: "=A1&\" World\"".to_string(),
            },
            Op::ClearCell {
                sheet: 0,
                row: 2,
                col: 0,
            },
        ],
        atomic: true,
        expected_revision: Some(42),
    });

    let serialized = serde_json::to_string(&msg).expect("serialization failed");

    assert_eq!(
        serialized, golden,
        "\nByte-exact serialization mismatch for ApplyOps!\n\
         Expected (golden): {}\n\
         Got (serialized):  {}\n\
         This indicates wire format drift. Check field order in ApplyOpsMessage or Op variants.",
        golden, serialized
    );
}

/// Byte-exact test for ClientMessage::Inspect (cell target) serialization.
#[test]
fn test_client_inspect_cell_byte_exact() {
    let golden = load_golden_lines("inspect.jsonl")[0].clone();

    let msg = ClientMessage::Inspect(InspectMessage {
        id: "req-5".to_string(),
        target: InspectTarget::Cell {
            sheet: 0,
            row: 0,
            col: 0,
        },
    });

    let serialized = serde_json::to_string(&msg).expect("serialization failed");

    assert_eq!(
        serialized, golden,
        "\nByte-exact serialization mismatch for Inspect (cell)!\n\
         Expected (golden): {}\n\
         Got (serialized):  {}\n\
         This indicates wire format drift. Check field order in InspectMessage or InspectTarget.",
        golden, serialized
    );
}

/// Byte-exact test for ClientMessage::Inspect (workbook target) serialization.
#[test]
fn test_client_inspect_workbook_byte_exact() {
    let golden = load_golden_lines("inspect.jsonl")[4].clone();

    let msg = ClientMessage::Inspect(InspectMessage {
        id: "req-7".to_string(),
        target: InspectTarget::Workbook,
    });

    let serialized = serde_json::to_string(&msg).expect("serialization failed");

    assert_eq!(
        serialized, golden,
        "\nByte-exact serialization mismatch for Inspect (workbook)!\n\
         Expected (golden): {}\n\
         Got (serialized):  {}\n\
         This indicates wire format drift. Check InspectTarget::Workbook serialization.",
        golden, serialized
    );
}

/// Byte-exact test for ServerMessage::Error serialization.
/// Tests server→CLI direction with optional fields (retry_after_ms).
#[test]
fn test_server_error_byte_exact() {
    // Line 10 in errors.jsonl: writer_conflict with retry_after_ms
    let golden = load_golden_lines("errors.jsonl")[9].clone();

    let msg = ServerMessage::Error(ErrorMessage {
        id: "req-10".to_string(),
        code: "writer_conflict".to_string(),
        message: "Write lease held by another connection".to_string(),
        retry_after_ms: Some(5000),
    });

    let serialized = serde_json::to_string(&msg).expect("serialization failed");

    assert_eq!(
        serialized, golden,
        "\nByte-exact serialization mismatch for Error (writer_conflict)!\n\
         Expected (golden): {}\n\
         Got (serialized):  {}\n\
         This indicates wire format drift. Check ErrorMessage field order or optional fields.",
        golden, serialized
    );
}
