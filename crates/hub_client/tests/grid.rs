//! The Grid client against a mock Grid: the credential split, the access
//! token clock, and the exact sheet routes and bodies Grid accepts.
use httpmock::prelude::*;
use serde_json::json;
use visigrid_hub_client::grid::{self, GridClient, GridCredentials, GridError};

const PID: &str = "10000000-0000-4000-8000-000000000001";
const SECRET: &str = "s3cretS3cretS3cretS3cretS3cretS3cretS3cretS3cretS3cretS3cretAbcd";

fn client(server: &MockServer) -> GridClient {
    GridClient::new(GridCredentials { secret: SECRET.into(), api_base: server.base_url() })
}

fn sheet(revision: i64) -> serde_json::Value {
    json!({"id": 3, "pid": PID, "name": "Budget", "revision": revision, "folder_id": null,
           "role": "owner", "can_manage_sharing": true, "created_at": "2026-10-01T00:00:00Z", "updated_at": "2026-10-01T00:00:00Z"})
}

fn token_mock<'a>(server: &'a MockServer, token: &str) -> httpmock::Mock<'a> {
    server.mock(|when, then| {
        when.method(POST).path("/api/device-sessions/token").json_body(json!({"secret": SECRET}));
        then.status(200).json_body(json!({"token": token, "expires_in": 600}));
    })
}

#[test]
fn the_grid_credential_is_its_own_0600_file_and_never_auth_json() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("visigrid/grid-auth.json");
    let creds = GridCredentials { secret: SECRET.into(), api_base: "https://grid.test".into() };
    grid::save_grid_auth_to(&path, &creds).unwrap();
    assert_eq!(grid::load_grid_auth_from(&path), Some(creds.clone()));
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        assert_eq!(std::fs::metadata(&path).unwrap().permissions().mode() & 0o777, 0o600);
    }
    // Not the Rails Hub token file, which Hub publish and `vgrid login` keep.
    assert_ne!(grid::grid_auth_path(), visigrid_hub_client::auth_file_path());
    assert!(grid::grid_auth_path().unwrap().ends_with("visigrid/grid-auth.json"));
    // The secret never reaches a log line.
    assert!(!format!("{creds:?}").contains(SECRET));
}

#[test]
fn sheet_calls_use_a_cached_short_lived_token_from_the_secret() {
    let server = MockServer::start();
    let token = token_mock(&server, "access-1");
    let list = server.mock(|when, then| {
        when.method(GET).path("/api/sheets").header("authorization", "Bearer access-1");
        then.status(200).json_body(json!([sheet(4)]));
    });
    let c = client(&server);
    assert_eq!(c.list().unwrap()[0].revision, 4);
    assert_eq!(c.list().unwrap()[0].pid, PID);
    token.assert_calls(1);
    list.assert_calls(2);
}

#[test]
fn a_rejected_access_token_is_refreshed_once_then_reported_signed_out() {
    let server = MockServer::start();
    let token = token_mock(&server, "access-1");
    let sheets = server.mock(|when, then| {
        when.method(GET).path(format!("/api/sheets/{PID}"));
        then.status(401).json_body(json!({"error": "unauthorized"}));
    });
    match client(&server).get(PID) {
        Err(GridError::SignedOut) => {}
        other => panic!("{other:?}"),
    }
    token.assert_calls(2);
    sheets.assert_calls(2);
}

#[test]
fn a_revoked_device_cannot_refresh() {
    let server = MockServer::start();
    server.mock(|when, then| {
        when.method(POST).path("/api/device-sessions/token");
        then.status(401).json_body(json!({"error": "unauthorized"}));
    });
    assert!(matches!(client(&server).list(), Err(GridError::SignedOut)));
    assert!(matches!(client(&server).verify(), Err(GridError::SignedOut)));
}

#[test]
fn small_saves_go_inline_with_exactly_the_fields_grid_accepts() {
    let server = MockServer::start();
    token_mock(&server, "access-1");
    let document = json!({"version": 1, "sheets": [{"name": "Sheet1"}]});
    let save = server.mock(|when, then| {
        when.method(POST)
            .path(format!("/api/sheets/{PID}/save"))
            .header("authorization", "Bearer access-1")
            .json_body(json!({"expected_revision": 4, "document": document.clone()}));
        then.status(200).json_body(sheet(5));
    });
    let saved = client(&server).save(PID, 4, &serde_json::to_vec(&document).unwrap()).unwrap();
    assert_eq!(saved.revision, 5);
    save.assert_calls(1);
}

#[test]
fn a_stale_save_is_a_conflict() {
    let server = MockServer::start();
    token_mock(&server, "access-1");
    server.mock(|when, then| {
        when.method(POST).path(format!("/api/sheets/{PID}/save"));
        then.status(409).json_body(json!({"error": "revision_conflict", "description": "The sheet changed; reload before saving"}));
    });
    assert!(matches!(client(&server).save(PID, 4, b"{}"), Err(GridError::Conflict)));
}

#[test]
fn large_saves_go_through_a_signed_upload_without_the_bearer() {
    let server = MockServer::start();
    token_mock(&server, "access-1");
    let document = serde_json::to_vec(&json!({"blob": "x".repeat(grid::INLINE_SAVE_BYTES)})).unwrap();
    let upload_id = "20000000-0000-4000-8000-000000000002";
    let reserve = server.mock(|when, then| {
        when.method(POST)
            .path(format!("/api/sheets/{PID}/uploads"))
            .header("authorization", "Bearer access-1")
            .json_body_includes(r#"{"expected_revision": 4}"#);
        then.status(201).json_body(json!({"upload_id": upload_id, "sheet_id": PID, "expected_revision": 4,
            "url": server.url(format!("/storage/{upload_id}.json")), "expires_in": 600}));
    });
    let put = server.mock(|when, then| {
        when.method(PUT).path(format!("/storage/{upload_id}.json")).header_missing("authorization");
        then.status(200);
    });
    let complete = server.mock(|when, then| {
        when.method(POST).path(format!("/api/sheets/{PID}/uploads/{upload_id}/complete"));
        then.status(200).json_body(sheet(5));
    });
    assert_eq!(client(&server).save(PID, 4, &document).unwrap().revision, 5);
    reserve.assert_calls(1);
    put.assert_calls(1);
    complete.assert_calls(1);
}

#[test]
fn old_rails_ids_resolve_through_the_legacy_route() {
    let server = MockServer::start();
    token_mock(&server, "access-1");
    for id in ["90", "AbCdEfGhIjKlMnOpQrSt_-"] {
        server.mock(|when, then| {
            when.method(GET).path(format!("/api/sheets/legacy/{id}"));
            then.status(200).json_body(json!({"pid": PID}));
        });
        assert_eq!(client(&server).resolve_legacy(id).unwrap(), PID);
    }
    // Anything that is not an id never reaches a URL.
    assert!(client(&server).resolve_legacy("../me").is_err());
}

#[test]
fn signing_out_revokes_the_device() {
    let server = MockServer::start();
    let revoke = server.mock(|when, then| {
        when.method(POST).path("/api/device-sessions/revoke").json_body(json!({"secret": SECRET}));
        then.status(200).json_body(json!({}));
    });
    client(&server).sign_out().unwrap();
    revoke.assert_calls(1);
}
