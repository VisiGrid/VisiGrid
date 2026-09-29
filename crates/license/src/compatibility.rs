//! Shared Ruby/Rust golden vectors. Keys are injected only inside this test module.
use super::*;
use ed25519_dalek::{Signer, SigningKey};

fn fixtures() -> serde_json::Value {
    serde_json::from_str(include_str!("../tests/fixtures/compatibility.json")).unwrap()
}

fn fixture_key(fixture: &serde_json::Value, id: &str) -> Option<[u8; 32]> {
    fixture["public_keys"][id].as_str().map(|key| {
        base64::engine::general_purpose::STANDARD
            .decode(key)
            .unwrap()
            .try_into()
            .unwrap()
    })
}

#[test]
fn ruby_rust_golden_signatures_and_validation() {
    let fixture = fixtures();
    let now = fixture["now"]
        .as_str()
        .unwrap()
        .parse::<DateTime<Utc>>()
        .unwrap();
    let seed = [0x42; 32]; // public test seed, matches generator and corpus
    assert_eq!(fixture["test_seed_hex"], "42".repeat(32));
    for case in fixture["cases"].as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let envelope = &case["envelope"];
        let message = canonical_json(&envelope["payload"]).unwrap();
        assert_eq!(
            message,
            case["canonical"].as_str().unwrap().as_bytes(),
            "{name}: bytes differ from Ruby"
        );
        let license = LicenseFile::from_json(&envelope.to_string()).unwrap();
        let key = fixture_key(&fixture, &license.key_id);
        assert_eq!(
            license.verify_signature(key).is_ok(),
            case["signature_valid"].as_bool().unwrap(),
            "{name}: signature"
        );
        let result = license.validate_at(now, key);
        assert_eq!(
            result.valid,
            case["license_valid"].as_bool().unwrap(),
            "{name}: {result:?}"
        );
        assert_eq!(
            result.in_grace_period,
            case["in_grace_period"].as_bool().unwrap(),
            "{name}: grace"
        );
        if case["signature_valid"] == true {
            // A future Rust issuer must produce the exact same Ed25519 signature.
            let signing_key = SigningKey::from_bytes(if license.key_id == "fixture-only-002" {
                &[0x43; 32]
            } else {
                &seed
            });
            let signature = signing_key.sign(&message);
            assert_eq!(
                base64::engine::general_purpose::STANDARD.encode(signature.to_bytes()),
                envelope["signature"],
                "{name}: Rust signing differs from Ruby"
            );
        }
        // Round trips preserve the original signed payload, including unknowns/nulls.
        assert_eq!(
            serde_json::to_value(&license).unwrap(),
            *envelope,
            "{name}: round trip"
        );
        assert!(
            !license.validate().valid,
            "{name}: test keys must never be trusted by normal validation"
        );
    }
}

#[test]
fn typed_payload_mutation_cannot_bypass_signature() {
    let fixture = fixtures();
    let mut license = LicenseFile::from_json(&fixture["cases"][0]["envelope"].to_string()).unwrap();
    let key = fixture_key(&fixture, &license.key_id);
    assert!(license.verify_signature(key).is_ok());
    license.payload.edition = Edition::ProPlus;
    assert!(license.verify_signature(key).is_err());
}

#[test]
fn malformed_license_envelopes_are_rejected() {
    for input in ["", "{}", "null", "[]", "{", r#"{"payload":null}"#] {
        assert!(LicenseFile::from_json(input).is_err());
    }
}

#[test]
fn shipped_trust_store_excludes_public_fixture_and_placeholder_keys() {
    for id in [
        "fixture-only-001",
        "fixture-only-002",
        "test-vector-001",
        "dev-001",
    ] {
        assert!(get_public_key(id).is_none());
    }
}
