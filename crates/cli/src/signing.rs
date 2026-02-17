//! Shared signing primitives for `vgrid sign` and `vgrid verify totals --sign`.
//!
//! Provides Ed25519 signing with BLAKE3 payload hashing, keypair management,
//! and a generic `SignedEnvelope` format for audit-grade proofs.

use std::path::{Path, PathBuf};

use base64::Engine;
use ed25519_dalek::Signer;
use serde::{Deserialize, Serialize};

use crate::CliError;

// ── Types ───────────────────────────────────────────────────────────

/// On-disk keypair format (JSON).
#[derive(Debug, Serialize, Deserialize)]
pub struct StoredKeypair {
    pub public_key: String,
    pub secret_key: String,
}

/// Generic signed envelope — wraps any JSON payload with Ed25519 signature.
#[derive(Debug, Serialize, Deserialize)]
pub struct SignedEnvelope {
    /// Schema identifier, e.g. "vgrid.file_proof.v1"
    pub schema: String,
    /// The signed payload (arbitrary JSON object)
    pub payload: serde_json::Value,
    /// BLAKE3 hash of the compact payload JSON bytes (debugging aid)
    pub payload_blake3: String,
    /// Base64-encoded Ed25519 signature of the compact payload JSON bytes
    pub signature: String,
    /// Base64-encoded Ed25519 verifying (public) key
    pub public_key: String,
    /// First 16 hex chars of BLAKE3(public_key_bytes) — identifies the signing key
    pub key_id: String,
}

// ── Functions ───────────────────────────────────────────────────────

/// Load an existing keypair from disk, or generate a new one.
///
/// Default path: `~/.config/vgrid/proof_key.json`
pub fn load_or_generate_key(
    key_path: &Option<PathBuf>,
) -> Result<(ed25519_dalek::SigningKey, ed25519_dalek::VerifyingKey), CliError> {
    let b64 = base64::engine::general_purpose::STANDARD;

    let path = match key_path {
        Some(p) => p.clone(),
        None => {
            let config_dir = dirs::config_dir()
                .ok_or_else(|| {
                    CliError::io("cannot determine config directory".to_string())
                })?
                .join("vgrid");
            config_dir.join("proof_key.json")
        }
    };

    if path.exists() {
        let data = std::fs::read_to_string(&path)
            .map_err(|e| CliError::io(format!("cannot read {}: {e}", path.display())))?;
        let stored: StoredKeypair = serde_json::from_str(&data)
            .map_err(|e| CliError::parse(format!("invalid key file {}: {e}", path.display())))?;
        let secret_bytes = b64
            .decode(&stored.secret_key)
            .map_err(|e| CliError::parse(format!("invalid secret key base64: {e}")))?;
        let secret_array: [u8; 32] = secret_bytes
            .try_into()
            .map_err(|_| CliError::parse("secret key must be 32 bytes".to_string()))?;
        let signing_key = ed25519_dalek::SigningKey::from_bytes(&secret_array);
        let verifying_key = signing_key.verifying_key();
        Ok((signing_key, verifying_key))
    } else {
        let mut rng = rand::thread_rng();
        let signing_key = ed25519_dalek::SigningKey::generate(&mut rng);
        let verifying_key = signing_key.verifying_key();

        let stored = StoredKeypair {
            public_key: b64.encode(verifying_key.to_bytes()),
            secret_key: b64.encode(signing_key.to_bytes()),
        };

        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)
                .map_err(|e| CliError::io(format!("cannot create {}: {e}", parent.display())))?;
        }

        let json = serde_json::to_string_pretty(&stored)
            .map_err(|e| CliError::io(format!("key serialization error: {e}")))?;
        std::fs::write(&path, json)
            .map_err(|e| CliError::io(format!("cannot write {}: {e}", path.display())))?;

        eprintln!("  generated new signing key: {}", path.display());
        Ok((signing_key, verifying_key))
    }
}

/// Sign a JSON payload and produce a `SignedEnvelope`.
///
/// 1. Serialize payload to compact JSON bytes (deterministic)
/// 2. BLAKE3 hash the payload bytes
/// 3. Ed25519 sign the payload bytes
/// 4. Base64 encode signature + public key
/// 5. Compute key_id from public key
pub fn sign_payload(
    schema: &str,
    payload: &serde_json::Value,
    signing_key: &ed25519_dalek::SigningKey,
    verifying_key: &ed25519_dalek::VerifyingKey,
) -> Result<SignedEnvelope, CliError> {
    let b64 = base64::engine::general_purpose::STANDARD;

    // Compact JSON for deterministic signing
    let payload_bytes = serde_json::to_vec(payload)
        .map_err(|e| CliError::io(format!("payload serialization error: {e}")))?;

    // BLAKE3 of payload bytes
    let payload_hash = blake3::hash(&payload_bytes);
    let payload_blake3 = payload_hash.to_hex().to_string();

    // Ed25519 sign
    let signature = signing_key.sign(&payload_bytes);

    Ok(SignedEnvelope {
        schema: schema.to_string(),
        payload: payload.clone(),
        payload_blake3,
        signature: b64.encode(signature.to_bytes()),
        public_key: b64.encode(verifying_key.to_bytes()),
        key_id: key_id(verifying_key),
    })
}

/// BLAKE3 hash of a file's contents, returned as a hex string.
pub fn hash_file_blake3(path: &Path) -> Result<String, CliError> {
    let bytes = std::fs::read(path)
        .map_err(|e| CliError::io(format!("cannot read {}: {e}", path.display())))?;
    Ok(blake3::hash(&bytes).to_hex().to_string())
}

/// Key identifier: first 16 hex characters of BLAKE3(public_key_bytes).
pub fn key_id(verifying_key: &ed25519_dalek::VerifyingKey) -> String {
    let hash = blake3::hash(&verifying_key.to_bytes());
    hash.to_hex()[..16].to_string()
}

/// Verify an Ed25519 signature in a `SignedEnvelope`.
///
/// Returns `Ok(())` if valid, `Err` with detail if not.
pub fn verify_envelope_signature(envelope: &SignedEnvelope) -> Result<(), String> {
    use ed25519_dalek::Verifier;
    let b64 = base64::engine::general_purpose::STANDARD;

    let pub_bytes = b64
        .decode(&envelope.public_key)
        .map_err(|e| format!("invalid public key base64: {e}"))?;
    let pub_array: [u8; 32] = pub_bytes
        .try_into()
        .map_err(|_| "public key must be 32 bytes".to_string())?;
    let verifying_key = ed25519_dalek::VerifyingKey::from_bytes(&pub_array)
        .map_err(|e| format!("invalid public key: {e}"))?;

    let sig_bytes = b64
        .decode(&envelope.signature)
        .map_err(|e| format!("invalid signature base64: {e}"))?;
    let sig_array: [u8; 64] = sig_bytes
        .try_into()
        .map_err(|_| "signature must be 64 bytes".to_string())?;
    let signature = ed25519_dalek::Signature::from_bytes(&sig_array);

    // Recompute compact JSON of payload (same as signing)
    let payload_bytes = serde_json::to_vec(&envelope.payload)
        .map_err(|e| format!("payload serialization error: {e}"))?;

    verifying_key
        .verify(&payload_bytes, &signature)
        .map_err(|e| format!("signature verification failed: {e}"))
}

// ── Tests ───────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    fn test_keypair() -> (ed25519_dalek::SigningKey, ed25519_dalek::VerifyingKey) {
        let mut rng = rand::thread_rng();
        let signing_key = ed25519_dalek::SigningKey::generate(&mut rng);
        let verifying_key = signing_key.verifying_key();
        (signing_key, verifying_key)
    }

    #[test]
    fn test_sign_and_verify_roundtrip() {
        let (sk, vk) = test_keypair();
        let payload = serde_json::json!({"hello": "world", "n": 42});
        let envelope = sign_payload("test.v1", &payload, &sk, &vk).unwrap();

        assert_eq!(envelope.schema, "test.v1");
        assert_eq!(envelope.payload, payload);
        assert!(!envelope.payload_blake3.is_empty());
        assert!(!envelope.signature.is_empty());
        assert!(!envelope.public_key.is_empty());
        assert_eq!(envelope.key_id.len(), 16);

        // Verify signature
        verify_envelope_signature(&envelope).unwrap();
    }

    #[test]
    fn test_deterministic_signing() {
        let (sk, vk) = test_keypair();
        let payload = serde_json::json!({"a": 1, "b": [2, 3]});

        let e1 = sign_payload("det.v1", &payload, &sk, &vk).unwrap();
        let e2 = sign_payload("det.v1", &payload, &sk, &vk).unwrap();

        assert_eq!(e1.signature, e2.signature);
        assert_eq!(e1.payload_blake3, e2.payload_blake3);
        assert_eq!(e1.key_id, e2.key_id);
    }

    #[test]
    fn test_key_id_from_pubkey() {
        let (_, vk) = test_keypair();
        let id = key_id(&vk);
        assert_eq!(id.len(), 16);
        // All hex chars
        assert!(id.chars().all(|c| c.is_ascii_hexdigit()));
        // Deterministic
        assert_eq!(id, key_id(&vk));
    }

    #[test]
    fn test_hash_file_blake3() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.txt");
        std::fs::write(&path, b"hello world").unwrap();

        let hash = hash_file_blake3(&path).unwrap();
        assert_eq!(hash.len(), 64); // BLAKE3 hex = 64 chars

        // Same content → same hash
        let hash2 = hash_file_blake3(&path).unwrap();
        assert_eq!(hash, hash2);

        // Known BLAKE3 hash of "hello world"
        let expected = blake3::hash(b"hello world").to_hex().to_string();
        assert_eq!(hash, expected);
    }

    #[test]
    fn test_tampered_payload_fails() {
        let (sk, vk) = test_keypair();
        let payload = serde_json::json!({"data": "original"});
        let mut envelope = sign_payload("tamper.v1", &payload, &sk, &vk).unwrap();

        // Tamper with payload
        envelope.payload = serde_json::json!({"data": "tampered"});

        let result = verify_envelope_signature(&envelope);
        assert!(result.is_err());
    }

    #[test]
    fn test_tampered_signature_bytes_fails() {
        let (sk, vk) = test_keypair();
        let payload = serde_json::json!({"data": "original"});
        let mut envelope = sign_payload("tamper.v1", &payload, &sk, &vk).unwrap();

        // Flip first character of base64 signature
        let mut sig_chars: Vec<char> = envelope.signature.chars().collect();
        sig_chars[0] = if sig_chars[0] == 'A' { 'B' } else { 'A' };
        envelope.signature = sig_chars.into_iter().collect();

        let result = verify_envelope_signature(&envelope);
        assert!(result.is_err());
    }

    #[test]
    fn test_wrong_key_fails() {
        let (sk, vk) = test_keypair();
        let (_, other_vk) = test_keypair();
        let payload = serde_json::json!({"data": "original"});
        let mut envelope = sign_payload("tamper.v1", &payload, &sk, &vk).unwrap();

        // Replace public key with a different key
        let b64 = base64::engine::general_purpose::STANDARD;
        envelope.public_key = b64.encode(other_vk.to_bytes());

        let result = verify_envelope_signature(&envelope);
        assert!(result.is_err());
    }

    #[test]
    fn test_file_proof_roundtrip() {
        // Full flow: create a file, sign it, verify the envelope matches
        let dir = tempfile::tempdir().unwrap();
        let file_path = dir.path().join("test_artifact.dat");
        let content = b"deterministic financial data\n";
        std::fs::write(&file_path, content).unwrap();

        let (sk, vk) = test_keypair();

        // Build payload matching cmd_sign structure
        let file_hash = hash_file_blake3(&file_path).unwrap();
        let file_size = std::fs::metadata(&file_path).unwrap().len();

        let payload = serde_json::json!({
            "signer": { "name": "vgrid", "version": "test" },
            "signed_at": "2026-01-01T00:00:00Z",
            "file": {
                "name": "test_artifact.dat",
                "blake3": file_hash,
                "size_bytes": file_size,
            }
        });

        let envelope = sign_payload("vgrid.file_proof.v1", &payload, &sk, &vk).unwrap();

        // Verify signature
        verify_envelope_signature(&envelope).unwrap();

        // Verify file hash matches
        assert_eq!(
            envelope.payload["file"]["blake3"].as_str().unwrap(),
            &file_hash
        );

        // Verify payload_blake3 is consistent
        let payload_bytes = serde_json::to_vec(&envelope.payload).unwrap();
        let expected_payload_hash = blake3::hash(&payload_bytes).to_hex().to_string();
        assert_eq!(envelope.payload_blake3, expected_payload_hash);

        // Now tamper with the file and verify hash no longer matches
        std::fs::write(&file_path, b"tampered data\n").unwrap();
        let tampered_hash = hash_file_blake3(&file_path).unwrap();
        assert_ne!(tampered_hash, file_hash, "tampered file should have different hash");
    }
}
