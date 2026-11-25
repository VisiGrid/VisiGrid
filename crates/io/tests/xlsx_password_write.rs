use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::process::Command;
use visigrid_io::xlsx_crypto::{decrypt_ooxml, encrypt_ooxml, CryptoError};

const REFERENCE: &[u8] = include_bytes!("fixtures/encrypted_msoffcrypto_agile.xlsx");

#[test]
fn round_trips_segment_boundaries() {
    for size in [1, 4095, 4096, 4097, 4096 * 5 + 15] {
        let plain: Vec<u8> = (0..size).map(|i| (i % 251) as u8).collect();
        let encrypted = encrypt_ooxml(&plain, "päss🔑").unwrap();
        assert_eq!(decrypt_ooxml(&encrypted, "päss🔑").unwrap(), plain);
    }
}

#[test]
fn fresh_randomness_password_verification_and_package_integrity() {
    let mut bytes = encrypt_ooxml(b"same input", "same password").unwrap();
    assert_ne!(
        bytes,
        encrypt_ooxml(b"same input", "same password").unwrap()
    );
    assert_eq!(
        decrypt_ooxml(&bytes, "wrong").unwrap_err(),
        CryptoError::IncorrectPassword
    );
    {
        let mut file = cfb::CompoundFile::open(Cursor::new(&mut bytes)).unwrap();
        let mut stream = file.open_stream("EncryptedPackage").unwrap();
        stream.seek(SeekFrom::Start(8)).unwrap();
        let mut byte = [0];
        stream.read_exact(&mut byte).unwrap();
        stream.seek(SeekFrom::Start(8)).unwrap();
        stream.write_all(&[byte[0] ^ 1]).unwrap();
    }
    assert_eq!(
        decrypt_ooxml(&bytes, "same password").unwrap_err(),
        CryptoError::IntegrityCheckFailed
    );
}

#[test]
fn output_password_limit_counts_utf16_units() {
    let at_limit = format!("{}x", "🔑".repeat(127));
    let bytes = encrypt_ooxml(b"x", &at_limit).unwrap();
    assert_eq!(decrypt_ooxml(&bytes, &at_limit).unwrap(), b"x");
    for invalid in ["", &"a".repeat(256), &"🔑".repeat(128)] {
        assert_eq!(
            encrypt_ooxml(b"x", invalid).unwrap_err(),
            CryptoError::InvalidPasswordLength
        );
    }
}

#[test]
fn dataspaces_layout_and_contents_match_reference() {
    let bytes = encrypt_ooxml(b"x", "test").unwrap();
    let mut actual = cfb::CompoundFile::open(Cursor::new(bytes)).unwrap();
    let mut reference = cfb::CompoundFile::open(Cursor::new(REFERENCE)).unwrap();
    let paths = |entry: cfb::Entry| (entry.path().to_owned(), entry.is_stream());
    let expected: std::collections::BTreeSet<_> = reference.walk().map(paths).collect();
    assert_eq!(
        actual
            .walk()
            .map(paths)
            .collect::<std::collections::BTreeSet<_>>(),
        expected
    );
    for (path, is_stream) in expected {
        if is_stream && path.components().count() > 2 {
            let mut new = Vec::new();
            let mut old = Vec::new();
            actual
                .open_stream(&path)
                .unwrap()
                .read_to_end(&mut new)
                .unwrap();
            reference
                .open_stream(&path)
                .unwrap()
                .read_to_end(&mut old)
                .unwrap();
            assert_eq!(new, old, "DataSpaces stream {path:?}");
        }
    }
}

#[test]
fn independent_msoffcrypto_decryption() {
    // Run with a Python virtualenv containing msoffcrypto-tool on PATH.
    let check = Command::new("python3").args(["-c",
        "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('msoffcrypto') else 77)"
    ]).status();
    match check {
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
            eprintln!("skipped: python3 unavailable");
            return;
        }
        Ok(status) if status.code() == Some(77) => {
            eprintln!("skipped: msoffcrypto-tool unavailable");
            return;
        }
        other => assert!(other.unwrap().success()),
    }

    let plain = decrypt_ooxml(REFERENCE, "hunter2").unwrap();
    let dir = tempfile::tempdir().unwrap();
    let protected = dir.path().join("protected.xlsx");
    let original = dir.path().join("original.xlsx");
    std::fs::write(
        &protected,
        encrypt_ooxml(&plain, "interoperability").unwrap(),
    )
    .unwrap();
    std::fs::write(&original, plain).unwrap();
    let status = Command::new("python3").arg("-c").arg(r#"
import hashlib, io, pathlib, sys, msoffcrypto
with open(sys.argv[1], 'rb') as source:
    book = msoffcrypto.OfficeFile(source)
    book.load_key(password='interoperability', verify_password=True)
    result = io.BytesIO()
    book.decrypt(result, verify_integrity=True)
data = result.getvalue()
assert data == pathlib.Path(sys.argv[2]).read_bytes()
print(f'msoffcrypto: password and HMAC verified; {len(data)} identical bytes; SHA-256 {hashlib.sha256(data).hexdigest()}')
"#).arg(protected).arg(original).status().unwrap();
    assert!(status.success());
}
