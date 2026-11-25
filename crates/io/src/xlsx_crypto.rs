//! MS-OFFCRYPTO (ECMA-376) encryption and decryption for password-protected OOXML workbooks.
//!
//! Writes Agile AES-256-CBC/SHA-512 with authenticated package data.
//! Reads:
//! - Agile Encryption (Excel 2010–present): AES-128/192/256-CBC with SHA-1,
//!   SHA-256, SHA-384, or SHA-512, constant-time verifier check, and
//!   constant-time `dataIntegrity` HMAC verification when present.
//! - Standard Encryption (Excel 2007): AES-128/192/256-ECB with SHA-1 and
//!   constant-time verifier check. (Standard encryption has no format-level
//!   HMAC integrity check.)
//!
//! Legacy binary `.xls` (BIFF8) encryption is out of scope and reported with a
//! dedicated error.

use std::fmt;
use std::io::{Cursor, Read};

use aes::cipher::{BlockDecrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use base64::{engine::general_purpose::STANDARD as BASE64, Engine as _};
use cbc::cipher::{BlockDecryptMut, KeyIvInit};
use hmac::{Hmac, Mac};
use quick_xml::events::Event;
use quick_xml::Reader;
use sha1::Sha1;
use sha2::{Digest, Sha256, Sha384, Sha512};
use subtle::ConstantTimeEq;
pub use zeroize::Zeroizing;

mod write;
pub use write::encrypt_ooxml;

const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Maximum accepted key-stretching spin count (bounds CPU DoS on hostile files).
pub const MAX_SPIN_COUNT: u32 = 1_000_000;
/// Maximum accepted `EncryptionInfo` stream size (1 MiB).
pub const MAX_ENCRYPTION_INFO_BYTES: u64 = 1024 * 1024;
/// Maximum accepted `EncryptedPackage` / declared plaintext size (512 MiB).
pub const MAX_ENCRYPTED_PACKAGE_BYTES: u64 = 512 * 1024 * 1024;

const AGILE_VERIFIER_HASH_INPUT_BLOCK_KEY: [u8; 8] =
    [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
const AGILE_VERIFIER_HASH_VALUE_BLOCK_KEY: [u8; 8] =
    [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
const AGILE_ENCRYPTED_KEY_VALUE_BLOCK_KEY: [u8; 8] =
    [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
const AGILE_HMAC_KEY_BLOCK_KEY: [u8; 8] = [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
const AGILE_HMAC_VALUE_BLOCK_KEY: [u8; 8] = [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];

const PASSWORD_KEY_ENCRYPTOR_URI: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

const STANDARD_SPIN_COUNT: u32 = 50_000;
const ALG_ID_AES_128: u32 = 0x0000_660E;
const ALG_ID_AES_192: u32 = 0x0000_660F;
const ALG_ID_AES_256: u32 = 0x0000_6610;
const ALG_ID_SHA1: u32 = 0x0000_8004;

/// Errors from MS-OFFCRYPTO detection, decryption and encryption.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CryptoError {
    /// A new workbook password must contain 1 through 255 UTF-16 code units.
    InvalidPasswordLength,
    /// The encrypted output could not be constructed.
    EncryptionFailed(String),
    /// Document is encrypted and no password was supplied.
    PasswordRequired,
    /// Supplied password did not match the document's verifier hash.
    IncorrectPassword,
    /// Agile `dataIntegrity` HMAC check failed (tampered or corrupted file).
    IntegrityCheckFailed,
    /// Legacy binary `.xls` encryption is not supported.
    UnsupportedLegacyXlsEncryption,
    /// Encryption scheme/parameters are valid MS-OFFCRYPTO but unsupported.
    UnsupportedEncryption(String),
    /// Malformed CFB container, EncryptionInfo, or EncryptedPackage, or a limit was exceeded.
    Malformed(String),
}

impl fmt::Display for CryptoError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CryptoError::InvalidPasswordLength => write!(f, "Excel output password must contain 1 to 255 UTF-16 code units"),
            CryptoError::EncryptionFailed(msg) => write!(f, "Failed to encrypt Excel workbook: {msg}"),
            CryptoError::PasswordRequired => write!(f, "Workbook is password protected"),
            CryptoError::IncorrectPassword => write!(f, "Incorrect password"),
            CryptoError::IntegrityCheckFailed => {
                write!(f, "Encrypted workbook failed integrity check (file may be tampered or corrupted)")
            }
            CryptoError::UnsupportedLegacyXlsEncryption => write!(
                f,
                "Password-protected legacy .xls files are not supported. Re-save as .xlsx in Excel or LibreOffice to open it."
            ),
            CryptoError::UnsupportedEncryption(msg) => {
                write!(f, "Unsupported Excel encryption: {msg}")
            }
            CryptoError::Malformed(msg) => {
                write!(f, "Malformed encrypted Excel file: {msg}")
            }
        }
    }
}

impl std::error::Error for CryptoError {}

/// Quick check: is `bytes` an OLE2 compound document containing an OOXML
/// `EncryptedPackage` stream (and `EncryptionInfo`)?
pub fn is_encrypted_ooxml(bytes: &[u8]) -> bool {
    if !bytes.starts_with(&OLE2_MAGIC) {
        return false;
    }
    let Ok(comp) = cfb::CompoundFile::open(Cursor::new(bytes)) else {
        return false;
    };
    comp.exists("EncryptedPackage") && comp.exists("EncryptionInfo")
}

/// Decrypt an MS-OFFCRYPTO encrypted OOXML package in memory.
///
/// Returns the decrypted OOXML ZIP bytes (`Zeroizing` is not returned because
/// callers hand the slice to the spreadsheet importer, which builds the
/// in-memory `Workbook`; key material and password UTF-16LE buffers are zeroized
/// before returning).
pub fn decrypt_ooxml(bytes: &[u8], password: &str) -> Result<Vec<u8>, CryptoError> {
    if !bytes.starts_with(&OLE2_MAGIC) {
        return Err(CryptoError::Malformed("not an OLE2 compound file".into()));
    }
    let mut comp = cfb::CompoundFile::open(Cursor::new(bytes))
        .map_err(|e| CryptoError::Malformed(format!("invalid OLE2 container: {e}")))?;

    let info_bytes = read_bounded_stream(&mut comp, "EncryptionInfo", MAX_ENCRYPTION_INFO_BYTES)?;
    let package_bytes =
        read_bounded_stream(&mut comp, "EncryptedPackage", MAX_ENCRYPTED_PACKAGE_BYTES)?;

    if info_bytes.len() < 8 {
        return Err(CryptoError::Malformed("EncryptionInfo header is truncated".into()));
    }
    let version_major = u16::from_le_bytes([info_bytes[0], info_bytes[1]]);
    let version_minor = u16::from_le_bytes([info_bytes[2], info_bytes[3]]);

    match (version_major, version_minor) {
        (4, 4) => decrypt_agile(&info_bytes[8..], &package_bytes, password),
        (2 | 3 | 4, 2) => decrypt_standard(&info_bytes[8..], &package_bytes, password),
        _ => Err(CryptoError::UnsupportedEncryption(format!(
            "unsupported EncryptionInfo version {version_major}.{version_minor}"
        ))),
    }
}

fn read_bounded_stream(
    comp: &mut cfb::CompoundFile<Cursor<&[u8]>>,
    name: &str,
    max_bytes: u64,
) -> Result<Vec<u8>, CryptoError> {
    let stream = comp
        .open_stream(name)
        .map_err(|e| CryptoError::Malformed(format!("missing {name} stream: {e}")))?;
    let declared = stream.len();
    if declared > max_bytes {
        return Err(CryptoError::Malformed(format!(
            "{name} stream ({declared} bytes) exceeds {max_bytes}-byte limit"
        )));
    }
    let mut out = Vec::with_capacity(declared as usize);
    stream
        .take(max_bytes + 1)
        .read_to_end(&mut out)
        .map_err(|e| CryptoError::Malformed(format!("failed reading {name} stream: {e}")))?;
    if out.len() as u64 > max_bytes {
        return Err(CryptoError::Malformed(format!(
            "{name} stream exceeds {max_bytes}-byte limit"
        )));
    }
    Ok(out)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum HashAlg {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlg {
    fn from_name(name: &str) -> Result<Self, CryptoError> {
        match name {
            "SHA1" | "SHA-1" => Ok(Self::Sha1),
            "SHA256" | "SHA-256" => Ok(Self::Sha256),
            "SHA384" | "SHA-384" => Ok(Self::Sha384),
            "SHA512" | "SHA-512" => Ok(Self::Sha512),
            other => Err(CryptoError::UnsupportedEncryption(format!(
                "unsupported hash algorithm '{other}'"
            ))),
        }
    }

    fn digest_len(self) -> usize {
        match self {
            Self::Sha1 => 20,
            Self::Sha256 => 32,
            Self::Sha384 => 48,
            Self::Sha512 => 64,
        }
    }

    fn hash(self, chunks: &[&[u8]]) -> Vec<u8> {
        match self {
            Self::Sha1 => {
                let mut h = Sha1::new();
                for c in chunks {
                    h.update(c);
                }
                h.finalize().to_vec()
            }
            Self::Sha256 => {
                let mut h = Sha256::new();
                for c in chunks {
                    h.update(c);
                }
                h.finalize().to_vec()
            }
            Self::Sha384 => {
                let mut h = Sha384::new();
                for c in chunks {
                    h.update(c);
                }
                h.finalize().to_vec()
            }
            Self::Sha512 => {
                let mut h = Sha512::new();
                for c in chunks {
                    h.update(c);
                }
                h.finalize().to_vec()
            }
        }
    }

    fn verify_hmac(self, key: &[u8], data: &[u8], expected: &[u8]) -> Result<(), CryptoError> {
        let ok = match self {
            Self::Sha1 => {
                let mut mac = <Hmac<Sha1> as Mac>::new_from_slice(key)
                    .map_err(|_| CryptoError::Malformed("invalid HMAC key".into()))?;
                mac.update(data);
                mac.verify_slice(expected).is_ok()
            }
            Self::Sha256 => {
                let mut mac = <Hmac<Sha256> as Mac>::new_from_slice(key)
                    .map_err(|_| CryptoError::Malformed("invalid HMAC key".into()))?;
                mac.update(data);
                mac.verify_slice(expected).is_ok()
            }
            Self::Sha384 => {
                let mut mac = <Hmac<Sha384> as Mac>::new_from_slice(key)
                    .map_err(|_| CryptoError::Malformed("invalid HMAC key".into()))?;
                mac.update(data);
                mac.verify_slice(expected).is_ok()
            }
            Self::Sha512 => {
                let mut mac = <Hmac<Sha512> as Mac>::new_from_slice(key)
                    .map_err(|_| CryptoError::Malformed("invalid HMAC key".into()))?;
                mac.update(data);
                mac.verify_slice(expected).is_ok()
            }
        };
        if ok {
            Ok(())
        } else {
            Err(CryptoError::IntegrityCheckFailed)
        }
    }
}

fn password_utf16le(password: &str) -> Zeroizing<Vec<u8>> {
    let mut out = Zeroizing::new(Vec::with_capacity(password.len() * 2));
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn pad_or_truncate(mut bytes: Vec<u8>, target_len: usize, pad_byte: u8) -> Vec<u8> {
    if bytes.len() < target_len {
        bytes.resize(target_len, pad_byte);
    } else {
        bytes.truncate(target_len);
    }
    bytes
}

fn aes_cbc_decrypt_in_place(key: &[u8], iv: &[u8], data: &mut [u8]) -> Result<(), CryptoError> {
    if iv.len() != 16 {
        return Err(CryptoError::Malformed(format!(
            "AES-CBC requires 16-byte IV, got {}",
            iv.len()
        )));
    }
    if data.len() % 16 != 0 {
        return Err(CryptoError::Malformed(format!(
            "AES-CBC ciphertext length ({}) is not a multiple of 16",
            data.len()
        )));
    }
    match key.len() {
        16 => {
            let mut dec = cbc::Decryptor::<Aes128>::new_from_slices(key, iv)
                .map_err(|e| CryptoError::Malformed(format!("AES-128-CBC init: {e}")))?;
            for block in data.chunks_exact_mut(16) {
                dec.decrypt_block_mut(block.into());
            }
        }
        24 => {
            let mut dec = cbc::Decryptor::<Aes192>::new_from_slices(key, iv)
                .map_err(|e| CryptoError::Malformed(format!("AES-192-CBC init: {e}")))?;
            for block in data.chunks_exact_mut(16) {
                dec.decrypt_block_mut(block.into());
            }
        }
        32 => {
            let mut dec = cbc::Decryptor::<Aes256>::new_from_slices(key, iv)
                .map_err(|e| CryptoError::Malformed(format!("AES-256-CBC init: {e}")))?;
            for block in data.chunks_exact_mut(16) {
                dec.decrypt_block_mut(block.into());
            }
        }
        other => {
            return Err(CryptoError::UnsupportedEncryption(format!(
                "unsupported AES key length {other} bytes"
            )));
        }
    }
    Ok(())
}

fn aes_ecb_decrypt_in_place(key: &[u8], data: &mut [u8]) -> Result<(), CryptoError> {
    if data.len() % 16 != 0 {
        return Err(CryptoError::Malformed(format!(
            "AES-ECB ciphertext length ({}) is not a multiple of 16",
            data.len()
        )));
    }
    match key.len() {
        16 => {
            let cipher = Aes128::new_from_slice(key)
                .map_err(|e| CryptoError::Malformed(format!("AES-128-ECB init: {e}")))?;
            for block in data.chunks_exact_mut(16) {
                cipher.decrypt_block(block.into());
            }
        }
        24 => {
            let cipher = Aes192::new_from_slice(key)
                .map_err(|e| CryptoError::Malformed(format!("AES-192-ECB init: {e}")))?;
            for block in data.chunks_exact_mut(16) {
                cipher.decrypt_block(block.into());
            }
        }
        32 => {
            let cipher = Aes256::new_from_slice(key)
                .map_err(|e| CryptoError::Malformed(format!("AES-256-ECB init: {e}")))?;
            for block in data.chunks_exact_mut(16) {
                cipher.decrypt_block(block.into());
            }
        }
        other => {
            return Err(CryptoError::UnsupportedEncryption(format!(
                "unsupported AES key length {other} bytes"
            )));
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Agile Encryption (MS-OFFCRYPTO 2.3.4.10 – 2.3.4.15)
// ---------------------------------------------------------------------------

struct AgileKeyData {
    block_size: usize,
    key_bytes: usize,
    hash_alg: HashAlg,
    salt: Vec<u8>,
}

struct AgileDataIntegrity {
    encrypted_hmac_key: Vec<u8>,
    encrypted_hmac_value: Vec<u8>,
}

struct AgilePasswordKey {
    spin_count: u32,
    block_size: usize,
    key_bytes: usize,
    hash_alg: HashAlg,
    salt: Vec<u8>,
    encrypted_verifier_hash_input: Vec<u8>,
    encrypted_verifier_hash_value: Vec<u8>,
    encrypted_key_value: Vec<u8>,
}

struct AgileDescriptor {
    key_data: AgileKeyData,
    data_integrity: Option<AgileDataIntegrity>,
    password_key: AgilePasswordKey,
}

fn parse_agile_xml(xml_bytes: &[u8]) -> Result<AgileDescriptor, CryptoError> {
    // Strip trailing NUL bytes if any (some writers pad XML).
    let trimmed = match xml_bytes.iter().rposition(|&b| b != 0) {
        Some(pos) => &xml_bytes[..=pos],
        None => return Err(CryptoError::Malformed("empty Agile EncryptionInfo XML".into())),
    };
    let xml = std::str::from_utf8(trimmed)
        .map_err(|e| CryptoError::Malformed(format!("invalid UTF-8 in EncryptionInfo XML: {e}")))?;

    let mut reader = Reader::from_str(xml);
    let mut key_data: Option<AgileKeyData> = None;
    let mut data_integrity: Option<AgileDataIntegrity> = None;
    let mut password_key: Option<AgilePasswordKey> = None;
    let mut in_password_encryptor = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"keyData" => {
                        key_data = Some(parse_key_data(&e)?);
                    }
                    b"dataIntegrity" => {
                        let enc_key = decode_b64_attr(&e, b"encryptedHmacKey")?;
                        let enc_val = decode_b64_attr(&e, b"encryptedHmacValue")?;
                        data_integrity = Some(AgileDataIntegrity {
                            encrypted_hmac_key: enc_key,
                            encrypted_hmac_value: enc_val,
                        });
                    }
                    b"keyEncryptor" => {
                        let uri = req_attr(&e, b"uri")?;
                        in_password_encryptor = uri == PASSWORD_KEY_ENCRYPTOR_URI;
                    }
                    b"encryptedKey" if in_password_encryptor => {
                        password_key = Some(parse_password_key(&e)?);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == b"keyEncryptor" {
                    in_password_encryptor = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(CryptoError::Malformed(format!(
                    "failed to parse Agile EncryptionInfo XML: {e}"
                )));
            }
            _ => {}
        }
    }

    Ok(AgileDescriptor {
        key_data: key_data
            .ok_or_else(|| CryptoError::Malformed("missing <keyData> in EncryptionInfo".into()))?,
        data_integrity,
        password_key: password_key.ok_or_else(|| {
            CryptoError::UnsupportedEncryption(
                "no password key encryptor found (certificate/IRM encryption is not supported)"
                    .into(),
            )
        })?,
    })
}

fn req_attr(e: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Result<String, CryptoError> {
    for attr in e.attributes() {
        let attr = attr
            .map_err(|err| CryptoError::Malformed(format!("malformed XML attribute: {err}")))?;
        if attr.key.local_name().as_ref() == name {
            let s = std::str::from_utf8(attr.value.as_ref())
                .map_err(|err| CryptoError::Malformed(format!("non-UTF-8 attribute: {err}")))?;
            return Ok(s.to_owned());
        }
    }
    Err(CryptoError::Malformed(format!(
        "missing attribute '{}'",
        String::from_utf8_lossy(name)
    )))
}

fn req_usize_attr(e: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Result<usize, CryptoError> {
    let s = req_attr(e, name)?;
    s.parse::<usize>().map_err(|_| {
        CryptoError::Malformed(format!(
            "invalid integer attribute '{}' = '{s}'",
            String::from_utf8_lossy(name)
        ))
    })
}

fn decode_b64_attr(
    e: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
) -> Result<Vec<u8>, CryptoError> {
    let s = req_attr(e, name)?;
    BASE64.decode(s.trim()).map_err(|err| {
        CryptoError::Malformed(format!(
            "invalid base64 in '{}': {err}",
            String::from_utf8_lossy(name)
        ))
    })
}

fn validate_aes_cbc_params(
    cipher_alg: &str,
    chaining: &str,
    block_size: usize,
    key_bits: usize,
    hash_alg: HashAlg,
    hash_size: usize,
    salt_size: usize,
    salt: &[u8],
) -> Result<usize, CryptoError> {
    if cipher_alg != "AES" {
        return Err(CryptoError::UnsupportedEncryption(format!(
            "unsupported cipher '{cipher_alg}'"
        )));
    }
    if chaining != "ChainingModeCBC" {
        return Err(CryptoError::UnsupportedEncryption(format!(
            "unsupported cipher chaining '{chaining}'"
        )));
    }
    if block_size != 16 {
        return Err(CryptoError::UnsupportedEncryption(format!(
            "unsupported AES block size {block_size}"
        )));
    }
    if !matches!(key_bits, 128 | 192 | 256) {
        return Err(CryptoError::UnsupportedEncryption(format!(
            "unsupported AES key size {key_bits} bits"
        )));
    }
    if hash_size != hash_alg.digest_len() {
        return Err(CryptoError::Malformed(format!(
            "hashSize ({hash_size}) does not match {hash_alg:?} digest length"
        )));
    }
    if salt.len() != salt_size || salt_size == 0 || salt_size > 64 {
        return Err(CryptoError::Malformed(format!(
            "invalid saltSize ({salt_size}) / salt length ({})",
            salt.len()
        )));
    }
    Ok(key_bits / 8)
}

fn parse_key_data(e: &quick_xml::events::BytesStart<'_>) -> Result<AgileKeyData, CryptoError> {
    let salt_size = req_usize_attr(e, b"saltSize")?;
    let block_size = req_usize_attr(e, b"blockSize")?;
    let key_bits = req_usize_attr(e, b"keyBits")?;
    let hash_size = req_usize_attr(e, b"hashSize")?;
    let cipher_alg = req_attr(e, b"cipherAlgorithm")?;
    let chaining = req_attr(e, b"cipherChaining")?;
    let hash_alg = HashAlg::from_name(&req_attr(e, b"hashAlgorithm")?)?;
    let salt = decode_b64_attr(e, b"saltValue")?;
    let key_bytes = validate_aes_cbc_params(
        &cipher_alg,
        &chaining,
        block_size,
        key_bits,
        hash_alg,
        hash_size,
        salt_size,
        &salt,
    )?;
    Ok(AgileKeyData {
        block_size,
        key_bytes,
        hash_alg,
        salt,
    })
}

fn parse_password_key(
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<AgilePasswordKey, CryptoError> {
    let spin_count_u64 = req_attr(e, b"spinCount")?
        .parse::<u64>()
        .map_err(|_| CryptoError::Malformed("invalid spinCount".into()))?;
    if spin_count_u64 > MAX_SPIN_COUNT as u64 {
        return Err(CryptoError::Malformed(format!(
            "spinCount ({spin_count_u64}) exceeds safety limit ({MAX_SPIN_COUNT})"
        )));
    }
    let spin_count = spin_count_u64 as u32;
    let salt_size = req_usize_attr(e, b"saltSize")?;
    let block_size = req_usize_attr(e, b"blockSize")?;
    let key_bits = req_usize_attr(e, b"keyBits")?;
    let hash_size = req_usize_attr(e, b"hashSize")?;
    let cipher_alg = req_attr(e, b"cipherAlgorithm")?;
    let chaining = req_attr(e, b"cipherChaining")?;
    let hash_alg = HashAlg::from_name(&req_attr(e, b"hashAlgorithm")?)?;
    let salt = decode_b64_attr(e, b"saltValue")?;
    let key_bytes = validate_aes_cbc_params(
        &cipher_alg,
        &chaining,
        block_size,
        key_bits,
        hash_alg,
        hash_size,
        salt_size,
        &salt,
    )?;
    let encrypted_verifier_hash_input = decode_b64_attr(e, b"encryptedVerifierHashInput")?;
    let encrypted_verifier_hash_value = decode_b64_attr(e, b"encryptedVerifierHashValue")?;
    let encrypted_key_value = decode_b64_attr(e, b"encryptedKeyValue")?;

    Ok(AgilePasswordKey {
        spin_count,
        block_size,
        key_bytes,
        hash_alg,
        salt,
        encrypted_verifier_hash_input,
        encrypted_verifier_hash_value,
        encrypted_key_value,
    })
}

fn agile_spin_hash(alg: HashAlg, salt: &[u8], password: &str, spin_count: u32) -> Zeroizing<Vec<u8>> {
    let pw = password_utf16le(password);
    let mut h = Zeroizing::new(alg.hash(&[salt, &pw]));
    for i in 0..spin_count {
        let iter = i.to_le_bytes();
        h = Zeroizing::new(alg.hash(&[&iter, &h]));
    }
    h
}

fn agile_derive_key(alg: HashAlg, spun: &[u8], block_key: &[u8], key_bytes: usize) -> Zeroizing<Vec<u8>> {
    let h = alg.hash(&[spun, block_key]);
    Zeroizing::new(pad_or_truncate(h, key_bytes, 0x36))
}

fn agile_derive_iv(alg: HashAlg, salt: &[u8], block_key: &[u8], block_size: usize) -> Vec<u8> {
    let h = alg.hash(&[salt, block_key]);
    pad_or_truncate(h, block_size, 0x36)
}

fn decrypt_agile(
    xml_bytes: &[u8],
    package_bytes: &[u8],
    password: &str,
) -> Result<Vec<u8>, CryptoError> {
    let desc = parse_agile_xml(xml_bytes)?;
    let pk = &desc.password_key;
    let kd = &desc.key_data;

    let spun = agile_spin_hash(pk.hash_alg, &pk.salt, password, pk.spin_count);
    let iv = pad_or_truncate(pk.salt.clone(), pk.block_size, 0x36);

    // 1. Constant-time password verifier check (MS-OFFCRYPTO 2.3.4.13)
    let verifier_input_key = agile_derive_key(
        pk.hash_alg,
        &spun,
        &AGILE_VERIFIER_HASH_INPUT_BLOCK_KEY,
        pk.key_bytes,
    );
    let mut verifier_input = Zeroizing::new(pk.encrypted_verifier_hash_input.clone());
    aes_cbc_decrypt_in_place(&verifier_input_key, &iv, &mut verifier_input)?;

    let verifier_value_key = agile_derive_key(
        pk.hash_alg,
        &spun,
        &AGILE_VERIFIER_HASH_VALUE_BLOCK_KEY,
        pk.key_bytes,
    );
    let mut verifier_hash_decrypted = Zeroizing::new(pk.encrypted_verifier_hash_value.clone());
    aes_cbc_decrypt_in_place(&verifier_value_key, &iv, &mut verifier_hash_decrypted)?;

    let hash_len = pk.hash_alg.digest_len();
    if verifier_hash_decrypted.len() < hash_len {
        return Err(CryptoError::Malformed(
            "encryptedVerifierHashValue is shorter than hashSize".into(),
        ));
    }
    let computed_hash = pk.hash_alg.hash(&[&verifier_input]);
    if !bool::from(computed_hash.as_slice().ct_eq(&verifier_hash_decrypted[..hash_len])) {
        return Err(CryptoError::IncorrectPassword);
    }

    // 2. Recover document secret key
    let enc_key_key = agile_derive_key(
        pk.hash_alg,
        &spun,
        &AGILE_ENCRYPTED_KEY_VALUE_BLOCK_KEY,
        pk.key_bytes,
    );
    let mut secret_key_buf = Zeroizing::new(pk.encrypted_key_value.clone());
    aes_cbc_decrypt_in_place(&enc_key_key, &iv, &mut secret_key_buf)?;
    if secret_key_buf.len() < kd.key_bytes {
        return Err(CryptoError::Malformed(
            "encryptedKeyValue is shorter than keyData keyBits".into(),
        ));
    }
    let secret_key = Zeroizing::new(secret_key_buf[..kd.key_bytes].to_vec());

    // 3. Verify Agile dataIntegrity HMAC over the entire EncryptedPackage stream
    if let Some(di) = &desc.data_integrity {
        let hmac_key_iv = agile_derive_iv(
            kd.hash_alg,
            &kd.salt,
            &AGILE_HMAC_KEY_BLOCK_KEY,
            kd.block_size,
        );
        let mut hmac_key_buf = Zeroizing::new(di.encrypted_hmac_key.clone());
        aes_cbc_decrypt_in_place(&secret_key, &hmac_key_iv, &mut hmac_key_buf)?;
        let kd_hash_len = kd.hash_alg.digest_len();
        if hmac_key_buf.len() < kd_hash_len {
            return Err(CryptoError::Malformed(
                "encryptedHmacKey is shorter than hashSize".into(),
            ));
        }
        let hmac_key = Zeroizing::new(hmac_key_buf[..kd_hash_len].to_vec());

        let hmac_val_iv = agile_derive_iv(
            kd.hash_alg,
            &kd.salt,
            &AGILE_HMAC_VALUE_BLOCK_KEY,
            kd.block_size,
        );
        let mut hmac_val_buf = di.encrypted_hmac_value.clone();
        aes_cbc_decrypt_in_place(&secret_key, &hmac_val_iv, &mut hmac_val_buf)?;
        if hmac_val_buf.len() < kd_hash_len {
            return Err(CryptoError::Malformed(
                "encryptedHmacValue is shorter than hashSize".into(),
            ));
        }
        kd.hash_alg
            .verify_hmac(&hmac_key, package_bytes, &hmac_val_buf[..kd_hash_len])?;
    }

    // 4. Decrypt EncryptedPackage in 4096-byte segments (MS-OFFCRYPTO 2.3.4.15)
    let (unencrypted_size, encrypted_payload) = parse_package_header(package_bytes)?;
    let mut output = Vec::with_capacity(encrypted_payload.len());
    for (segment_idx, chunk) in encrypted_payload.chunks(4096).enumerate() {
        let seg_u32 = u32::try_from(segment_idx)
            .map_err(|_| CryptoError::Malformed("EncryptedPackage has too many segments".into()))?;
        let seg_iv = agile_derive_iv(
            kd.hash_alg,
            &kd.salt,
            &seg_u32.to_le_bytes(),
            kd.block_size,
        );
        let mut buf = chunk.to_vec();
        aes_cbc_decrypt_in_place(&secret_key, &seg_iv, &mut buf)?;
        output.extend_from_slice(&buf);
    }
    output.truncate(unencrypted_size);
    Ok(output)
}

// ---------------------------------------------------------------------------
// Standard Encryption (Excel 2007, MS-OFFCRYPTO 2.3.4.5 – 2.3.4.10)
// ---------------------------------------------------------------------------

fn decrypt_standard(
    rest: &[u8],
    package_bytes: &[u8],
    password: &str,
) -> Result<Vec<u8>, CryptoError> {
    if rest.len() < 4 {
        return Err(CryptoError::Malformed("Standard EncryptionInfo truncated".into()));
    }
    let header_size = u32::from_le_bytes([rest[0], rest[1], rest[2], rest[3]]) as usize;
    let header_end = 4usize
        .checked_add(header_size)
        .ok_or_else(|| CryptoError::Malformed("Standard header size overflow".into()))?;
    if header_size < 32 || rest.len() < header_end {
        return Err(CryptoError::Malformed("Standard EncryptionHeader truncated".into()));
    }
    let header = &rest[4..header_end];
    let alg_id = u32::from_le_bytes([header[8], header[9], header[10], header[11]]);
    let alg_id_hash = u32::from_le_bytes([header[12], header[13], header[14], header[15]]);
    let key_bits = u32::from_le_bytes([header[16], header[17], header[18], header[19]]);

    let key_bytes = match (alg_id, key_bits) {
        (ALG_ID_AES_128, 0 | 128) => 16usize,
        (ALG_ID_AES_192, 192) => 24usize,
        (ALG_ID_AES_256, 256) => 32usize,
        (0, 0 | 128) => 16usize,
        _ => {
            return Err(CryptoError::UnsupportedEncryption(format!(
                "unsupported Standard cipher algId=0x{alg_id:08x} keyBits={key_bits}"
            )));
        }
    };
    if alg_id_hash != 0 && alg_id_hash != ALG_ID_SHA1 {
        return Err(CryptoError::UnsupportedEncryption(format!(
            "unsupported Standard hash algIdHash=0x{alg_id_hash:08x}"
        )));
    }

    // Parse EncryptionVerifier (MS-OFFCRYPTO 2.3.3):
    // saltSize (4) + salt (16) + encryptedVerifier (16) + verifierHashSize (4) + encryptedVerifierHash (32)
    let verifier = &rest[header_end..];
    if verifier.len() < 4 + 16 + 16 + 4 + 32 {
        return Err(CryptoError::Malformed("Standard EncryptionVerifier truncated".into()));
    }
    let salt_size = u32::from_le_bytes([verifier[0], verifier[1], verifier[2], verifier[3]]);
    if salt_size != 16 {
        return Err(CryptoError::Malformed(format!(
            "Standard EncryptionVerifier saltSize must be 16, got {salt_size}"
        )));
    }
    let salt = &verifier[4..20];
    let encrypted_verifier = &verifier[20..36];
    let verifier_hash_size =
        u32::from_le_bytes([verifier[36], verifier[37], verifier[38], verifier[39]]);
    if verifier_hash_size != 20 {
        return Err(CryptoError::Malformed(format!(
            "Standard verifierHashSize must be 20 (SHA-1), got {verifier_hash_size}"
        )));
    }
    let encrypted_verifier_hash = &verifier[40..72];

    // Key derivation (MS-OFFCRYPTO 2.3.4.7)
    let key = standard_derive_key(salt, password, key_bytes);

    // Constant-time password verification (MS-OFFCRYPTO 2.3.4.9)
    let mut dec_verifier = Zeroizing::new(encrypted_verifier.to_vec());
    aes_ecb_decrypt_in_place(&key, &mut dec_verifier)?;
    let mut dec_verifier_hash = Zeroizing::new(encrypted_verifier_hash.to_vec());
    aes_ecb_decrypt_in_place(&key, &mut dec_verifier_hash)?;

    let computed_hash = Sha1::digest(&*dec_verifier);
    if !bool::from(computed_hash.as_slice().ct_eq(&dec_verifier_hash[..20])) {
        return Err(CryptoError::IncorrectPassword);
    }

    // Decrypt EncryptedPackage with AES-ECB (MS-OFFCRYPTO 2.3.4.10)
    // Note: Standard encryption has no format-level HMAC integrity check.
    let (unencrypted_size, encrypted_payload) = parse_package_header(package_bytes)?;
    let mut output = encrypted_payload.to_vec();
    aes_ecb_decrypt_in_place(&key, &mut output)?;
    output.truncate(unencrypted_size);
    Ok(output)
}

fn standard_derive_key(salt: &[u8], password: &str, key_bytes: usize) -> Zeroizing<Vec<u8>> {
    let pw = password_utf16le(password);
    let mut h = Sha1::new();
    h.update(salt);
    h.update(&*pw);
    let mut digest = Zeroizing::new(h.finalize().to_vec());

    for i in 0..STANDARD_SPIN_COUNT {
        let mut h = Sha1::new();
        h.update(i.to_le_bytes());
        h.update(&*digest);
        *digest = h.finalize().to_vec();
    }

    // Block 0 hash
    let mut h = Sha1::new();
    h.update(&*digest);
    h.update(0u32.to_le_bytes());
    let h_final = Zeroizing::new(h.finalize().to_vec());

    // CryptDeriveKey (MS-OFFCRYPTO 2.3.4.7)
    let mut buf1 = Zeroizing::new([0x36u8; 64]);
    let mut buf2 = Zeroizing::new([0x5Cu8; 64]);
    for i in 0..20 {
        buf1[i] ^= h_final[i];
        buf2[i] ^= h_final[i];
    }
    let x1 = Sha1::digest(&*buf1);
    let x2 = Sha1::digest(&*buf2);
    let mut derived = Zeroizing::new(Vec::with_capacity(40));
    derived.extend_from_slice(&x1);
    derived.extend_from_slice(&x2);
    derived.truncate(key_bytes);
    derived
}

fn parse_package_header(package_bytes: &[u8]) -> Result<(usize, &[u8]), CryptoError> {
    if package_bytes.len() < 8 {
        return Err(CryptoError::Malformed("EncryptedPackage stream is truncated".into()));
    }
    let unencrypted_size_u64 = u64::from_le_bytes(package_bytes[0..8].try_into().unwrap());
    let encrypted_payload = &package_bytes[8..];
    if unencrypted_size_u64 > MAX_ENCRYPTED_PACKAGE_BYTES {
        return Err(CryptoError::Malformed(format!(
            "declared unencrypted size ({unencrypted_size_u64} bytes) exceeds limit"
        )));
    }
    if unencrypted_size_u64 > encrypted_payload.len() as u64 {
        return Err(CryptoError::Malformed(format!(
            "declared unencrypted size ({unencrypted_size_u64}) exceeds encrypted payload ({})",
            encrypted_payload.len()
        )));
    }
    if encrypted_payload.len() % 16 != 0 {
        return Err(CryptoError::Malformed(
            "EncryptedPackage payload length is not a multiple of 16 bytes".into(),
        ));
    }
    Ok((unencrypted_size_u64 as usize, encrypted_payload))
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::cipher::BlockEncrypt;
    use std::io::Write;
    use visigrid_engine::workbook::Workbook;

    fn aes_ecb_encrypt_in_place(key: &[u8], data: &mut [u8]) {
        match key.len() {
            16 => {
                let cipher = Aes128::new_from_slice(key).unwrap();
                for block in data.chunks_exact_mut(16) {
                    cipher.encrypt_block(block.into());
                }
            }
            24 => {
                let cipher = Aes192::new_from_slice(key).unwrap();
                for block in data.chunks_exact_mut(16) {
                    cipher.encrypt_block(block.into());
                }
            }
            32 => {
                let cipher = Aes256::new_from_slice(key).unwrap();
                for block in data.chunks_exact_mut(16) {
                    cipher.encrypt_block(block.into());
                }
            }
            _ => unreachable!(),
        }
    }

    fn wrap_cfb(encryption_info: &[u8], encrypted_package: &[u8]) -> Vec<u8> {
        let mut comp = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
        comp.create_stream("EncryptionInfo")
            .unwrap()
            .write_all(encryption_info)
            .unwrap();
        comp.create_stream("EncryptedPackage")
            .unwrap()
            .write_all(encrypted_package)
            .unwrap();
        comp.into_inner().into_inner()
    }

    fn encrypt_agile_fixture(
        plaintext: &[u8],
        password: &str,
        hash_alg: HashAlg,
        key_bits: usize,
        spin_count: u32,
    ) -> Vec<u8> {
        let keys = write::AgileKeys {
            data_salt: [0x11; 16],
            password_salt: [0x22; 16],
            package_key: Zeroizing::new(vec![0x33; key_bits / 8]),
            verifier: Zeroizing::new([0x44; 16]),
            hmac_key: Zeroizing::new(vec![0x55; hash_alg.digest_len()]),
        };
        write::encrypt_agile(plaintext, password, hash_alg, spin_count, &keys).unwrap()
    }

    /// Build a spec-compliant Excel 2007 Standard MS-OFFCRYPTO container (AES-128-ECB + SHA-1).
    fn encrypt_standard_fixture(plaintext: &[u8], password: &str) -> Vec<u8> {
        let salt = [0x77u8; 16];
        let verifier_plain = [0x88u8; 16];
        let key = standard_derive_key(&salt, password, 16);

        let mut enc_verifier = verifier_plain.to_vec();
        aes_ecb_encrypt_in_place(&key, &mut enc_verifier);

        let mut enc_verifier_hash = Sha1::digest(&verifier_plain).to_vec();
        enc_verifier_hash.resize(32, 0);
        aes_ecb_encrypt_in_place(&key, &mut enc_verifier_hash);

        let mut package = Vec::new();
        package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        let mut payload = plaintext.to_vec();
        let rem = payload.len() % 16;
        if rem != 0 {
            payload.resize(payload.len() + (16 - rem), 0);
        }
        aes_ecb_encrypt_in_place(&key, &mut payload);
        package.extend_from_slice(&payload);

        let mut header = Vec::new();
        header.extend_from_slice(&0x24u32.to_le_bytes()); // flags (fCryptoAPI | fAES)
        header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
        header.extend_from_slice(&ALG_ID_AES_128.to_le_bytes());
        header.extend_from_slice(&ALG_ID_SHA1.to_le_bytes());
        header.extend_from_slice(&128u32.to_le_bytes()); // keySize
        header.extend_from_slice(&0x18u32.to_le_bytes()); // providerType
        header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        header.extend_from_slice(&0u32.to_le_bytes()); // reserved2
        for unit in "Microsoft Enhanced RSA and AES Cryptographic Provider\0".encode_utf16() {
            header.extend_from_slice(&unit.to_le_bytes());
        }

        let mut info = Vec::new();
        info.extend_from_slice(&3u16.to_le_bytes()); // major (Excel 2007)
        info.extend_from_slice(&2u16.to_le_bytes()); // minor (Standard)
        info.extend_from_slice(&0x24u32.to_le_bytes()); // flags
        info.extend_from_slice(&(header.len() as u32).to_le_bytes());
        info.extend_from_slice(&header);
        // EncryptionVerifier
        info.extend_from_slice(&16u32.to_le_bytes()); // saltSize
        info.extend_from_slice(&salt);
        info.extend_from_slice(&enc_verifier);
        info.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize
        info.extend_from_slice(&enc_verifier_hash);

        wrap_cfb(&info, &package)
    }

    fn sample_xlsx_bytes() -> Vec<u8> {
        let mut wb = Workbook::new();
        let sheet = wb.active_sheet_mut();
        sheet.set_value(0, 0, "Revenue");
        sheet.set_value(0, 1, "12345");
        sheet.set_value(1, 0, "Double");
        sheet.set_value(1, 1, "=B1*2");
        let (bytes, _) = crate::xlsx::export_to_buffer(&wb, None).unwrap();
        bytes
    }

    #[test]
    fn writer_uses_fixed_agile_parameters() {
        let encrypted = encrypt_ooxml(b"test", "workbook").unwrap();
        let mut file = cfb::CompoundFile::open(Cursor::new(encrypted.as_slice())).unwrap();
        let info = read_bounded_stream(&mut file, "EncryptionInfo", MAX_ENCRYPTION_INFO_BYTES).unwrap();
        assert_eq!(&info[..8], &[4, 0, 4, 0, 0x40, 0, 0, 0]);
        let desc = parse_agile_xml(&info[8..]).unwrap();
        let pk = desc.password_key;
        let kd = desc.key_data;
        assert_eq!(pk.spin_count, 100_000);
        assert_eq!((pk.hash_alg, pk.key_bytes, pk.block_size, pk.salt.len()), (HashAlg::Sha512, 32, 16, 16));
        assert_eq!((kd.hash_alg, kd.key_bytes, kd.block_size, kd.salt.len()), (HashAlg::Sha512, 32, 16, 16));
        assert_ne!(pk.salt, kd.salt);
        assert!(desc.data_integrity.is_some());
    }

    #[test]
    fn agile_sha512_aes256_roundtrip_and_xlsx_import() {
        let xlsx = sample_xlsx_bytes();
        let encrypted = encrypt_agile_fixture(&xlsx, "secret-pass", HashAlg::Sha512, 256, 100);
        assert!(is_encrypted_ooxml(&encrypted));

        // Without password -> PasswordRequired
        let err = crate::xlsx::import_bytes_with_options_typed(
            &encrypted,
            Some(crate::xlsx::WorkbookFormat::Xlsx),
            &crate::xlsx::ImportOptions::default(),
        )
        .unwrap_err();
        assert_eq!(err, crate::xlsx::ImportError::PasswordRequired);

        // Wrong password -> IncorrectPassword
        let err = crate::xlsx::import_bytes_with_options_typed(
            &encrypted,
            Some(crate::xlsx::WorkbookFormat::Xlsx),
            &crate::xlsx::ImportOptions {
                password: Some(Zeroizing::new("wrong-pass".into())),
                ..Default::default()
            },
        )
        .unwrap_err();
        assert_eq!(err, crate::xlsx::ImportError::IncorrectPassword);

        // Correct password -> opens cleanly and sets was_encrypted = true
        let (wb, result) = crate::xlsx::import_bytes_with_options_typed(
            &encrypted,
            Some(crate::xlsx::WorkbookFormat::Xlsx),
            &crate::xlsx::ImportOptions {
                password: Some(Zeroizing::new("secret-pass".into())),
                ..Default::default()
            },
        )
        .unwrap();
        assert!(result.was_encrypted);
        assert_eq!(wb.active_sheet().get_display(0, 0), "Revenue");
        assert_eq!(wb.active_sheet().get_display(0, 1), "12345");
        assert_eq!(wb.active_sheet().get_display(1, 1), "24690");
    }

    #[test]
    fn agile_sha1_excel2010_and_sha256_decrypt() {
        let xlsx = sample_xlsx_bytes();
        for (alg, bits) in [(HashAlg::Sha1, 128), (HashAlg::Sha256, 128), (HashAlg::Sha384, 256)] {
            let encrypted = encrypt_agile_fixture(&xlsx, "pässwörd🔑", alg, bits, 50);
            let decrypted = decrypt_ooxml(&encrypted, "pässwörd🔑").unwrap();
            assert_eq!(decrypted, xlsx, "failed for {alg:?}/{bits}");
            assert_eq!(
                decrypt_ooxml(&encrypted, "wrong").unwrap_err(),
                CryptoError::IncorrectPassword
            );
        }
    }

    #[test]
    fn agile_empty_password_works() {
        let xlsx = sample_xlsx_bytes();
        let encrypted = encrypt_agile_fixture(&xlsx, "", HashAlg::Sha512, 256, 10);
        assert_eq!(decrypt_ooxml(&encrypted, "").unwrap(), xlsx);
        assert_eq!(
            decrypt_ooxml(&encrypted, "nonempty").unwrap_err(),
            CryptoError::IncorrectPassword
        );
    }

    #[test]
    fn agile_tampered_package_fails_hmac_integrity_check() {
        let xlsx = sample_xlsx_bytes();
        let encrypted = encrypt_agile_fixture(&xlsx, "secret", HashAlg::Sha512, 256, 10);

        // Open CFB, flip one bit inside EncryptedPackage, re-wrap.
        let mut comp = cfb::CompoundFile::open(Cursor::new(encrypted.as_slice())).unwrap();
        let mut info = Vec::new();
        comp.open_stream("EncryptionInfo").unwrap().read_to_end(&mut info).unwrap();
        let mut pkg = Vec::new();
        comp.open_stream("EncryptedPackage").unwrap().read_to_end(&mut pkg).unwrap();
        pkg[16] ^= 0x01;
        let tampered = wrap_cfb(&info, &pkg);

        let err = decrypt_ooxml(&tampered, "secret").unwrap_err();
        assert_eq!(err, CryptoError::IntegrityCheckFailed);
    }

    #[test]
    fn standard_excel2007_aes128_roundtrip_and_wrong_password() {
        let xlsx = sample_xlsx_bytes();
        let encrypted = encrypt_standard_fixture(&xlsx, "excel2007!");
        assert!(is_encrypted_ooxml(&encrypted));
        assert_eq!(
            decrypt_ooxml(&encrypted, "nope").unwrap_err(),
            CryptoError::IncorrectPassword
        );
        let decrypted = decrypt_ooxml(&encrypted, "excel2007!").unwrap();
        assert_eq!(decrypted, xlsx);
    }

    #[test]
    fn hostile_spin_count_is_rejected_before_hashing() {
        let xlsx = sample_xlsx_bytes();
        // Build with small spinCount then patch XML spinCount to 99_999_999.
        let encrypted = encrypt_agile_fixture(&xlsx, "pw", HashAlg::Sha512, 256, 10);
        let mut comp = cfb::CompoundFile::open(Cursor::new(encrypted.as_slice())).unwrap();
        let mut info = Vec::new();
        comp.open_stream("EncryptionInfo").unwrap().read_to_end(&mut info).unwrap();
        let mut pkg = Vec::new();
        comp.open_stream("EncryptedPackage").unwrap().read_to_end(&mut pkg).unwrap();
        let xml = String::from_utf8(info[8..].to_vec()).unwrap();
        let patched_xml = xml.replace("spinCount=\"10\"", "spinCount=\"99999999\"");
        let mut patched_info = info[..8].to_vec();
        patched_info.extend_from_slice(patched_xml.as_bytes());
        let hostile = wrap_cfb(&patched_info, &pkg);

        let err = decrypt_ooxml(&hostile, "pw").unwrap_err();
        assert!(matches!(err, CryptoError::Malformed(ref m) if m.contains("spinCount")));
    }

    #[test]
    fn hostile_declared_package_size_is_rejected() {
        let xlsx = sample_xlsx_bytes();
        let encrypted = encrypt_standard_fixture(&xlsx, "pw");
        let mut comp = cfb::CompoundFile::open(Cursor::new(encrypted.as_slice())).unwrap();
        let mut info = Vec::new();
        comp.open_stream("EncryptionInfo").unwrap().read_to_end(&mut info).unwrap();
        let mut pkg = Vec::new();
        comp.open_stream("EncryptedPackage").unwrap().read_to_end(&mut pkg).unwrap();
        // Overwrite declared size with 10 GiB
        pkg[0..8].copy_from_slice(&(10u64 * 1024 * 1024 * 1024).to_le_bytes());
        let hostile = wrap_cfb(&info, &pkg);

        let err = decrypt_ooxml(&hostile, "pw").unwrap_err();
        assert!(matches!(err, CryptoError::Malformed(ref m) if m.contains("exceeds")));
    }

    #[test]
    fn real_msoffcrypto_tool_agile_fixture_opens_and_rejects_wrong_password_or_tampering() {
        let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/encrypted_msoffcrypto_agile.xlsx");

        // No password -> PasswordRequired
        let err = crate::xlsx::import_with_options_typed(
            &fixture,
            &crate::xlsx::ImportOptions::default(),
        )
        .unwrap_err();
        assert_eq!(err, crate::xlsx::ImportError::PasswordRequired);

        // Wrong password -> IncorrectPassword
        let err = crate::xlsx::import_with_options_typed(
            &fixture,
            &crate::xlsx::ImportOptions {
                password: Some(Zeroizing::new("wrong".into())),
                ..Default::default()
            },
        )
        .unwrap_err();
        assert_eq!(err, crate::xlsx::ImportError::IncorrectPassword);

        // Right password -> imports cleanly with was_encrypted = true
        let (wb, result) = crate::xlsx::import_with_options_typed(
            &fixture,
            &crate::xlsx::ImportOptions {
                password: Some(Zeroizing::new("hunter2".into())),
                ..Default::default()
            },
        )
        .unwrap();
        assert!(result.was_encrypted);
        assert!(wb.sheet_count() >= 1);

        // Flip a byte inside EncryptedPackage -> IntegrityCheckFailed
        let raw = std::fs::read(&fixture).unwrap();
        let mut comp = cfb::CompoundFile::open(Cursor::new(raw.as_slice())).unwrap();
        let mut info = Vec::new();
        comp.open_stream("EncryptionInfo").unwrap().read_to_end(&mut info).unwrap();
        let mut pkg = Vec::new();
        comp.open_stream("EncryptedPackage").unwrap().read_to_end(&mut pkg).unwrap();
        pkg[16] ^= 0x01;
        let tampered = wrap_cfb(&info, &pkg);
        let err = decrypt_ooxml(&tampered, "hunter2").unwrap_err();
        assert_eq!(err, CryptoError::IntegrityCheckFailed);
    }
}
