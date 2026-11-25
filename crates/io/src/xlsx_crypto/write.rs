//! Agile encryption for newly written OOXML packages (MS-OFFCRYPTO 2.3.4).
use super::*;
use cbc::cipher::BlockEncryptMut;
use std::io::{self, Write};

pub(super) struct AgileKeys {
    pub(super) data_salt: [u8; 16],
    pub(super) password_salt: [u8; 16],
    pub(super) package_key: Zeroizing<Vec<u8>>,
    pub(super) verifier: Zeroizing<[u8; 16]>,
    pub(super) hmac_key: Zeroizing<Vec<u8>>,
}

/// Wrap finished ZIP bytes using Agile AES-256-CBC, SHA-512 and data integrity.
pub fn encrypt_ooxml(package: &[u8], password: &str) -> Result<Vec<u8>, CryptoError> {
    if !(1..=255).contains(&password.encode_utf16().take(256).count()) {
        return Err(CryptoError::InvalidPasswordLength);
    }
    // The reader's stream limit includes the length prefix and AES padding.
    if package.len() as u64 > (MAX_ENCRYPTED_PACKAGE_BYTES - 8) / 16 * 16 {
        return Err(CryptoError::EncryptionFailed(
            "package exceeds the encrypted stream size limit".into(),
        ));
    }
    let mut keys = AgileKeys {
        data_salt: [0; 16],
        password_salt: [0; 16],
        package_key: Zeroizing::new(vec![0; 32]),
        verifier: Zeroizing::new([0; 16]),
        hmac_key: Zeroizing::new(vec![0; 64]),
    };
    for secret in [
        &mut keys.data_salt[..],
        &mut keys.password_salt,
        &mut keys.package_key[..],
        &mut keys.verifier[..],
        &mut keys.hmac_key[..],
    ] {
        getrandom::fill(secret)
            .map_err(|_| CryptoError::EncryptionFailed("OS randomness unavailable".into()))?;
    }
    encrypt_agile(package, password, HashAlg::Sha512, 100_000, &keys)
}

fn aes_cbc_encrypt_in_place(key: &[u8], iv: &[u8], data: &mut [u8]) {
    match key.len() {
        16 => {
            let mut enc = cbc::Encryptor::<Aes128>::new_from_slices(key, iv).unwrap();
            for block in data.chunks_exact_mut(16) {
                enc.encrypt_block_mut(block.into());
            }
        }
        24 => {
            let mut enc = cbc::Encryptor::<Aes192>::new_from_slices(key, iv).unwrap();
            for block in data.chunks_exact_mut(16) {
                enc.encrypt_block_mut(block.into());
            }
        }
        32 => {
            let mut enc = cbc::Encryptor::<Aes256>::new_from_slices(key, iv).unwrap();
            for block in data.chunks_exact_mut(16) {
                enc.encrypt_block_mut(block.into());
            }
        }
        _ => unreachable!(),
    }
}

fn compute_hmac(alg: HashAlg, key: &[u8], data: &[u8]) -> Vec<u8> {
    match alg {
        HashAlg::Sha1 => {
            let mut mac = <Hmac<Sha1> as Mac>::new_from_slice(key).unwrap();
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlg::Sha256 => {
            let mut mac = <Hmac<Sha256> as Mac>::new_from_slice(key).unwrap();
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlg::Sha384 => {
            let mut mac = <Hmac<Sha384> as Mac>::new_from_slice(key).unwrap();
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlg::Sha512 => {
            let mut mac = <Hmac<Sha512> as Mac>::new_from_slice(key).unwrap();
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
    }
}

fn encrypted_field(key: &[u8], iv: &[u8], value: &[u8]) -> Vec<u8> {
    let mut field = vec![0; value.len().next_multiple_of(16)];
    field[..value.len()].copy_from_slice(value);
    aes_cbc_encrypt_in_place(key, iv, &mut field);
    field
}

pub(super) fn encrypt_agile(
    plaintext: &[u8],
    password: &str,
    hash_alg: HashAlg,
    spin_count: u32,
    keys: &AgileKeys,
) -> Result<Vec<u8>, CryptoError> {
    let secret_key = &keys.package_key[..];
    let key_bits = secret_key.len() * 8;
    let hash_size = hash_alg.digest_len();
    let hash_name = match hash_alg {
        HashAlg::Sha1 => "SHA1",
        HashAlg::Sha256 => "SHA256",
        HashAlg::Sha384 => "SHA384",
        HashAlg::Sha512 => "SHA512",
    };
    let mut package = Vec::with_capacity(8 + plaintext.len().next_multiple_of(16));
    package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    for (index, chunk) in plaintext.chunks(4096).enumerate() {
        let iv = agile_derive_iv(hash_alg, &keys.data_salt, &(index as u32).to_le_bytes(), 16);
        let start = package.len();
        package.extend_from_slice(chunk);
        package.resize(start + chunk.len().next_multiple_of(16), 0);
        aes_cbc_encrypt_in_place(secret_key, &iv, &mut package[start..]);
    }

    let integrity_field = |block_key: &[u8], value: &[u8]| {
        let iv = agile_derive_iv(hash_alg, &keys.data_salt, block_key, 16);
        encrypted_field(secret_key, &iv, value)
    };
    let enc_hmac_key = integrity_field(&AGILE_HMAC_KEY_BLOCK_KEY, &keys.hmac_key);
    let mac = compute_hmac(hash_alg, &keys.hmac_key, &package);
    let enc_hmac_val = integrity_field(&AGILE_HMAC_VALUE_BLOCK_KEY, &mac);

    let spun = agile_spin_hash(hash_alg, &keys.password_salt, password, spin_count);
    let password_field = |block_key: &[u8], value: &[u8]| {
        let key = agile_derive_key(hash_alg, &spun, block_key, secret_key.len());
        encrypted_field(&key, &keys.password_salt, value)
    };
    let enc_vi = password_field(&AGILE_VERIFIER_HASH_INPUT_BLOCK_KEY, &keys.verifier[..]);
    let verifier_hash = hash_alg.hash(&[&keys.verifier[..]]);
    let enc_vv = password_field(&AGILE_VERIFIER_HASH_VALUE_BLOCK_KEY, &verifier_hash);
    let enc_ek = password_field(&AGILE_ENCRYPTED_KEY_VALUE_BLOCK_KEY, secret_key);
    let xml = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">"#,
            r#"<keyData saltSize="16" blockSize="16" keyBits="{key_bits}" hashSize="{hash_size}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="{hash_name}" saltValue="{kd_salt}"/>"#,
            r#"<dataIntegrity encryptedHmacKey="{enc_hmac_key}" encryptedHmacValue="{enc_hmac_val}"/>"#,
            r#"<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">"#,
            r#"<p:encryptedKey spinCount="{spin_count}" saltSize="16" blockSize="16" keyBits="{key_bits}" hashSize="{hash_size}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="{hash_name}" saltValue="{pk_salt}" encryptedVerifierHashInput="{enc_vi}" encryptedVerifierHashValue="{enc_vv}" encryptedKeyValue="{enc_ek}"/>"#,
            r#"</keyEncryptor></keyEncryptors></encryption>"#
        ),
        key_bits = key_bits,
        hash_size = hash_size,
        hash_name = hash_name,
        spin_count = spin_count,
        kd_salt = BASE64.encode(keys.data_salt),
        enc_hmac_key = BASE64.encode(&enc_hmac_key),
        enc_hmac_val = BASE64.encode(&enc_hmac_val),
        pk_salt = BASE64.encode(keys.password_salt),
        enc_vi = BASE64.encode(&enc_vi),
        enc_vv = BASE64.encode(&enc_vv),
        enc_ek = BASE64.encode(&enc_ek),
    );

    let mut info = Vec::new();
    info.extend_from_slice(&4u16.to_le_bytes()); // major
    info.extend_from_slice(&4u16.to_le_bytes()); // minor
    info.extend_from_slice(&0x40u32.to_le_bytes()); // flags
    info.extend_from_slice(xml.as_bytes());

    container(&info, &package).map_err(|e| CryptoError::EncryptionFailed(e.to_string()))
}

// MS-OFFCRYPTO 2.1.5–2.1.8: byte-counted UTF-16LE strings, padded to 4 bytes.
fn unicode_lpp4(out: &mut Vec<u8>, value: &str) {
    let units = value.encode_utf16();
    out.extend_from_slice(&((units.clone().count() * 2) as u32).to_le_bytes());
    for unit in units {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out.resize(out.len().next_multiple_of(4), 0);
}

fn words(out: &mut Vec<u8>, fields: &[u32]) {
    for field in fields {
        out.extend_from_slice(&field.to_le_bytes());
    }
}

fn container(info: &[u8], package: &[u8]) -> io::Result<Vec<u8>> {
    let mut file =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new()))?;
    let root = "\x06DataSpaces";
    file.create_storage_all(format!("{root}/DataSpaceInfo"))?;
    file.create_storage_all(format!("{root}/TransformInfo/StrongEncryptionTransform"))?;
    let mut stream = |path: &str, data: &[u8]| {
        let mut entry = file.create_stream(path)?;
        entry.write_all(data)?;
        entry.flush()
    };
    stream("EncryptionInfo", info)?;
    stream("EncryptedPackage", package)?;

    let mut version = Vec::new();
    unicode_lpp4(&mut version, "Microsoft.Container.DataSpaces");
    words(&mut version, &[1, 1, 1]); // reader, updater and writer version 1.0
    stream(&format!("{root}/Version"), &version)?;

    let mut map = Vec::new();
    words(&mut map, &[8, 1, 104, 1, 0]); // one 104-byte entry, one stream reference
    unicode_lpp4(&mut map, "EncryptedPackage");
    unicode_lpp4(&mut map, "StrongEncryptionDataSpace");
    stream(&format!("{root}/DataSpaceMap"), &map)?;

    let mut definition = Vec::new();
    words(&mut definition, &[8, 1]);
    unicode_lpp4(&mut definition, "StrongEncryptionTransform");
    stream(
        &format!("{root}/DataSpaceInfo/StrongEncryptionDataSpace"),
        &definition,
    )?;

    let mut transform = Vec::new();
    words(&mut transform, &[88, 1]); // TransformInfoHeader length, ID is a GUID
    unicode_lpp4(&mut transform, "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}");
    unicode_lpp4(&mut transform, "Microsoft.Container.EncryptionTransform");
    words(&mut transform, &[1, 1, 1]); // reader, updater and writer version 1.0
    words(&mut transform, &[0, 0, 0, 4]); // Agile transform trailer; EncryptionInfo supplies the cipher parameters
    stream(
        &format!("{root}/TransformInfo/StrongEncryptionTransform/\x06Primary"),
        &transform,
    )?;
    file.flush()?;
    Ok(file.into_inner().into_inner())
}
