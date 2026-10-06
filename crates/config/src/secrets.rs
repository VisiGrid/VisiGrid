//! Secrets in the OS keychain (service `visigrid`), by account name: AI
//! provider keys live under `ai/<provider>`, VisiBooks keys under
//! `visibooks/<server origin>`. Never written to settings or workbooks.

const KEYCHAIN_SERVICE: &str = "visigrid";

/// The secret stored for `account`, if any.
pub fn get(account: &str) -> Option<String> {
    #[cfg(feature = "keychain")]
    {
        let entry = keyring::Entry::new(KEYCHAIN_SERVICE, account).ok()?;
        entry.get_password().ok().filter(|s| !s.is_empty())
    }
    #[cfg(not(feature = "keychain"))]
    {
        let _ = account;
        None
    }
}

/// Store `secret` for `account`, replacing any earlier one.
pub fn set(account: &str, secret: &str) -> Result<(), String> {
    #[cfg(feature = "keychain")]
    {
        keyring::Entry::new(KEYCHAIN_SERVICE, account)
            .and_then(|e| e.set_password(secret))
            .map_err(|e| format!("couldn't store the key in the system keychain: {e}"))
    }
    #[cfg(not(feature = "keychain"))]
    {
        let _ = (account, secret);
        Err("this build has no keychain support".into())
    }
}

/// Remove the secret for `account`; removing one that isn't there is fine.
pub fn delete(account: &str) -> Result<(), String> {
    #[cfg(feature = "keychain")]
    {
        match keyring::Entry::new(KEYCHAIN_SERVICE, account).and_then(|e| e.delete_credential()) {
            Ok(()) | Err(keyring::Error::NoEntry) => Ok(()),
            Err(e) => Err(format!("couldn't remove the key from the system keychain: {e}")),
        }
    }
    #[cfg(not(feature = "keychain"))]
    {
        let _ = account;
        Ok(())
    }
}
