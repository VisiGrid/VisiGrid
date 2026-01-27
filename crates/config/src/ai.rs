// AI configuration and secrets management
//
// API keys are stored securely using:
// 1. System keychain (preferred)
// 2. Environment variables (fallback for CI/headless)
//
// Keys are NEVER stored in settings.json

use std::env;

/// Service name for keychain storage
const KEYCHAIN_SERVICE: &str = "visigrid";

/// Source of an API key
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum KeySource {
    /// Key retrieved from system keychain
    Keychain,
    /// Key retrieved from environment variable
    Environment,
    /// No key found
    None,
}

impl KeySource {
    pub fn as_str(&self) -> &'static str {
        match self {
            KeySource::Keychain => "keychain",
            KeySource::Environment => "environment",
            KeySource::None => "none",
        }
    }
}

/// Result of key lookup
#[derive(Debug, Clone)]
pub struct KeyLookup {
    pub key: Option<String>,
    pub source: KeySource,
}

/// Get the environment variable name for a provider
fn env_var_name(provider: &str) -> String {
    format!("VISIGRID_{}_KEY", provider.to_uppercase())
}

/// Get the keychain account name for a provider
fn keychain_account(provider: &str) -> String {
    format!("ai/{}", provider.to_lowercase())
}

/// Get an API key for the specified provider
///
/// Checks in order:
/// 1. System keychain
/// 2. Environment variable (VISIGRID_OPENAI_KEY, etc.)
pub fn get_api_key(provider: &str) -> KeyLookup {
    // Try keychain first
    #[cfg(feature = "keychain")]
    {
        if let Ok(entry) = keyring::Entry::new(KEYCHAIN_SERVICE, &keychain_account(provider)) {
            if let Ok(key) = entry.get_password() {
                return KeyLookup {
                    key: Some(key),
                    source: KeySource::Keychain,
                };
            }
        }
    }

    // Fall back to environment variable
    let env_name = env_var_name(provider);
    if let Ok(key) = env::var(&env_name) {
        if !key.is_empty() {
            return KeyLookup {
                key: Some(key),
                source: KeySource::Environment,
            };
        }
    }

    KeyLookup {
        key: None,
        source: KeySource::None,
    }
}

/// Store an API key in the system keychain
#[cfg(feature = "keychain")]
pub fn set_api_key(provider: &str, key: &str) -> Result<(), String> {
    let entry = keyring::Entry::new(KEYCHAIN_SERVICE, &keychain_account(provider))
        .map_err(|e| format!("Failed to create keychain entry: {}", e))?;

    entry
        .set_password(key)
        .map_err(|e| format!("Failed to store key in keychain: {}", e))
}

#[cfg(not(feature = "keychain"))]
pub fn set_api_key(_provider: &str, _key: &str) -> Result<(), String> {
    Err("Keychain support not enabled. Set VISIGRID_<PROVIDER>_KEY environment variable instead.".to_string())
}

/// Delete an API key from the system keychain
#[cfg(feature = "keychain")]
pub fn delete_api_key(provider: &str) -> Result<(), String> {
    let entry = keyring::Entry::new(KEYCHAIN_SERVICE, &keychain_account(provider))
        .map_err(|e| format!("Failed to access keychain entry: {}", e))?;

    entry
        .delete_credential()
        .map_err(|e| format!("Failed to delete key from keychain: {}", e))
}

#[cfg(not(feature = "keychain"))]
pub fn delete_api_key(_provider: &str) -> Result<(), String> {
    Err("Keychain support not enabled.".to_string())
}

/// Check if keychain support is available
pub fn keychain_available() -> bool {
    #[cfg(feature = "keychain")]
    {
        // Try to create a test entry to verify keychain access
        keyring::Entry::new(KEYCHAIN_SERVICE, "test").is_ok()
    }
    #[cfg(not(feature = "keychain"))]
    {
        false
    }
}

/// Diagnostic information about AI configuration
#[derive(Debug)]
pub struct AIDiagnostics {
    pub provider: String,
    pub model: String,
    pub key_present: bool,
    pub key_source: KeySource,
    pub keychain_available: bool,
    pub endpoint: Option<String>,
    pub privacy_mode: bool,
    pub allow_proposals: bool,
}

impl AIDiagnostics {
    /// Create diagnostics from current settings
    pub fn from_settings(settings: &crate::settings::AISettings) -> Self {
        let provider_str = match settings.provider {
            crate::settings::AIProvider::None => "none",
            crate::settings::AIProvider::Local => "local",
            crate::settings::AIProvider::OpenAI => "openai",
            crate::settings::AIProvider::Anthropic => "anthropic",
        };

        let key_lookup = if settings.provider.is_enabled() {
            get_api_key(provider_str)
        } else {
            KeyLookup {
                key: None,
                source: KeySource::None,
            }
        };

        Self {
            provider: provider_str.to_string(),
            model: settings.effective_model().to_string(),
            key_present: key_lookup.key.is_some(),
            key_source: key_lookup.source,
            keychain_available: keychain_available(),
            endpoint: if matches!(settings.provider, crate::settings::AIProvider::Local) {
                Some(settings.effective_endpoint().to_string())
            } else {
                None
            },
            privacy_mode: settings.privacy_mode,
            allow_proposals: settings.allow_proposals,
        }
    }
}

impl std::fmt::Display for AIDiagnostics {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "AI Configuration")?;
        writeln!(f, "──────────────────────────────")?;
        writeln!(f, "Provider:          {}", self.provider)?;
        writeln!(f, "Model:             {}", self.model)?;
        writeln!(f, "Key present:       {}", if self.key_present { "yes" } else { "no" })?;
        writeln!(f, "Key source:        {}", self.key_source.as_str())?;
        writeln!(f, "Keychain available:{}", if self.keychain_available { "yes" } else { "no" })?;
        if let Some(endpoint) = &self.endpoint {
            writeln!(f, "Endpoint:          {}", endpoint)?;
        }
        writeln!(f, "Privacy mode:      {}", if self.privacy_mode { "on" } else { "off" })?;
        writeln!(f, "Allow proposals:   {}", if self.allow_proposals { "yes" } else { "no" })?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_env_var_name() {
        assert_eq!(env_var_name("openai"), "VISIGRID_OPENAI_KEY");
        assert_eq!(env_var_name("anthropic"), "VISIGRID_ANTHROPIC_KEY");
        assert_eq!(env_var_name("OpenAI"), "VISIGRID_OPENAI_KEY");
    }

    #[test]
    fn test_keychain_account() {
        assert_eq!(keychain_account("openai"), "ai/openai");
        assert_eq!(keychain_account("OpenAI"), "ai/openai");
        assert_eq!(keychain_account("anthropic"), "ai/anthropic");
    }

    #[test]
    fn test_key_lookup_from_env() {
        // Set a test env var
        env::set_var("VISIGRID_TESTPROVIDER_KEY", "test-key-123");

        let lookup = get_api_key("testprovider");
        assert_eq!(lookup.source, KeySource::Environment);
        assert_eq!(lookup.key, Some("test-key-123".to_string()));

        // Clean up
        env::remove_var("VISIGRID_TESTPROVIDER_KEY");
    }

    #[test]
    fn test_key_lookup_missing() {
        let lookup = get_api_key("nonexistent_provider_xyz");
        assert_eq!(lookup.source, KeySource::None);
        assert!(lookup.key.is_none());
    }
}
