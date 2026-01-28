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

// ============================================================================
// Resolved AI Configuration (single source of truth)
// ============================================================================

/// The effective AI configuration, fully resolved from all sources.
/// This is the single source of truth for runtime AI behavior.
#[derive(Debug, Clone)]
pub struct ResolvedAIConfig {
    /// Effective provider (None, Local, OpenAI, Anthropic, Gemini, Grok)
    pub provider: crate::settings::AIProvider,
    /// Effective model (resolved from settings or provider default)
    pub model: String,
    /// Endpoint for Local provider (resolved with default)
    pub endpoint: Option<String>,
    /// Privacy mode setting
    pub privacy_mode: bool,
    /// Whether AI can propose cell changes
    pub allow_proposals: bool,
    /// API key (if available and provider needs one)
    pub api_key: Option<String>,
    /// Source of the API key
    pub key_source: KeySource,
    /// Capabilities implemented for this provider
    pub capabilities: crate::settings::ProviderCapabilities,
    /// Overall status
    pub status: AIConfigStatus,
    /// Human-readable reason if not ready
    pub blocking_reason: Option<String>,
}

/// Status of the AI configuration
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AIConfigStatus {
    /// AI is disabled (provider = None)
    Disabled,
    /// Configuration is valid and provider has implemented capabilities
    Ready,
    /// Configuration is valid but provider not yet implemented
    NotImplemented,
    /// Provider is configured but API key is missing
    MissingKey,
    /// Configuration error (e.g., keychain access failed)
    Error,
}

impl AIConfigStatus {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Disabled => "disabled",
            Self::Ready => "ready",
            Self::NotImplemented => "not_implemented",
            Self::MissingKey => "missing_key",
            Self::Error => "error",
        }
    }

    pub fn is_ready(&self) -> bool {
        matches!(self, Self::Ready)
    }

    /// Returns true if configuration is valid (key present if needed)
    /// but provider may not have capabilities implemented yet
    pub fn is_configured(&self) -> bool {
        matches!(self, Self::Ready | Self::NotImplemented)
    }
}

impl ResolvedAIConfig {
    /// Resolve the effective AI configuration from settings.
    /// This is the single entry point for all AI config resolution.
    pub fn from_settings(settings: &crate::settings::AISettings) -> Self {
        use crate::settings::AIProvider;

        let provider = settings.provider;
        let capabilities = provider.capabilities();

        // If disabled, return early
        if !provider.is_enabled() {
            return Self {
                provider,
                model: String::new(),
                endpoint: None,
                privacy_mode: settings.privacy_mode,
                allow_proposals: settings.allow_proposals,
                api_key: None,
                key_source: KeySource::None,
                capabilities,
                status: AIConfigStatus::Disabled,
                blocking_reason: None,
            };
        }

        // Resolve model (use default if not specified)
        let model = settings.effective_model().to_string();

        // Resolve endpoint (for Local provider)
        let endpoint = if matches!(provider, AIProvider::Local) {
            Some(settings.effective_endpoint().to_string())
        } else {
            None
        };

        // Get API key if needed
        let (api_key, key_source, key_status, key_reason) = if provider.needs_api_key() {
            let provider_name = provider.name();
            let lookup = get_api_key(provider_name);

            match lookup.key {
                Some(key) => (Some(key), lookup.source, None, None),
                None => (
                    None,
                    KeySource::None,
                    Some(AIConfigStatus::MissingKey),
                    Some(format!(
                        "No API key found. Set via keychain or {}",
                        env_var_name(provider_name)
                    )),
                ),
            }
        } else {
            // Local provider doesn't need a key
            (None, KeySource::None, None, None)
        };

        // Determine final status:
        // 1. If key is missing, that's the blocking issue
        // 2. If key is present but no capabilities implemented, status is NotImplemented
        // 3. If key is present and capabilities exist, status is Ready
        let (status, blocking_reason) = if let Some(s) = key_status {
            (s, key_reason)
        } else if !capabilities.any_implemented() {
            (
                AIConfigStatus::NotImplemented,
                Some(format!(
                    "Provider {} is configured but not yet implemented",
                    provider.name()
                )),
            )
        } else {
            (AIConfigStatus::Ready, None)
        };

        Self {
            provider,
            model,
            endpoint,
            privacy_mode: settings.privacy_mode,
            allow_proposals: settings.allow_proposals,
            api_key,
            key_source,
            capabilities,
            status,
            blocking_reason,
        }
    }

    /// Load settings and resolve in one call (convenience method)
    pub fn load() -> Self {
        let settings = crate::settings::Settings::load();
        Self::from_settings(&settings.ai)
    }

    /// Context policy description based on privacy mode
    pub fn context_policy(&self) -> &'static str {
        if self.privacy_mode {
            "minimal"
        } else {
            "extended"
        }
    }

    /// Provider display name
    pub fn provider_name(&self) -> &'static str {
        self.provider.name()
    }
}

// ============================================================================
// Configuration Validation
// ============================================================================

/// Result of configuration validation
#[derive(Debug, Clone)]
pub enum ValidationResult {
    /// Configuration is valid
    Valid(String),
    /// Configuration has issues
    Invalid(String),
    /// Validation was skipped (AI disabled)
    Skipped(String),
}

impl ValidationResult {
    pub fn as_str(&self) -> &str {
        match self {
            Self::Valid(msg) => msg,
            Self::Invalid(msg) => msg,
            Self::Skipped(msg) => msg,
        }
    }

    pub fn is_valid(&self) -> bool {
        matches!(self, Self::Valid(_))
    }
}

// Keep the old name as an alias for compatibility
pub type AITestResult = ValidationResult;

impl ResolvedAIConfig {
    /// Validate the AI configuration.
    /// This checks credentials and basic reachability, NOT feature functionality.
    /// For Local provider: pings Ollama endpoint.
    /// For cloud providers: confirms key is present (no network call).
    pub fn validate_config(&self) -> ValidationResult {
        use crate::settings::AIProvider;

        match self.status {
            AIConfigStatus::Disabled => {
                ValidationResult::Skipped("AI is disabled".to_string())
            }
            AIConfigStatus::MissingKey => {
                ValidationResult::Invalid("No API key configured".to_string())
            }
            AIConfigStatus::Error => {
                ValidationResult::Invalid(
                    self.blocking_reason.clone().unwrap_or_else(|| "Configuration error".to_string())
                )
            }
            AIConfigStatus::NotImplemented => {
                // Config is valid, but capabilities not implemented yet
                let msg = if self.api_key.is_some() {
                    format!("API key present ({})", self.key_source.as_str())
                } else {
                    "Configured".to_string()
                };
                ValidationResult::Valid(format!("{} - provider not yet implemented", msg))
            }
            AIConfigStatus::Ready => {
                match self.provider {
                    AIProvider::Local => {
                        // Try to reach Ollama endpoint
                        let endpoint = self.endpoint.as_deref().unwrap_or("http://localhost:11434");
                        let url = format!("{}/api/tags", endpoint);

                        // Simple HTTP check with timeout using curl
                        match std::process::Command::new("curl")
                            .args(["-s", "-o", "/dev/null", "-w", "%{http_code}", "--max-time", "3", &url])
                            .output()
                        {
                            Ok(output) => {
                                let code = String::from_utf8_lossy(&output.stdout);
                                if code.trim() == "200" {
                                    ValidationResult::Valid("Ollama reachable".to_string())
                                } else {
                                    ValidationResult::Invalid(format!("Ollama returned HTTP {}", code.trim()))
                                }
                            }
                            Err(e) => {
                                ValidationResult::Invalid(format!("Connection failed: {}", e))
                            }
                        }
                    }
                    AIProvider::OpenAI | AIProvider::Anthropic | AIProvider::Gemini | AIProvider::Grok => {
                        // For cloud providers, confirm key is present
                        // Actual API validation happens on first use
                        ValidationResult::Valid(format!(
                            "API key present ({})",
                            self.key_source.as_str()
                        ))
                    }
                    AIProvider::None => {
                        ValidationResult::Skipped("AI is disabled".to_string())
                    }
                }
            }
        }
    }

    /// Backward-compatible alias for validate_config()
    #[deprecated(since = "0.3.5", note = "Use validate_config() instead")]
    pub fn test_connection(&self) -> ValidationResult {
        self.validate_config()
    }
}

// ============================================================================
// Diagnostics (for CLI doctor and debugging)
// ============================================================================

/// Diagnostic information about AI configuration
#[derive(Debug)]
pub struct AIDiagnostics {
    pub provider: String,
    pub model: String,
    pub status: AIConfigStatus,
    pub key_present: bool,
    pub key_source: KeySource,
    pub keychain_available: bool,
    pub endpoint: Option<String>,
    pub privacy_mode: bool,
    pub allow_proposals: bool,
    /// Capabilities implemented for this provider
    pub capabilities: crate::settings::ProviderCapabilities,
}

impl AIDiagnostics {
    /// Create diagnostics from resolved config (preferred)
    pub fn from_resolved(config: &ResolvedAIConfig) -> Self {
        Self {
            provider: config.provider.name().to_string(),
            model: config.model.clone(),
            status: config.status,
            key_present: config.api_key.is_some(),
            key_source: config.key_source,
            keychain_available: keychain_available(),
            endpoint: config.endpoint.clone(),
            privacy_mode: config.privacy_mode,
            allow_proposals: config.allow_proposals,
            capabilities: config.capabilities,
        }
    }

    /// Create diagnostics from current settings
    pub fn from_settings(settings: &crate::settings::AISettings) -> Self {
        let config = ResolvedAIConfig::from_settings(settings);
        Self::from_resolved(&config)
    }
}

impl std::fmt::Display for AIDiagnostics {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "AI Configuration")?;
        writeln!(f, "──────────────────────────────")?;
        writeln!(f, "Provider:          {}", self.provider)?;
        writeln!(f, "Status:            {}", self.status.as_str())?;
        writeln!(f, "Model:             {}", self.model)?;
        writeln!(f, "Key present:       {}", if self.key_present { "yes" } else { "no" })?;
        writeln!(f, "Key source:        {}", self.key_source.as_str())?;
        writeln!(f, "Keychain available:{}", if self.keychain_available { "yes" } else { "no" })?;
        if let Some(endpoint) = &self.endpoint {
            writeln!(f, "Endpoint:          {}", endpoint)?;
        }
        writeln!(f, "Privacy mode:      {}", if self.privacy_mode { "on" } else { "off" })?;
        writeln!(f, "Allow proposals:   {}", if self.allow_proposals { "yes" } else { "no" })?;

        // Show capabilities
        writeln!(f, "Capabilities:")?;
        writeln!(f, "  Ask AI:          {}", if self.capabilities.ask { "yes" } else { "no" })?;
        writeln!(f, "  Explain diffs:   {}", if self.capabilities.explain_diffs { "yes" } else { "no" })?;
        writeln!(f, "  Propose changes: {}", if self.capabilities.propose { "yes" } else { "no" })?;

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
