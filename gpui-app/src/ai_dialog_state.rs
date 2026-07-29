//! Cycle banner, AI Settings, and Ask AI dialog state.
//! Extracted from app.rs 2026-07-29 (pure move).

use gpui::*;

// ============================================================================
// Cycle Banner State
// ============================================================================

/// State for the dismissible cycle-status banner.
///
/// Shown after import or iteration toggle when cycles, freeze, or iteration is active.
/// Dismissed by user click; stays dismissed until next import or explicit toggle.
#[derive(Clone, Debug, Default)]
pub struct CycleBannerState {
    /// Whether the banner is currently visible.
    pub visible: bool,
    /// Whether the user has explicitly dismissed the banner.
    /// When true, don't re-show on recalc — only on new import or explicit toggle.
    pub dismissed: bool,
}

impl CycleBannerState {
    /// Show the banner (only if not already dismissed by user).
    pub fn show(&mut self) {
        if !self.dismissed {
            self.visible = true;
        }
    }

    /// Show the banner unconditionally (after import or explicit toggle — overrides dismiss).
    pub fn show_force(&mut self) {
        self.dismissed = false;
        self.visible = true;
    }

    /// User dismisses the banner. Stays dismissed until next import/reimport.
    pub fn dismiss(&mut self) {
        self.visible = false;
        self.dismissed = true;
    }

    /// Reset for new file / new import. Clears dismissed state.
    pub fn reset_for_new_file(&mut self) {
        self.visible = false;
        self.dismissed = false;
    }
}

/// Lightweight UTC timestamp without external chrono dependency.
/// Returns ISO 8601 format: "2024-01-15T14:30:00Z"
pub fn chrono_lite_utc() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};

    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default();
    let secs = now.as_secs();

    // Calculate date/time components from Unix timestamp
    // Days since epoch
    let days = secs / 86400;
    let time_secs = secs % 86400;

    let hours = time_secs / 3600;
    let minutes = (time_secs % 3600) / 60;
    let seconds = time_secs % 60;

    // Year calculation (simplified leap year handling)
    let mut year = 1970;
    let mut remaining_days = days as i64;

    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if remaining_days < days_in_year {
            break;
        }
        remaining_days -= days_in_year;
        year += 1;
    }

    // Month calculation
    let month_days: &[i64] = if is_leap_year(year) {
        &[31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        &[31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month = 1;
    for &days in month_days {
        if remaining_days < days {
            break;
        }
        remaining_days -= days;
        month += 1;
    }

    let day = remaining_days + 1;

    format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
        year, month, day, hours, minutes, seconds
    )
}

pub fn is_leap_year(year: i64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

// ============================================================================
// AI Settings Dialog State
// ============================================================================

/// Selected AI provider in the settings dialog
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AIProviderOption {
    #[default]
    None,
    Local,
    OpenAI,
    Anthropic,
    Gemini,
    Grok,
}

impl AIProviderOption {
    pub fn label(&self) -> &'static str {
        match self {
            AIProviderOption::None => "Disabled",
            AIProviderOption::Local => "Local (Ollama)",
            AIProviderOption::OpenAI => "OpenAI",
            AIProviderOption::Anthropic => "Anthropic",
            AIProviderOption::Gemini => "Google Gemini",
            AIProviderOption::Grok => "xAI Grok",
        }
    }

    pub fn to_config(&self) -> visigrid_config::settings::AIProvider {
        match self {
            AIProviderOption::None => visigrid_config::settings::AIProvider::None,
            AIProviderOption::Local => visigrid_config::settings::AIProvider::Local,
            AIProviderOption::OpenAI => visigrid_config::settings::AIProvider::OpenAI,
            AIProviderOption::Anthropic => visigrid_config::settings::AIProvider::Anthropic,
            AIProviderOption::Gemini => visigrid_config::settings::AIProvider::Gemini,
            AIProviderOption::Grok => visigrid_config::settings::AIProvider::Grok,
        }
    }

    pub fn from_config(provider: visigrid_config::settings::AIProvider) -> Self {
        match provider {
            visigrid_config::settings::AIProvider::None => AIProviderOption::None,
            visigrid_config::settings::AIProvider::Local => AIProviderOption::Local,
            visigrid_config::settings::AIProvider::OpenAI => AIProviderOption::OpenAI,
            visigrid_config::settings::AIProvider::Anthropic => AIProviderOption::Anthropic,
            visigrid_config::settings::AIProvider::Gemini => AIProviderOption::Gemini,
            visigrid_config::settings::AIProvider::Grok => AIProviderOption::Grok,
        }
    }
}

/// Focus state for AI settings dialog
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AISettingsFocus {
    #[default]
    Provider,
    Model,
    Endpoint,
    KeyInput,
}

/// Test status for AI key verification
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum AITestStatus {
    #[default]
    Idle,
    Testing,
    Success(String),  // Model name or version returned
    Error(String),    // Error message
}

/// State for the AI Settings dialog
#[derive(Debug, Clone, Default)]
pub struct AISettingsDialogState {
    /// Selected AI provider
    pub provider: AIProviderOption,
    /// Whether the provider dropdown is open
    pub provider_dropdown_open: bool,

    /// Model identifier (empty = use provider default)
    pub model: String,

    /// Custom endpoint for Local provider (Ollama)
    pub endpoint: String,

    /// Privacy mode: minimize data sent to AI
    pub privacy_mode: bool,

    /// Allow AI to propose cell changes
    pub allow_proposals: bool,

    /// Which field has focus
    pub focus: AISettingsFocus,

    /// API key input (for setting new key)
    pub key_input: String,
    /// Whether key is currently stored
    pub key_present: bool,
    /// Source of current key ("keychain", "environment", "none")
    pub key_source: String,

    /// Test status
    pub test_status: AITestStatus,

    /// Error message
    pub error: Option<String>,
}

impl AISettingsDialogState {
    /// Reset to defaults
    pub fn reset(&mut self) {
        *self = Self::default();
    }

    /// Load current settings from config crate using ResolvedAIConfig
    pub fn load_from_config(&mut self) {
        use visigrid_config::ai::ResolvedAIConfig;
        use visigrid_config::settings::Settings;

        // Use the single source of truth
        let config = ResolvedAIConfig::load();
        let settings = Settings::load();

        self.provider = AIProviderOption::from_config(config.provider);
        // Load model from settings (not resolved) so user sees what they actually set
        self.model = settings.ai.model.clone();
        self.endpoint = config.endpoint.unwrap_or_default();
        self.privacy_mode = config.privacy_mode;
        self.allow_proposals = config.allow_proposals;

        // Key status from resolved config
        self.key_present = config.api_key.is_some();
        self.key_source = config.key_source.as_str().to_string();

        self.key_input.clear();
        self.test_status = AITestStatus::Idle;
        self.error = None;
    }

    /// Save current state to config
    pub fn save_to_config(&self) -> Result<(), String> {
        let mut settings = visigrid_config::settings::Settings::load();

        settings.ai.provider = self.provider.to_config();
        settings.ai.model = self.model.clone();
        settings.ai.endpoint = if self.endpoint.is_empty() {
            None
        } else {
            Some(self.endpoint.clone())
        };
        settings.ai.privacy_mode = self.privacy_mode;
        settings.ai.allow_proposals = self.allow_proposals;

        settings.save()
    }

    /// Get the effective model name (user-specified or provider default)
    pub fn effective_model(&self) -> &str {
        if self.model.is_empty() {
            match self.provider {
                AIProviderOption::None => "",
                AIProviderOption::Local => "llama3:8b",
                AIProviderOption::OpenAI => "gpt-4o-mini",
                AIProviderOption::Anthropic => "claude-sonnet-4-20250514",
                AIProviderOption::Gemini => "gemini-2.0-flash",
                AIProviderOption::Grok => "grok-3-mini-fast-beta",
            }
        } else {
            &self.model
        }
    }

    /// Check if the current provider requires an API key
    pub fn needs_api_key(&self) -> bool {
        matches!(
            self.provider,
            AIProviderOption::OpenAI | AIProviderOption::Anthropic | AIProviderOption::Gemini | AIProviderOption::Grok
        )
    }

    /// Get context policy description based on privacy mode
    pub fn context_policy(&self) -> &'static str {
        if self.privacy_mode {
            "Minimal: only current cell and selection"
        } else {
            "Extended: includes nearby cells and context"
        }
    }
}

// ============================================================================
// Ask AI Dialog State
// ============================================================================

/// Which AI verb is active in the dialog
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AiVerb {
    /// Insert Formula with AI (Ctrl+Shift+A) — single-cell write contract
    #[default]
    InsertFormula,
    /// Analyze with AI (Ctrl+Shift+E) — read-only contract
    Analyze,
}

/// Status of an Ask AI request
#[derive(Debug, Clone, Default, PartialEq)]
pub enum AskAIStatus {
    #[default]
    Idle,
    Loading,
    Success,
    Error(String),
}

/// How context range is selected
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum AskAIContextMode {
    /// Use current cell selection
    #[default]
    CurrentSelection,
    /// Use contiguous region around active cell
    CurrentRegion,
    /// Use entire used range of sheet
    EntireUsedRange,
}

impl AskAIContextMode {
    pub fn label(&self) -> &'static str {
        match self {
            Self::CurrentSelection => "Current selection",
            Self::CurrentRegion => "Current region",
            Self::EntireUsedRange => "Entire used range",
        }
    }
}

/// What was actually sent to the AI (for transparency)
#[derive(Debug, Clone, Default)]
pub struct AskAISentContext {
    /// Provider used
    pub provider: String,
    /// Model used
    pub model: String,
    /// Privacy mode enabled
    pub privacy_mode: bool,
    /// Effective range after truncation
    pub range_display: String,
    /// Rows actually sent
    pub rows_sent: usize,
    /// Columns actually sent
    pub cols_sent: usize,
    /// Total cells sent
    pub total_cells: usize,
    /// Whether headers were detected/included
    pub headers_included: bool,
    /// Truncation applied
    pub truncation: AskAITruncation,
}

/// What truncation was applied
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum AskAITruncation {
    #[default]
    None,
    Rows,
    Cols,
    Both,
}

impl AskAITruncation {
    pub fn label(&self) -> &'static str {
        match self {
            Self::None => "None",
            Self::Rows => "Rows truncated",
            Self::Cols => "Columns truncated",
            Self::Both => "Rows and columns truncated",
        }
    }
}

/// State for the AI dialog (shared by InsertFormula and Analyze verbs)
#[derive(Debug, Clone, Default)]
pub struct AskAIDialogState {
    /// Which verb is active
    pub verb: AiVerb,

    /// User's question
    pub question: String,

    /// Context selection mode
    pub context_mode: AskAIContextMode,

    /// Context summary (e.g., "Sheet1!A1:F200")
    pub context_summary: String,

    /// Context range info for insertion (top-left cell)
    pub context_top_left: Option<(usize, usize)>,

    /// Selected range bounds (start_row, start_col, end_row, end_col)
    pub selected_range: Option<(usize, usize, usize, usize)>,

    /// What was actually sent to AI (populated after request)
    pub sent_context: Option<AskAISentContext>,

    /// Request status
    pub status: AskAIStatus,

    /// Unique request ID for tracking
    pub request_id: Option<String>,

    /// AI explanation (when successful)
    pub explanation: Option<String>,

    /// Proposed formula (when AI suggests one)
    pub formula: Option<String>,

    /// Whether formula is valid for insertion
    pub formula_valid: bool,

    /// Formula validation error (if any)
    pub formula_error: Option<String>,

    /// Warnings from context extraction or AI
    pub warnings: Vec<String>,

    /// Error message (if failed)
    pub error: Option<String>,

    /// Raw model response (for debugging)
    pub raw_response: Option<String>,

    /// Analysis text from Analyze verb (read-only response)
    pub response_text: Option<String>,

    /// Last insertion confirmation (e.g., "Inserted into A1")
    pub last_insertion: Option<String>,

    /// Whether current formula was already inserted (prevent double-insert)
    pub inserted: bool,

    /// Whether "Sent to AI" panel is expanded
    pub sent_panel_expanded: bool,

    /// Whether context selector is open
    pub context_selector_open: bool,
}

impl AskAIDialogState {
    /// Reset to initial state (verb is set by the caller before reset)
    pub fn reset(&mut self) {
        // Note: verb is NOT reset here — it's set by show_ask_ai/show_analyze before reset
        self.question.clear();
        self.context_mode = AskAIContextMode::CurrentSelection;
        self.context_summary.clear();
        self.context_top_left = None;
        self.selected_range = None;
        self.sent_context = None;
        self.status = AskAIStatus::Idle;
        self.request_id = None;
        self.explanation = None;
        self.formula = None;
        self.formula_valid = false;
        self.formula_error = None;
        self.response_text = None;
        self.warnings.clear();
        self.error = None;
        self.raw_response = None;
        self.last_insertion = None;
        self.inserted = false;
        self.sent_panel_expanded = false;
        self.context_selector_open = false;
    }

    /// Clear response state (keep question, context, and verb)
    pub fn clear_response(&mut self) {
        self.sent_context = None;
        self.status = AskAIStatus::Idle;
        self.request_id = None;
        self.explanation = None;
        self.formula = None;
        self.formula_valid = false;
        self.formula_error = None;
        self.response_text = None;
        self.error = None;
        self.raw_response = None;
        self.last_insertion = None;
        self.inserted = false;
    }

    /// Check if Insert Formula button should be enabled
    pub fn can_insert(&self) -> bool {
        matches!(self.status, AskAIStatus::Success)
            && self.formula.is_some()
            && self.formula_valid
            && !self.inserted
    }

    /// Check if a request is in flight
    pub fn is_loading(&self) -> bool {
        matches!(self.status, AskAIStatus::Loading)
    }

    /// Check if retry is available
    pub fn can_retry(&self) -> bool {
        !self.is_loading() && !self.question.is_empty()
    }
}
