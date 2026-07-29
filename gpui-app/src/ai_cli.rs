//! AI CLI detection — which agent runtime (claude / codex / gemini) is
//! installed, and the context files we generate for each. Extracted from
//! app.rs 2026-07-29 (pure move).

// ============================================================================
// AI CLI detection
// ============================================================================

/// Detected AI CLI runtime.
pub(crate) enum AiCli {
    Claude, // `claude` — Anthropic Claude Code
    Codex,  // `codex` — OpenAI Codex CLI
    Gemini, // `gemini` — Google Gemini CLI
}

/// All known AI CLI variants, for iterating when generating context files.
pub(crate) const ALL_AI_CLIS: &[AiCli] = &[AiCli::Claude, AiCli::Codex, AiCli::Gemini];

impl AiCli {
    pub(crate) fn binary(&self) -> &'static str {
        match self {
            AiCli::Claude => "claude",
            AiCli::Codex => "codex",
            AiCli::Gemini => "gemini",
        }
    }
    pub(crate) fn display_name(&self) -> &'static str {
        match self {
            AiCli::Claude => "Claude Code",
            AiCli::Codex => "Codex CLI",
            AiCli::Gemini => "Gemini CLI",
        }
    }
    pub(crate) fn install_hint(&self) -> &'static str {
        match self {
            AiCli::Claude => "npm i -g @anthropic-ai/claude-code",
            AiCli::Codex => "npm i -g @openai/codex",
            AiCli::Gemini => "npm i -g @google/gemini-cli",
        }
    }
    /// The filename each CLI natively discovers for project context.
    pub(crate) fn context_filename(&self) -> &'static str {
        match self {
            AiCli::Claude => "CLAUDE.md",
            AiCli::Codex => "AGENTS.md",
            AiCli::Gemini => "GEMINI.md",
        }
    }
}

/// Check PATH for known AI CLIs.
/// Respects user preference; falls back to auto-detect (Claude → Codex → Gemini).
///
/// When launched from a desktop entry the process PATH may not include
/// directories added by node version managers (nvm, fnm, volta, etc.).
/// We augment the search with common locations so detection works
/// regardless of how the app was started.
pub(crate) fn detect_ai_cli(pref: crate::settings::PreferredAiCli) -> Option<AiCli> {
    use crate::settings::PreferredAiCli;

    // If user has a preference, try that first
    let preferred = match pref {
        PreferredAiCli::Auto => None,
        PreferredAiCli::Claude => Some(AiCli::Claude),
        PreferredAiCli::Codex => Some(AiCli::Codex),
        PreferredAiCli::Gemini => Some(AiCli::Gemini),
    };
    if let Some(cli) = preferred {
        if find_ai_binary(cli.binary()).is_some() {
            return Some(cli);
        }
    }

    // Auto-detect fallback
    if find_ai_binary("claude").is_some() {
        return Some(AiCli::Claude);
    }
    if find_ai_binary("codex").is_some() {
        return Some(AiCli::Codex);
    }
    if find_ai_binary("gemini").is_some() {
        return Some(AiCli::Gemini);
    }
    None
}

/// Find an AI CLI binary on PATH or in common node-manager locations.
pub(crate) fn find_ai_binary(name: &str) -> Option<std::path::PathBuf> {
    // Fast path: already on PATH
    if let Ok(p) = which::which(name) {
        return Some(p);
    }

    // Check common node version manager bin dirs
    if let Some(home) = dirs::home_dir() {
        let candidates: Vec<Option<std::path::PathBuf>> = vec![
            // nvm: ~/.nvm/versions/node/*/bin/
            glob_first(&home.join(".nvm/versions/node"), name),
            // fnm: ~/.local/share/fnm/node-versions/*/installation/bin/
            glob_first(&home.join(".local/share/fnm/node-versions"), name),
            // volta: ~/.volta/bin/
            Some(home.join(".volta/bin").join(name)),
            // global npm/yarn: ~/.local/bin/
            Some(home.join(".local/bin").join(name)),
            // Homebrew (macOS/Linux): /opt/homebrew/bin/, /home/linuxbrew/.linuxbrew/bin/
            Some(std::path::PathBuf::from("/opt/homebrew/bin").join(name)),
            Some(std::path::PathBuf::from("/home/linuxbrew/.linuxbrew/bin").join(name)),
        ];
        for candidate in candidates.into_iter().flatten() {
            if candidate.is_file() {
                return Some(candidate);
            }
        }
    }
    None
}

/// Find the latest node version dir containing `bin/<name>`.
pub(crate) fn glob_first(versions_dir: &std::path::Path, name: &str) -> Option<std::path::PathBuf> {
    let rd = std::fs::read_dir(versions_dir).ok()?;
    let mut best: Option<std::path::PathBuf> = None;
    for entry in rd.flatten() {
        let bin = entry.path().join("bin").join(name);
        if bin.is_file() {
            // Take the lexicographically last (highest version)
            if best.as_ref().is_none_or(|b| entry.path() > *b) {
                best = Some(bin);
            }
        }
    }
    best
}

/// Default content for the system-level AI context file (~/.config/visigrid/ai/).
pub(crate) fn system_ai_context() -> &'static str {
    r#"# VisiGrid — AI Context (System)

You are working with data from VisiGrid, a native GPU-accelerated spreadsheet.

## Key facts
- File formats: .xlsx, .csv, .tsv, .ods, .json (VisiGrid native)
- Formulas use `=` prefix, Excel-compatible syntax (96+ functions)
- Cell references: A1-style (e.g., A1, B2:D10, Sheet2!A1)
- Row/column indices are 1-based in the UI, 0-based internally

## When the user pastes data
- TSV blocks from VisiGrid are tab-separated, one row per line
- Headers are auto-detected and listed as comments above the data
- "Truncated" means the selection was capped; ask about the full range if needed

## Best practices
- Prefer formulas over manual calculations when the user needs repeatability
- When generating CSV/TSV output, match the column order from the headers
- If you need more data than what was pasted, ask the user to select a larger range
  and use Ctrl+Shift+S (or the palette command "Paste Selection to Terminal")
"#
}

/// Default content for the project-level AI context file (<workbook_dir>/.visigrid/).
pub(crate) fn project_ai_context_template(workbook_name: &str) -> String {
    format!(
        r#"# VisiGrid — AI Context (Project)

## Workbook: {}

<!-- Add project-specific instructions below. -->
<!-- This file is read by AI CLIs when launched from VisiGrid. -->
<!-- Examples: column meanings, business rules, expected output formats. -->
"#,
        workbook_name
    )
}

/// Shell-quote a string with single quotes (POSIX-safe).
/// Handles embedded single quotes by ending the quote, escaping, and restarting.
pub(crate) fn shell_quote(s: &str) -> String {
    format!("'{}'", s.replace('\'', "'\\''"))
}
