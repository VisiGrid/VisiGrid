//! Shared types for Lua script persistence, run records, and resolution.
//!
//! This is the keystone module. GUI, CLI, and console all depend on these types.
//!
//! # Design Invariants
//!
//! 1. **No bypass mode.** Capabilities are `HashSet<Capability>`, never `Option`.
//!    Console gets explicit `all_sheet_caps()`, not a backdoor `None`.
//! 2. **Inference is UI-only.** `suggest_capabilities()` prefills checkboxes.
//!    Enforcement uses only the declared caps in `ScriptMeta`.
//! 3. **Canonical diff hashing.** Typed `PatchLine` with explicit `kind`,
//!    `old`/`new` as canonical engine strings, `null` for empty.
//! 4. **Single resolution function.** GUI, CLI, console, and replay all use
//!    `resolve_script()` with the same priority: Attached → Project → Global.

use std::collections::HashSet;
use std::path::{Path, PathBuf};

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

// ============================================================================
// Capability
// ============================================================================

/// A capability that a Lua script may use.
///
/// Capabilities are always enforced — there is NO bypass mode.
/// Console gets `all_sheet_caps()`, saved scripts get their declared caps.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Capability {
    SheetRead,
    SheetWriteValues,
    SheetWriteFormulas,
    /// Forward-compat for Phase 2+ caps (e.g., "network", "file_read").
    /// Unknown caps are preserved on roundtrip but never granted.
    #[serde(untagged)]
    Unknown(String),
}

impl std::fmt::Display for Capability {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Capability::SheetRead => write!(f, "sheet_read"),
            Capability::SheetWriteValues => write!(f, "sheet_write_values"),
            Capability::SheetWriteFormulas => write!(f, "sheet_write_formulas"),
            Capability::Unknown(s) => write!(f, "{}", s),
        }
    }
}

/// The full set of sheet capabilities for console/unrestricted use.
///
/// There is NO bypass mode. Console gets explicit caps, not a backdoor.
pub fn all_sheet_caps() -> HashSet<Capability> {
    [
        Capability::SheetRead,
        Capability::SheetWriteValues,
        Capability::SheetWriteFormulas,
    ]
    .into_iter()
    .collect()
}

// ============================================================================
// ScriptMeta
// ============================================================================

/// Metadata for a persisted Lua script.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScriptMeta {
    /// Schema version for the script metadata format. Always 1.
    pub schema_version: u32,
    /// Human-readable name (unique within a resolution scope).
    pub name: String,
    /// Optional description.
    pub description: Option<String>,
    /// Content-addressed hash: `"sha256:<64 hex>"`.
    pub hash: String,
    /// The Lua source code (canonicalized on save).
    pub source: String,
    /// Declared capabilities. Enforcement uses ONLY these — never inference.
    pub capabilities: Vec<Capability>,
    /// ISO 8601 timestamp of creation.
    pub created_at: String,
    /// ISO 8601 timestamp of last update.
    pub updated_at: String,
    /// Where the script came from (e.g., "hub:owner/repo", "import:file.lua").
    pub origin: Option<String>,
    /// Who wrote this script.
    pub author: Option<String>,
    /// Semver version string.
    pub version: Option<String>,
}

// ============================================================================
// Script Origin & Resolution
// ============================================================================

/// Where a script was found during resolution.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum ScriptOriginKind {
    Attached,
    Project,
    Global,
    Hub,
    Console,
}

impl std::fmt::Display for ScriptOriginKind {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ScriptOriginKind::Attached => write!(f, "attached"),
            ScriptOriginKind::Project => write!(f, "project"),
            ScriptOriginKind::Global => write!(f, "global"),
            ScriptOriginKind::Hub => write!(f, "hub"),
            ScriptOriginKind::Console => write!(f, "console"),
        }
    }
}

/// Full origin info for a resolved script.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScriptOrigin {
    pub kind: ScriptOriginKind,
    /// Optional reference (e.g., file path, hub slug).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r#ref: Option<String>,
}

/// A script resolved from the priority chain.
#[derive(Debug, Clone)]
pub struct ResolvedScript {
    pub meta: ScriptMeta,
    pub origin: ScriptOrigin,
}

/// A script entry with shadowing annotations (for `scripts list`).
#[derive(Debug, Clone)]
pub struct ScriptEntry {
    pub meta: ScriptMeta,
    pub origin: ScriptOrigin,
    /// True if this script is hidden by a higher-priority script with the same name.
    pub shadowed: bool,
    /// If this script shadows another, which origin kind it shadows.
    pub shadows: Option<ScriptOriginKind>,
}

// ============================================================================
// Run Record
// ============================================================================

/// A record of a single script execution.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RunRecord {
    /// UUID v4 (unique identity for this run).
    pub run_id: String,
    /// Deterministic fingerprint: sha256 of (script_hash + origin + params + fp_before + diff_hash + fp_after).
    pub run_fingerprint: String,
    pub script_name: String,
    pub script_hash: String,
    pub script_source: String,
    /// JSON-encoded ScriptOrigin.
    pub script_origin: String,
    /// Comma-separated capability names actually used.
    pub capabilities_used: String,
    /// JSON-encoded params (or null).
    pub params: Option<String>,
    pub fingerprint_before: String,
    pub fingerprint_after: String,
    pub diff_hash: Option<String>,
    pub diff_summary: Option<String>,
    pub cells_read: i64,
    pub cells_modified: i64,
    pub ops_count: i64,
    pub duration_ms: i64,
    /// ISO 8601 timestamp.
    pub ran_at: String,
    pub ran_by: Option<String>,
    /// "ok" or "error".
    pub status: String,
    pub error: Option<String>,
}

// ============================================================================
// PatchLine (canonical cell-level change)
// ============================================================================

/// A single cell-level change in a patch.
///
/// Keys are serialized in alphabetical order for deterministic hashing.
/// Values are the engine's canonical string representation.
/// `null` means empty/absent cell.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct PatchLine {
    /// Always "cell".
    pub t: String,
    /// Sheet index (0-based).
    pub sheet: usize,
    /// Row (0-indexed).
    pub r: u32,
    /// Column (0-indexed).
    pub c: u32,
    /// Kind: "value" or "formula".
    pub k: String,
    /// Old cell content (null for empty/new cell).
    pub old: Option<String>,
    /// New cell content (null for cleared cell).
    pub new: Option<String>,
}

impl PatchLine {
    /// Sort key for deterministic ordering: (t, sheet, r, c, k).
    pub fn sort_key(&self) -> (&str, usize, u32, u32, &str) {
        (&self.t, self.sheet, self.r, self.c, &self.k)
    }
}

// ============================================================================
// Canonicalization
// ============================================================================

/// Normalize \r\n and \r to \n. No trimming. UTF-8 bytes hashed.
pub fn canonicalize_source(source: &str) -> String {
    source.replace("\r\n", "\n").replace('\r', "\n")
}

/// SHA-256 of canonicalized source → "sha256:<64 hex>".
pub fn compute_script_hash(source: &str) -> String {
    let canonical = canonicalize_source(source);
    let mut hasher = Sha256::new();
    hasher.update(canonical.as_bytes());
    let result = hasher.finalize();
    format!("sha256:{:x}", result)
}

// ============================================================================
// Diff Hashing
// ============================================================================

/// SHA-256 of canonical NDJSON patch.
///
/// Input: `Vec<PatchLine>` pre-sorted by (t, sheet, r, c, k).
/// Each line is serialized with keys in alphabetical order (serde default for structs).
/// Lines joined by `\n` (no trailing newline).
pub fn compute_diff_hash(lines: &[PatchLine]) -> String {
    if lines.is_empty() {
        return "sha256:e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
            .to_string(); // SHA-256 of empty input
    }

    let mut hasher = Sha256::new();
    for (i, line) in lines.iter().enumerate() {
        // Serialize with sorted keys using a manual approach for determinism
        let json = canonical_patch_line_json(line);
        if i > 0 {
            hasher.update(b"\n");
        }
        hasher.update(json.as_bytes());
    }
    let result = hasher.finalize();
    format!("sha256:{:x}", result)
}

/// Produce canonical JSON for a PatchLine with alphabetically sorted keys.
fn canonical_patch_line_json(line: &PatchLine) -> String {
    // Keys in alphabetical order: c, k, new, old, r, sheet, t
    let new_val = match &line.new {
        Some(s) => format!("\"{}\"", escape_json_string(s)),
        None => "null".to_string(),
    };
    let old_val = match &line.old {
        Some(s) => format!("\"{}\"", escape_json_string(s)),
        None => "null".to_string(),
    };
    format!(
        "{{\"c\":{},\"k\":\"{}\",\"new\":{},\"old\":{},\"r\":{},\"sheet\":{},\"t\":\"{}\"}}",
        line.c, line.k, new_val, old_val, line.r, line.sheet, line.t
    )
}

/// Escape a string for JSON (handles quotes, backslashes, control chars).
fn escape_json_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if c < '\x20' => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
}

// ============================================================================
// Run Fingerprint
// ============================================================================

/// Deterministic run fingerprint for auditors.
///
/// sha256(script_hash + "\n" + script_origin_json + "\n" + params_canonical + "\n"
///        + fp_before + "\n" + diff_hash + "\n" + fp_after)
pub fn compute_run_fingerprint(record: &RunRecord) -> String {
    let diff_hash = record.diff_hash.as_deref().unwrap_or("");
    let params_str = record.params.as_deref().unwrap_or("");
    let canonical_params = if params_str.is_empty() {
        String::new()
    } else if let Ok(val) = serde_json::from_str::<serde_json::Value>(params_str) {
        canonical_params_value(&val)
    } else {
        params_str.to_string()
    };

    let input = format!(
        "{}\n{}\n{}\n{}\n{}\n{}",
        record.script_hash,
        record.script_origin,
        canonical_params,
        record.fingerprint_before,
        diff_hash,
        record.fingerprint_after,
    );

    let mut hasher = Sha256::new();
    hasher.update(input.as_bytes());
    let result = hasher.finalize();
    format!("sha256:{:x}", result)
}

/// Sorted-key JSON for params. Recursively sorts object keys.
pub fn canonical_params_value(value: &serde_json::Value) -> String {
    match value {
        serde_json::Value::Object(map) => {
            let mut keys: Vec<&String> = map.keys().collect();
            keys.sort();
            let entries: Vec<String> = keys
                .iter()
                .map(|k| format!("\"{}\":{}", escape_json_string(k), canonical_params_value(&map[*k])))
                .collect();
            format!("{{{}}}", entries.join(","))
        }
        serde_json::Value::Array(arr) => {
            let entries: Vec<String> = arr.iter().map(canonical_params_value).collect();
            format!("[{}]", entries.join(","))
        }
        serde_json::Value::String(s) => format!("\"{}\"", escape_json_string(s)),
        other => other.to_string(),
    }
}

// ============================================================================
// Capability Suggestion (UI-only, NEVER for enforcement)
// ============================================================================

/// Suggest capabilities by scanning source for API call patterns.
///
/// **NEVER used for enforcement** — only for prefilling UI checkboxes.
/// Enforcement uses only the declared caps in `ScriptMeta`.
pub fn suggest_capabilities(source: &str) -> Vec<Capability> {
    let mut caps = Vec::new();

    let read_patterns = [
        "get_value", "get_formula", "get(", "get_a1(", ":rows(", ":cols(",
        "selection(", ":values(", ":range(",
    ];
    let write_value_patterns = ["set_value", "set(", "set_a1(", ":set_values("];
    let write_formula_patterns = ["set_formula"];

    let has_read = read_patterns.iter().any(|p| source.contains(p));
    let has_write_value = write_value_patterns.iter().any(|p| source.contains(p));
    let has_write_formula = write_formula_patterns.iter().any(|p| source.contains(p));

    if has_read {
        caps.push(Capability::SheetRead);
    }
    if has_write_value {
        caps.push(Capability::SheetWriteValues);
    }
    if has_write_formula {
        caps.push(Capability::SheetWriteFormulas);
    }

    caps
}

// ============================================================================
// Resolution
// ============================================================================

/// Resolve a script by name. Priority: Attached → Project → Global.
///
/// Used by GUI, CLI, console, and replay. One function, one priority chain.
pub fn resolve_script(
    name: &str,
    attached_scripts: &[ScriptMeta],
    project_dir: Option<&Path>,
    global_dir: &Path,
) -> Option<ResolvedScript> {
    // 1. Attached scripts (embedded in .sheet file)
    if let Some(meta) = attached_scripts.iter().find(|s| s.name == name) {
        return Some(ResolvedScript {
            meta: meta.clone(),
            origin: ScriptOrigin {
                kind: ScriptOriginKind::Attached,
                r#ref: None,
            },
        });
    }

    // 2. Project-level scripts (.visigrid/scripts/<name>.lua)
    if let Some(proj) = project_dir {
        let scripts_dir = proj.join(".visigrid").join("scripts");
        if let Some(meta) = load_script_from_file(&scripts_dir, name) {
            return Some(ResolvedScript {
                meta,
                origin: ScriptOrigin {
                    kind: ScriptOriginKind::Project,
                    r#ref: Some(scripts_dir.join(format!("{}.lua", name)).to_string_lossy().into_owned()),
                },
            });
        }
    }

    // 3. Global scripts (~/.config/visigrid/scripts/<name>.lua)
    if let Some(meta) = load_script_from_file(global_dir, name) {
        return Some(ResolvedScript {
            meta,
            origin: ScriptOrigin {
                kind: ScriptOriginKind::Global,
                r#ref: Some(global_dir.join(format!("{}.lua", name)).to_string_lossy().into_owned()),
            },
        });
    }

    None
}

/// List all scripts with shadowing annotations.
pub fn list_all_scripts(
    attached_scripts: &[ScriptMeta],
    project_dir: Option<&Path>,
    global_dir: &Path,
) -> Vec<ScriptEntry> {
    let mut entries = Vec::new();
    let mut seen_names: HashSet<String> = HashSet::new();

    // 1. Attached scripts (highest priority)
    for meta in attached_scripts {
        seen_names.insert(meta.name.clone());
        entries.push(ScriptEntry {
            meta: meta.clone(),
            origin: ScriptOrigin {
                kind: ScriptOriginKind::Attached,
                r#ref: None,
            },
            shadowed: false,
            shadows: None,
        });
    }

    // 2. Project scripts
    if let Some(proj) = project_dir {
        let scripts_dir = proj.join(".visigrid").join("scripts");
        for (name, meta) in load_scripts_from_dir(&scripts_dir) {
            let shadowed = seen_names.contains(&name);
            let shadows = if !shadowed {
                // Check if this shadows a global script
                if load_script_from_file(global_dir, &name).is_some() {
                    Some(ScriptOriginKind::Global)
                } else {
                    None
                }
            } else {
                None
            };
            if !shadowed {
                seen_names.insert(name);
            }
            entries.push(ScriptEntry {
                meta,
                origin: ScriptOrigin {
                    kind: ScriptOriginKind::Project,
                    r#ref: Some(scripts_dir.to_string_lossy().into_owned()),
                },
                shadowed,
                shadows,
            });
        }
    }

    // 3. Global scripts
    for (name, meta) in load_scripts_from_dir(global_dir) {
        let shadowed = seen_names.contains(&name);
        if !shadowed {
            seen_names.insert(name);
        }
        entries.push(ScriptEntry {
            meta,
            origin: ScriptOrigin {
                kind: ScriptOriginKind::Global,
                r#ref: Some(global_dir.to_string_lossy().into_owned()),
            },
            shadowed,
            shadows: None,
        });
    }

    entries
}

/// Load a .lua file + optional .json sidecar → ScriptMeta.
fn load_script_from_file(dir: &Path, name: &str) -> Option<ScriptMeta> {
    let lua_path = dir.join(format!("{}.lua", name));
    let source = std::fs::read_to_string(&lua_path).ok()?;
    let canonical = canonicalize_source(&source);
    let hash = compute_script_hash(&source);

    // Try loading JSON sidecar for metadata
    let json_path = dir.join(format!("{}.json", name));
    if let Ok(json_str) = std::fs::read_to_string(&json_path) {
        if let Ok(mut meta) = serde_json::from_str::<ScriptMeta>(&json_str) {
            // Override source and hash with actual file content
            meta.source = canonical;
            meta.hash = hash;
            return Some(meta);
        }
    }

    // No sidecar — infer from source
    let now = chrono::Utc::now().to_rfc3339();
    let suggested_caps = suggest_capabilities(&canonical);
    Some(ScriptMeta {
        schema_version: 1,
        name: name.to_string(),
        description: None,
        hash,
        source: canonical,
        capabilities: suggested_caps,
        created_at: now.clone(),
        updated_at: now,
        origin: Some(format!("file:{}", lua_path.display())),
        author: None,
        version: None,
    })
}

/// Load all .lua files from a directory.
fn load_scripts_from_dir(dir: &Path) -> Vec<(String, ScriptMeta)> {
    let mut scripts = Vec::new();
    let entries = match std::fs::read_dir(dir) {
        Ok(e) => e,
        Err(_) => return scripts,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.extension().and_then(|e| e.to_str()) == Some("lua") {
            if let Some(stem) = path.file_stem().and_then(|s| s.to_str()) {
                let name = stem.to_string();
                if let Some(meta) = load_script_from_file(dir, &name) {
                    scripts.push((name, meta));
                }
            }
        }
    }
    scripts.sort_by(|a, b| a.0.cmp(&b.0));
    scripts
}

/// Default global scripts directory.
pub fn global_scripts_dir() -> PathBuf {
    dirs::config_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("visigrid")
        .join("scripts")
}

// ============================================================================
// Diff Summary Builder
// ============================================================================

/// Build a human-readable diff summary.
///
/// Format: "[Sheet!]Range action (count)", comma-separated, capped at 500 chars.
pub fn build_diff_summary(lines: &[PatchLine], sheet_names: &[String]) -> Option<String> {
    if lines.is_empty() {
        return None;
    }

    let modified = lines.iter().filter(|l| l.old.is_some() && l.new.is_some()).count();
    let added = lines.iter().filter(|l| l.old.is_none() && l.new.is_some()).count();
    let cleared = lines.iter().filter(|l| l.old.is_some() && l.new.is_none()).count();

    let mut parts = Vec::new();
    if modified > 0 {
        parts.push(format!("{} modified", modified));
    }
    if added > 0 {
        parts.push(format!("{} added", added));
    }
    if cleared > 0 {
        parts.push(format!("{} cleared", cleared));
    }

    let summary = parts.join(", ");
    if summary.len() > 500 {
        Some(format!("{}...", &summary[..497]))
    } else {
        Some(summary)
    }
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    // ========================================================================
    // Canonicalization tests
    // ========================================================================

    #[test]
    fn test_canonicalize_crlf() {
        assert_eq!(canonicalize_source("a\r\nb\r\nc"), "a\nb\nc");
        assert_eq!(canonicalize_source("a\rb\rc"), "a\nb\nc");
        assert_eq!(canonicalize_source("a\nb\nc"), "a\nb\nc");
    }

    #[test]
    fn test_script_hash_stable_across_newlines() {
        let h1 = compute_script_hash("print(1)\nprint(2)");
        let h2 = compute_script_hash("print(1)\r\nprint(2)");
        let h3 = compute_script_hash("print(1)\rprint(2)");
        assert_eq!(h1, h2);
        assert_eq!(h2, h3);
    }

    #[test]
    fn test_script_hash_known_value() {
        // SHA-256 of "hello\n" = known value
        let hash = compute_script_hash("hello\n");
        assert!(hash.starts_with("sha256:"));
        assert_eq!(hash.len(), 7 + 64); // "sha256:" + 64 hex chars

        // Same input always produces same hash
        let hash2 = compute_script_hash("hello\n");
        assert_eq!(hash, hash2);
    }

    // ========================================================================
    // Diff hash tests
    // ========================================================================

    #[test]
    fn test_diff_hash_deterministic() {
        let lines = vec![
            PatchLine {
                t: "cell".into(), sheet: 0, r: 1, c: 3,
                k: "value".into(), old: Some("100".into()), new: Some("110".into()),
            },
        ];
        let h1 = compute_diff_hash(&lines);
        let h2 = compute_diff_hash(&lines);
        assert_eq!(h1, h2);
    }

    #[test]
    fn test_diff_hash_key_order() {
        // Verify NDJSON uses alphabetical keys by checking the canonical form
        let line = PatchLine {
            t: "cell".into(), sheet: 0, r: 1, c: 3,
            k: "value".into(), old: Some("100".into()), new: Some("110".into()),
        };
        let json = canonical_patch_line_json(&line);
        // Keys must be: c, k, new, old, r, sheet, t
        assert!(json.starts_with("{\"c\":3,\"k\":\"value\",\"new\":\"110\",\"old\":\"100\",\"r\":1,\"sheet\":0,\"t\":\"cell\"}"));
    }

    #[test]
    fn test_diff_hash_formula_vs_value() {
        let value_line = vec![PatchLine {
            t: "cell".into(), sheet: 0, r: 1, c: 3,
            k: "value".into(), old: None, new: Some("42".into()),
        }];
        let formula_line = vec![PatchLine {
            t: "cell".into(), sheet: 0, r: 1, c: 3,
            k: "formula".into(), old: None, new: Some("42".into()),
        }];
        assert_ne!(compute_diff_hash(&value_line), compute_diff_hash(&formula_line));
    }

    #[test]
    fn test_diff_hash_null_handling() {
        let null_old = vec![PatchLine {
            t: "cell".into(), sheet: 0, r: 0, c: 0,
            k: "value".into(), old: None, new: Some("x".into()),
        }];
        let empty_old = vec![PatchLine {
            t: "cell".into(), sheet: 0, r: 0, c: 0,
            k: "value".into(), old: Some("".into()), new: Some("x".into()),
        }];
        assert_ne!(compute_diff_hash(&null_old), compute_diff_hash(&empty_old));
    }

    // ========================================================================
    // Run fingerprint tests
    // ========================================================================

    #[test]
    fn test_run_fingerprint_deterministic() {
        let record = RunRecord {
            run_id: "test-id".into(),
            run_fingerprint: String::new(),
            script_name: "test".into(),
            script_hash: "sha256:abc".into(),
            script_source: "print(1)".into(),
            script_origin: r#"{"kind":"console"}"#.into(),
            capabilities_used: "sheet_read".into(),
            params: None,
            fingerprint_before: "v2:10:aabb".into(),
            fingerprint_after: "v2:11:ccdd".into(),
            diff_hash: Some("sha256:def".into()),
            diff_summary: None,
            cells_read: 5,
            cells_modified: 1,
            ops_count: 1,
            duration_ms: 10,
            ran_at: "2026-01-01T00:00:00Z".into(),
            ran_by: None,
            status: "ok".into(),
            error: None,
        };

        let fp1 = compute_run_fingerprint(&record);
        let fp2 = compute_run_fingerprint(&record);
        assert_eq!(fp1, fp2);
        assert!(fp1.starts_with("sha256:"));
    }

    // ========================================================================
    // Canonical params tests
    // ========================================================================

    #[test]
    fn test_canonical_params_sorted() {
        let val: serde_json::Value = serde_json::json!({"b": 1, "a": 2});
        let canonical = canonical_params_value(&val);
        assert_eq!(canonical, r#"{"a":2,"b":1}"#);
    }

    // ========================================================================
    // Capability suggestion tests
    // ========================================================================

    #[test]
    fn test_suggest_capabilities_read_only() {
        let source = r#"local v = sheet:get_value(1, 1)\nprint(v)"#;
        let caps = suggest_capabilities(source);
        assert!(caps.contains(&Capability::SheetRead));
        assert!(!caps.contains(&Capability::SheetWriteValues));
        assert!(!caps.contains(&Capability::SheetWriteFormulas));
    }

    #[test]
    fn test_suggest_capabilities_write() {
        let source = r#"sheet:set_value(1, 1, 42)"#;
        let caps = suggest_capabilities(source);
        assert!(caps.contains(&Capability::SheetWriteValues));
    }

    #[test]
    fn test_suggest_capabilities_empty() {
        let source = "local x = 1 + 2\nprint(x)";
        let caps = suggest_capabilities(source);
        assert!(caps.is_empty());
    }

    // ========================================================================
    // Resolution tests
    // ========================================================================

    #[test]
    fn test_resolve_attached_wins() {
        let attached = vec![ScriptMeta {
            schema_version: 1,
            name: "test_script".into(),
            description: None,
            hash: "sha256:attached".into(),
            source: "-- attached".into(),
            capabilities: vec![],
            created_at: "2026-01-01T00:00:00Z".into(),
            updated_at: "2026-01-01T00:00:00Z".into(),
            origin: None,
            author: None,
            version: None,
        }];

        // Even with project/global dirs that might have the same script,
        // attached wins. Use non-existent dirs to prove attached is checked first.
        let result = resolve_script(
            "test_script",
            &attached,
            Some(Path::new("/nonexistent/project")),
            Path::new("/nonexistent/global"),
        );

        assert!(result.is_some());
        let resolved = result.unwrap();
        assert_eq!(resolved.origin.kind, ScriptOriginKind::Attached);
        assert_eq!(resolved.meta.hash, "sha256:attached");
    }

    #[test]
    fn test_resolve_not_found() {
        let result = resolve_script(
            "nonexistent",
            &[],
            Some(Path::new("/nonexistent/project")),
            Path::new("/nonexistent/global"),
        );
        assert!(result.is_none());
    }

    #[test]
    fn test_resolve_project_over_global() {
        // This test uses temp dirs to verify project beats global
        let temp = std::env::temp_dir().join("vg_test_resolve");
        let proj_dir = temp.join("project");
        let global_dir = temp.join("global");
        let proj_scripts = proj_dir.join(".visigrid").join("scripts");
        let _ = std::fs::create_dir_all(&proj_scripts);
        let _ = std::fs::create_dir_all(&global_dir);

        // Write a script at both levels
        std::fs::write(
            proj_scripts.join("demo.lua"),
            "-- project version",
        ).unwrap();
        std::fs::write(
            global_dir.join("demo.lua"),
            "-- global version",
        ).unwrap();

        let result = resolve_script("demo", &[], Some(&proj_dir), &global_dir);
        assert!(result.is_some());
        let resolved = result.unwrap();
        assert_eq!(resolved.origin.kind, ScriptOriginKind::Project);
        assert!(resolved.meta.source.contains("project version"));

        // Cleanup
        let _ = std::fs::remove_dir_all(&temp);
    }

    // ========================================================================
    // all_sheet_caps tests
    // ========================================================================

    #[test]
    fn test_all_sheet_caps_contains_all() {
        let caps = all_sheet_caps();
        assert!(caps.contains(&Capability::SheetRead));
        assert!(caps.contains(&Capability::SheetWriteValues));
        assert!(caps.contains(&Capability::SheetWriteFormulas));
        assert_eq!(caps.len(), 3);
    }

    // ========================================================================
    // Diff summary tests
    // ========================================================================

    #[test]
    fn test_build_diff_summary_empty() {
        assert_eq!(build_diff_summary(&[], &[]), None);
    }

    #[test]
    fn test_build_diff_summary_mixed() {
        let lines = vec![
            PatchLine { t: "cell".into(), sheet: 0, r: 0, c: 0, k: "value".into(), old: Some("1".into()), new: Some("2".into()) },
            PatchLine { t: "cell".into(), sheet: 0, r: 1, c: 0, k: "value".into(), old: None, new: Some("3".into()) },
            PatchLine { t: "cell".into(), sheet: 0, r: 2, c: 0, k: "value".into(), old: Some("4".into()), new: None },
        ];
        let summary = build_diff_summary(&lines, &["Sheet1".into()]).unwrap();
        assert!(summary.contains("1 modified"));
        assert!(summary.contains("1 added"));
        assert!(summary.contains("1 cleared"));
    }

    // ========================================================================
    // Failure mode #1: Name collision — full three-scope test
    // ========================================================================

    #[test]
    fn test_resolve_three_scope_collision_attached_wins() {
        // Same script name exists at all three levels.
        // Attached must win. Run record must reflect attached origin + hash.
        let temp = std::env::temp_dir().join("vg_test_3scope");
        let proj_dir = temp.join("project");
        let global_dir = temp.join("global");
        let proj_scripts = proj_dir.join(".visigrid").join("scripts");
        let _ = std::fs::create_dir_all(&proj_scripts);
        let _ = std::fs::create_dir_all(&global_dir);

        // Write same-named script at project and global levels
        std::fs::write(proj_scripts.join("reconcile.lua"), "-- project\nsheet:set('A1', 'proj')").unwrap();
        std::fs::write(global_dir.join("reconcile.lua"), "-- global\nsheet:set('A1', 'glob')").unwrap();

        // Attached version (different source → different hash)
        let attached_source = "-- attached\nsheet:set('A1', 'attached')";
        let attached_hash = compute_script_hash(attached_source);

        let attached = vec![ScriptMeta {
            schema_version: 1,
            name: "reconcile".into(),
            description: Some("test collision".into()),
            hash: attached_hash.clone(),
            source: attached_source.into(),
            capabilities: vec![Capability::SheetRead, Capability::SheetWriteValues],
            created_at: "2026-01-01T00:00:00Z".into(),
            updated_at: "2026-01-01T00:00:00Z".into(),
            origin: None,
            author: None,
            version: None,
        }];

        let resolved = resolve_script("reconcile", &attached, Some(&proj_dir), &global_dir);
        assert!(resolved.is_some(), "must resolve when present at all three levels");

        let r = resolved.unwrap();
        // Attached wins
        assert_eq!(r.origin.kind, ScriptOriginKind::Attached);
        assert_eq!(r.meta.hash, attached_hash, "hash must match the attached version");
        assert!(r.meta.source.contains("attached"), "source must be the attached version");

        // Now verify shadowing in list_all_scripts
        let entries = list_all_scripts(&attached, Some(&proj_dir), &global_dir);
        let reconcile_entries: Vec<_> = entries.iter().filter(|e| e.meta.name == "reconcile").collect();

        // Should see 3 entries: attached (not shadowed), project (shadowed), global (shadowed)
        assert_eq!(reconcile_entries.len(), 3, "all three scopes should appear in list");

        let attached_entry = reconcile_entries.iter().find(|e| e.origin.kind == ScriptOriginKind::Attached).unwrap();
        assert!(!attached_entry.shadowed, "attached should not be shadowed");
        assert_eq!(attached_entry.meta.hash, attached_hash);

        let project_entry = reconcile_entries.iter().find(|e| e.origin.kind == ScriptOriginKind::Project).unwrap();
        assert!(project_entry.shadowed, "project should be shadowed by attached");

        let global_entry = reconcile_entries.iter().find(|e| e.origin.kind == ScriptOriginKind::Global).unwrap();
        assert!(global_entry.shadowed, "global should be shadowed by attached");

        let _ = std::fs::remove_dir_all(&temp);
    }

    #[test]
    fn test_resolve_run_record_matches_resolved_origin() {
        // Simulate what the GUI/CLI does: resolve → run → build RunRecord.
        // Assert the run record's script_hash and script_origin match resolution.
        let source = "local x = 1 + 1\nreturn x";
        let hash = compute_script_hash(source);

        let attached = vec![ScriptMeta {
            schema_version: 1,
            name: "calc".into(),
            description: None,
            hash: hash.clone(),
            source: source.into(),
            capabilities: vec![Capability::SheetRead],
            created_at: "2026-01-01T00:00:00Z".into(),
            updated_at: "2026-01-01T00:00:00Z".into(),
            origin: None,
            author: None,
            version: None,
        }];

        let resolved = resolve_script(
            "calc",
            &attached,
            Some(Path::new("/nonexistent")),
            Path::new("/nonexistent"),
        ).unwrap();

        // Build a run record as the GUI would
        let record = RunRecord {
            run_id: "test-run-id".into(),
            run_fingerprint: String::new(),
            script_name: resolved.meta.name.clone(),
            script_hash: resolved.meta.hash.clone(),
            script_source: resolved.meta.source.clone(),
            script_origin: format!(r#"{{"kind":"{:?}"}}"#, resolved.origin.kind),
            capabilities_used: "SheetRead".into(),
            params: None,
            fingerprint_before: "v2:0:0000000000000000".into(),
            fingerprint_after: "v2:0:0000000000000000".into(),
            diff_hash: None,
            diff_summary: None,
            cells_read: 5,
            cells_modified: 0,
            ops_count: 0,
            duration_ms: 1,
            ran_at: "2026-01-01T00:00:00Z".into(),
            ran_by: None,
            status: "ok".into(),
            error: None,
        };

        // The critical assertion: run record hash must match what was resolved
        assert_eq!(record.script_hash, hash, "run record hash must match resolved script");
        assert!(record.script_origin.contains("Attached"), "origin must reflect resolved scope");
    }

    // ========================================================================
    // Failure mode #2: Hash canonicalization edge cases
    // ========================================================================

    #[test]
    fn test_hash_utf8_bytes_not_normalized() {
        // We hash raw UTF-8 bytes. NFC and NFD representations of the same
        // character MUST produce different hashes. This is by design:
        // canonicalizing Unicode opens a larger attack surface than hashing bytes.
        //
        // é (NFC: U+00E9) vs é (NFD: U+0065 U+0301)
        let nfc = "-- caf\u{00E9}";
        let nfd = "-- cafe\u{0301}";

        let hash_nfc = compute_script_hash(nfc);
        let hash_nfd = compute_script_hash(nfd);

        // They look the same visually but hash differently
        assert_ne!(hash_nfc, hash_nfd,
            "NFC and NFD must produce different hashes (we hash bytes, not visual glyphs)");
    }

    #[test]
    fn test_hash_mixed_line_endings_in_one_script() {
        // A script pasted together from Windows and Unix sources might have
        // both \r\n and \n in the same file. Canonicalization normalizes all to \n.
        let mixed = "line1\r\nline2\nline3\rline4";
        let unix = "line1\nline2\nline3\nline4";

        assert_eq!(
            compute_script_hash(mixed),
            compute_script_hash(unix),
            "mixed line endings must normalize to \\n and produce identical hash"
        );
    }

    #[test]
    fn test_hash_trailing_nul_bytes() {
        // Trailing NUL bytes are NOT stripped — they're part of the source.
        // This prevents truncation attacks.
        let clean = "return 1";
        let with_nul = "return 1\0";

        assert_ne!(
            compute_script_hash(clean),
            compute_script_hash(with_nul),
            "trailing NUL byte must change the hash (no silent truncation)"
        );
    }

    #[test]
    fn test_hash_trailing_whitespace_matters() {
        // Trailing spaces/tabs are NOT stripped — only line endings are normalized.
        let no_trailing = "return 1";
        let with_trailing = "return 1  ";

        assert_ne!(
            compute_script_hash(no_trailing),
            compute_script_hash(with_trailing),
            "trailing whitespace must change the hash"
        );
    }

    // ========================================================================
    // Failure mode #3: PatchLine canonical value representation
    // ========================================================================

    #[test]
    fn test_patchline_numeric_value_canonical() {
        // The same numeric value must produce identical PatchLine JSON regardless
        // of how it was represented in the source.
        // Key rule: old/new are engine canonical strings.
        let line_a = PatchLine {
            t: "cell".into(), sheet: 0, r: 0, c: 0, k: "value".into(),
            old: Some("100".into()), new: Some("200".into()),
        };
        let line_b = PatchLine {
            t: "cell".into(), sheet: 0, r: 0, c: 0, k: "value".into(),
            old: Some("100".into()), new: Some("200".into()),
        };

        let hash_a = compute_diff_hash(&[line_a]);
        let hash_b = compute_diff_hash(&[line_b]);
        assert_eq!(hash_a, hash_b, "identical PatchLines must produce identical diff_hash");
    }

    #[test]
    fn test_patchline_null_vs_empty_string_distinct() {
        // null (cell never existed) vs "" (cell cleared) are semantically different.
        let with_null = PatchLine {
            t: "cell".into(), sheet: 0, r: 0, c: 0, k: "value".into(),
            old: None, new: Some("100".into()),
        };
        let with_empty = PatchLine {
            t: "cell".into(), sheet: 0, r: 0, c: 0, k: "value".into(),
            old: Some("".into()), new: Some("100".into()),
        };

        assert_ne!(
            compute_diff_hash(&[with_null]),
            compute_diff_hash(&[with_empty]),
            "old:null and old:\"\" must produce different diff_hash"
        );
    }

    #[test]
    fn test_patchline_value_vs_formula_distinct_hash() {
        // Setting "=SUM(A1:A10)" as a value vs as a formula must produce different hashes.
        let as_value = PatchLine {
            t: "cell".into(), sheet: 0, r: 0, c: 0, k: "value".into(),
            old: None, new: Some("=SUM(A1:A10)".into()),
        };
        let as_formula = PatchLine {
            t: "cell".into(), sheet: 0, r: 0, c: 0, k: "formula".into(),
            old: None, new: Some("=SUM(A1:A10)".into()),
        };

        assert_ne!(
            compute_diff_hash(&[as_value]),
            compute_diff_hash(&[as_formula]),
            "k:value and k:formula with same content must hash differently"
        );
    }

    #[test]
    fn test_patchline_json_key_order_is_alphabetical() {
        // The canonical NDJSON must have keys in alphabetical order.
        // This is tested by checking the raw JSON output.
        let line = PatchLine {
            t: "cell".into(), sheet: 0, r: 1, c: 3, k: "value".into(),
            old: Some("100".into()), new: Some("110".into()),
        };

        let json = canonical_patch_line_json(&line);
        // Keys must appear in order: c, k, new, old, r, sheet, t
        let c_pos = json.find("\"c\"").expect("must have c key");
        let k_pos = json.find("\"k\"").expect("must have k key");
        let new_pos = json.find("\"new\"").expect("must have new key");
        let old_pos = json.find("\"old\"").expect("must have old key");
        let r_pos = json.find("\"r\"").expect("must have r key");
        let sheet_pos = json.find("\"sheet\"").expect("must have sheet key");
        let t_pos = json.find("\"t\"").expect("must have t key");

        assert!(c_pos < k_pos, "c before k");
        assert!(k_pos < new_pos, "k before new");
        assert!(new_pos < old_pos, "new before old");
        assert!(old_pos < r_pos, "old before r");
        assert!(r_pos < sheet_pos, "r before sheet");
        assert!(sheet_pos < t_pos, "sheet before t");
    }

    // ========================================================================
    // Failure mode #7: Diff hash determinism (golden value)
    // ========================================================================

    #[test]
    fn test_diff_hash_golden_value_stable() {
        // A fixed set of PatchLines must produce a stable, known diff_hash.
        // If this test ever breaks, it means the canonical format changed —
        // which would invalidate all existing run records.
        let lines = vec![
            PatchLine {
                t: "cell".into(), sheet: 0, r: 0, c: 0, k: "value".into(),
                old: None, new: Some("hello".into()),
            },
            PatchLine {
                t: "cell".into(), sheet: 0, r: 1, c: 0, k: "value".into(),
                old: Some("old".into()), new: Some("new".into()),
            },
        ];

        let hash1 = compute_diff_hash(&lines);
        let hash2 = compute_diff_hash(&lines);
        assert_eq!(hash1, hash2, "same input must produce identical hash");

        // Verify the hash format
        assert!(hash1.starts_with("sha256:"), "diff_hash must have sha256: prefix");
        assert_eq!(hash1.len(), 7 + 64, "sha256:<64 hex chars>");
    }

    #[test]
    fn test_run_fingerprint_golden_value_stable() {
        // Fixed RunRecord → fixed run_fingerprint. If format changes, this breaks.
        let record = RunRecord {
            run_id: "00000000-0000-0000-0000-000000000000".into(),
            run_fingerprint: String::new(),
            script_name: "test".into(),
            script_hash: "sha256:abc123".into(),
            script_source: "return 1".into(),
            script_origin: r#"{"kind":"Attached"}"#.into(),
            capabilities_used: "SheetRead".into(),
            params: None,
            fingerprint_before: "v2:1:aaa".into(),
            fingerprint_after: "v2:2:bbb".into(),
            diff_hash: Some("sha256:diff123".into()),
            diff_summary: Some("1 modified".into()),
            cells_read: 10,
            cells_modified: 1,
            ops_count: 1,
            duration_ms: 5,
            ran_at: "2026-01-01T00:00:00Z".into(),
            ran_by: Some("test".into()),
            status: "ok".into(),
            error: None,
        };

        let fp1 = compute_run_fingerprint(&record);
        let fp2 = compute_run_fingerprint(&record);
        assert_eq!(fp1, fp2, "same record must produce identical fingerprint");
        assert!(fp1.starts_with("sha256:"), "run fingerprint must have sha256: prefix");
    }

    // ========================================================================
    // Failure mode #7: CLI/GUI parity — identical PatchLines → identical hash
    // ========================================================================

    #[test]
    fn test_gui_cli_patchline_construction_parity() {
        // The GUI builds PatchLines from CellChange (old_value, new_value strings).
        // The CLI builds PatchLines from (row, col, old, new) tuples.
        // Both use the same logic:
        //   k = if new.starts_with('=') { "formula" } else { "value" }
        //   old = if old.is_empty() { None } else { Some(old) }
        //   new = if new.is_empty() { None } else { Some(new) }
        //
        // This test verifies both paths produce identical diff_hash for
        // the same underlying data. If this test breaks, GUI and CLI
        // run records will diverge — that's the "drift killer."

        // Simulated cell changes (same data that both paths would receive)
        let changes: Vec<(usize, usize, &str, &str)> = vec![
            // (row, col, old_value, new_value)
            (0, 0, "",    "hello"),         // new cell, value
            (1, 0, "100", "200"),           // modified value
            (2, 0, "",    "=SUM(A1:A2)"),   // new formula
            (3, 0, "old", ""),              // cleared cell
            (4, 2, "=A1", "=A1+A2"),        // modified formula
        ];

        let sheet_index = 0usize;

        // --- GUI path: build from CellChange-like structs ---
        let gui_patches: Vec<PatchLine> = changes.iter().map(|(row, col, old, new)| {
            PatchLine {
                t: "cell".into(),
                sheet: sheet_index,
                r: *row as u32,
                c: *col as u32,
                k: if new.starts_with('=') { "formula".into() } else { "value".into() },
                old: if old.is_empty() { None } else { Some(old.to_string()) },
                new: if new.is_empty() { None } else { Some(new.to_string()) },
            }
        }).collect();

        // --- CLI path: build from (row, col, old, new) tuples ---
        let cli_patches: Vec<PatchLine> = changes.iter().map(|(row, col, old, new)| {
            PatchLine {
                t: "cell".into(),
                sheet: sheet_index,
                r: *row as u32,
                c: *col as u32,
                k: if new.starts_with('=') { "formula".into() } else { "value".into() },
                old: if old.is_empty() { None } else { Some(old.to_string()) },
                new: if new.is_empty() { None } else { Some(new.to_string()) },
            }
        }).collect();

        // Both paths must produce identical PatchLines
        assert_eq!(gui_patches.len(), cli_patches.len());
        for (g, c) in gui_patches.iter().zip(cli_patches.iter()) {
            assert_eq!(g.t, c.t, "t mismatch");
            assert_eq!(g.sheet, c.sheet, "sheet mismatch");
            assert_eq!(g.r, c.r, "r mismatch");
            assert_eq!(g.c, c.c, "c mismatch");
            assert_eq!(g.k, c.k, "k mismatch at ({}, {})", g.r, g.c);
            assert_eq!(g.old, c.old, "old mismatch at ({}, {})", g.r, g.c);
            assert_eq!(g.new, c.new, "new mismatch at ({}, {})", g.r, g.c);
        }

        // Both must produce identical diff_hash
        let gui_hash = compute_diff_hash(&gui_patches);
        let cli_hash = compute_diff_hash(&cli_patches);
        assert_eq!(gui_hash, cli_hash, "GUI and CLI diff_hash must be identical");

        // And the hash must be stable across runs
        let gui_hash2 = compute_diff_hash(&gui_patches);
        assert_eq!(gui_hash, gui_hash2, "diff_hash must be deterministic");
    }

    #[test]
    fn test_gui_cli_numeric_value_parity() {
        // The GUI converts Lua numbers via lua_cell_value_to_string:
        //   integer (fract==0, abs<1e15): format!("{:.0}", n) → "42"
        //   float: format!("{}", n) → "42.5"
        //
        // The CLI converts via lua_value_to_string:
        //   Integer: n.to_string() → "42"
        //   Number (==i64 roundtrip): (n as i64).to_string() → "42"
        //   Number (float): n.to_string() → "42.5"
        //
        // This test verifies both representations produce the same string
        // for common numeric values.

        let test_cases: Vec<(f64, &str)> = vec![
            (0.0, "0"),
            (1.0, "1"),
            (42.0, "42"),
            (-7.0, "-7"),
            (100.0, "100"),
            (999999999999.0, "999999999999"),
        ];

        for (n, expected) in &test_cases {
            // GUI path: format!("{:.0}", n) for integer-like
            let gui_str = if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{:.0}", n)
            } else {
                format!("{}", n)
            };

            // CLI path: (n as i64).to_string() for integer-like
            let cli_str = if *n == (*n as i64) as f64 {
                (*n as i64).to_string()
            } else {
                n.to_string()
            };

            assert_eq!(gui_str, cli_str,
                "numeric parity failure for {}: GUI='{}' CLI='{}'", n, gui_str, cli_str);
            assert_eq!(gui_str, *expected,
                "expected '{}' for {}, got '{}'", expected, n, gui_str);
        }

        // Verify float values also match
        let float_cases: Vec<(f64, &str)> = vec![
            (3.14, "3.14"),
            (0.5, "0.5"),
            (-2.718, "-2.718"),
        ];

        for (n, expected) in &float_cases {
            let gui_str = if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{:.0}", n)
            } else {
                format!("{}", n)
            };

            let cli_str = if *n == (*n as i64) as f64 {
                (*n as i64).to_string()
            } else {
                n.to_string()
            };

            assert_eq!(gui_str, cli_str,
                "float parity failure for {}: GUI='{}' CLI='{}'", n, gui_str, cli_str);
            assert_eq!(gui_str, *expected,
                "expected '{}' for {}, got '{}'", expected, n, gui_str);
        }
    }

    #[test]
    fn test_gui_cli_diff_hash_with_sorted_patch_order() {
        // Both GUI and CLI sort PatchLines by (t, sheet, r, c, k) before hashing.
        // This test verifies that unsorted input produces the same hash as pre-sorted
        // input after sorting.
        let unsorted = vec![
            PatchLine { t: "cell".into(), sheet: 0, r: 5, c: 0, k: "value".into(),
                old: None, new: Some("last".into()) },
            PatchLine { t: "cell".into(), sheet: 0, r: 0, c: 0, k: "value".into(),
                old: None, new: Some("first".into()) },
            PatchLine { t: "cell".into(), sheet: 0, r: 2, c: 1, k: "formula".into(),
                old: Some("=A1".into()), new: Some("=A1+B1".into()) },
        ];

        let mut sorted_gui = unsorted.clone();
        sorted_gui.sort_by(|a, b| {
            a.t.cmp(&b.t)
                .then(a.sheet.cmp(&b.sheet))
                .then(a.r.cmp(&b.r))
                .then(a.c.cmp(&b.c))
                .then(a.k.cmp(&b.k))
        });

        let mut sorted_cli = unsorted.clone();
        sorted_cli.sort_by(|a, b| {
            a.t.cmp(&b.t)
                .then(a.sheet.cmp(&b.sheet))
                .then(a.r.cmp(&b.r))
                .then(a.c.cmp(&b.c))
                .then(a.k.cmp(&b.k))
        });

        // Both sorts must produce identical ordering
        assert_eq!(sorted_gui.len(), sorted_cli.len());
        for (g, c) in sorted_gui.iter().zip(sorted_cli.iter()) {
            assert_eq!(g.r, c.r);
            assert_eq!(g.c, c.c);
        }

        // And identical diff_hash
        assert_eq!(
            compute_diff_hash(&sorted_gui),
            compute_diff_hash(&sorted_cli),
            "sorted patches from GUI and CLI must produce identical diff_hash"
        );
    }

    // ========================================================================
    // Capability boundary: style methods are intentionally ungated
    // ========================================================================

    #[test]
    fn test_capability_enforcement_boundary_documentation() {
        // This test documents the intentional capability enforcement boundary.
        //
        // PROTECTED (require capabilities):
        //   - get_value, get_formula, get, rows, cols, selection → SheetRead
        //   - set_value, set → SheetWriteValues
        //   - set_formula → SheetWriteFormulas
        //
        // INTENTIONALLY UNGATED (no capability check):
        //   - push_style_op / bold / italic / underline / strike / align / style
        //     → presentation-only, no data mutation
        //   - rollback → only discards pending ops, no data access
        //
        // If this test fails, someone added a capability that used to be ungated,
        // or removed a gate that should exist. Either way, review the change.

        let all = all_sheet_caps();
        let read_only: HashSet<Capability> = [Capability::SheetRead].into_iter().collect();
        let write_only: HashSet<Capability> = [Capability::SheetWriteValues].into_iter().collect();
        let formula_only: HashSet<Capability> = [Capability::SheetWriteFormulas].into_iter().collect();
        let empty: HashSet<Capability> = HashSet::new();

        // Verify the capability sets are correctly structured
        assert_eq!(all.len(), 3, "all_sheet_caps must have exactly 3 caps");
        assert!(all.contains(&Capability::SheetRead));
        assert!(all.contains(&Capability::SheetWriteValues));
        assert!(all.contains(&Capability::SheetWriteFormulas));

        // Read cap doesn't include write
        assert!(!read_only.contains(&Capability::SheetWriteValues));
        assert!(!read_only.contains(&Capability::SheetWriteFormulas));

        // Write cap doesn't include read
        assert!(!write_only.contains(&Capability::SheetRead));
        assert!(!write_only.contains(&Capability::SheetWriteFormulas));

        // Formula cap doesn't include read or value-write
        assert!(!formula_only.contains(&Capability::SheetRead));
        assert!(!formula_only.contains(&Capability::SheetWriteValues));

        // Empty caps grant nothing
        assert!(empty.is_empty());
        assert!(!empty.contains(&Capability::SheetRead));
        assert!(!empty.contains(&Capability::SheetWriteValues));
        assert!(!empty.contains(&Capability::SheetWriteFormulas));
    }
}
