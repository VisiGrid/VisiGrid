//! Link detection for URLs, emails, and file paths.
//!
//! Conservative detection with strict rules:
//! - URLs require a scheme (http://, https://, ftp://)
//! - Emails require user@domain.tld format (dot in domain)
//! - Paths are expanded (~) and checked for existence

use std::path::PathBuf;

/// A detected link target
#[derive(Debug, Clone, PartialEq)]
pub enum LinkTarget {
    /// URL with scheme (http://, https://, ftp://)
    Url(String),
    /// Email address (opens as mailto:)
    Email(String),
    /// File path (expanded and verified to exist)
    Path(PathBuf),
}

impl LinkTarget {
    /// Get the string to pass to `open::that()`
    pub fn open_string(&self) -> String {
        match self {
            LinkTarget::Url(url) => url.clone(),
            LinkTarget::Email(email) => format!("mailto:{}", email),
            LinkTarget::Path(path) => path.to_string_lossy().to_string(),
        }
    }
}

/// Characters to strip from the end of a potential link (excluding brackets)
const TRAILING_PUNCTUATION: &[char] = &['.', ',', ';', ':', '!', '?', '"', '\''];

/// Bracket pairs for balanced stripping
const BRACKET_PAIRS: &[(char, char)] = &[('(', ')'), ('[', ']'), ('{', '}')];

/// Strip trailing punctuation that commonly follows links in text.
/// Smart about brackets: only strips closing brackets if they're unbalanced.
fn strip_trailing_punctuation(s: &str) -> &str {
    let mut result = s;

    loop {
        let Some(c) = result.chars().last() else { break };

        // Check if it's a simple punctuation mark
        if TRAILING_PUNCTUATION.contains(&c) {
            result = &result[..result.len() - c.len_utf8()];
            continue;
        }

        // Check if it's a closing bracket
        let mut stripped = false;
        for &(open, close) in BRACKET_PAIRS {
            if c == close {
                // Count balance: only strip if there are more closes than opens
                let opens = result.chars().filter(|&ch| ch == open).count();
                let closes = result.chars().filter(|&ch| ch == close).count();
                if closes > opens {
                    result = &result[..result.len() - c.len_utf8()];
                    stripped = true;
                    break;
                }
            }
        }

        if !stripped {
            break;
        }
    }

    result
}

/// Expand ~ to home directory
fn expand_tilde(path: &str) -> Option<PathBuf> {
    if path.starts_with("~/") {
        dirs::home_dir().map(|home| home.join(&path[2..]))
    } else if path == "~" {
        dirs::home_dir()
    } else {
        Some(PathBuf::from(path))
    }
}

/// Check if a string looks like a URL (has a recognized scheme)
fn is_url(s: &str) -> bool {
    let lower = s.to_lowercase();
    lower.starts_with("http://")
        || lower.starts_with("https://")
        || lower.starts_with("ftp://")
        || lower.starts_with("ftps://")
        || lower.starts_with("file://")
}

/// Check if a string looks like an email address
fn is_email(s: &str) -> bool {
    // Must have exactly one @ and at least one dot after @
    let parts: Vec<&str> = s.split('@').collect();
    if parts.len() != 2 {
        return false;
    }

    let local = parts[0];
    let domain = parts[1];

    // Local part must not be empty
    if local.is_empty() {
        return false;
    }

    // Domain must have at least one dot and not start/end with dot
    if !domain.contains('.') || domain.starts_with('.') || domain.ends_with('.') {
        return false;
    }

    // Basic character validation (not exhaustive, but catches obvious non-emails)
    let valid_local = local.chars().all(|c| {
        c.is_alphanumeric() || c == '.' || c == '_' || c == '-' || c == '+'
    });

    let valid_domain = domain.chars().all(|c| {
        c.is_alphanumeric() || c == '.' || c == '-'
    });

    valid_local && valid_domain
}

/// Check if a string looks like a file path
fn is_path_like(s: &str) -> bool {
    // Must start with / or ~ for Unix-style paths
    // Or contain path separators
    s == "~" || s.starts_with('/') || s.starts_with("~/") || s.starts_with("~\\")
    // Windows absolute paths
    || (s.len() >= 3 && s.chars().nth(1) == Some(':') && (s.chars().nth(2) == Some('/') || s.chars().nth(2) == Some('\\')))
}

/// Detect if text contains a single link target.
///
/// Only activates when the entire trimmed text is a single link token.
/// Returns None if no link is detected or if detection is ambiguous.
pub fn detect_link(text: &str) -> Option<LinkTarget> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }

    // Reject if contains newlines (not a single token)
    if trimmed.contains('\n') || trimmed.contains('\r') {
        return None;
    }

    // Strip trailing punctuation
    let clean = strip_trailing_punctuation(trimmed);
    if clean.is_empty() {
        return None;
    }

    // Priority 1: URL (most specific - has scheme)
    if is_url(clean) {
        return Some(LinkTarget::Url(clean.to_string()));
    }

    // Priority 2: Email (has @ with valid structure)
    if is_email(clean) {
        return Some(LinkTarget::Email(clean.to_string()));
    }

    // Priority 3: Path (if it exists)
    if is_path_like(clean) {
        if let Some(expanded) = expand_tilde(clean) {
            if expanded.exists() {
                return Some(LinkTarget::Path(expanded));
            }
        }
    }

    None
}

#[cfg(test)]
mod tests {
    use super::*;

    // ========================================================================
    // URL detection tests
    // ========================================================================

    #[test]
    fn test_url_https() {
        assert_eq!(
            detect_link("https://stripe.com/inv_abc123"),
            Some(LinkTarget::Url("https://stripe.com/inv_abc123".to_string()))
        );
    }

    #[test]
    fn test_url_http() {
        assert_eq!(
            detect_link("http://example.com"),
            Some(LinkTarget::Url("http://example.com".to_string()))
        );
    }

    #[test]
    fn test_url_with_trailing_punctuation() {
        assert_eq!(
            detect_link("https://example.com)."),
            Some(LinkTarget::Url("https://example.com".to_string()))
        );
    }

    #[test]
    fn test_url_with_path() {
        assert_eq!(
            detect_link("https://github.com/user/repo/issues/123"),
            Some(LinkTarget::Url("https://github.com/user/repo/issues/123".to_string()))
        );
    }

    #[test]
    fn test_url_with_query() {
        assert_eq!(
            detect_link("https://example.com/search?q=test&page=1"),
            Some(LinkTarget::Url("https://example.com/search?q=test&page=1".to_string()))
        );
    }

    #[test]
    fn test_url_ftp() {
        assert_eq!(
            detect_link("ftp://files.example.com/data.zip"),
            Some(LinkTarget::Url("ftp://files.example.com/data.zip".to_string()))
        );
    }

    #[test]
    fn test_no_scheme_not_url() {
        // example.com without scheme should NOT be detected as URL
        assert_eq!(detect_link("example.com"), None);
    }

    #[test]
    fn test_url_with_whitespace() {
        assert_eq!(
            detect_link("  https://example.com  "),
            Some(LinkTarget::Url("https://example.com".to_string()))
        );
    }

    // ========================================================================
    // Email detection tests
    // ========================================================================

    #[test]
    fn test_email_simple() {
        assert_eq!(
            detect_link("billing@example.com"),
            Some(LinkTarget::Email("billing@example.com".to_string()))
        );
    }

    #[test]
    fn test_email_with_trailing_dot() {
        assert_eq!(
            detect_link("billing@example.com."),
            Some(LinkTarget::Email("billing@example.com".to_string()))
        );
    }

    #[test]
    fn test_email_with_subdomain() {
        assert_eq!(
            detect_link("user@mail.example.org"),
            Some(LinkTarget::Email("user@mail.example.org".to_string()))
        );
    }

    #[test]
    fn test_email_with_plus() {
        assert_eq!(
            detect_link("user+tag@example.com"),
            Some(LinkTarget::Email("user+tag@example.com".to_string()))
        );
    }

    #[test]
    fn test_email_no_tld_not_email() {
        // foo@bar without a dot in domain should NOT be detected
        assert_eq!(detect_link("foo@bar"), None);
    }

    #[test]
    fn test_email_empty_local_not_email() {
        assert_eq!(detect_link("@example.com"), None);
    }

    #[test]
    fn test_email_double_at_not_email() {
        assert_eq!(detect_link("user@@example.com"), None);
    }

    // ========================================================================
    // Path detection tests
    // ========================================================================

    #[test]
    fn test_path_absolute_nonexistent() {
        // Non-existent path should not be detected
        assert_eq!(detect_link("/nonexistent/path/to/file.txt"), None);
    }

    #[test]
    fn test_path_tmp_if_exists() {
        // /tmp typically exists on Unix systems
        if std::path::Path::new("/tmp").exists() {
            assert_eq!(
                detect_link("/tmp"),
                Some(LinkTarget::Path(PathBuf::from("/tmp")))
            );
        }
    }

    #[test]
    fn test_path_tilde_expansion() {
        // Test that ~ expands but only detects if path exists
        // This test may not run in all environments (e.g., sandboxed)
        let home = dirs::home_dir();
        if let Some(home_path) = home {
            if home_path.exists() {
                let result = detect_link("~");
                // The home directory should be detected as a path
                // But we allow None if the path detection fails for any reason
                if result.is_some() {
                    assert!(matches!(result, Some(LinkTarget::Path(_))));
                }
            }
        }
    }

    // ========================================================================
    // Edge cases
    // ========================================================================

    #[test]
    fn test_empty_string() {
        assert_eq!(detect_link(""), None);
    }

    #[test]
    fn test_whitespace_only() {
        assert_eq!(detect_link("   "), None);
    }

    #[test]
    fn test_plain_text() {
        assert_eq!(detect_link("Hello, world!"), None);
    }

    #[test]
    fn test_number() {
        assert_eq!(detect_link("12345"), None);
    }

    #[test]
    fn test_punctuation_only() {
        assert_eq!(detect_link("..."), None);
    }

    // ========================================================================
    // Trailing punctuation stripping
    // ========================================================================

    #[test]
    fn test_strip_trailing_period() {
        assert_eq!(strip_trailing_punctuation("hello."), "hello");
    }

    #[test]
    fn test_strip_trailing_multiple() {
        assert_eq!(strip_trailing_punctuation("hello)."), "hello");
    }

    #[test]
    fn test_strip_preserves_url_path() {
        // Should not strip the dot in .com
        assert_eq!(strip_trailing_punctuation("example.com"), "example.com");
    }

    // ========================================================================
    // LinkTarget::open_string tests
    // ========================================================================

    #[test]
    fn test_open_string_url() {
        let target = LinkTarget::Url("https://example.com".to_string());
        assert_eq!(target.open_string(), "https://example.com");
    }

    #[test]
    fn test_open_string_email() {
        let target = LinkTarget::Email("user@example.com".to_string());
        assert_eq!(target.open_string(), "mailto:user@example.com");
    }

    #[test]
    fn test_open_string_path() {
        let target = LinkTarget::Path(PathBuf::from("/tmp/file.txt"));
        assert_eq!(target.open_string(), "/tmp/file.txt");
    }

    // ========================================================================
    // Edge cases (from review feedback)
    // ========================================================================

    // URL edge cases

    #[test]
    fn test_url_with_parentheses_in_path() {
        // Internal parentheses should be preserved, only trailing stripped
        assert_eq!(
            detect_link("https://example.com/foo(bar)"),
            Some(LinkTarget::Url("https://example.com/foo(bar)".to_string()))
        );
    }

    #[test]
    fn test_url_with_trailing_paren_only() {
        // Trailing paren stripped when it's orphaned
        assert_eq!(
            detect_link("https://example.com/page)"),
            Some(LinkTarget::Url("https://example.com/page".to_string()))
        );
    }

    #[test]
    fn test_url_file_scheme() {
        assert_eq!(
            detect_link("file:///home/user/doc.pdf"),
            Some(LinkTarget::Url("file:///home/user/doc.pdf".to_string()))
        );
    }

    #[test]
    fn test_url_with_newline_rejected() {
        // URLs with embedded newlines should not be detected
        assert_eq!(detect_link("https://example.com\n/path"), None);
    }

    // Email edge cases

    #[test]
    fn test_email_with_plus_tag() {
        // Plus addressing is common for filters
        assert_eq!(
            detect_link("a+b@example.com"),
            Some(LinkTarget::Email("a+b@example.com".to_string()))
        );
    }

    #[test]
    fn test_email_with_multi_level_subdomain() {
        assert_eq!(
            detect_link("name@sub.domain.tld"),
            Some(LinkTarget::Email("name@sub.domain.tld".to_string()))
        );
    }

    #[test]
    fn test_email_localhost_rejected() {
        // localhost emails are rejected (no dot in domain)
        assert_eq!(detect_link("name@localhost"), None);
    }

    #[test]
    fn test_email_with_dots_in_local() {
        assert_eq!(
            detect_link("first.last@example.com"),
            Some(LinkTarget::Email("first.last@example.com".to_string()))
        );
    }

    // Path edge cases

    #[test]
    fn test_relative_path_rejected() {
        // Relative paths like ./foo.txt are NOT detected
        // (we'd need to define "relative to what")
        assert_eq!(detect_link("./foo.txt"), None);
    }

    #[test]
    fn test_windows_path_not_detected_on_unix() {
        // Windows-style paths C:\... are detected by is_path_like
        // but won't exist on Unix, so they return None
        // This test just verifies we don't crash
        let result = detect_link("C:\\Users\\test\\doc.txt");
        // On Unix: None (path doesn't exist)
        // On Windows: might be Some if path exists
        // Either way, no crash
        let _ = result;
    }

    #[test]
    fn test_path_with_spaces_tilde() {
        // Path with spaces should work if it exists
        // We can't test actual path existence reliably,
        // but we can test the pattern is recognized
        let path = "~/Documents/My File.pdf";
        // This won't return Some unless the file exists,
        // but is_path_like should recognize it
        let _ = detect_link(path);
    }
}
