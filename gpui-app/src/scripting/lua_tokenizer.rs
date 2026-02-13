//! Single-pass Lua tokenizer for syntax highlighting.
//!
//! Returns `Vec<(Range<usize>, LuaTokenType)>` with byte offsets aligned to
//! UTF-8 char boundaries. Gaps (whitespace) are NOT emitted; the renderer
//! fills them as plain spans.

use std::ops::Range;

/// Token types for Lua syntax highlighting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LuaTokenType {
    Keyword,
    Boolean,
    String,
    Number,
    Comment,
    Operator,
    Punctuation,
    Identifier,
}

/// Tokenize Lua source code into a list of (byte_range, token_type) pairs.
///
/// All range endpoints are guaranteed to be valid UTF-8 char boundaries.
/// Tokens never overlap. Whitespace gaps are not emitted.
pub fn tokenize_lua(input: &str) -> Vec<(Range<usize>, LuaTokenType)> {
    let mut tokens = Vec::new();
    let mut chars = input.char_indices().peekable();

    while let Some(&(i, ch)) = chars.peek() {
        // Skip whitespace
        if ch.is_ascii_whitespace() {
            chars.next();
            continue;
        }

        // Comment: -- (two contiguous hyphens)
        if ch == '-' {
            let next = peek_ahead(input, i, 1);
            if next == Some('-') {
                let start = i;
                chars.next(); // first -
                chars.next(); // second -
                // Check for long bracket comment: --[=*[
                if let Some(level) = try_long_bracket_open(input, chars.peek().map(|&(j, _)| j).unwrap_or(input.len()), input) {
                    let bracket_start = chars.peek().map(|&(j, _)| j).unwrap_or(input.len());
                    // Skip past the opening [=*[
                    let open_len = 2 + level; // [ + N equals + [
                    for _ in 0..open_len {
                        chars.next();
                    }
                    // Find matching close ]=*]
                    let end = find_long_bracket_close(input, chars.peek().map(|&(j, _)| j).unwrap_or(input.len()), level);
                    // Advance chars past end
                    while chars.peek().map(|&(j, _)| j).unwrap_or(input.len()) < end {
                        chars.next();
                    }
                    tokens.push((start..end, LuaTokenType::Comment));
                } else {
                    // Line comment: consume to end of line
                    while let Some(&(_, c)) = chars.peek() {
                        if c == '\n' {
                            break;
                        }
                        chars.next();
                    }
                    let end = chars.peek().map(|&(j, _)| j).unwrap_or(input.len());
                    tokens.push((start..end, LuaTokenType::Comment));
                }
                continue;
            }
        }

        // Strings: "..." or '...'
        if ch == '"' || ch == '\'' {
            let start = i;
            let quote = ch;
            chars.next(); // opening quote
            loop {
                match chars.next() {
                    Some((_, '\\')) => {
                        // Skip escaped character
                        chars.next();
                    }
                    Some((j, c)) if c == quote => {
                        // Find the byte position after this closing quote
                        let end = next_char_boundary(input, j);
                        tokens.push((start..end, LuaTokenType::String));
                        break;
                    }
                    Some((_, '\n')) | None => {
                        // Unterminated string extends to EOF
                        let end = input.len();
                        tokens.push((start..end, LuaTokenType::String));
                        break;
                    }
                    Some(_) => {}
                }
            }
            continue;
        }

        // Long bracket strings: [=*[..]=*]
        if ch == '[' {
            if let Some(level) = try_long_bracket_open(input, i, input) {
                let start = i;
                let open_len = 2 + level;
                for _ in 0..open_len {
                    chars.next();
                }
                let end = find_long_bracket_close(input, chars.peek().map(|&(j, _)| j).unwrap_or(input.len()), level);
                while chars.peek().map(|&(j, _)| j).unwrap_or(input.len()) < end {
                    chars.next();
                }
                tokens.push((start..end, LuaTokenType::String));
                continue;
            }
        }

        // Numbers: digits, hex (0x/0X), floats, exponents
        if ch.is_ascii_digit() || (ch == '.' && matches!(peek_ahead(input, i, 1), Some(c) if c.is_ascii_digit())) {
            let start = i;
            let end = scan_number(input, &mut chars);
            tokens.push((start..end, LuaTokenType::Number));
            continue;
        }

        // Identifiers and keywords
        if ch.is_ascii_alphabetic() || ch == '_' {
            let start = i;
            while let Some(&(_, c)) = chars.peek() {
                if c.is_ascii_alphanumeric() || c == '_' {
                    chars.next();
                } else {
                    break;
                }
            }
            let end = chars.peek().map(|&(j, _)| j).unwrap_or(input.len());
            let word = &input[start..end];
            let tt = classify_word(word);
            tokens.push((start..end, tt));
            continue;
        }

        // Multi-char and single-char operators
        if is_operator_char(ch) {
            let start = i;
            let end = scan_operator(input, &mut chars);
            tokens.push((start..end, LuaTokenType::Operator));
            continue;
        }

        // Punctuation characters
        if is_punct_char(ch) {
            let start = i;
            chars.next();
            let end = chars.peek().map(|&(j, _)| j).unwrap_or(input.len());
            tokens.push((start..end, LuaTokenType::Punctuation));
            continue;
        }

        // Unknown character (e.g. unicode outside strings) - skip
        chars.next();
    }

    tokens
}

/// Peek at the character `offset` positions ahead from byte position `pos`.
fn peek_ahead(input: &str, pos: usize, offset: usize) -> Option<char> {
    input[pos..].char_indices().nth(offset).map(|(_, c)| c)
}

/// Get the byte position of the next character after byte position `pos`.
fn next_char_boundary(input: &str, pos: usize) -> usize {
    input[pos..]
        .char_indices()
        .nth(1)
        .map(|(i, _)| pos + i)
        .unwrap_or(input.len())
}

/// Try to parse a long bracket open at byte position `pos`.
/// Returns `Some(level)` where level is the number of `=` signs, or None.
/// Long bracket open: `[` + N `=` + `[`
fn try_long_bracket_open(input: &str, pos: usize, _src: &str) -> Option<usize> {
    let remaining = &input[pos..];
    let mut iter = remaining.char_indices();

    match iter.next() {
        Some((_, '[')) => {}
        _ => return None,
    }

    let mut level = 0;
    loop {
        match iter.next() {
            Some((_, '=')) => level += 1,
            Some((_, '[')) => return Some(level),
            _ => return None,
        }
    }
}

/// Find the end of a long bracket close `]=*]` with matching level.
/// Returns byte position after the close bracket.
fn find_long_bracket_close(input: &str, from: usize, level: usize) -> usize {
    let remaining = &input[from..];
    let mut iter = remaining.char_indices().peekable();

    while let Some((offset, ch)) = iter.next() {
        if ch == ']' {
            let close_start = from + offset;
            // Try to match =*]
            let mut eq_count = 0;
            let mut inner = remaining[offset..].char_indices().skip(1); // skip the first ]
            loop {
                match inner.next() {
                    Some((_, '=')) => eq_count += 1,
                    Some((off, ']')) if eq_count == level => {
                        // Matched! Return byte after the closing ]
                        return from + offset + off + 1;
                    }
                    _ => break,
                }
            }
        }
    }

    // Unterminated: extend to end of input
    input.len()
}

const LUA_KEYWORDS: &[&str] = &[
    "local", "function", "if", "then", "else", "elseif", "end",
    "for", "while", "do", "return", "repeat", "until", "break",
    "in", "and", "or", "not", "goto",
];

const LUA_BOOLEANS: &[&str] = &["true", "false", "nil"];

fn classify_word(word: &str) -> LuaTokenType {
    if LUA_BOOLEANS.contains(&word) {
        LuaTokenType::Boolean
    } else if LUA_KEYWORDS.contains(&word) {
        LuaTokenType::Keyword
    } else {
        LuaTokenType::Identifier
    }
}

/// Scan a number literal starting at the current char_indices position.
/// Handles decimal, hex (0x), floats, exponents.
fn scan_number(
    input: &str,
    chars: &mut std::iter::Peekable<std::str::CharIndices>,
) -> usize {
    // Check for hex prefix
    let is_hex = {
        let &(i, ch) = chars.peek().unwrap();
        if ch == '0' {
            matches!(peek_ahead(input, i, 1), Some('x') | Some('X'))
        } else {
            false
        }
    };

    if is_hex {
        chars.next(); // '0'
        chars.next(); // 'x' or 'X'
        // Hex digits + optional dot + hex digits + optional pN exponent
        while let Some(&(_, c)) = chars.peek() {
            if c.is_ascii_hexdigit() || c == '_' {
                chars.next();
            } else {
                break;
            }
        }
        // Optional hex float part: .hex_digits
        if let Some(&(_, '.')) = chars.peek() {
            // Only consume if followed by hex digit (avoid consuming `..` operator)
            let pos = chars.peek().map(|&(j, _)| j).unwrap_or(input.len());
            if matches!(peek_ahead(input, pos, 1), Some(c) if c.is_ascii_hexdigit()) {
                chars.next(); // '.'
                while let Some(&(_, c)) = chars.peek() {
                    if c.is_ascii_hexdigit() || c == '_' {
                        chars.next();
                    } else {
                        break;
                    }
                }
            }
        }
        // Optional binary exponent: p/P [+-] digits
        if let Some(&(_, c)) = chars.peek() {
            if c == 'p' || c == 'P' {
                chars.next();
                if let Some(&(_, c2)) = chars.peek() {
                    if c2 == '+' || c2 == '-' {
                        chars.next();
                    }
                }
                while let Some(&(_, c2)) = chars.peek() {
                    if c2.is_ascii_digit() {
                        chars.next();
                    } else {
                        break;
                    }
                }
            }
        }
    } else {
        // Decimal number
        // Leading dot already handled by caller for .5 style numbers
        let &(_, first) = chars.peek().unwrap();
        if first == '.' {
            chars.next(); // '.'
        }
        while let Some(&(_, c)) = chars.peek() {
            if c.is_ascii_digit() || c == '_' {
                chars.next();
            } else {
                break;
            }
        }
        // Optional decimal point
        if let Some(&(_, '.')) = chars.peek() {
            let pos = chars.peek().map(|&(j, _)| j).unwrap_or(input.len());
            // Don't consume if it's `..` (concat operator)
            if !matches!(peek_ahead(input, pos, 1), Some('.')) {
                chars.next(); // '.'
                while let Some(&(_, c)) = chars.peek() {
                    if c.is_ascii_digit() || c == '_' {
                        chars.next();
                    } else {
                        break;
                    }
                }
            }
        }
        // Optional exponent: e/E [+-] digits
        if let Some(&(_, c)) = chars.peek() {
            if c == 'e' || c == 'E' {
                chars.next();
                if let Some(&(_, c2)) = chars.peek() {
                    if c2 == '+' || c2 == '-' {
                        chars.next();
                    }
                }
                while let Some(&(_, c2)) = chars.peek() {
                    if c2.is_ascii_digit() {
                        chars.next();
                    } else {
                        break;
                    }
                }
            }
        }
    }

    chars.peek().map(|&(j, _)| j).unwrap_or(input.len())
}

/// Characters that start operator tokens.
fn is_operator_char(ch: char) -> bool {
    matches!(ch, '+' | '-' | '*' | '/' | '%' | '^' | '#' | '=' | '~' | '<' | '>' | '&' | '|' | '.')
}

/// Characters that are always single-char punctuation.
fn is_punct_char(ch: char) -> bool {
    matches!(ch, '(' | ')' | '{' | '}' | '[' | ']' | ',' | ';' | ':')
}

/// Scan a multi-char or single-char operator starting at the current position.
fn scan_operator(
    input: &str,
    chars: &mut std::iter::Peekable<std::str::CharIndices>,
) -> usize {
    let &(i, ch) = chars.peek().unwrap();
    chars.next();

    match ch {
        '.' => {
            // Could be `.`, `..`, or `...`
            if let Some(&(_, '.')) = chars.peek() {
                chars.next();
                if let Some(&(_, '.')) = chars.peek() {
                    chars.next(); // `...`
                }
                // `..` or `...`
            }
            // lone `.` -- emit as operator (field access)
        }
        '=' | '~' | '<' | '>' | '/' => {
            // Could be ==, ~=, <=, >=, //, <<, >>
            if let Some(&(_, next)) = chars.peek() {
                match (ch, next) {
                    ('=', '=') | ('~', '=') | ('<', '=') | ('>', '=') | ('/', '/') => {
                        chars.next();
                    }
                    ('<', '<') | ('>', '>') => {
                        chars.next();
                    }
                    _ => {}
                }
            }
        }
        // Single-char operators: + - * % ^ # & |
        _ => {}
    }

    chars.peek().map(|&(j, _)| j).unwrap_or(input.len())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn tok(input: &str) -> Vec<(&str, LuaTokenType)> {
        tokenize_lua(input)
            .into_iter()
            .map(|(range, tt)| (&input[range], tt))
            .collect()
    }

    #[test]
    fn test_keywords() {
        let result = tok("local function if then else end");
        assert_eq!(result, vec![
            ("local", LuaTokenType::Keyword),
            ("function", LuaTokenType::Keyword),
            ("if", LuaTokenType::Keyword),
            ("then", LuaTokenType::Keyword),
            ("else", LuaTokenType::Keyword),
            ("end", LuaTokenType::Keyword),
        ]);
    }

    #[test]
    fn test_identifiers() {
        let result = tok("foo bar_baz _x");
        assert_eq!(result, vec![
            ("foo", LuaTokenType::Identifier),
            ("bar_baz", LuaTokenType::Identifier),
            ("_x", LuaTokenType::Identifier),
        ]);
    }

    #[test]
    fn test_booleans() {
        let result = tok("true false nil");
        assert_eq!(result, vec![
            ("true", LuaTokenType::Boolean),
            ("false", LuaTokenType::Boolean),
            ("nil", LuaTokenType::Boolean),
        ]);
    }

    #[test]
    fn test_numbers() {
        let result = tok("123 3.14 0xFF 1e10 1E-3");
        assert_eq!(result, vec![
            ("123", LuaTokenType::Number),
            ("3.14", LuaTokenType::Number),
            ("0xFF", LuaTokenType::Number),
            ("1e10", LuaTokenType::Number),
            ("1E-3", LuaTokenType::Number),
        ]);
    }

    #[test]
    fn test_hex_float() {
        let result = tok("0xA.Fp2");
        assert_eq!(result, vec![
            ("0xA.Fp2", LuaTokenType::Number),
        ]);
    }

    #[test]
    fn test_strings() {
        let result = tok(r#""hello" 'world'"#);
        assert_eq!(result, vec![
            ("\"hello\"", LuaTokenType::String),
            ("'world'", LuaTokenType::String),
        ]);
    }

    #[test]
    fn test_string_escapes() {
        let result = tok(r#""he\"llo""#);
        assert_eq!(result, vec![
            (r#""he\"llo""#, LuaTokenType::String),
        ]);
    }

    #[test]
    fn test_long_bracket_string() {
        let result = tok("[==[multi-line\nlong string]==]");
        assert_eq!(result, vec![
            ("[==[multi-line\nlong string]==]", LuaTokenType::String),
        ]);
    }

    #[test]
    fn test_zero_equals_long_bracket() {
        let result = tok("[[multi-line]]");
        assert_eq!(result, vec![
            ("[[multi-line]]", LuaTokenType::String),
        ]);
    }

    #[test]
    fn test_line_comment() {
        let result = tok("x -- a comment\ny");
        assert_eq!(result, vec![
            ("x", LuaTokenType::Identifier),
            ("-- a comment", LuaTokenType::Comment),
            ("y", LuaTokenType::Identifier),
        ]);
    }

    #[test]
    fn test_block_comment() {
        let result = tok("--[==[block comment]==]");
        assert_eq!(result, vec![
            ("--[==[block comment]==]", LuaTokenType::Comment),
        ]);
    }

    #[test]
    fn test_zero_equals_block_comment() {
        let result = tok("--[[block comment]]");
        assert_eq!(result, vec![
            ("--[[block comment]]", LuaTokenType::Comment),
        ]);
    }

    #[test]
    fn test_not_a_comment_two_minus_ops() {
        // `a - -b` should be identifier, operator, operator, identifier
        let result = tok("a - -b");
        assert_eq!(result, vec![
            ("a", LuaTokenType::Identifier),
            ("-", LuaTokenType::Operator),
            ("-", LuaTokenType::Operator),
            ("b", LuaTokenType::Identifier),
        ]);
    }

    #[test]
    fn test_indexing_not_long_bracket() {
        // t[i] should not be treated as a long bracket
        let result = tok("t[i]");
        assert_eq!(result, vec![
            ("t", LuaTokenType::Identifier),
            ("[", LuaTokenType::Punctuation),
            ("i", LuaTokenType::Identifier),
            ("]", LuaTokenType::Punctuation),
        ]);
    }

    #[test]
    fn test_varargs_vs_concat_vs_field() {
        let result = tok("... .. .");
        assert_eq!(result, vec![
            ("...", LuaTokenType::Operator),
            ("..", LuaTokenType::Operator),
            (".", LuaTokenType::Operator),
        ]);
    }

    #[test]
    fn test_local_function_foo() {
        let result = tok("local function foo()");
        assert_eq!(result, vec![
            ("local", LuaTokenType::Keyword),
            ("function", LuaTokenType::Keyword),
            ("foo", LuaTokenType::Identifier),
            ("(", LuaTokenType::Punctuation),
            (")", LuaTokenType::Punctuation),
        ]);
    }

    #[test]
    fn test_operators() {
        let result = tok("== ~= <= >= // << >>");
        assert_eq!(result, vec![
            ("==", LuaTokenType::Operator),
            ("~=", LuaTokenType::Operator),
            ("<=", LuaTokenType::Operator),
            (">=", LuaTokenType::Operator),
            ("//", LuaTokenType::Operator),
            ("<<", LuaTokenType::Operator),
            (">>", LuaTokenType::Operator),
        ]);
    }

    #[test]
    fn test_single_char_operators() {
        let result = tok("+ - * / % ^ # & |");
        assert_eq!(result, vec![
            ("+", LuaTokenType::Operator),
            ("-", LuaTokenType::Operator),
            ("*", LuaTokenType::Operator),
            ("/", LuaTokenType::Operator),
            ("%", LuaTokenType::Operator),
            ("^", LuaTokenType::Operator),
            ("#", LuaTokenType::Operator),
            ("&", LuaTokenType::Operator),
            ("|", LuaTokenType::Operator),
        ]);
    }

    #[test]
    fn test_punctuation() {
        let result = tok("( ) { } [ ] , ; :");
        assert_eq!(result, vec![
            ("(", LuaTokenType::Punctuation),
            (")", LuaTokenType::Punctuation),
            ("{", LuaTokenType::Punctuation),
            ("}", LuaTokenType::Punctuation),
            ("[", LuaTokenType::Punctuation),
            ("]", LuaTokenType::Punctuation),
            (",", LuaTokenType::Punctuation),
            (";", LuaTokenType::Punctuation),
            (":", LuaTokenType::Punctuation),
        ]);
    }

    #[test]
    fn test_mixed() {
        let result = tok("local x = 42 -- a number");
        assert_eq!(result, vec![
            ("local", LuaTokenType::Keyword),
            ("x", LuaTokenType::Identifier),
            ("=", LuaTokenType::Operator),
            ("42", LuaTokenType::Number),
            ("-- a number", LuaTokenType::Comment),
        ]);
    }

    #[test]
    fn test_unicode_in_string() {
        let result = tok(r#""héllo 🌍""#);
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].1, LuaTokenType::String);
    }

    #[test]
    fn test_unicode_in_comment() {
        let result = tok("-- café ☕");
        assert_eq!(result.len(), 1);
        assert_eq!(result[0].1, LuaTokenType::Comment);
    }

    #[test]
    fn test_unterminated_string() {
        let result = tok("\"hello");
        assert_eq!(result, vec![
            ("\"hello", LuaTokenType::String),
        ]);
    }

    #[test]
    fn test_unterminated_long_bracket() {
        let result = tok("[==[unterminated");
        assert_eq!(result, vec![
            ("[==[unterminated", LuaTokenType::String),
        ]);
    }

    #[test]
    fn test_empty_input() {
        let result = tok("");
        assert!(result.is_empty());
    }

    #[test]
    fn test_whitespace_only() {
        let result = tok("   \n\t  ");
        assert!(result.is_empty());
    }

    #[test]
    fn test_char_boundary_invariant() {
        let input = "local ñ = \"héllo\" -- café";
        let tokens = tokenize_lua(input);
        for (range, _) in &tokens {
            assert!(input.is_char_boundary(range.start),
                "start {} is not a char boundary", range.start);
            assert!(input.is_char_boundary(range.end),
                "end {} is not a char boundary", range.end);
        }
    }

    #[test]
    fn test_no_overlapping_tokens() {
        let input = "local function foo(x, y)\n  return x + y\nend";
        let tokens = tokenize_lua(input);
        for window in tokens.windows(2) {
            assert!(window[0].0.end <= window[1].0.start,
                "tokens overlap: {:?} and {:?}", window[0], window[1]);
        }
    }

    #[test]
    fn test_number_before_concat() {
        // 42..x should parse as number 42, concat .., identifier x
        let result = tok("42 ..x");
        assert_eq!(result, vec![
            ("42", LuaTokenType::Number),
            ("..", LuaTokenType::Operator),
            ("x", LuaTokenType::Identifier),
        ]);
    }

    #[test]
    fn test_dot_number() {
        let result = tok(".5");
        assert_eq!(result, vec![
            (".5", LuaTokenType::Number),
        ]);
    }

    /// Stress test: 300-line realistic Lua script tokenizes correctly.
    #[test]
    fn test_stress_300_lines() {
        // Build a ~300-line script with all token types
        let mut script = String::new();
        for i in 0..100 {
            script.push_str(&format!(
                "local x_{i} = {i} + 3.14 -- iteration {i}\n\
                 if x_{i} > 0 then\n\
                   print(\"value: \" .. tostring(x_{i}))\n\
                 end\n"
            ));
        }
        let tokens = tokenize_lua(&script);

        // Verify invariants
        assert!(!tokens.is_empty());
        for (range, _) in &tokens {
            assert!(script.is_char_boundary(range.start));
            assert!(script.is_char_boundary(range.end));
        }
        for window in tokens.windows(2) {
            assert!(window[0].0.end <= window[1].0.start,
                "overlap at {:?} / {:?}", window[0], window[1]);
        }

        // Verify some expected tokens exist
        let has_keyword = tokens.iter().any(|(_, tt)| *tt == LuaTokenType::Keyword);
        let has_number = tokens.iter().any(|(_, tt)| *tt == LuaTokenType::Number);
        let has_string = tokens.iter().any(|(_, tt)| *tt == LuaTokenType::String);
        let has_comment = tokens.iter().any(|(_, tt)| *tt == LuaTokenType::Comment);
        assert!(has_keyword && has_number && has_string && has_comment);
    }

    /// Micro-benchmark: tokenize 10k chars in well under 10ms.
    /// Not a real bench harness, but catches catastrophic regressions.
    #[test]
    fn test_perf_10k_chars() {
        // ~10k chars of mixed Lua
        let chunk = "local x = 42 + 3.14 -- comment\n\
                     if x > 0 then print(\"hello\") end\n\
                     for i = 1, 100 do x = x + i end\n\
                     local t = {1, 2, 3, [\"key\"] = true}\n";
        let mut script = String::with_capacity(chunk.len() * 80);
        for _ in 0..80 {
            script.push_str(chunk);
        }
        assert!(script.len() >= 10_000, "script is {} bytes", script.len());

        let start = std::time::Instant::now();
        let iterations = 100;
        for _ in 0..iterations {
            let tokens = tokenize_lua(&script);
            // Prevent optimizer from eliding
            assert!(!tokens.is_empty());
        }
        let elapsed = start.elapsed();
        let per_call = elapsed / iterations;

        // 10k chars should tokenize in well under 10ms even on slow machines
        assert!(per_call.as_millis() < 10,
            "tokenize_lua took {:?} per call for {} bytes — too slow",
            per_call, script.len());
    }
}
