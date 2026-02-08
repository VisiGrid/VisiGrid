use unicode_width::UnicodeWidthStr;

/// Display width of a string, accounting for CJK double-width, emoji, etc.
pub(crate) fn display_width(s: &str) -> usize {
    UnicodeWidthStr::width(s)
}

/// Truncate a string to fit within `width` display columns, adding ".." if truncated.
/// Uses Unicode display width so CJK/emoji alignment stays correct.
pub(crate) fn truncate_display(s: &str, width: usize) -> String {
    if width < 3 {
        // Just return the first char if it fits, else empty
        for ch in s.chars() {
            let cw = unicode_width::UnicodeWidthChar::width(ch).unwrap_or(0);
            if cw <= width {
                return ch.to_string();
            }
        }
        return String::new();
    }

    let str_width = UnicodeWidthStr::width(s);
    if str_width <= width {
        return s.to_string();
    }

    // Walk chars, accumulating display width, stop at width - 2 to leave room for ".."
    let budget = width - 2;
    let mut used = 0;
    let mut end_byte = 0;
    for (i, ch) in s.char_indices() {
        let cw = unicode_width::UnicodeWidthChar::width(ch).unwrap_or(0);
        if used + cw > budget {
            end_byte = i;
            break;
        }
        used += cw;
        end_byte = i + ch.len_utf8();
    }

    format!("{}..", &s[..end_byte])
}

/// Pad or truncate a string to exactly `width` display columns.
/// If shorter, right-pads with spaces. If longer, truncates with "..".
pub(crate) fn pad_right(s: &str, width: usize) -> String {
    let sw = UnicodeWidthStr::width(s);
    if sw > width {
        truncate_display(s, width)
    } else {
        format!("{}{}", s, " ".repeat(width - sw))
    }
}

/// Convert column index to letter (0 -> A, 1 -> B, 26 -> AA, etc.)
pub(crate) fn col_to_letter(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;
    loop {
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn display_width_ascii() {
        assert_eq!(display_width("hello"), 5);
        assert_eq!(display_width(""), 0);
    }

    #[test]
    fn display_width_cjk() {
        // CJK characters are 2 display columns each
        assert_eq!(display_width("\u{4e16}\u{754c}"), 4); // "世界"
    }

    #[test]
    fn truncate_fits() {
        assert_eq!(truncate_display("abc", 5), "abc");
        assert_eq!(truncate_display("abc", 3), "abc");
    }

    #[test]
    fn truncate_cuts() {
        assert_eq!(truncate_display("abcdef", 5), "abc..");
        assert_eq!(truncate_display("abcdef", 4), "ab..");
    }

    #[test]
    fn truncate_narrow() {
        assert_eq!(truncate_display("abc", 2), "a");
        assert_eq!(truncate_display("abc", 1), "a");
    }

    #[test]
    fn truncate_empty() {
        assert_eq!(truncate_display("", 5), "");
        assert_eq!(truncate_display("", 0), "");
    }

    #[test]
    fn truncate_cjk_boundary() {
        // "世界你好" is 8 display cols; truncate to 6 should cut at char boundary
        let s = "\u{4e16}\u{754c}\u{4f60}\u{597d}";
        let t = truncate_display(s, 6);
        // Budget is 4 display cols for text + 2 for ".."
        // "世界" = 4 display cols, fits in budget of 4
        assert_eq!(t, "\u{4e16}\u{754c}..");
        assert!(display_width(&t) <= 6);
    }

    #[test]
    fn pad_right_short() {
        assert_eq!(pad_right("ab", 5), "ab   ");
    }

    #[test]
    fn pad_right_exact() {
        assert_eq!(pad_right("abcde", 5), "abcde");
    }

    #[test]
    fn pad_right_long() {
        assert_eq!(pad_right("abcdef", 5), "abc..");
    }

    #[test]
    fn col_letters() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(701), "ZZ");
    }
}
