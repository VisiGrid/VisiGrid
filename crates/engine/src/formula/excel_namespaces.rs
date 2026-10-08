//! Excel function namespaces are syntax wrappers, not part of function names.
// Excel namespaces function identifiers, not text, sheet names or Table fields.
// Scan before parsing: combined _xlfn._xlws identifiers exceed the native
// parser's grammar. Preserve all other bytes, including formula grouping.
pub fn normalize_formula(source: &str) -> String {
    if !source
        .as_bytes()
        .windows(6)
        .any(|s| s.eq_ignore_ascii_case(b"_xlfn.") || s.eq_ignore_ascii_case(b"_xlws."))
    {
        return source.to_owned();
    }
    let mut output = String::with_capacity(source.len());
    let mut i = 0;
    let mut brackets = 0usize;
    while i < source.len() {
        let start = i;
        let ch = source[i..].chars().next().unwrap();
        i += ch.len_utf8();
        if brackets > 0 {
            match ch {
                // Apostrophes escape special characters in structured headers.
                '\'' => {
                    if let Some(next) = source[i..].chars().next() {
                        i += next.len_utf8();
                    }
                }
                '[' => brackets += 1,
                ']' => brackets -= 1,
                _ => {}
            }
        } else if ch == '[' {
            brackets = 1;
        } else if ch == '"' || ch == '\'' {
            // Strings and quoted worksheet names use doubled delimiter escapes.
            while i < source.len() {
                let next = source[i..].chars().next().unwrap();
                i += next.len_utf8();
                if next == ch {
                    if source[i..].starts_with(ch) {
                        i += ch.len_utf8();
                    } else {
                        break;
                    }
                }
            }
        } else if ch.is_alphanumeric() || ch == '_' || ch == '\\' || ch == '.' {
            while let Some(next) = source[i..].chars().next() {
                if next.is_alphanumeric() || next == '_' || next == '\\' || next == '.' {
                    i += next.len_utf8();
                } else {
                    break;
                }
            }
            let mut name = &source[start..i];
            if source[i..].trim_start().starts_with('(') {
                for prefix in ["_xlfn.", "_xlws."] {
                    if name
                        .get(..prefix.len())
                        .is_some_and(|s| s.eq_ignore_ascii_case(prefix))
                    {
                        name = &name[prefix.len()..];
                    }
                }
            }
            output.push_str(name);
            continue;
        }
        output.push_str(&source[start..i]);
    }
    output
}
