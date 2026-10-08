//! Function names and LET/LAMBDA names as Excel stores them (#89).
//!
//! rust_xlsxwriter adds `_xlfn.` only to uppercase function names, so a
//! formula typed as `=xlookup(...)` opened as `#NAME?`. Excel also requires
//! `_xlpm.` on every LET and LAMBDA name and on each use of one; without it
//! Excel offers to repair the file. Names are scoped as the engine evaluates
//! them: a LET name covers the arguments after it, a LAMBDA parameter covers
//! the calculation, and a use outside that scope is left alone.

#[derive(Clone, Copy, PartialEq)]
enum Kind {
    Ident,
    Open,
    Close,
    Comma,
    Space,
    // Strings, quoted sheet names, structured references, array constants,
    // numbers and operators: copied unchanged.
    Other,
}

struct Token<'a> {
    kind: Kind,
    text: &'a str,
}

/// Rewrite a formula (with or without its leading `=`) for an .xlsx writer.
pub(crate) fn excel_function_names(source: &str) -> String {
    let tokens = tokenize(source);
    let mut out = String::with_capacity(source.len() + 16);
    let mut bound = Vec::new();
    let mut i = 0;
    while i < tokens.len() {
        i = expr_token(&tokens, i, &mut bound, &mut out);
    }
    out
}

/// The reverse, for import: drop `_xlfn.`, `_xlws.` and `_xlpm.` from names.
/// Strings, quoted sheet names and Table fields keep them.
pub(crate) fn from_excel(source: &str) -> String {
    if !source.as_bytes().windows(3).any(|w| w.eq_ignore_ascii_case(b"_xl")) {
        return source.to_owned();
    }
    let mut out = String::with_capacity(source.len());
    for token in tokenize(source) {
        let mut text = token.text;
        if token.kind == Kind::Ident {
            while let Some(rest) = ["_xlfn.", "_xlws.", "_xlpm."].iter().find_map(|prefix| {
                text.get(..prefix.len()).filter(|head| head.eq_ignore_ascii_case(prefix)).map(|_| &text[prefix.len()..])
            }) {
                text = rest;
            }
        }
        out.push_str(text);
    }
    out
}

fn tokenize(source: &str) -> Vec<Token<'_>> {
    let bytes = source.as_bytes();
    let mut tokens = Vec::new();
    let mut i = 0;
    while i < bytes.len() {
        let start = i;
        let ch = source[i..].chars().next().unwrap();
        i += ch.len_utf8();
        let kind = match ch {
            '(' => Kind::Open,
            ')' => Kind::Close,
            ',' => Kind::Comma,
            c if c.is_whitespace() => {
                while let Some(next) = source[i..].chars().next().filter(|c| c.is_whitespace()) {
                    i += next.len_utf8();
                }
                Kind::Space
            }
            // Doubled delimiters escape strings and quoted sheet names.
            '"' | '\'' => {
                while i < bytes.len() {
                    let next = source[i..].chars().next().unwrap();
                    i += next.len_utf8();
                    if next == ch {
                        if source[i..].starts_with(ch) {
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
                Kind::Other
            }
            // Structured references nest; an apostrophe escapes the next character.
            '[' => {
                let mut depth = 1usize;
                while i < bytes.len() && depth > 0 {
                    let next = source[i..].chars().next().unwrap();
                    i += next.len_utf8();
                    match next {
                        '\'' => {
                            if let Some(escaped) = source[i..].chars().next() {
                                i += escaped.len_utf8();
                            }
                        }
                        '[' => depth += 1,
                        ']' => depth -= 1,
                        _ => {}
                    }
                }
                Kind::Other
            }
            // Array constants hold no names, and their commas separate items.
            '{' => {
                while i < bytes.len() {
                    let next = source[i..].chars().next().unwrap();
                    i += next.len_utf8();
                    if next == '"' {
                        while i < bytes.len() {
                            let s = source[i..].chars().next().unwrap();
                            i += s.len_utf8();
                            if s == '"' {
                                if source[i..].starts_with('"') { i += 1; } else { break; }
                            }
                        }
                    } else if next == '}' {
                        break;
                    }
                }
                Kind::Other
            }
            // Error literals such as #N/A and #REF! are not names.
            '#' => {
                while let Some(next) = source[i..].chars().next().filter(|c| c.is_alphanumeric() || matches!(c, '/' | '!' | '?')) {
                    i += next.len_utf8();
                }
                Kind::Other
            }
            c if c.is_ascii_digit() => {
                while let Some(next) = source[i..].chars().next().filter(|c| c.is_ascii_alphanumeric() || *c == '.') {
                    i += next.len_utf8();
                }
                // Exponent sign: 1E+5.
                if matches!(bytes[i - 1], b'e' | b'E') && matches!(bytes.get(i), Some(b'+' | b'-')) {
                    i += 1;
                    while bytes.get(i).is_some_and(u8::is_ascii_digit) {
                        i += 1;
                    }
                }
                Kind::Other
            }
            c if c.is_alphabetic() || c == '_' || c == '\\' || c == '$' => {
                while let Some(next) = source[i..].chars().next().filter(|c| c.is_alphanumeric() || matches!(c, '_' | '.' | '\\' | '$')) {
                    i += next.len_utf8();
                }
                Kind::Ident
            }
            _ => Kind::Other,
        };
        tokens.push(Token { kind, text: &source[start..i] });
    }
    tokens
}

fn next_solid(tokens: &[Token], mut i: usize) -> Option<usize> {
    while tokens.get(i)?.kind == Kind::Space {
        i += 1;
    }
    Some(i)
}

/// Emit the token at `i` (a call consumes through its closing parenthesis)
/// and return the index after it.
fn expr_token(tokens: &[Token], i: usize, bound: &mut Vec<String>, out: &mut String) -> usize {
    let token = &tokens[i];
    if token.kind != Kind::Ident {
        out.push_str(token.text);
        return i + 1;
    }
    let is_bound = bound.iter().any(|name| name.eq_ignore_ascii_case(token.text));
    // `Sheet!A1` and `Table[Col]` qualify what follows; they are not uses.
    let qualifies = tokens.get(i + 1).is_some_and(|t| t.text.starts_with('!') || t.text.starts_with('['));
    let qualified = i > 0 && tokens[i - 1].text.ends_with('!');
    let call = tokens.get(i + 1).is_some_and(|t| t.kind == Kind::Open);
    if !call {
        if is_bound && !qualifies && !qualified {
            out.push_str("_xlpm.");
        }
        out.push_str(token.text);
        return i + 1;
    }
    let binds = token.text.eq_ignore_ascii_case("LET") || token.text.eq_ignore_ascii_case("LAMBDA");
    if is_bound && !binds {
        // A LET name holding a LAMBDA, called: `f(3)`.
        out.push_str("_xlpm.");
        out.push_str(token.text);
    } else {
        out.push_str(&token.text.to_ascii_uppercase());
    }
    out.push('(');
    let args = arguments(tokens, i + 2);
    let depth = bound.len();
    let is_let = token.text.eq_ignore_ascii_case("LET");
    // A LET name is bound after its value: `LET(x, x+1, x)` reads the
    // outer x in the value, as the engine evaluates it.
    let mut pending: Option<String> = None;
    let mut j = i + 2;
    for (index, &(start, end)) in args.ranges.iter().enumerate() {
        let last = args.ranges.len() - 1;
        let name_position = binds && index < last && args.ranges.len() >= 2
            && (token.text.eq_ignore_ascii_case("LAMBDA") || index % 2 == 0);
        let single = next_solid(tokens, start)
            .filter(|&k| k < end && tokens[k].kind == Kind::Ident)
            .filter(|&k| next_solid(tokens, k + 1).is_none_or(|n| n >= end));
        while j < start {
            out.push_str(tokens[j].text);
            j += 1;
        }
        if let (true, Some(k)) = (name_position, single) {
            for t in &tokens[start..k] { out.push_str(t.text); }
            out.push_str("_xlpm.");
            out.push_str(tokens[k].text);
            for t in &tokens[k + 1..end] { out.push_str(t.text); }
            if is_let {
                pending = Some(tokens[k].text.to_string());
            } else {
                // Parameters cover only the calculation, which is last.
                bound.push(tokens[k].text.to_string());
            }
            j = end;
        } else {
            while j < end {
                j = expr_token(tokens, j, bound, out);
            }
            bound.extend(pending.take());
        }
    }
    bound.truncate(depth);
    // Emit any trailing tokens up to and including the closing parenthesis.
    while j < args.close {
        out.push_str(tokens[j].text);
        j += 1;
    }
    if args.close < tokens.len() {
        out.push(')');
        args.close + 1
    } else {
        args.close
    }
}

struct Arguments {
    ranges: Vec<(usize, usize)>,
    close: usize,
}

/// Top-level argument token ranges of a call whose arguments start at `i`.
/// The commas between them stay outside the ranges.
fn arguments(tokens: &[Token], i: usize) -> Arguments {
    let mut ranges = Vec::new();
    let mut start = i;
    let mut depth = 0usize;
    let mut j = i;
    while j < tokens.len() {
        match tokens[j].kind {
            Kind::Open => depth += 1,
            Kind::Close if depth == 0 => break,
            Kind::Close => depth -= 1,
            Kind::Comma if depth == 0 => {
                ranges.push((start, j));
                start = j + 1;
            }
            _ => {}
        }
        j += 1;
    }
    if start < j || !ranges.is_empty() {
        ranges.push((start, j));
    }
    Arguments { ranges, close: j }
}

#[cfg(test)]
mod tests {
    use super::excel_function_names as x;

    #[test]
    fn function_names_are_uppercased_so_the_writer_prefixes_them() {
        assert_eq!(x("=xlookup(\"c\",H1:H3,I1:I3)"), "=XLOOKUP(\"c\",H1:H3,I1:I3)");
        assert_eq!(x("=sum(a1:a3)+Stdev.S(B1:B2)"), "=SUM(a1:a3)+STDEV.S(B1:B2)");
        assert_eq!(x("=\"xlookup(\"&lower(A1)"), "=\"xlookup(\"&LOWER(A1)");
    }

    #[test]
    fn let_and_lambda_names_get_the_parameter_prefix() {
        assert_eq!(x("=LET(x,5,y,7,x*y)"), "=LET(_xlpm.x,5,_xlpm.y,7,_xlpm.x*_xlpm.y)");
        assert_eq!(x("=LAMBDA(a,b,a+b)(2,3)"), "=LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)(2,3)");
        assert_eq!(x("=let( x , 1 , x )"), "=LET( _xlpm.x , 1 , _xlpm.x )");
        assert_eq!(
            x("=LET(f,LAMBDA(n,n*2),f(3))"),
            "=LET(_xlpm.f,LAMBDA(_xlpm.n,_xlpm.n*2),_xlpm.f(3))"
        );
    }

    #[test]
    fn names_are_prefixed_only_inside_their_scope() {
        // The first value is computed before x exists, so it reads the
        // defined name x; the same holds after the call closes.
        assert_eq!(x("=LET(x,x+1,x)+x"), "=LET(_xlpm.x,x+1,_xlpm.x)+x");
        assert_eq!(x("=LET(a,b,b,2,a+b)"), "=LET(_xlpm.a,b,_xlpm.b,2,_xlpm.a+_xlpm.b)");
        assert_eq!(x("=SUM(LAMBDA(v,v)(1),v)"), "=SUM(LAMBDA(_xlpm.v,_xlpm.v)(1),v)");
    }

    #[test]
    fn strings_sheets_tables_and_errors_are_not_names() {
        assert_eq!(x("=LET(x,\"x\",x&\"x\")"), "=LET(_xlpm.x,\"x\",_xlpm.x&\"x\")");
        assert_eq!(x("=LET(x,1,x!A1+'x'!A1+x)"), "=LET(_xlpm.x,1,x!A1+'x'!A1+_xlpm.x)");
        assert_eq!(x("=LET(x,1,Sales[x]+x)"), "=LET(_xlpm.x,1,Sales[x]+_xlpm.x)");
        assert_eq!(x("=LET(N,1,IFERROR(#N/A,N))"), "=LET(_xlpm.N,1,IFERROR(#N/A,_xlpm.N))");
        assert_eq!(x("=LET(x,{1,2;3,4},SUM(x))"), "=LET(_xlpm.x,{1,2;3,4},SUM(_xlpm.x))");
        assert_eq!(x("=LET(E,1E+5,E)"), "=LET(_xlpm.E,1E+5,_xlpm.E)");
    }

    #[test]
    fn import_drops_the_prefixes_outside_strings_and_fields() {
        use super::from_excel;
        assert_eq!(
            from_excel("=_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2),_xlpm.f(3))"),
            "=LET(f,LAMBDA(n,n*2),f(3))"
        );
        assert_eq!(from_excel("=_xlfn._xlws.FILTER(A1:A2,A1:A2>0)"), "=FILTER(A1:A2,A1:A2>0)");
        assert_eq!(
            from_excel("=_xlfn.XLOOKUP(\"_xlfn.x\",'_xlfn.S'!A1:A2,Sales[_xlpm.c])"),
            "=XLOOKUP(\"_xlfn.x\",'_xlfn.S'!A1:A2,Sales[_xlpm.c])"
        );
        for f in ["=SUM(A1:A3)", "=LET(x,5,x)", "=A1&\"_xl\""] {
            assert_eq!(from_excel(f), f);
        }
        for f in ["=xlookup(1,{1},{2})", "=LET(x,x+1,LAMBDA(a,a)(x))", "=LET(N,1,IFERROR(#N/A,N))"] {
            assert_eq!(from_excel(&x(f)).to_ascii_uppercase(), f.to_ascii_uppercase(), "{f}");
        }
    }

    #[test]
    fn formulas_without_names_are_unchanged() {
        for f in ["=A1+B1", "=SUM(A1:A3)", "='My Sheet'!A1*2", "=Sales[[#This Row],[Amount]]*2", "=IF(A1,\"a,b\",{1,2})", "=SUM(A1"] {
            assert_eq!(x(f), f, "{f}");
        }
    }
}
