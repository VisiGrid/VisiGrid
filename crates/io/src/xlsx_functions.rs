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
    let tokens = tokenize(source);
    let mut out = String::with_capacity(source.len());
    for (i, token) in tokens.iter().enumerate() {
        let mut text = token.text;
        if token.kind == Kind::Ident {
            // Function namespaces only wrap calls; a parameter prefix marks
            // every use of a LET or LAMBDA name.
            let call = tokens.get(i + 1).is_some_and(|t| t.kind == Kind::Open);
            let prefixes: &[&str] = if call { &["_xlfn.", "_xlws.", "_xlpm."] } else { &["_xlpm."] };
            while let Some(rest) = prefixes.iter().find_map(|prefix| {
                text.get(..prefix.len()).filter(|head| head.eq_ignore_ascii_case(prefix)).map(|_| &text[prefix.len()..])
            }) {
                text = rest;
            }
        }
        out.push_str(text);
    }
    out
}

/// Prefix the future functions rust_xlsxwriter 0.79.4 would prefix, including
/// on the table-XML path that never reaches that writer. A name that already
/// starts with `_xlfn.` is left alone. `FILTER`, `SORT` and `PY` take
/// `_xlfn._xlws.`; every other future function takes `_xlfn.`.
pub(crate) fn prefix_future_functions(source: &str) -> String {
    let tokens = tokenize(source);
    let mut out = String::with_capacity(source.len() + 16);
    for (i, token) in tokens.iter().enumerate() {
        if token.kind == Kind::Ident
            && tokens.get(i + 1).is_some_and(|next| next.kind == Kind::Open)
            && !token.text.get(..6).is_some_and(|head| head.eq_ignore_ascii_case("_xlfn."))
        {
            if let Some(prefix) = future_prefix(&token.text.to_ascii_uppercase()) {
                out.push_str(prefix);
            }
        }
        out.push_str(token.text);
    }
    out
}

fn future_prefix(name: &str) -> Option<&'static str> {
    use std::collections::HashMap;
    use std::sync::OnceLock;
    static FUTURE: OnceLock<HashMap<&'static str, u8>> = OnceLock::new();
    // Same classification as rust_xlsxwriter 0.79.4: 2 is `_xlfn._xlws.`.
    let kind = *FUTURE.get_or_init(|| {
        HashMap::from([
            ("ACOTH", 0), ("ACOT", 0), ("AGGREGATE", 0), ("ARABIC", 0), ("ARRAYTOTEXT", 0),
            ("BASE", 0), ("BETA.DIST", 0), ("BETA.INV", 0), ("BINOM.DIST.RANGE", 0),
            ("BINOM.DIST", 0), ("BINOM.INV", 0), ("BITAND", 0), ("BITLSHIFT", 0), ("BITOR", 0),
            ("BITRSHIFT", 0), ("BITXOR", 0), ("CEILING.MATH", 0), ("CEILING.PRECISE", 0),
            ("CHISQ.DIST.RT", 0), ("CHISQ.DIST", 0), ("CHISQ.INV.RT", 0), ("CHISQ.INV", 0),
            ("CHISQ.TEST", 0), ("COMBINA", 0), ("CONCAT", 0), ("CONFIDENCE.NORM", 0),
            ("CONFIDENCE.T", 0), ("COTH", 0), ("COT", 0), ("COVARIANCE.P", 0), ("COVARIANCE.S", 0),
            ("CSCH", 0), ("CSC", 0), ("DAYS", 0), ("DECIMAL", 0), ("ERF.PRECISE", 0),
            ("ERFC.PRECISE", 0), ("EXPON.DIST", 0), ("F.DIST.RT", 0), ("F.DIST", 0),
            ("F.INV.RT", 0), ("F.INV", 0), ("F.TEST", 0), ("FIELDVALUE", 0), ("FILTERXML", 0),
            ("FLOOR.MATH", 0), ("FLOOR.PRECISE", 0), ("FORECAST.ETS.CONFINT", 0),
            ("FORECAST.ETS.SEASONALITY", 0), ("FORECAST.ETS.STAT", 0), ("FORECAST.ETS", 0),
            ("FORECAST.LINEAR", 0), ("FORMULATEXT", 0), ("GAMMA.DIST", 0), ("GAMMA.INV", 0),
            ("GAMMALN.PRECISE", 0), ("GAMMA", 0), ("GAUSS", 0), ("HYPGEOM.DIST", 0), ("IFNA", 0),
            ("IFS", 0), ("IMAGE", 0), ("IMCOSH", 0), ("IMCOT", 0), ("IMCSCH", 0), ("IMCSC", 0),
            ("IMSECH", 0), ("IMSEC", 0), ("IMSINH", 0), ("IMTAN", 0), ("ISFORMULA", 0),
            ("ISOMITTED", 0), ("ISOWEEKNUM", 0), ("LET", 0), ("LOGNORM.DIST", 0),
            ("LOGNORM.INV", 0), ("MAXIFS", 0), ("MINIFS", 0), ("MODE.MULT", 0), ("MODE.SNGL", 0),
            ("MUNIT", 0), ("NEGBINOM.DIST", 0), ("NORM.DIST", 0), ("NORM.INV", 0),
            ("NORM.S.DIST", 0), ("NORM.S.INV", 0), ("NUMBERVALUE", 0), ("PDURATION", 0),
            ("PERCENTILE.EXC", 0), ("PERCENTILE.INC", 0), ("PERCENTRANK.EXC", 0),
            ("PERCENTRANK.INC", 0), ("PERMUTATIONA", 0), ("PHI", 0), ("POISSON.DIST", 0),
            ("PQSOURCE", 0), ("PYTHON_STR", 0), ("PYTHON_TYPE", 0), ("PYTHON_TYPENAME", 0),
            ("QUARTILE.EXC", 0), ("QUARTILE.INC", 0), ("QUERYSTRING", 0), ("RANK.AVG", 0),
            ("RANK.EQ", 0), ("RRI", 0), ("SECH", 0), ("SEC", 0), ("SHEETS", 0), ("SHEET", 0),
            ("SKEW.P", 0), ("STDEV.P", 0), ("STDEV.S", 0), ("T.DIST.2T", 0), ("T.DIST.RT", 0),
            ("T.DIST", 0), ("T.INV.2T", 0), ("T.INV", 0), ("T.TEST", 0), ("TEXTAFTER", 0),
            ("TEXTBEFORE", 0), ("TEXTJOIN", 0), ("UNICHAR", 0), ("UNICODE", 0), ("VALUETOTEXT", 0),
            ("VAR.P", 0), ("VAR.S", 0), ("WEBSERVICE", 0), ("WEIBULL.DIST", 0), ("XMATCH", 0),
            ("XOR", 0), ("Z.TEST", 0), ("ANCHORARRAY", 1), ("BYCOL", 1), ("BYROW", 1),
            ("CHOOSECOLS", 1), ("CHOOSEROWS", 1), ("DROP", 1), ("EXPAND", 1), ("HSTACK", 1),
            ("LAMBDA", 1), ("MAKEARRAY", 1), ("MAP", 1), ("RANDARRAY", 1), ("REDUCE", 1),
            ("SCAN", 1), ("SEQUENCE", 1), ("SINGLE", 1), ("SORTBY", 1), ("SWITCH", 1),
            ("TAKE", 1), ("TEXTSPLIT", 1), ("TOCOL", 1), ("TOROW", 1), ("UNIQUE", 1),
            ("VSTACK", 1), ("WRAPCOLS", 1), ("WRAPROWS", 1), ("XLOOKUP", 1),
            ("FILTER", 2), ("SORT", 2), ("PY", 2),
        ])
    }).get(name)?;
    Some(if kind == 2 { "_xlfn._xlws." } else { "_xlfn." })
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
        // Keep an existing _xlfn./_xlws. prefix as Excel spells it.
        let mut name = token.text;
        while let Some(prefix) = ["_xlfn.", "_xlws."].iter().find(|p| name.get(..p.len()).is_some_and(|h| h.eq_ignore_ascii_case(p))) {
            out.push_str(prefix);
            name = &name[prefix.len()..];
        }
        out.push_str(&name.to_ascii_uppercase());
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
        assert_eq!(x("=_xlfn.xlookup(1,A1,B1)"), "=_xlfn.XLOOKUP(1,A1,B1)");
        assert_eq!(x("=_XLFN._xlws.filter(A1,A1>0)"), "=_xlfn._xlws.FILTER(A1,A1>0)");
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
        // Not calls: a sheet, a Table field and a defined name keep the prefix.
        assert_eq!(
            from_excel("='_xlfn.SEQUENCE'!A1+Sales[_xlfn.SEQUENCE]+_xlfn.Name"),
            "='_xlfn.SEQUENCE'!A1+Sales[_xlfn.SEQUENCE]+_xlfn.Name"
        );
        assert_eq!(from_excel("=(_xlfn.SEQUENCE(2)+1)*2"), "=(SEQUENCE(2)+1)*2");
        assert_eq!(
            from_excel("=_xlfn.SEQUENCE(2)+LEN(\"é\"\"_xlfn.SEQUENCE(2)\")+Sales[a']_xlfn.SEQUENCE(2)]"),
            "=SEQUENCE(2)+LEN(\"é\"\"_xlfn.SEQUENCE(2)\")+Sales[a']_xlfn.SEQUENCE(2)]"
        );
        for f in ["=SUM(A1:A3)", "=LET(x,5,x)", "=A1&\"_xl\""] {
            assert_eq!(from_excel(f), f);
        }
        for f in ["=xlookup(1,{1},{2})", "=LET(x,x+1,LAMBDA(a,a)(x))", "=LET(N,1,IFERROR(#N/A,N))"] {
            assert_eq!(from_excel(&x(f)).to_ascii_uppercase(), f.to_ascii_uppercase(), "{f}");
        }
    }

    #[test]
    fn future_functions_are_prefixed_once_and_ordinary_names_are_not() {
        use super::prefix_future_functions as p;
        assert_eq!(p("=XLOOKUP(1,A1,B1)"), "=_xlfn.XLOOKUP(1,A1,B1)");
        assert_eq!(p("=FILTER(A1:A2,A1:A2>0)"), "=_xlfn._xlws.FILTER(A1:A2,A1:A2>0)");
        assert_eq!(p("=SORT(A1:A2)"), "=_xlfn._xlws.SORT(A1:A2)");
        assert_eq!(p("=LET(_xlpm.x,5,_xlpm.x)"), "=_xlfn.LET(_xlpm.x,5,_xlpm.x)");
        assert_eq!(p("=SUM(A1)+[[#This Row],[Qty]]*C4"), "=SUM(A1)+[[#This Row],[Qty]]*C4");
        assert_eq!(p("=_xlfn.XLOOKUP(1,A1,B1)"), "=_xlfn.XLOOKUP(1,A1,B1)");
        assert_eq!(p("=_xlfn._xlws.FILTER(A1,A1>0)"), "=_xlfn._xlws.FILTER(A1,A1>0)");
        assert_eq!(p("=_xlfn.Name"), "=_xlfn.Name");
    }

    #[test]
    fn formulas_without_names_are_unchanged() {
        for f in ["=A1+B1", "=SUM(A1:A3)", "='My Sheet'!A1*2", "=Sales[[#This Row],[Amount]]*2", "=IF(A1,\"a,b\",{1,2})", "=SUM(A1"] {
            assert_eq!(x(f), f, "{f}");
        }
    }
}
