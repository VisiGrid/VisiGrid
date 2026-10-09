// Text functions: CONCATENATE, TEXTJOIN, LEFT, RIGHT, MID, LEN, UPPER, LOWER,
// TRIM, TEXT, VALUE, FIND, SUBSTITUTE, REPT, HYPERLINK, TEXTSPLIT, SPLIT, CHAR,
// CODE, CLEAN, FIXED, DOLLAR

use super::eval::{evaluate, Array2D, CellLookup, EvalResult, Value};
use super::eval_helpers::excel_error;
use super::parser::{BoundExpr, Expr};


/// Position of `needle` in `haystack`, counted in characters from `start`.
///
/// Excel measures strings in characters, so this does too — a byte offset
/// disagrees with LEN, LEFT, MID and RIGHT as soon as a string contains an
/// accent or an emoji, and silently returns a position one or more too far
/// along rather than failing.
fn find_chars(needle: &str, haystack: &str, start: usize, ignore_case: bool) -> Option<usize> {
    let hay: Vec<char> = if ignore_case {
        haystack.to_lowercase().chars().collect()
    } else {
        haystack.chars().collect()
    };
    let pat: Vec<char> = if ignore_case {
        needle.to_lowercase().chars().collect()
    } else {
        needle.chars().collect()
    };

    if start > hay.len() {
        return None;
    }
    // Excel finds an empty needle at the start position rather than nowhere.
    if pat.is_empty() {
        return Some(start);
    }
    if pat.len() > hay.len() {
        return None;
    }
    (start..=hay.len() - pat.len()).find(|&i| hay[i..i + pat.len()] == pat[..])
}

/// Position of the first place `pattern` matches, counted in characters from
/// `start`, with `*`, `?` and `~` behaving as they do everywhere else.
///
/// `wildcard_match` anchors at both ends because criteria in COUNTIF and
/// friends match a whole cell. SEARCH matches anywhere, so this asks the same
/// matcher whether the pattern matches a *prefix* of each tail — appending `*`
/// is what turns one question into the other. Reusing it is the point: the
/// wildcard rules then cannot drift between the functions that offer them.
fn search_chars(pattern: &str, haystack: &str, start: usize) -> Option<usize> {
    let hay: Vec<char> = haystack.chars().collect();
    if start > hay.len() {
        return None;
    }
    let prefix_pattern = format!("{pattern}*");
    (start..=hay.len()).find(|&i| {
        let tail: String = hay[i..].iter().collect();
        crate::formula::eval_helpers::wildcard_match(&prefix_pattern, &tail)
    })
}

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "CONCATENATE" | "CONCAT" => {
            let mut result = String::new();
            for arg in args {
                result.push_str(&evaluate(arg, lookup).to_text());
            }
            EvalResult::Text(result)
        }
        "TEXTJOIN" => {
            // TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
            if args.len() < 3 {
                return Some(EvalResult::Error("TEXTJOIN requires at least 3 arguments".to_string()));
            }
            let delimiter = evaluate(&args[0], lookup).to_text();
            let ignore_empty = evaluate(&args[1], lookup).to_bool().unwrap_or(true);

            let mut parts: Vec<String> = Vec::new();

            for arg in &args[2..] {
                match arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                        // Collect all values from range
                        let (min_row, min_col, max_row, max_col) = (
                            (*start_row).min(*end_row), (*start_col).min(*end_col),
                            (*start_row).max(*end_row), (*start_col).max(*end_col)
                        );
                        for r in min_row..=max_row {
                            for c in min_col..=max_col {
                                let text = lookup.get_text(r, c);
                                if !ignore_empty || !text.is_empty() {
                                    parts.push(text);
                                }
                            }
                        }
                    }
                    _ => {
                        let text = evaluate(arg, lookup).to_text();
                        if !ignore_empty || !text.is_empty() {
                            parts.push(text);
                        }
                    }
                }
            }

            EvalResult::Text(parts.join(&delimiter))
        }
        "LEFT" => {
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("LEFT requires 1 or 2 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let num_chars = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as usize,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1
            };
            EvalResult::Text(text.chars().take(num_chars).collect())
        }
        "RIGHT" => {
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("RIGHT requires 1 or 2 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let num_chars = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as usize,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1
            };
            let len = text.chars().count();
            let start = len.saturating_sub(num_chars);
            EvalResult::Text(text.chars().skip(start).collect())
        }
        "MID" => {
            if args.len() != 3 {
                return Some(EvalResult::Error("MID requires exactly 3 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let start = match evaluate(&args[1], lookup).to_number() {
                Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => (n as usize).saturating_sub(1), // 1-indexed
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let num_chars = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            EvalResult::Text(text.chars().skip(start).take(num_chars).collect())
        }
        "LEN" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("LEN requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Number(text.chars().count() as f64)
        }
        "UPPER" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("UPPER requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Text(text.to_uppercase())
        }
        "LOWER" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("LOWER requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Text(text.to_lowercase())
        }
        "TRIM" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("TRIM requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            // TRIM removes leading/trailing spaces and collapses internal spaces
            let trimmed: String = text.split_whitespace().collect::<Vec<_>>().join(" ");
            EvalResult::Text(trimmed)
        }
        "TEXT" => {
            // TEXT(value, format_text): a number rendered with an Excel format code.
            //
            // This used to be a stub that understood only "0.00"-style decimals and a
            // crude "%", so TEXT(date,"yyyy-mm-dd") returned the serial number and
            // "#,##0" dropped the separator. Cells already render custom formats with
            // ssfmt; TEXT now uses the same renderer, so a format code means the same
            // thing in a cell and in a formula.
            if args.len() != 2 {
                return Some(EvalResult::Error("TEXT requires exactly 2 arguments".to_string()));
            }
            let format = match evaluate(&args[1], lookup) {
                EvalResult::Error(e) => return Some(EvalResult::Error(e)),
                other => other.to_text(),
            };
            let value = match evaluate(&args[0], lookup) {
                EvalResult::Error(e) => return Some(EvalResult::Error(e)),
                EvalResult::Number(n) => n,
                EvalResult::Empty => 0.0,
                // Excel passes booleans and non-numeric text through unformatted.
                EvalResult::Boolean(b) => return Some(EvalResult::Text(if b { "TRUE" } else { "FALSE" }.to_string())),
                EvalResult::Text(t) => match crate::cell::parse_finite(t.trim()) {
                    Some(n) => n,
                    None => return Some(EvalResult::Text(t)),
                },
                other => match other.to_number() {
                    Ok(n) => n,
                    Err(e) => return Some(EvalResult::Error(e)),
                },
            };
            if format.is_empty() {
                return Some(EvalResult::Text(String::new()));
            }
            match ssfmt::format_default(value, &format) {
                Ok(text) => EvalResult::Text(text),
                Err(_) => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "VALUE" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("VALUE requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            match text.replace(',', "").trim().parse::<f64>() {
                Ok(n) => EvalResult::Number(n),
                Err(_) => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "FIND" | "SEARCH" => {
            // One implementation, because these differ only in whether case
            // matters. FIND used to be separate and indexed by bytes, which
            // made it disagree with LEN, LEFT, MID and RIGHT the moment a
            // string held anything outside ASCII.
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error(format!("{name} requires 2 or 3 arguments")));
            }
            let needle = evaluate(&args[0], lookup).to_text();
            let haystack = evaluate(&args[1], lookup).to_text();
            let start = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                    Ok(n) => (n as usize) - 1,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0
            };
            // FIND is literal; SEARCH takes Excel's wildcards. Both use the
            // matcher COUNTIF and XLOOKUP already share, rather than a second
            // one written to look the same.
            let found = if name == "SEARCH" {
                search_chars(&needle, &haystack, start)
            } else {
                find_chars(&needle, &haystack, start, false)
            };
            match found {
                Some(pos) => EvalResult::Number((pos + 1) as f64),
                None => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        // Excel counts characters, not bytes. These index by chars throughout,
        // so a name with an accent in it behaves the same as one without.
        "REPLACE" => {
            if args.len() != 4 {
                return Some(EvalResult::Error("REPLACE requires exactly 4 arguments".to_string()));
            }
            let text: Vec<char> = evaluate(&args[0], lookup).to_text().chars().collect();
            let start = match evaluate(&args[1], lookup).to_number() {
                Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => (n as usize) - 1,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let count = match evaluate(&args[2], lookup).to_number() {
                Ok(n) if n < 0.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let new_text = evaluate(&args[3], lookup).to_text();
            // Starting past the end appends, which is what Excel does rather
            // than erroring.
            let head: String = text.iter().take(start).collect();
            let tail: String = text.iter().skip(start.saturating_add(count)).collect();
            EvalResult::Text(format!("{head}{new_text}{tail}"))
        }
        "PROPER" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("PROPER requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            // A letter starts a word when what precedes it is not a letter, so
            // "o'neill" becomes "O'Neill" and "2nd place" becomes "2Nd Place",
            // both of which are what Excel produces.
            let mut out = String::with_capacity(text.len());
            let mut prev_alpha = false;
            for ch in text.chars() {
                if ch.is_alphabetic() {
                    if prev_alpha {
                        out.extend(ch.to_lowercase());
                    } else {
                        out.extend(ch.to_uppercase());
                    }
                    prev_alpha = true;
                } else {
                    out.push(ch);
                    prev_alpha = false;
                }
            }
            EvalResult::Text(out)
        }
        "EXACT" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("EXACT requires exactly 2 arguments".to_string()));
            }
            let a = evaluate(&args[0], lookup).to_text();
            let b = evaluate(&args[1], lookup).to_text();
            EvalResult::Boolean(a == b)
        }
        "TEXTBEFORE" | "TEXTAFTER" => {
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error(format!("{name} requires 2 or 3 arguments")));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let delimiter = evaluate(&args[1], lookup).to_text();
            let instance = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) => n as i64,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1
            };
            if delimiter.is_empty() || instance == 0 {
                return Some(EvalResult::Error("#VALUE!".to_string()));
            }
            let hits: Vec<usize> = text.match_indices(&delimiter).map(|(i, _)| i).collect();
            // A negative instance counts from the end, as Excel does.
            let chosen = if instance > 0 {
                hits.get((instance - 1) as usize).copied()
            } else {
                let from_end = (-instance) as usize;
                hits.len().checked_sub(from_end).and_then(|i| hits.get(i).copied())
            };
            match chosen {
                None => EvalResult::Error("#N/A".to_string()),
                Some(at) if name == "TEXTBEFORE" => EvalResult::Text(text[..at].to_string()),
                Some(at) => EvalResult::Text(text[at + delimiter.len()..].to_string()),
            }
        }
        "SUBSTITUTE" => {
            if args.len() < 3 || args.len() > 4 {
                return Some(EvalResult::Error("SUBSTITUTE requires 3 or 4 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let old_text = evaluate(&args[1], lookup).to_text();
            let new_text = evaluate(&args[2], lookup).to_text();
            let instance = if args.len() == 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => Some(n as usize),
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                None
            };

            let result = if let Some(n) = instance {
                // Replace only the nth instance
                let mut count = 0;
                let mut result = String::new();
                let mut remaining = text.as_str();
                while let Some(pos) = remaining.find(&old_text) {
                    count += 1;
                    if count == n {
                        result.push_str(&remaining[..pos]);
                        result.push_str(&new_text);
                        result.push_str(&remaining[pos + old_text.len()..]);
                        break;
                    } else {
                        result.push_str(&remaining[..pos + old_text.len()]);
                        remaining = &remaining[pos + old_text.len()..];
                    }
                }
                if count < n {
                    text // Not enough instances found
                } else {
                    result
                }
            } else {
                // Replace all instances
                text.replace(&old_text, &new_text)
            };
            EvalResult::Text(result)
        }
        "REPT" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("REPT requires exactly 2 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let times = match evaluate(&args[1], lookup).to_number() {
                Ok(n) if n < 0.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            EvalResult::Text(text.repeat(times))
        }
        "HYPERLINK" => {
            // HYPERLINK(link_location, [friendly_name]): a cell shows the
            // friendly name, or the link itself when there is none. Only the
            // value is computed here; following the link is the UI's business.
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("HYPERLINK requires 1 or 2 arguments".to_string()));
            }
            let link = evaluate(&args[0], lookup);
            if let EvalResult::Error(e) = link {
                return Some(EvalResult::Error(e));
            }
            match args.get(1) {
                Some(friendly) if !matches!(friendly, Expr::Empty) => evaluate(friendly, lookup),
                _ => EvalResult::Text(link.to_text()),
            }
        }
        "TEXTSPLIT" => {
            // TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty],
            // [match_mode], [pad_with]): split text into a grid that spills.
            // Either delimiter may be a list ({",",";"}); the earliest match
            // wins, the longest at a tie. ignore_empty drops empty pieces;
            // match_mode 1 ignores case. Short rows are padded with pad_with,
            // #N/A by default.
            if args.len() < 2 || args.len() > 6 {
                return Some(EvalResult::Error("TEXTSPLIT requires 2 to 6 arguments".to_string()));
            }
            let text = match evaluate(&args[0], lookup) {
                EvalResult::Error(e) => return Some(EvalResult::Error(e)),
                other => other.to_text(),
            };
            let delimiters = |arg: Option<&BoundExpr>| -> Result<Vec<String>, String> {
                let Some(arg) = arg.filter(|a| !matches!(a, Expr::Empty)) else { return Ok(Vec::new()) };
                let list: Vec<String> = match evaluate(arg, lookup) {
                    EvalResult::Error(e) => return Err(e),
                    EvalResult::Array(a) => (0..a.rows())
                        .flat_map(|r| (0..a.cols()).map(move |c| (r, c)))
                        .map(|(r, c)| a.get(r, c).map(|v| v.to_text()).unwrap_or_default())
                        .collect(),
                    other => vec![other.to_text()],
                };
                if list.iter().any(|d| d.is_empty()) {
                    return Err("#VALUE!".to_string());
                }
                Ok(list)
            };
            let cols = match delimiters(args.get(1)) {
                Ok(d) => d,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let rows = match delimiters(args.get(2)) {
                Ok(d) => d,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if cols.is_empty() && rows.is_empty() {
                return Some(EvalResult::Error("#VALUE!".to_string()));
            }
            let flag = |i: usize| -> Result<bool, String> {
                match args.get(i).filter(|a| !matches!(a, Expr::Empty)) {
                    None => Ok(false),
                    Some(a) => evaluate(a, lookup).to_bool(),
                }
            };
            let ignore_empty = match flag(3) {
                Ok(b) => b,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let ignore_case = match flag(4) {
                Ok(b) => b,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let pad = match args.get(5).filter(|a| !matches!(a, Expr::Empty)) {
                None => Value::Error("#N/A".to_string()),
                Some(a) => evaluate(a, lookup).to_value(),
            };
            let mut grid: Vec<Vec<String>> = Vec::new();
            for line in split_any(&text, &rows, ignore_case) {
                if ignore_empty && line.is_empty() {
                    continue;
                }
                let mut pieces = split_any(&line, &cols, ignore_case);
                if ignore_empty {
                    pieces.retain(|p| !p.is_empty());
                    if pieces.is_empty() {
                        continue;
                    }
                }
                grid.push(pieces);
            }
            if grid.is_empty() {
                return Some(EvalResult::Error("#CALC! Empty array".to_string()));
            }
            let width = grid.iter().map(Vec::len).max().unwrap_or(1);
            if grid.len() == 1 && width == 1 {
                return Some(EvalResult::Text(grid.remove(0).remove(0)));
            }
            if let Err(error) = super::eval_budget::array(grid.len(), width) { return Some(EvalResult::Error(error)); }
            let mut out = Array2D::new(grid.len(), width);
            for (r, row) in grid.into_iter().enumerate() {
                let filled = row.len();
                for (c, piece) in row.into_iter().enumerate() {
                    out.set(r, c, Value::Text(piece));
                }
                for c in filled..width {
                    out.set(r, c, pad.clone());
                }
            }
            EvalResult::Array(out)
        }
        "SPLIT" => {
            // SPLIT(text, delimiter, [split_by_each], [remove_empty_text]):
            // Google Sheets' one-row split, which spills across. By default
            // each character of delimiter splits on its own ("-/" splits at
            // either), and empty pieces are dropped; FALSE for split_by_each
            // splits only at the whole delimiter, FALSE for remove_empty_text
            // keeps the empty pieces. As in Sheets, a piece that reads as a
            // number becomes one ("007" is 7). Empty text or an empty
            // delimiter is #VALUE!.
            if args.len() < 2 || args.len() > 4 {
                return Some(EvalResult::Error("SPLIT requires 2 to 4 arguments".to_string()));
            }
            let text_arg = |i: usize| -> Result<String, String> {
                match evaluate(&args[i], lookup) {
                    EvalResult::Error(e) => Err(e),
                    EvalResult::Array(a) if a.rows() * a.cols() > 1 => Err("#VALUE!".to_string()),
                    other => Ok(other.to_text()),
                }
            };
            let text = match text_arg(0) {
                Ok(t) if t.is_empty() => return Some(EvalResult::Error("#VALUE! SPLIT needs text to split".to_string())),
                Ok(t) => t,
                Err(e) => return Some(EvalResult::Error(excel_error(e))),
            };
            let delimiter = match text_arg(1) {
                Ok(d) if d.is_empty() => return Some(EvalResult::Error("#VALUE! SPLIT needs a delimiter".to_string())),
                Ok(d) => d,
                Err(e) => return Some(EvalResult::Error(excel_error(e))),
            };
            let flag = |i: usize| -> Result<bool, String> {
                match args.get(i).filter(|a| !matches!(a, Expr::Empty)) {
                    None => Ok(true),
                    Some(a) => evaluate(a, lookup).to_bool(),
                }
            };
            let each = match flag(2) {
                Ok(b) => b,
                Err(e) => return Some(EvalResult::Error(excel_error(e))),
            };
            let remove_empty = match flag(3) {
                Ok(b) => b,
                Err(e) => return Some(EvalResult::Error(excel_error(e))),
            };
            let delimiters: Vec<String> = if each {
                delimiter.chars().map(String::from).collect()
            } else {
                vec![delimiter]
            };
            let mut pieces = split_any(&text, &delimiters, false);
            if remove_empty {
                pieces.retain(|p| !p.is_empty());
            }
            let piece = |p: String| match crate::cell::parse_finite(p.trim()) {
                Some(n) => Value::Number(n),
                None => Value::Text(p),
            };
            match pieces.len() {
                // Nothing but delimiters: a blank, as Sheets shows.
                0 => EvalResult::Empty,
                1 => EvalResult::from_value(&piece(pieces.remove(0))),
                n => {
                    if let Err(error) = super::eval_budget::array(1, n) { return Some(EvalResult::Error(error)); }
                    let mut out = Array2D::new(1, n);
                    for (c, p) in pieces.into_iter().enumerate() {
                        out.set(0, c, piece(p));
                    }
                    EvalResult::Array(out)
                }
            }
        }
        "CHAR" => {
            // CHAR(number): the character with that code. 1 to 255 are
            // Windows-1252, as in Excel, so CHAR(128) is "€"; above 255 the
            // number is a Unicode code point, as in Google Sheets. 0, a
            // negative number, or a value that is not a character is #VALUE!.
            if args.len() != 1 {
                return Some(EvalResult::Error("CHAR requires exactly one argument".to_string()));
            }
            let code = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n.trunc(),
                Err(e) => return Some(EvalResult::Error(excel_error(e))),
            };
            if !(1.0..=0x10FFFF as f64).contains(&code) {
                return Some(EvalResult::Error("#VALUE!".to_string()));
            }
            match char_from_code(code as u32) {
                Some(ch) => EvalResult::Text(ch.to_string()),
                None => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "CODE" => {
            // CODE(text): the code of the first character — CHAR's inverse,
            // so Windows-1252 for the characters it has and the Unicode code
            // point for the rest. Empty text is #VALUE!.
            if args.len() != 1 {
                return Some(EvalResult::Error("CODE requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            match text.chars().next() {
                Some(ch) => EvalResult::Number(code_of_char(ch) as f64),
                None => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "CLEAN" => {
            // CLEAN(text): text without the control characters 0 to 31 (line
            // breaks, tabs, bells). Other characters are kept, as in Excel.
            if args.len() != 1 {
                return Some(EvalResult::Error("CLEAN requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Text(text.chars().filter(|c| (*c as u32) >= 32).collect())
        }
        "FIXED" | "DOLLAR" => {
            // FIXED(number, [decimals], [no_commas]) and DOLLAR(number,
            // [decimals]): a number rounded to decimals places (2 when
            // omitted; negative rounds left of the point) and written as text
            // with thousands separators. FIXED drops the separators when
            // no_commas is TRUE. DOLLAR writes Excel's currency format,
            // $#,##0.00_);($#,##0.00): a dollar sign, and a negative amount
            // in parentheses. Decimals above 127 are #VALUE!.
            let max_args = if name == "FIXED" { 3 } else { 2 };
            if args.is_empty() || args.len() > max_args {
                return Some(EvalResult::Error(format!("{name} requires 1 to {max_args} arguments")));
            }
            let number = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(excel_error(e))),
            };
            let decimals = match args.get(1) {
                None => 2.0,
                Some(a) => match evaluate(a, lookup).to_number() {
                    Ok(n) => n.trunc(),
                    Err(e) => return Some(EvalResult::Error(excel_error(e))),
                },
            };
            if decimals > 127.0 {
                return Some(EvalResult::Error("#VALUE!".to_string()));
            }
            let no_commas = match args.get(2) {
                None => false,
                Some(a) => match evaluate(a, lookup).to_bool() {
                    Ok(b) => b,
                    Err(e) => return Some(EvalResult::Error(excel_error(e))),
                },
            };
            let decimals = decimals.max(-308.0) as i32;
            let rounded = super::eval_helpers::round_to_digits(number, decimals, super::eval_helpers::RoundMode::Nearest);
            let body = fixed_digits(rounded.abs(), decimals.max(0) as usize, !no_commas);
            let negative = rounded < 0.0;
            EvalResult::Text(match (name, negative) {
                ("FIXED", true) => format!("-{body}"),
                ("FIXED", false) => body,
                (_, true) => format!("(${body})"),
                (_, false) => format!("${body}"),
            })
        }
        _ => return None,
    };
    Some(result)
}

/// Windows-1252's characters for 128 to 159, where it differs from Latin-1.
/// The five codes it leaves undefined map to the matching control character.
const CP1252_HIGH: [char; 32] = [
    '\u{20AC}', '\u{0081}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}', '\u{2021}',
    '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008D}', '\u{017D}', '\u{008F}',
    '\u{0090}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}', '\u{2022}', '\u{2013}', '\u{2014}',
    '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}', '\u{0153}', '\u{009D}', '\u{017E}', '\u{0178}',
];

/// CHAR's character for a code: Windows-1252 up to 255, Unicode above.
fn char_from_code(code: u32) -> Option<char> {
    match code {
        128..=159 => Some(CP1252_HIGH[(code - 128) as usize]),
        _ => char::from_u32(code),
    }
}

/// CODE's number for a character: the inverse of `char_from_code`.
fn code_of_char(ch: char) -> u32 {
    match CP1252_HIGH.iter().position(|c| *c == ch) {
        Some(i) => 128 + i as u32,
        None => ch as u32,
    }
}

/// `value` (already rounded, not negative) with `decimals` places, and
/// thousands separators in the whole part when `commas` is set.
fn fixed_digits(value: f64, decimals: usize, commas: bool) -> String {
    let text = format!("{value:.decimals$}");
    let (whole, fraction) = match text.split_once('.') {
        Some((w, f)) => (w, Some(f)),
        None => (text.as_str(), None),
    };
    let mut out = String::with_capacity(text.len() + whole.len() / 3);
    for (i, digit) in whole.chars().enumerate() {
        if commas && i > 0 && (whole.len() - i) % 3 == 0 {
            out.push(',');
        }
        out.push(digit);
    }
    if let Some(f) = fraction {
        out.push('.');
        out.push_str(f);
    }
    out
}

#[cfg(test)]
mod newly_added_tests {
    use crate::formula::eval::{evaluate, CellLookup, EvalResult};
    use crate::formula::parser::{bind_expr_same_sheet, parse};

    struct Empty;
    impl CellLookup for Empty {
        fn get_value(&self, _r: usize, _c: usize) -> f64 { 0.0 }
        fn get_text(&self, _r: usize, _c: usize) -> String { String::new() }
    }

    fn eval(formula: &str) -> EvalResult {
        evaluate(&bind_expr_same_sheet(&parse(formula).unwrap()), &Empty)
    }

    fn text(formula: &str) -> String {
        match eval(formula) {
            EvalResult::Text(t) => t,
            other => panic!("{formula} gave {other:?}, expected text"),
        }
    }

    fn number(formula: &str) -> f64 {
        match eval(formula) {
            EvalResult::Number(n) => n,
            other => panic!("{formula} gave {other:?}, expected a number"),
        }
    }

    fn is_error(formula: &str) -> bool {
        matches!(eval(formula), EvalResult::Error(_))
    }

    /// SEARCH takes wildcards; FIND does not. That is Excel's split, and
    /// half of why the two functions both exist.
    ///
    /// The rules come from the matcher COUNTIF and XLOOKUP already use. When
    /// SEARCH was added it matched patterns literally instead, so the engine
    /// had wildcards that worked in one place and not another — the same rule
    /// with two implementations, which is how they drift.
    #[test]
    fn search_takes_wildcards_and_find_does_not() {
        assert_eq!(number(r#"=SEARCH("s?eet","Spreadsheet")"#), 7.0);
        assert_eq!(number(r#"=SEARCH("*sheet","Spreadsheet")"#), 1.0);
        assert_eq!(number(r#"=SEARCH("sp*et","Spreadsheet")"#), 1.0);
        // ~ escapes, so this looks for a literal question mark.
        assert_eq!(number(r#"=SEARCH("~?","a?b")"#), 2.0);
        assert!(is_error(r#"=SEARCH("z?z","abc")"#));

        // FIND treats the same pattern as characters to find.
        assert!(is_error(r#"=FIND("s?eet","Spreadsheet")"#));

        // Wildcards did not cost the plain cases their behaviour.
        assert_eq!(number(r#"=SEARCH("e","Spreadsheet",5)"#), 9.0);
        assert_eq!(number(r#"=SEARCH("é","caFÉ")"#), 4.0);
    }

    /// FIND counts characters, like every other function in this family.
    ///
    /// It used to slice the string by bytes. That made it disagree with LEN,
    /// LEFT, MID and RIGHT the moment a string held anything outside ASCII —
    /// so the standard split idiom, MID(text, FIND("-", text) + 1, ...),
    /// returned a slice starting one or more characters too far along. No
    /// error, just the wrong field. It could also panic outright when a start
    /// position landed inside a multi-byte character, which in the browser
    /// takes down the whole wasm instance rather than reddening one cell.
    #[test]
    fn find_counts_characters_not_bytes() {
        assert_eq!(number(r#"=FIND("-","aé-cd")"#), 3.0);
        // The composition is the point: these are what the split idiom does.
        assert_eq!(text(r#"=MID("aé-cd",FIND("-","aé-cd")+1,10)"#), "cd");
        assert_eq!(text(r#"=LEFT("aé-cd",FIND("-","aé-cd")-1)"#), "aé");
        // Agrees with LEN, which was always character-based.
        assert_eq!(number(r#"=LEN("café")"#), 4.0);
        assert_eq!(number(r#"=FIND("é","café")"#), 4.0);
    }

    /// A start position inside a multi-byte character returns an error.
    #[test]
    fn find_does_not_panic_on_a_character_boundary() {
        assert!(is_error(r#"=FIND("x","héllo",3)"#));
        assert_eq!(number(r#"=FIND("l","héllo",3)"#), 3.0);
        assert_eq!(text(r#"=IFERROR(FIND("x","héllo",3),"handled")"#), "handled");
    }

    /// FIND is case-sensitive and SEARCH is not; nothing else differs.
    #[test]
    fn find_and_search_differ_only_in_case() {
        assert!(is_error(r#"=FIND("B","abc")"#));
        assert_eq!(number(r#"=SEARCH("B","abc")"#), 2.0);
        assert_eq!(number(r#"=FIND("b","abc")"#), 2.0);
    }

    /// SEARCH is FIND without the case sensitivity.
    #[test]
    fn search_ignores_case_and_counts_from_one() {
        assert_eq!(number(r#"=SEARCH("b","ABC")"#), 2.0);
        assert_eq!(number(r#"=SEARCH("B","abc")"#), 2.0);
        assert_eq!(number(r#"=SEARCH("c","abcabc",4)"#), 6.0);
        // Excel reports #VALUE! when there is no match, not zero.
        assert!(is_error(r#"=SEARCH("z","abc")"#));
        // Counted in characters, so an accent earlier in the string does not
        // shift the answer the way byte offsets would.
        assert_eq!(number(r#"=SEARCH("é","caFÉ")"#), 4.0);
    }

    #[test]
    fn replace_works_on_character_positions() {
        assert_eq!(text(r#"=REPLACE("abcdef",2,3,"XY")"#), "aXYef");
        // Zero characters is an insert.
        assert_eq!(text(r#"=REPLACE("abcdef",2,0,"XY")"#), "aXYbcdef");
        // Starting past the end appends rather than erroring.
        assert_eq!(text(r#"=REPLACE("abc",10,2,"Z")"#), "abcZ");
    }

    #[test]
    fn proper_capitalises_after_every_non_letter() {
        assert_eq!(text(r#"=PROPER("hello world")"#), "Hello World");
        assert_eq!(text(r#"=PROPER("o'neill")"#), "O'Neill");
        // Excel really does produce "2Nd" here; the digit ends the word.
        assert_eq!(text(r#"=PROPER("2nd PLACE")"#), "2Nd Place");
        assert_eq!(text(r#"=PROPER("")"#), "");
    }

    #[test]
    fn exact_is_the_case_sensitive_comparison() {
        assert_eq!(eval(r#"=EXACT("a","A")"#), EvalResult::Boolean(false));
        assert_eq!(eval(r#"=EXACT("abc","abc")"#), EvalResult::Boolean(true));
    }

    #[test]
    fn textbefore_and_textafter_take_an_instance() {
        assert_eq!(text(r#"=TEXTBEFORE("a-b-c","-")"#), "a");
        assert_eq!(text(r#"=TEXTAFTER("a-b-c","-")"#), "b-c");
        assert_eq!(text(r#"=TEXTBEFORE("a-b-c","-",2)"#), "a-b");
        // A negative instance counts back from the end.
        assert_eq!(text(r#"=TEXTAFTER("a-b-c","-",-1)"#), "c");
        // Missing delimiter is #N/A, not an empty string — an empty string
        // would be indistinguishable from a delimiter at position one.
        assert!(is_error(r#"=TEXTBEFORE("abc","-")"#));
    }
}

/// Split `text` at every occurrence of any delimiter. With no delimiters the
/// text is one piece. At one position the longest matching delimiter wins.
/// Works on characters, so case-insensitive matching never splits a
/// multi-byte character.
fn split_any(text: &str, delimiters: &[String], ignore_case: bool) -> Vec<String> {
    if delimiters.is_empty() {
        return vec![text.to_string()];
    }
    let fold = |c: char| if ignore_case { c.to_lowercase().next().unwrap_or(c) } else { c };
    let chars: Vec<char> = text.chars().collect();
    let folded: Vec<char> = chars.iter().map(|&c| fold(c)).collect();
    let mut delims: Vec<Vec<char>> = delimiters.iter().map(|d| d.chars().map(fold).collect()).collect();
    delims.sort_by_key(|d| std::cmp::Reverse(d.len()));
    let mut pieces = Vec::new();
    let mut start = 0;
    let mut i = 0;
    while i < chars.len() {
        if let Some(d) = delims.iter().find(|d| folded[i..].starts_with(d)) {
            pieces.push(chars[start..i].iter().collect());
            i += d.len();
            start = i;
        } else {
            i += 1;
        }
    }
    pieces.push(chars[start..].iter().collect());
    pieces
}
