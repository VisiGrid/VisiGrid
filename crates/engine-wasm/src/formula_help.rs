//! Formula-editing help for the web grid: the same function table, caret
//! analysis, diagnostics and highlight tokens the desktop's formula bar uses
//! (`visigrid_engine::formula::help`). Positions cross as UTF-16 offsets, as
//! JavaScript strings count them.

use serde_json::{json, Value};
use visigrid_engine::formula::help::{analyze, check_errors, get_function, get_functions_by_prefix, tokenize_for_highlight, DiagnosticKind, FormulaEditMode, FunctionInfo, TokenType};
use wasm_bindgen::prelude::*;

/// UTF-16 offset of char index `i` in `s`.
fn u16_at(s: &str, i: usize) -> usize {
    s.chars().take(i).map(char::len_utf16).sum()
}

/// Char index of UTF-16 offset `o` in `s`.
fn char_at(s: &str, o: usize) -> usize {
    let mut n = 0;
    for (i, c) in s.chars().enumerate() {
        if n >= o {
            return i;
        }
        n += c.len_utf16();
    }
    s.chars().count()
}

fn function_json(f: &FunctionInfo) -> Value {
    json!({
        "name": f.name,
        "signature": f.signature,
        "description": f.description,
        "params": f.parameters.iter().map(|p| json!({"name": p.name, "description": p.description, "optional": p.optional, "repeatable": p.repeatable})).collect::<Vec<_>>(),
    })
}

fn token_kind(t: TokenType) -> &'static str {
    match t {
        TokenType::Function => "function",
        TokenType::CellRef => "cell",
        TokenType::Range => "range",
        TokenType::NamedRange => "name",
        TokenType::StructuredRef => "structured",
        TokenType::Number => "number",
        TokenType::String => "string",
        TokenType::Boolean => "boolean",
        TokenType::Operator => "operator",
        TokenType::Comparison => "comparison",
        TokenType::Paren => "paren",
        TokenType::Comma => "comma",
        TokenType::Colon => "colon",
        _ => "other",
    }
}

/// What the editor needs at `cursor` (UTF-16 offset) in `formula` (with its
/// leading `=`): edit mode, the function the caret is in and which argument,
/// autocomplete suggestions for the name being typed and the span they
/// replace, a definite mistake to show (if any), and every token's span.
pub(crate) fn help(formula: &str, cursor: usize) -> Value {
    let cur = char_at(formula, cursor);
    let ctx = analyze(formula, cur);
    let mode = match ctx.mode {
        FormulaEditMode::Start => "start",
        FormulaEditMode::Identifier => "identifier",
        FormulaEditMode::ArgList => "args",
        FormulaEditMode::String => "string",
        FormulaEditMode::Reference => "reference",
        FormulaEditMode::Operator => "operator",
        FormulaEditMode::Number => "number",
        FormulaEditMode::Complete => "complete",
    };
    let suggestions: Vec<Value> = match (&ctx.mode, &ctx.identifier_text) {
        (FormulaEditMode::Identifier, Some(prefix)) if !prefix.is_empty() => get_functions_by_prefix(prefix).into_iter().take(12).map(function_json).collect(),
        _ => Vec::new(),
    };
    let diagnostic = check_errors(formula, cur, &[]).filter(|d| matches!(d.kind, DiagnosticKind::Hard)).map(|d| {
        json!({"message": d.message, "span": d.span.map(|r| [u16_at(formula, r.start), u16_at(formula, r.end)])})
    });
    let tokens: Vec<Value> = tokenize_for_highlight(formula)
        .into_iter()
        .map(|(r, t)| json!({"start": u16_at(formula, r.start), "end": u16_at(formula, r.end), "kind": token_kind(t)}))
        .collect();
    json!({
        "mode": mode,
        "function": ctx.current_function.map(function_json),
        "arg_index": ctx.current_arg_index,
        "identifier": ctx.identifier_text,
        "replace": [u16_at(formula, ctx.replace_range.start), u16_at(formula, ctx.replace_range.end)],
        "suggestions": suggestions,
        "diagnostic": diagnostic,
        "tokens": tokens,
    })
}

/// See `help`.
#[wasm_bindgen]
pub fn formula_help(formula: &str, cursor: usize) -> Result<JsValue, JsValue> {
    use serde::Serialize as _;
    help(formula, cursor).serialize(&serde_wasm_bindgen::Serializer::json_compatible()).map_err(|e| JsValue::from_str(&e.to_string()))
}

/// A function's entry (signature, description, parameters), or null.
#[wasm_bindgen]
pub fn function_info(name: &str) -> Result<JsValue, JsValue> {
    use serde::Serialize as _;
    let v = get_function(name).map(function_json).unwrap_or(Value::Null);
    v.serialize(&serde_wasm_bindgen::Serializer::json_compatible()).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn suggests_functions_and_knows_the_argument() {
        let h = help("=SU", 3);
        assert_eq!(h["mode"], "identifier");
        let names: Vec<&str> = h["suggestions"].as_array().unwrap().iter().map(|s| s["name"].as_str().unwrap()).collect();
        assert!(names.contains(&"SUM") && names.contains(&"SUMIF"), "{names:?}");
        assert_eq!(h["replace"], json!([1, 3]));

        let h = help("=SUMIF(A1:A9, \">3\", ", 21);
        assert_eq!(h["function"]["name"], "SUMIF");
        assert_eq!(h["arg_index"], 2);
        assert!(!h["function"]["params"].as_array().unwrap().is_empty());

        let h = help("=A1+B2:C3", 9);
        let kinds: Vec<&str> = h["tokens"].as_array().unwrap().iter().map(|t| t["kind"].as_str().unwrap()).collect();
        // A range comes as cell, colon, cell (the page joins them).
        assert_eq!(kinds, ["operator", "cell", "operator", "cell", "colon", "cell"]);
    }

    #[test]
    fn utf16_offsets() {
        // "é" is one UTF-16 unit; an emoji is two.
        let f = "=\"😀\"&SU";
        let h = help(f, f.encode_utf16().count());
        assert_eq!(h["mode"], "identifier");
        let end = f.encode_utf16().count();
        assert_eq!(h["replace"], json!([end - 2, end]));
    }

    #[test]
    fn unknown_functions_are_a_definite_mistake() {
        let h = help("=SUMM(1)", 8);
        assert!(h["diagnostic"]["message"].as_str().unwrap_or("").to_lowercase().contains("unknown"), "{h}");
    }
}
