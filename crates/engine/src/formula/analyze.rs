// Formula analysis utilities
//
// Provides static analysis of formula ASTs without evaluation.
// Used during import to identify unsupported functions.
//
// NOTE: Custom functions (user-defined Lua functions loaded from functions.lua)
// are not known to the engine's static analysis. They will appear as "unknown"
// in import tallies. This is acceptable because the analyzer only runs on the
// import path — it does not block formula evaluation, which dispatches custom
// functions via CellLookup::try_custom_function() at runtime.

use std::collections::HashMap;

use super::functions::is_known_function;
use super::parser::Expr;

/// Walk a formula AST and tally unknown function names.
///
/// This directly populates the provided HashMap, avoiding allocations.
/// Function names are already uppercase from the parser.
///
/// # Arguments
/// * `expr` - The parsed formula AST
/// * `counts` - HashMap to accumulate unknown function counts
///
/// # Example
/// ```ignore
/// let mut counts = HashMap::new();
/// tally_unknown_functions(&ast, &mut counts);
/// // counts now contains {"XLOOKUP": 2, "TEXTJOIN": 1}
/// ```
pub fn tally_unknown_functions<S>(expr: &Expr<S>, counts: &mut HashMap<String, usize>) {
    walk_expr(expr, &mut |name| {
        if !is_known_function(name) {
            *counts.entry(name.to_string()).or_insert(0) += 1;
        }
    });
}

/// Walk the AST and call the visitor for each function name encountered.
///
/// Calls of names bound by an enclosing LET or LAMBDA (`f(3)` in
/// `LET(f, LAMBDA(x, x*2), f(3))`) are not function names, and neither are
/// the parser's internal call names (underscore-prefixed); neither is
/// visited.
fn walk_expr<S, F: FnMut(&str)>(expr: &Expr<S>, visitor: &mut F) {
    walk_scoped(expr, &mut Vec::new(), visitor);
}

fn walk_scoped<S, F: FnMut(&str)>(expr: &Expr<S>, bound: &mut Vec<String>, visitor: &mut F) {
    match expr {
        Expr::RefError => {}
        Expr::Function { name, args } => {
            if !name.starts_with('_') && !bound.contains(name) {
                visitor(name);
            }
            let binds_names = (name == "LET" && args.len() >= 3) || (name == "LAMBDA" && !args.is_empty());
            if !binds_names {
                for arg in args {
                    walk_scoped(arg, bound, visitor);
                }
                return;
            }
            // LET(n1, v1, ..., calc) binds each name for what follows it;
            // LAMBDA(p1, ..., calc) binds its parameters for the calculation.
            let depth = bound.len();
            let last = args.len() - 1;
            for (i, arg) in args.iter().enumerate() {
                let is_name = if name == "LET" { i < last && i % 2 == 0 } else { i < last };
                match arg {
                    Expr::NamedRange(n) if is_name => bound.push(n.clone()),
                    _ => walk_scoped(arg, bound, visitor),
                }
            }
            bound.truncate(depth);
        }
        Expr::BinaryOp { left, right, .. } => {
            walk_scoped(left, bound, visitor);
            walk_scoped(right, bound, visitor);
        }
        // Leaf nodes - no functions to visit
        Expr::Number(_) |
        Expr::Text(_) |
        Expr::Boolean(_) |
        Expr::CellRef { .. } |
        Expr::Range { .. } |
        Expr::WholeRange { .. } |
        Expr::StructuredRef(_) | Expr::EmptyRange { .. } | Expr::ReferenceError(_) |
        Expr::NamedRange(_) |
        Expr::Empty => {}
    }
}

/// Check if a formula contains any unknown functions.
///
/// Returns true if at least one function in the AST is not known.
/// More efficient than tally_unknown_functions when you only need a boolean.
pub fn has_unknown_functions<S>(expr: &Expr<S>) -> bool {
    let mut found = false;
    walk_expr(expr, &mut |name| {
        if !found && !is_known_function(name) {
            found = true;
        }
    });
    found
}

/// Collect all function names used in a formula (known and unknown).
///
/// Useful for debugging or displaying formula dependencies.
pub fn collect_function_names<S>(expr: &Expr<S>) -> Vec<String> {
    let mut names = Vec::new();
    walk_expr(expr, &mut |name| {
        if !names.contains(&name.to_string()) {
            names.push(name.to_string());
        }
    });
    names
}

/// Functions that have dynamic/runtime-dependent references.
///
/// These functions compute their target references at evaluation time,
/// making static dependency analysis incomplete.
const DYNAMIC_REF_FUNCTIONS: &[&str] = &[
    "INDIRECT", // Converts text to cell reference
    "OFFSET",   // Returns reference offset from a starting point
];

/// Check if a formula contains functions with dynamic references.
///
/// Returns true if the formula contains INDIRECT, OFFSET, or similar
/// functions whose target cells cannot be determined statically.
///
/// Formulas with dynamic deps must be conservatively recomputed in
/// full ordered mode since their dependencies are incomplete.
/// Functions Excel treats as volatile: their result can change without any
/// cell they reference changing — the clock, the random generator, and
/// references resolved at evaluation time (INDIRECT and OFFSET read cells the
/// dependency graph cannot see). A formula using any of them is recalculated
/// on every recalculation, not only when its static inputs change.
const VOLATILE_FUNCTIONS: &[&str] = &["NOW", "TODAY", "RAND", "RANDBETWEEN", "INDIRECT", "OFFSET"];

/// Whether a formula calls a volatile function anywhere (see
/// `VOLATILE_FUNCTIONS`).
pub fn is_volatile<S>(expr: &Expr<S>) -> bool {
    let mut found = false;
    walk_expr(expr, &mut |name| {
        if !found && VOLATILE_FUNCTIONS.contains(&name) {
            found = true;
        }
    });
    found
}

pub fn has_dynamic_deps<S>(expr: &Expr<S>) -> bool {
    let mut found = false;
    walk_expr(expr, &mut |name| {
        if !found && DYNAMIC_REF_FUNCTIONS.contains(&name) {
            found = true;
        }
    });
    found
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::parser::parse;

    #[test]
    fn test_known_functions_no_unknowns() {
        let expr = parse("=SUM(A1:A10)").unwrap();
        let mut counts = HashMap::new();
        tally_unknown_functions(&expr, &mut counts);
        assert!(counts.is_empty());
    }

    #[test]
    fn test_single_unknown_function() {
        // DGET is not implemented (database functions are not)
        let expr = parse("=DGET(A1, B1:B10, C1:C10)").unwrap();
        let mut counts = HashMap::new();
        tally_unknown_functions(&expr, &mut counts);
        assert_eq!(counts.get("DGET"), Some(&1));
        assert_eq!(counts.len(), 1);
    }

    #[test]
    fn test_unknown_function_multiple_occurrences() {
        let expr = parse("=DGET(A1, B1:B10, C1:C10) + DGET(A2, B1:B10, C1:C10)").unwrap();
        let mut counts = HashMap::new();
        tally_unknown_functions(&expr, &mut counts);
        assert_eq!(counts.get("DGET"), Some(&2));
    }

    #[test]
    fn test_mixed_known_and_unknown() {
        // SUM is known, DGET and CUBEVALUE are unknown
        let expr = parse("=SUM(DGET(A1, B1:B10, C1:C10), CUBEVALUE(A1, A2, A3))").unwrap();
        let mut counts = HashMap::new();
        tally_unknown_functions(&expr, &mut counts);
        assert_eq!(counts.get("DGET"), Some(&1));
        assert_eq!(counts.get("CUBEVALUE"), Some(&1));
        assert!(counts.get("SUM").is_none()); // SUM is known
        assert_eq!(counts.len(), 2);
    }

    #[test]
    fn test_nested_unknown_functions() {
        // IF is known, CUBEVALUE and DGET are unknown
        let expr = parse("=IF(CUBEVALUE(5) > 10, DGET(A1, B1:B10, C1:C10), 0)").unwrap();
        let mut counts = HashMap::new();
        tally_unknown_functions(&expr, &mut counts);
        assert_eq!(counts.get("CUBEVALUE"), Some(&1));
        assert_eq!(counts.get("DGET"), Some(&1));
        assert!(counts.get("IF").is_none());
    }

    #[test]
    fn test_has_unknown_functions() {
        let known = parse("=SUM(A1:A10)").unwrap();
        let unknown = parse("=DGET(A1, B1:B10, C1:C10)").unwrap();

        assert!(!has_unknown_functions(&known));
        assert!(has_unknown_functions(&unknown));
    }

    #[test]
    fn test_collect_function_names() {
        // Test with mix of known (SUM, IF, XLOOKUP) functions
        let expr = parse("=SUM(IF(A1>0, XLOOKUP(A1, B1:B10, C1:C10), 0))").unwrap();
        let names = collect_function_names(&expr);

        assert!(names.contains(&"SUM".to_string()));
        assert!(names.contains(&"IF".to_string()));
        assert!(names.contains(&"XLOOKUP".to_string()));
        assert_eq!(names.len(), 3);
    }

    #[test]
    fn test_has_dynamic_deps_indirect() {
        let expr = parse("=INDIRECT(A1)").unwrap();
        assert!(has_dynamic_deps(&expr));
    }

    #[test]
    fn test_has_dynamic_deps_offset() {
        let expr = parse("=OFFSET(A1, 1, 1)").unwrap();
        assert!(has_dynamic_deps(&expr));
    }

    #[test]
    fn test_has_dynamic_deps_nested() {
        let expr = parse("=SUM(INDIRECT(A1))").unwrap();
        assert!(has_dynamic_deps(&expr));
    }

    #[test]
    fn test_has_dynamic_deps_none() {
        let expr = parse("=SUM(A1:A10) + AVERAGE(B1:B10)").unwrap();
        assert!(!has_dynamic_deps(&expr));
    }

    #[test]
    fn test_has_dynamic_deps_cell_ref_only() {
        let expr = parse("=A1+B1").unwrap();
        assert!(!has_dynamic_deps(&expr));
    }
}
