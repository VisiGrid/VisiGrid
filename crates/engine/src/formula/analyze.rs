// Formula analysis utilities
//
// Provides static analysis of formula ASTs without evaluation.
// Used during import to identify unsupported functions.

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
fn walk_expr<S, F: FnMut(&str)>(expr: &Expr<S>, visitor: &mut F) {
    match expr {
        Expr::Function { name, args } => {
            // Visit this function
            visitor(name);
            // Recurse into arguments
            for arg in args {
                walk_expr(arg, visitor);
            }
        }
        Expr::BinaryOp { left, right, .. } => {
            walk_expr(left, visitor);
            walk_expr(right, visitor);
        }
        // Leaf nodes - no functions to visit
        Expr::Number(_) |
        Expr::Text(_) |
        Expr::Boolean(_) |
        Expr::CellRef { .. } |
        Expr::Range { .. } |
        Expr::NamedRange(_) => {}
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
        // LET is not implemented (it's a lambda-like function)
        let expr = parse("=LET(A1, B1:B10, C1:C10)").unwrap();
        let mut counts = HashMap::new();
        tally_unknown_functions(&expr, &mut counts);
        assert_eq!(counts.get("LET"), Some(&1));
        assert_eq!(counts.len(), 1);
    }

    #[test]
    fn test_unknown_function_multiple_occurrences() {
        let expr = parse("=LET(A1, B1:B10, C1:C10) + LET(A2, B1:B10, C1:C10)").unwrap();
        let mut counts = HashMap::new();
        tally_unknown_functions(&expr, &mut counts);
        assert_eq!(counts.get("LET"), Some(&2));
    }

    #[test]
    fn test_mixed_known_and_unknown() {
        // SUM is known, LET and LAMBDA are unknown
        let expr = parse("=SUM(LET(A1, B1:B10, C1:C10), LAMBDA(A1, A2, A3))").unwrap();
        let mut counts = HashMap::new();
        tally_unknown_functions(&expr, &mut counts);
        assert_eq!(counts.get("LET"), Some(&1));
        assert_eq!(counts.get("LAMBDA"), Some(&1));
        assert!(counts.get("SUM").is_none()); // SUM is known
        assert_eq!(counts.len(), 2);
    }

    #[test]
    fn test_nested_unknown_functions() {
        // IF is known, LAMBDA and LET are unknown
        let expr = parse("=IF(LAMBDA(5) > 10, LET(A1, B1:B10, C1:C10), 0)").unwrap();
        let mut counts = HashMap::new();
        tally_unknown_functions(&expr, &mut counts);
        assert_eq!(counts.get("LAMBDA"), Some(&1));
        assert_eq!(counts.get("LET"), Some(&1));
        assert!(counts.get("IF").is_none());
    }

    #[test]
    fn test_has_unknown_functions() {
        let known = parse("=SUM(A1:A10)").unwrap();
        let unknown = parse("=LET(A1, B1:B10, C1:C10)").unwrap();

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
