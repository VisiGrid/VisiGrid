//! The README states how many functions the engine has. It cannot import the
//! number, so this checks it: every "N built-in functions" or "N formula
//! functions" in the README must equal the engine's own function list.
//!
//! The count drifted before — the README said 123 in one place and 138 in
//! another at the same time.

use visigrid_engine::formula::functions::list_functions;

#[test]
fn readme_function_counts_match_the_engine() {
    let readme = include_str!("../../../README.md");
    let actual = list_functions().len();
    let mut found = 0;
    for (i, line) in readme.lines().enumerate() {
        let words: Vec<&str> = line.split_whitespace().collect();
        for w in words.windows(3) {
            let is_count = matches!(w[1], "built-in" | "formula") && w[2].starts_with("functions");
            let Some(n) = w[0].trim_start_matches(['-', '*', '(']).parse::<usize>().ok() else { continue };
            if is_count {
                found += 1;
                assert_eq!(n, actual, "README.md line {}: says {n} functions, the engine has {actual}", i + 1);
            }
        }
    }
    assert!(found > 0, "no function count found in README.md; if it was removed, delete this test");
}
