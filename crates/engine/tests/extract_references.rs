use visigrid_engine::formula::extract::{
    first_reference_literal, reference_literal_count, replace_reference_literal,
};

#[test]
fn extraction_preserves_source_and_uses_complete_reference_tokens() {
    for (source, reference, name, expected, count) in [
        (
            "=SUM( A1 : B10 )+\"A1:B10 é\"",
            "A1:B10",
            "Values",
            "=SUM( Values )+\"A1:B10 é\"",
            1,
        ),
        (
            "=( $a$1 + $A$1 )*2 + A1",
            "$A$1",
            "Rate",
            "=( Rate + Rate )*2 + A1",
            2,
        ),
        (
            "=SUM(A1:B10)+A1+AA1+A11+_A1+A1_extra",
            "A1",
            "Rate",
            "=SUM(A1:B10)+Rate+AA1+A11+_A1+A1_extra",
            1,
        ),
        (
            "=A1+Sheet1!A1+'Other! O''Brien'!A1+Sales[A1]",
            "A1",
            "Rate",
            "=Rate+Sheet1!A1+'Other! O''Brien'!A1+Sales[A1]",
            1,
        ),
        (
            "='Other! O''Brien'!A1+\"'Other! O''Brien'!A1\"",
            "'other! o''brien'!a1",
            "Rate",
            "=Rate+\"'Other! O''Brien'!A1\"",
            1,
        ),
        (
            "=LOG10(100)+1E3+E3",
            "E3",
            "Rate",
            "=LOG10(100)+1E3+Rate",
            1,
        ),
        ("=\"a\"\"A1\"&A1", "A1", "Rate", "=\"a\"\"A1\"&Rate", 1),
        ("=SUM(B10:A1)", "B10:A1", "Values", "=SUM(Values)", 1),
    ] {
        assert_eq!(
            reference_literal_count(source, reference).unwrap(),
            count,
            "{source}"
        );
        assert_eq!(
            replace_reference_literal(source, reference, name).unwrap(),
            (expected.into(), count),
            "{source}"
        );
    }
}

#[test]
fn detection_ignores_strings_headers_exponents_and_function_names() {
    for (formula, expected) in [
        ("=\"A1:B10\"&SUM(Sales[A1])+LOG10(1E3)+$C$2", Some("$C$2")),
        (
            "=SUM('Other! O''Brien'!A1 : B4)",
            Some("'Other! O''Brien'!A1 : B4"),
        ),
        ("=SUM(A:A)+1E3+LOG10(100)", None),
        ("=SUM(A1:B10)+A1", Some("A1:B10")),
    ] {
        assert_eq!(
            first_reference_literal(formula).unwrap().as_deref(),
            expected
        );
    }
}

#[test]
fn local_bindings_cannot_capture_an_extracted_name() {
    for source in [
        "=LET(Rate,99,A1*Rate)",
        "=LAMBDA(Rate,A1*Rate)(2)",
        "=LET(Rate,99,LET(Other,0,A1))",
    ] {
        assert!(replace_reference_literal(source, "A1", "Rate")
            .unwrap_err()
            .contains("LET or LAMBDA"));
    }
    assert_eq!(
        replace_reference_literal("=LET(Rate,A1,Rate*2)", "A1", "Rate")
            .unwrap()
            .0,
        "=LET(Rate,Rate,Rate*2)"
    );
    assert_eq!(
        replace_reference_literal("=LET(Rate,99,Rate)+A1", "A1", "Rate")
            .unwrap()
            .0,
        "=LET(Rate,99,Rate)+Rate"
    );
    assert!(replace_reference_literal("=A1+@", "A1", "Rate").is_err());
    assert_eq!(
        replace_reference_literal("=UNSUPPORTED(@)", "A1", "Rate")
            .unwrap()
            .0,
        "=UNSUPPORTED(@)"
    );
    assert!(replace_reference_literal("=A1", "A1", "B2").is_err());
}
