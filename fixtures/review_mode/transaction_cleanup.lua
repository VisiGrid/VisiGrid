-- Fixed Phase 1.5 dogfood fixture. Run against the transaction workbook
-- assembled by the scripting plan integration test.
-- All row numbers are source-workbook coordinates. Deletions are resolved
-- together, so the formula written at source row 7 appears at preview row 5.

sheet:verify({
    id = "retained_payments",
    kind = "gross_minus_group_equals_preview",
    label = "Retained payments",
    source_range = "A2:C6",
    amount_column = "C",
    excluded_group = "exact_duplicates",
    tolerance = 0.01,
    currency = "USD",
})

sheet:review({
    group = "normalize_vendors",
    title = "Normalize vendor names",
    reason = "Canonicalize two known Amazon labels.",
    sources = { "B2", "B3" },
})
sheet:set("B2", "Amazon")
sheet:set("B3", "Amazon")

sheet:review({
    group = "exact_duplicates",
    title = "Remove exact duplicates",
    reason = "The transaction identifier, vendor, and amount match row 2.",
    sources = { "A2:C2", "A4:C4" },
})
sheet:delete_rows(4, 1)

sheet:review({
    group = "empty_rows",
    title = "Remove empty rows",
    reason = "The row contains no transaction data.",
    sources = { "A5:C5" },
})
sheet:delete_rows(5, 1)

sheet:review({
    group = "rebuild_total",
    title = "Rebuild the total",
    reason = "Sum the retained payment rows.",
    sources = { "C2:C6" },
})
sheet:set_formula(7, 3, "=SUM(C2:C6)")
