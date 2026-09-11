-- Negative dogfood fixture. The duplicate deletion is an allowed exclusion,
-- but deleting tx-003 is not. Retained payments must fail: expected 350,
-- actual 300.

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
    group = "exact_duplicates",
    title = "Remove exact duplicates",
    reason = "The transaction identifier, vendor, and amount match row 2.",
    sources = { "A2:C2", "A4:C4" },
})
sheet:delete_rows(4, 1)

sheet:review({
    group = "unclassified_removal",
    title = "Remove an unclassified payment",
    reason = "This deliberate mistake must be caught by retained-total verification.",
    sources = { "A6:C6" },
})
sheet:delete_rows(6, 1)
