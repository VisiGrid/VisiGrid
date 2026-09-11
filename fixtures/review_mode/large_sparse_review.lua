-- Sparse changes across the 30,000-row generated performance fixture.

sheet:review({
    group = "sparse_vendor_review",
    title = "Normalize sampled vendors",
    reason = "Exercise overlays, navigation, and the overview rail across a large sheet.",
    sources = { "B2", "B10002", "B20002", "B30000" },
})

for _, row in ipairs({ 2, 10002, 20002, 30000 }) do
    sheet:set("B" .. row, "Reviewed Vendor")
end
