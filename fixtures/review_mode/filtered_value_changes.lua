-- Value-only plan used to dogfood change navigation with an active filter.

sheet:review({
    group = "normalize_amazon",
    title = "Normalize Amazon",
    reason = "Use the canonical vendor label.",
    sources = { "B2" },
})
sheet:set("B2", "Amazon")

sheet:review({
    group = "normalize_acme",
    title = "Normalize Acme",
    reason = "Use the canonical vendor label.",
    sources = { "B6" },
})
sheet:set("B6", "Acme Corporation")
