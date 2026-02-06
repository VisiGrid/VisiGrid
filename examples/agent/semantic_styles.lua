-- VisiGrid Agent Demo: Semantic Cell Styles
-- Demonstrates intent-driven styling that agents should prefer over
-- low-level formatting (fill colors, font colors, borders).
--
-- Run in the Lua console (sheet:* API).
-- Styles convey meaning. The theme resolves them to visuals.

-- ============================================================================
-- Sugar methods (one per style)
-- ============================================================================

-- Mark input cells so users know what to edit
sheet:input("B2:B5")

-- Mark a totals row
sheet:total("A7:D7")

-- Flag problems
sheet:error("C3")
sheet:warning("C4")

-- Positive outcome / informational
sheet:success("D2")
sheet:note("E1")

-- Clear style back to default
sheet:clear_style("E1")

-- ============================================================================
-- General method: sheet:style(range, name_or_constant)
-- ============================================================================

-- String name (case-insensitive, aliases accepted)
sheet:style("A1:C5", "Error")
sheet:style("A1:C5", "warn")      -- alias for Warning
sheet:style("A1", "ok")           -- alias for Success
sheet:style("A1", "default")      -- clears style

-- Integer constant via the styles table
sheet:style("A1:C5", styles.Warning)
sheet:style("A1", styles.Default)

-- ============================================================================
-- Constants reference
-- ============================================================================
-- styles.Default  = 0
-- styles.Error    = 1
-- styles.Warning  = 2
-- styles.Success  = 3
-- styles.Input    = 4
-- styles.Total    = 5
-- styles.Note     = 6

print("Applied semantic styles")
