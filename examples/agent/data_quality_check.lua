-- VisiGrid Agent Demo: Data Quality Check with Semantic Styles
-- Scans a data range, validates values, and marks cells with
-- semantic styles so the user sees intent at a glance.
--
-- Run in the Lua console (sheet:* API).
-- Assumes data in A1:C20 with headers in row 1.

-- Mark headers
sheet:note("A1:C1")

local errors = 0
local warnings = 0

for i = 2, 20 do
    local name = sheet:get("A" .. i)
    local amount = sheet:get("B" .. i)
    local status = sheet:get("C" .. i)

    -- Skip empty rows
    if name == nil and amount == nil then
        goto continue
    end

    -- Missing name = error
    if name == nil or name == "" then
        sheet:error("A" .. i)
        errors = errors + 1
    end

    -- Negative amount = warning, zero = error
    if type(amount) == "number" then
        if amount < 0 then
            sheet:warning("B" .. i)
            warnings = warnings + 1
        elseif amount == 0 then
            sheet:error("B" .. i)
            errors = errors + 1
        else
            sheet:success("B" .. i)
        end
    elseif amount ~= nil then
        -- Non-numeric in amount column
        sheet:error("B" .. i)
        errors = errors + 1
    end

    -- Status column: mark "Pending" as input (editable)
    if status == "Pending" then
        sheet:input("C" .. i)
    end

    ::continue::
end

print(string.format("Scan complete: %d errors, %d warnings", errors, warnings))
