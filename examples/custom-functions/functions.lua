-- Custom functions for VisiGrid
-- Install: cp functions.lua ~/.config/visigrid/functions.lua
-- Reload: Command palette → "Reload Custom Functions"

-- ACCRUED_INTEREST: Simple day-count interest (ACT/365)
--   principal × rate × days / 365
function ACCRUED_INTEREST(principal, rate, days)
    if principal == nil or rate == nil or days == nil then
        error("requires 3 arguments: principal, rate, days")
    end
    return principal * rate * days / 365
end

-- HAIRCUT: Apply a risk haircut to a position value
--   value × (1 - haircut_pct)
function HAIRCUT(value, haircut_pct)
    if value == nil or haircut_pct == nil then
        error("requires 2 arguments: value, haircut_pct")
    end
    return value * (1 - haircut_pct)
end

-- WEIGHTED_AVG: Weighted average over two ranges
--   Σ(values[i] × weights[i]) / Σ(weights[i])
function WEIGHTED_AVG(values, weights)
    if values == nil or weights == nil then
        error("requires 2 range arguments")
    end
    local sum = 0
    local weight_sum = 0
    for i = 1, values.n do
        local v = values:get(i)
        local w = weights:get(i)
        if v ~= nil and w ~= nil then
            sum = sum + v * w
            weight_sum = weight_sum + w
        end
    end
    if weight_sum == 0 then
        error("total weight is zero")
    end
    return sum / weight_sum
end
