//! Canonical Lua script examples for the console.
//!
//! These examples demonstrate the power of VisiGrid's scripting capabilities.
//! Each script is designed to be immediately useful and educational.

/// A single example script
#[derive(Debug, Clone)]
pub struct LuaExample {
    /// Short name for the menu
    pub name: &'static str,
    /// One-line description
    pub description: &'static str,
    /// The actual Lua code
    pub code: &'static str,
}

/// All canonical examples
pub const EXAMPLES: &[LuaExample] = &[
    LuaExample {
        name: "Fill Series",
        description: "Fill column A with a doubling sequence (1, 2, 4, 8...)",
        code: r#"-- Fill A1:A10 with powers of 2
for i = 1, 10 do
    sheet:set("A" .. i, 2 ^ (i - 1))
end
print("Filled A1:A10 with powers of 2")"#,
    },
    LuaExample {
        name: "Trim Whitespace",
        description: "Remove leading/trailing spaces from column A",
        code: r#"-- Trim whitespace from A1:A20
local trimmed = 0
for i = 1, 20 do
    local val = sheet:get("A" .. i)
    if type(val) == "string" then
        local clean = val:match("^%s*(.-)%s*$")
        if clean ~= val then
            sheet:set("A" .. i, clean)
            trimmed = trimmed + 1
        end
    end
end
print("Trimmed " .. trimmed .. " cells")"#,
    },
    LuaExample {
        name: "Find Duplicates",
        description: "Find duplicate values in column A and list them",
        code: r#"-- Find duplicates in A1:A50
local seen = {}
local dupes = {}

for i = 1, 50 do
    local val = sheet:get("A" .. i)
    if val ~= nil then
        local key = tostring(val)
        if seen[key] then
            if not dupes[key] then
                dupes[key] = true
                print("Duplicate: " .. key .. " (first at row " .. seen[key] .. ")")
            end
        else
            seen[key] = i
        end
    end
end

local count = 0
for _ in pairs(dupes) do count = count + 1 end
print("Found " .. count .. " duplicate values")"#,
    },
    LuaExample {
        name: "Normalize Dates",
        description: "Convert various date formats to YYYY-MM-DD",
        code: r#"-- Normalize dates in A1:A20 to YYYY-MM-DD format
local patterns = {
    -- MM/DD/YYYY
    { pat = "^(%d%d?)/(%d%d?)/(%d%d%d%d)$", fmt = function(m,d,y) return string.format("%s-%02d-%02d", y, m, d) end },
    -- DD-MM-YYYY
    { pat = "^(%d%d?)-(%d%d?)-(%d%d%d%d)$", fmt = function(d,m,y) return string.format("%s-%02d-%02d", y, m, d) end },
    -- YYYY/MM/DD
    { pat = "^(%d%d%d%d)/(%d%d?)/(%d%d?)$", fmt = function(y,m,d) return string.format("%s-%02d-%02d", y, m, d) end },
}

local converted = 0
for i = 1, 20 do
    local val = sheet:get("A" .. i)
    if type(val) == "string" then
        for _, p in ipairs(patterns) do
            local a, b, c = val:match(p.pat)
            if a then
                sheet:set("A" .. i, p.fmt(tonumber(a), tonumber(b), tonumber(c)))
                converted = converted + 1
                break
            end
        end
    end
end
print("Converted " .. converted .. " dates to YYYY-MM-DD")"#,
    },
    LuaExample {
        name: "Compare Columns",
        description: "Find mismatches between columns A and B",
        code: r#"-- Compare A1:A20 with B1:B20, report mismatches
local mismatches = 0

for i = 1, 20 do
    local a = sheet:get("A" .. i)
    local b = sheet:get("B" .. i)

    -- Convert to comparable form
    local a_str = a == nil and "" or tostring(a)
    local b_str = b == nil and "" or tostring(b)

    if a_str ~= b_str then
        mismatches = mismatches + 1
        print(string.format("Row %d: A='%s' vs B='%s'", i, a_str, b_str))
    end
end

if mismatches == 0 then
    print("All values match!")
else
    print(string.format("\n%d mismatches found", mismatches))
end"#,
    },
    LuaExample {
        name: "Generate Multiplication Table",
        description: "Create a 10x10 multiplication table starting at A1",
        code: r#"-- Generate 10x10 multiplication table
for row = 1, 10 do
    for col = 1, 10 do
        sheet:set_value(row, col, row * col)
    end
end
print("Created 10x10 multiplication table")"#,
    },
    LuaExample {
        name: "Sum Column",
        description: "Calculate sum of numbers in column A",
        code: r#"-- Sum all numbers in A1:A100
local sum = 0
local count = 0

for i = 1, 100 do
    local val = sheet:get("A" .. i)
    if type(val) == "number" then
        sum = sum + val
        count = count + 1
    end
end

print(string.format("Sum: %g (%d values)", sum, count))
print(string.format("Average: %g", count > 0 and sum / count or 0))"#,
    },
];

/// Get all example names for menu display
pub fn example_names() -> Vec<&'static str> {
    EXAMPLES.iter().map(|e| e.name).collect()
}

/// Get an example by index
pub fn get_example(index: usize) -> Option<&'static LuaExample> {
    EXAMPLES.get(index)
}

/// Get an example by name
pub fn find_example(name: &str) -> Option<&'static LuaExample> {
    EXAMPLES.iter().find(|e| e.name == name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_examples_exist() {
        assert!(EXAMPLES.len() >= 5, "Should have at least 5 examples");
    }

    #[test]
    fn test_examples_valid() {
        for example in EXAMPLES {
            assert!(!example.name.is_empty(), "Example name should not be empty");
            assert!(!example.description.is_empty(), "Example description should not be empty");
            assert!(!example.code.is_empty(), "Example code should not be empty");
            // Basic syntax check - should start with comment or code
            assert!(
                example.code.starts_with("--") || example.code.starts_with("local") || example.code.starts_with("for"),
                "Example should start with comment or code: {}", example.name
            );
        }
    }
}
