# Lua Scripting: VisiGrid gpui

Planned scripting system using embedded Lua, accessible via `Ctrl+`` (backtick).

---

## Status: NOT YET IMPLEMENTED

Lua scripting is a future feature. This document outlines the design for when it's built.

---

## Why Lua?

| Benefit | Details |
|---------|---------|
| Tiny runtime | ~200KB vs ~30MB Python |
| Fast startup | <1ms vs ~100ms Python |
| Designed for embedding | Clean FFI, sandboxing |
| No dependency hell | Single static library |

---

## Planned Architecture

```
┌─────────────────────────────────────────────┐
│  App                                        │
│    │                                        │
│    ├── LuaRuntime (sandboxed)               │
│    │     ├── "sheet" global → SheetApi      │
│    │     └── "app" global → AppApi          │
│    │                                        │
│    ├── CommandQueue                         │
│    │     └── Vec<LuaOp> (batched edits)     │
│    │                                        │
│    └── REPL Panel                           │
│          ├── Output history (scrollable)    │
│          └── Input prompt (multiline)       │
└─────────────────────────────────────────────┘
```

**Execution model:** Lua calls queue operations. After the chunk finishes, ops are applied as a batch on the UI thread. This avoids borrow panics and enables single-undo grouping.

---

## Planned API

### Cell Access (Typed Values)

```lua
-- Get typed value (not display string!)
sheet:get_value(row, col)
sheet:get_value(1, 1)      -- A1 (1-indexed!)

-- Get formatted display string
sheet:get_display(row, col)

-- Get raw formula (or nil if not a formula)
sheet:get_formula(row, col)

-- Set typed value
sheet:set_value(row, col, value)
sheet:set_value(1, 1, 42)          -- number
sheet:set_value(1, 2, "Hello")     -- string
sheet:set_value(1, 3, true)        -- boolean
sheet:set_value(1, 4, nil)         -- clear cell

-- Set formula
sheet:set_formula(row, col, formula)
sheet:set_formula(1, 5, "=A1*2")
```

### A1-Style Addressing

```lua
sheet:get_a1("A1")
sheet:set_a1("B2", 100)
sheet:set_formula_a1("C3", "=A1+B2")

-- Ranges
local r = sheet:range("A1:C10")
r:values()                         -- 2D table
r:set_values({{1,2,3}, {4,5,6}})   -- bulk set
r:map(function(v) return v * 2 end) -- transform
```

### Selection

```lua
local sel = sheet:selection()  -- {start_row, start_col, end_row, end_col}
sheet:select(1, 1, 10, 1)      -- Select A1:A10
sheet:select_a1("A1:C10")
```

---

## Example Use Cases

### Bulk Operations
```lua
-- Fill column A with row numbers
for i = 1, 10 do
  sheet:set_value(i, 1, i)
end

-- Sum values in column A
local sum = 0
for i = 1, 10 do
  local v = sheet:get_value(i, 1)
  if type(v) == "number" then
    sum = sum + v
  end
end
print(sum)  -- 55
```

### Data Transformation
```lua
-- Uppercase all text in column B
for i = 1, sheet:rows() do
  local val = sheet:get_value(i, 2)
  if type(val) == "string" then
    sheet:set_value(i, 2, string.upper(val))
  end
end
```

---

## REPL Panel Design

```
┌─────────────────────────────────────────────┐
│ Lua REPL                                  × │
├─────────────────────────────────────────────┤
│ > sheet:get_a1("A1")                        │
│ 42                                          │
│ > for i = 1, 10 do                          │
│ ...   sheet:set_value(i, 1, i)              │
│ ... end                                     │
│ nil                                         │
├─────────────────────────────────────────────┤
│ > _                                         │
└─────────────────────────────────────────────┘
```

**Input modes:**
- `Enter` - Execute when chunk complete
- `Shift+Enter` - Insert newline (multiline)
- `Esc` - Cancel running script

---

## Keyboard Shortcuts (Planned)

| Action | Shortcut |
|--------|----------|
| Toggle REPL | `Ctrl+`` (backtick) |
| Execute code | `Enter` |
| Newline | `Shift+Enter` |
| Previous command | `Up Arrow` |
| Next command | `Down Arrow` |
| Cancel execution | `Escape` |
| Clear history | Type `clear` |

---

## Sandboxing (Required)

### Disabled Libraries

| Library | Status | Reason |
|---------|--------|--------|
| `table` | Enabled | Safe |
| `string` | Enabled | Safe |
| `math` | Enabled | Safe |
| `os` | **Disabled** | System calls |
| `io` | **Disabled** | File I/O |
| `debug` | **Disabled** | Introspection |
| `require` | **Disabled** | Module loading |

### Resource Limits

| Limit | Value | Purpose |
|-------|-------|---------|
| Memory | 10MB | Prevent huge allocations |
| Instructions | 10M | Prevent infinite loops |
| Wall-clock | 200ms interactive | UI responsiveness |
| Recursion | 200 | Prevent stack overflow |

---

## Implementation Priority

### Phase 1: Basic REPL
1. Toggle REPL panel with Ctrl+`
2. Execute Lua expressions
3. sheet:get_value / set_value
4. Command queue (batch ops)
5. Single undo step per script

### Phase 2: Range Operations
1. A1-style addressing
2. Range object with map/set_values
3. Multiline input
4. Command history

### Phase 3: Integration
1. app:execute() for VisiGrid commands
2. Bind scripts to shortcuts
3. Script files per workbook
4. Syntax highlighting

---

## Dependencies

```toml
# Cargo.toml (when implemented)
mlua = { version = "0.11", features = ["lua54"] }
```

---

## Design Decisions

### Why 1-indexed?
Spreadsheet users think in A1, not (0,0). The Lua API matches user mental model.

### Why command queue?
1. No borrow panics at runtime
2. One recalc at end of script
3. Clean undo grouping
4. Future async/background scripts

### Why typed values?
Returning display strings forces parsing `"$1,234.56"` back to numbers. Typed values are unambiguous.
