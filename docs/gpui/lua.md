# Lua Scripting: VisiGrid

A transactional, sandboxed scripting system for spreadsheet automation via `Ctrl+Shift+L`.

---

## Status: IMPLEMENTED (v1 Stable)

The Lua Console is a platform primitive: sandboxed, deterministic, fully undoable.

---

## Why Lua?

| Benefit | Details |
|---------|---------|
| Tiny runtime | ~200KB vs ~30MB Python |
| Fast startup | <1ms vs ~100ms Python |
| Designed for embedding | Clean FFI, sandboxing |
| No dependency hell | Single static library |

---

## Architecture

```
┌─────────────────────────────────────────────┐
│  Spreadsheet                                │
│    │                                        │
│    ├── LuaRuntime (sandboxed)               │
│    │     └── "sheet" global → DynOpSink     │
│    │                                        │
│    ├── Operation Journal                    │
│    │     ├── Vec<LuaOp> (batched edits)     │
│    │     └── Pending shadow map (R-A-W)     │
│    │                                        │
│    └── Console Panel                        │
│          ├── Output (virtual scroll)        │
│          ├── Input (multiline)              │
│          └── Examples system                │
└─────────────────────────────────────────────┘
```

**Execution model:** Lua calls queue operations into a journal. The pending map handles read-after-write. After the chunk finishes, ops are applied as a batch with a single undo entry.

---

## Sheet API (v1 Stable)

All coordinates are **1-indexed** for Lua convention.

### Cell Access

```lua
-- Read/write by row, col
sheet:get_value(row, col)      -- → value or nil
sheet:set_value(row, col, val) -- sets cell value
sheet:get_formula(row, col)    -- → formula string or nil
sheet:set_formula(row, col, f) -- sets formula

-- Read/write by A1 notation (shorthand)
sheet:get("A1")                -- → value at A1
sheet:set("A1", val)           -- sets value at A1
sheet:get_a1("B2")             -- same as get()
sheet:set_a1("B2", val)        -- same as set()
```

### Sheet Info

```lua
sheet:rows()      -- → number of rows with data
sheet:cols()      -- → number of columns with data
sheet:selection() -- → {start_row, start_col, end_row, end_col, range}
```

**Selection example:**
```lua
local sel = sheet:selection()
print(sel.range)      -- "A1:C5"
print(sel.start_row)  -- 1
print(sel.end_col)    -- 3
```

### Range Operations (Bulk Read/Write)

```lua
local r = sheet:range("A1:C5")

-- Read all values as 2D table
local data = r:values()
print(data[1][1])  -- value at A1
print(data[2][3])  -- value at C2

-- Write 2D table to range
r:set_values({
    {1, 2, 3},
    {4, 5, 6},
    {7, 8, 9}
})

-- Range info
r:rows()     -- → 5
r:cols()     -- → 3
r:address()  -- → "A1:C5"
```

### Transactions

```lua
sheet:begin()    -- start transaction (noop, for clarity)
sheet:rollback() -- discard pending changes, returns count
sheet:commit()   -- commit changes (noop, auto-committed at script end)
```

**Rollback example:**
```lua
sheet:set("A1", 100)
sheet:set("A2", 200)
local discarded = sheet:rollback()  -- → 2
-- No changes will be applied
```

---

## Console Commands

Type these directly in the console input:

| Command | Description |
|---------|-------------|
| `help` | Show all commands and shortcuts |
| `examples` | List available example scripts |
| `example N` | Load example N (or name) into input |
| `clear` | Clear output history |

### First-Open Experience

On first open, the console pre-fills `examples` in the input. Press Enter to see the list of available scripts.

---

## Built-in Examples

| Name | Description |
|------|-------------|
| Fill Series | Fill A1:A10 with powers of 2 |
| Trim Whitespace | Remove leading/trailing spaces from column A |
| Find Duplicates | Find and report duplicate values in column A |
| Normalize Dates | Convert date formats to YYYY-MM-DD |
| Compare Columns | Find mismatches between columns A and B |
| Generate Multiplication Table | Create 10x10 table starting at A1 |
| Sum Column | Calculate sum and average of numbers in column A |

Load any example: `example 1` or `example "Fill Series"`

---

## Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Toggle console | `Ctrl+Shift+L` |
| Execute | `Enter` or `Ctrl+Enter` |
| Newline (multiline) | `Shift+Enter` |
| Previous command | `Up Arrow` |
| Next command | `Down Arrow` |
| Scroll output up | `Page Up` |
| Scroll output down | `Page Down` |
| Scroll to top | `Ctrl+Home` |
| Scroll to bottom | `Ctrl+End` |
| Clear output | `Ctrl+L` |
| Close console | `Escape` |

---

## Sandboxing

### Disabled Libraries

| Library | Status | Reason |
|---------|--------|--------|
| `table` | Enabled | Safe |
| `string` | Enabled | Safe |
| `math` | Enabled | Safe |
| `utf8` | Enabled | Safe |
| `os` | **Disabled** | System calls |
| `io` | **Disabled** | File I/O |
| `debug` | **Disabled** | Introspection |
| `package` | **Disabled** | Module loading |
| `require` | **Disabled** | Module loading |
| `load` | **Disabled** | Bytecode execution |
| `loadfile` | **Disabled** | File loading |
| `dofile` | **Disabled** | File execution |

### Resource Limits

| Limit | Value | Purpose |
|-------|-------|---------|
| Operations | 1,000,000 | Prevent runaway scripts |
| Output lines | 5,000 | Prevent memory exhaustion |
| Instructions | 100,000,000 | Prevent infinite loops |
| Wall-clock | 30 seconds | Catch pathological patterns |

### Safety Guarantees

- **No filesystem access** - Scripts cannot read or write files
- **No network access** - Scripts cannot make HTTP requests
- **No system calls** - Scripts cannot execute commands
- **Single undo** - All changes from one script = one Ctrl+Z
- **Deterministic** - Same input always produces same output

---

## Execution Stats

After each script, the console shows:
```
ops: 100 | cells: 50 | time: 12.5ms
```

- **ops**: Total operations queued
- **cells**: Unique cells modified (deduplicated)
- **time**: Wall-clock execution time

---

## Example Scripts

### Fill Column with Sequence
```lua
for i = 1, 100 do
    sheet:set("A" .. i, i * 2)
end
print("Filled 100 cells")
```

### Process Selection
```lua
local sel = sheet:selection()
for row = sel.start_row, sel.end_row do
    for col = sel.start_col, sel.end_col do
        local val = sheet:get_value(row, col)
        if type(val) == "number" then
            sheet:set_value(row, col, val * 1.1)  -- +10%
        end
    end
end
```

### Bulk Transform with Range
```lua
local r = sheet:range("A1:C10")
local data = r:values()

-- Double all numbers
for i, row in ipairs(data) do
    for j, val in ipairs(row) do
        if type(val) == "number" then
            data[i][j] = val * 2
        end
    end
end

r:set_values(data)
```

### Conditional Rollback
```lua
sheet:begin()

for i = 1, 100 do
    sheet:set("A" .. i, math.random(1, 100))
end

-- Check if sum exceeds threshold
local sum = 0
for i = 1, 100 do
    sum = sum + (sheet:get("A" .. i) or 0)
end

if sum > 5000 then
    local discarded = sheet:rollback()
    print("Rolled back " .. discarded .. " ops (sum too high)")
else
    print("Sum: " .. sum)
end
```

---

## Design Decisions

### Why 1-indexed?
Spreadsheet users think in A1, not (0,0). The Lua API matches user mental model.

### Why command queue?
1. No borrow panics at runtime
2. One recalc at end of script
3. Clean undo grouping
4. Read-after-write consistency via pending map

### Why typed values?
Returning display strings forces parsing `"$1,234.56"` back to numbers. Typed values are unambiguous.

### Why no persistent scripts?
Trust. Every script is explicit, visible, and immediately undoable. No "run on open" surprises.

---

## Dependencies

```toml
# Cargo.toml
mlua = { version = "0.11", features = ["lua54"] }
```

---

## API Stability

The Sheet API documented here is **v1 Stable**. Breaking changes require major version bumps.

| API | Status |
|-----|--------|
| `sheet:get/set` | Stable |
| `sheet:get_value/set_value` | Stable |
| `sheet:get_formula/set_formula` | Stable |
| `sheet:rows/cols` | Stable |
| `sheet:selection` | Stable |
| `sheet:range` | Stable |
| `range:values/set_values` | Stable |
| `sheet:begin/rollback/commit` | Stable |
