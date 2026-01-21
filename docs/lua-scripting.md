# Lua Scripting REPL

Integrated scripting console using Lua, accessible via `Ctrl+Shift+L`.

---

## Overview

VisiGrid embeds Lua for automating spreadsheet tasks. Unlike hidden Excel macros, the REPL provides visible, debuggable scripting.

**Why Lua?**
- ~200KB runtime (vs ~30MB Python)
- <1ms startup (vs ~100ms Python)
- Designed for embedding
- No dependency hell

---

## Architecture

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

**Execution model:** Lua calls queue operations (SetValue, SetFormula, Select). After the chunk finishes, ops are applied as a batch on the UI thread. This avoids borrow panics and enables single-undo grouping.

---

## Lua API

### Cell Access (Typed Values)

```lua
-- Get typed value (not display string!)
-- Returns: nil, number, string, boolean, or error table
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

-- Set formula (separate from value)
sheet:set_formula(row, col, formula)
sheet:set_formula(1, 5, "=A1*2")
```

### A1-Style Addressing (Recommended)

```lua
-- More natural for spreadsheet users
sheet:get_a1("A1")                 -- get_value with A1 notation
sheet:set_a1("B2", 100)            -- set_value with A1 notation
sheet:set_formula_a1("C3", "=A1+B2")

-- Ranges
local r = sheet:range("A1:C10")
r:values()                         -- 2D table of typed values
r:set_values({{1,2,3}, {4,5,6}})   -- bulk set
r:map(function(v) return v * 2 end) -- transform
```

### Value Types

| Lua Type | Example | Notes |
|----------|---------|-------|
| `nil` | `nil` | Empty cell |
| `number` | `42`, `3.14` | Numeric value |
| `string` | `"Hello"` | Text value |
| `boolean` | `true`, `false` | Boolean value |
| `table` (error) | `{kind="DIV0", message="Division by zero"}` | Cell error |

### Sheet Info

```lua
sheet:rows()           -- number of rows
sheet:cols()           -- number of columns
sheet:name()           -- sheet name
```

### Selection

```lua
-- Get current selection
local sel = sheet:selection()
-- Returns: { start_row, start_col, end_row, end_col } (1-indexed)

-- Set selection
sheet:select(r1, c1, r2, c2)
sheet:select(1, 1, 10, 1)   -- Select A1:A10

-- A1 style
sheet:select_a1("A1:C10")
```

### App Commands (Future)

```lua
-- Execute VisiGrid commands
app:execute("edit.copy")
app:execute("navigate.goto", "A1")

-- Show status message
app:message("Done!")

-- Prompt for input (future)
local name = app:prompt("Enter name:")
```

---

## Examples

### Basic Arithmetic
```lua
> 1 + 1
2

> math.sqrt(144)
12

> string.upper("hello")
HELLO

> _              -- last result
HELLO
```

### Cell Operations
```lua
> sheet:get_a1("A1")
42

> sheet:set_a1("A1", 100)
nil

> sheet:get_a1("A1")
100

> type(sheet:get_a1("A1"))
number
```

### Bulk Operations
```lua
-- Fill column A with row numbers (1-indexed)
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

### Range Operations (Efficient)
```lua
-- Get all values in range as 2D table
local r = sheet:range("A1:C10")
local data = r:values()

-- Transform and write back
r:map(function(v)
  if type(v) == "number" then
    return v * 2
  end
  return v
end)

-- Bulk set from table
sheet:range("D1:F3"):set_values({
  {1, 2, 3},
  {4, 5, 6},
  {7, 8, 9}
})
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

## REPL Panel UI

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
│ > invalid code                              │
│ Error: syntax error near 'code'             │
├─────────────────────────────────────────────┤
│ > _                                         │
└─────────────────────────────────────────────┘
```

**Input modes:**
- `Enter` - Execute when chunk is complete
- `Shift+Enter` - Insert newline (multiline input)
- `Esc` - Cancel running script

---

## Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Toggle REPL | `Ctrl+Shift+L` |
| Execute code | `Enter` (when complete) |
| Newline | `Shift+Enter` |
| Previous command | `Up Arrow` |
| Next command | `Down Arrow` |
| Cancel execution | `Escape` |
| Clear history | Type `clear` |

**Special variables:**
- `_` - Last expression result
- `ans` - Alias for `_`

---

## Sandboxing

### Disabled Libraries

| Library | Status | Reason |
|---------|--------|--------|
| `table` | Enabled | Safe table manipulation |
| `string` | Enabled | Safe string functions |
| `math` | Enabled | Safe math functions |
| `utf8` | Enabled | Unicode support |
| `os` | **Disabled** | System calls |
| `io` | **Disabled** | File I/O |
| `debug` | **Disabled** | Introspection |
| `loadfile` | **Disabled** | Arbitrary code execution |
| `dofile` | **Disabled** | Arbitrary code execution |
| `require` | **Disabled** | Module loading |

### Resource Limits (Mandatory)

| Limit | Value | Purpose |
|-------|-------|---------|
| Memory | 10MB | Prevent huge allocations |
| Instructions | 10M | Prevent infinite loops |
| Wall-clock | 200ms interactive, 5s batch | UI responsiveness |
| Recursion depth | 200 | Prevent stack overflow |
| Output size | 1MB | Prevent REPL spam |

**Cancel:** `Esc` key aborts running Lua execution.

---

## Execution Model

### Command Queue (Not Live Mutation)

Lua doesn't directly mutate the sheet. Instead:

```rust
enum LuaOp {
    SetValue { row: usize, col: usize, value: LuaValue },
    SetFormula { row: usize, col: usize, formula: String },
    Select { r1: usize, c1: usize, r2: usize, c2: usize },
}

struct CommandQueue {
    ops: Vec<LuaOp>,
}
```

1. Lua calls `sheet:set_value(1, 1, 42)`
2. SheetApi pushes `SetValue{row:0, col:0, value:42}` to queue
3. After chunk completes, App applies all ops as batch
4. Single recalc at end
5. Single undo entry: "Lua: modified 120 cells"

**Benefits:**
- No `Rc<RefCell<>>` borrow panics
- One recalc per script (not per cell)
- Clean undo grouping
- Future: async scripts, background jobs

### Reads vs Writes

- **Reads** (`get_value`, `get_display`) read from current sheet state
- **Writes** (`set_value`, `set_formula`) queue operations
- Writes within same script see queued values (shadow state)

---

## Undo Integration

All Lua edits are grouped as a single undo step:

```
Undo: "Lua script (modified 120 cells)"
```

Status bar shows: `Lua: modified 120 cells in 45ms`

---

## Implementation Details

### Dependencies

```toml
# Cargo.toml
mlua = { version = "0.11", features = ["lua54"] }
```

### Files

| File | Purpose |
|------|---------|
| `src/scripting/mod.rs` | Module root |
| `src/scripting/lua_runtime.rs` | Sandboxed Lua runtime |
| `src/scripting/sheet_api.rs` | SheetApi UserData |
| `src/scripting/command_queue.rs` | Batched operations |
| `src/app.rs` | REPL UI panel, apply ops |
| `src/config/keybindings.rs` | Ctrl+` binding |

### Key Types

```rust
// Value types exposed to Lua
pub enum LuaValue {
    Nil,
    Number(f64),
    String(String),
    Boolean(bool),
    Error { kind: String, message: String },
}

// Queued operation
pub enum LuaOp {
    SetValue { row: usize, col: usize, value: LuaValue },
    SetFormula { row: usize, col: usize, formula: String },
    Select { r1: usize, c1: usize, r2: usize, c2: usize },
}

// Command queue
pub struct CommandQueue {
    ops: Vec<LuaOp>,
    shadow: HashMap<(usize, usize), LuaValue>, // for reads-after-writes
}

// REPL history entry
pub struct ReplEntry {
    pub input: String,
    pub output: String,
    pub is_error: bool,
    pub cells_modified: usize,
    pub duration_ms: u64,
}
```

---

## v1 Priorities

**Must have:**
- [x] Split `get_value` vs `get_display` (typed values)
- [x] A1 helpers (`get_a1`, `set_a1`, `range`)
- [x] Instruction limit + wall-clock timeout + Esc cancel
- [x] Command queue (batch ops, not live mutation)
- [x] Multiline REPL (Shift+Enter)
- [x] 1-indexed API (matches spreadsheet convention)
- [x] Single undo step per script

**Nice to have (v1.1):**
- [ ] `range:map(fn)` for efficient transforms
- [ ] `app:execute(cmd)` for VisiGrid commands
- [ ] Run script against selection
- [ ] Bind script to shortcut

**Future (v2):**
- [ ] Script files (`.lua` per workbook)
- [ ] Syntax highlighting
- [ ] Autocomplete
- [ ] Background/async scripts
- [ ] Date type support

---

## Design Decisions

### Why 1-indexed?

Spreadsheet users think in A1, not (0,0). The internal 0-indexed representation is an implementation detail. The Lua API should match user mental model.

### Why command queue instead of Rc<RefCell<>>?

1. No borrow panics at runtime
2. One recalc at end of script (performance)
3. Clean undo grouping
4. Enables future async/background scripts
5. Reads can see pending writes (shadow state)

### Why instruction limits are mandatory?

Memory limits alone don't prevent `while true do end`. Without instruction limits, any infinite loop hangs the UI with no recovery except force-quit.

### Why typed values instead of strings?

Returning display strings forces users to parse `"$1,234.56"` back to numbers. This creates bugs when:
- Locale changes decimal separator
- User changes number format
- Percentages, currencies, dates

Typed values (`number`, `string`, `boolean`, `nil`, error table) are unambiguous.
