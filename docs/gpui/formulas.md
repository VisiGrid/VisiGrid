# VisiGrid Formula System (gpui)

> **This document defines formula UI behavior for gpui. The formula engine is shared with the iced version and is complete. This spec covers the UI layer that must be rebuilt.**
>
> **If observed behavior differs from this document, the behavior is a bug.**

---

## Current State

### What Works

| Component | Status | Location |
|-----------|--------|----------|
| Formula engine | Complete | `crates/engine/src/formula/` |
| Parser (tokenizer, AST) | Complete | `parser.rs` |
| Evaluator (96 functions) | Complete | `eval.rs` |
| Cell references (A1, $A$1) | Complete | parser + eval |
| Range references (A1:B10) | Complete | parser + eval |
| Dependency tracking | Complete | `extract_cell_refs()` |
| Array formulas with spill | Complete | eval.rs |
| Formula entry in formula bar | Working | `views/formula_bar.rs` |
| Formula evaluation on Enter | Working | `app.rs` |
| Error display (#DIV/0!, etc.) | Working | cell rendering |

### What's Missing (This Spec)

| Feature | Priority | Section |
|---------|----------|---------|
| Formula context analyzer | P0 | [Context Analyzer](#formula-context-analyzer) |
| Function autocomplete | P1 | [Autocomplete](#autocomplete-popup) |
| Signature help | P1 | [Signature Help](#signature-help) |
| Error banner | P1 | [Error Highlighting](#error-highlighting) |
| F4 cycle reference type | P1 | [Quick Actions](#quick-actions) |
| Syntax highlighting | P2 | [Syntax Highlighting](#syntax-highlighting) |
| Hover docs | P2 | [Hover Documentation](#hover-documentation) |
| Alt+= AutoSum | P2 | [Quick Actions](#quick-actions) |
| Spill visualization | P3 | [Array Visualization](#array-visualization) |
| Ctrl+` formula view | P3 | [Formula View](#formula-view-toggle) |

---

## Formula Context Analyzer

> **This is the foundation. All formula UI features route through one canonical analyzer.**

### The Problem

Without a unified context model, each feature (autocomplete, signature help, F4 cycling, highlighting, error squiggles) will parse the formula differently, leading to:
- Inconsistent behavior at edge cases
- Duplicate parsing logic
- Bugs where features disagree about cursor position

### The Solution

One structure, one function, used by everything:

```rust
/// Describes the editing context at a specific cursor position within a formula.
pub struct FormulaContext {
    /// What kind of position the cursor is in
    pub mode: FormulaEditMode,

    /// Cursor position (char index from start of formula, including '=')
    /// NOTE: This is char index, not byte offset. See "Cursor Indexing Policy".
    pub cursor: usize,

    /// If inside a function's argument list, which function
    pub current_function: Option<&'static FunctionInfo>,

    /// If inside argument list, which argument (0-indexed)
    pub current_arg_index: Option<usize>,

    /// The token spanning the cursor position, if any
    pub token_at_cursor: Option<TokenSpan>,

    /// The primary span for operations (F4 cycles this, hover attaches here, squiggles here)
    /// This is what all features should default to operating on.
    pub primary_span: Option<Range<usize>>,

    /// What range autocomplete should replace when accepting a suggestion
    pub replace_range: Option<Range<usize>>,

    /// Nesting depth of parentheses at cursor
    pub paren_depth: usize,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FormulaEditMode {
    /// Right after '=' with nothing typed
    Start,
    /// Typing a function name (or could be cell ref)
    Identifier,
    /// Inside a function's argument list
    ArgList,
    /// Inside a string literal
    String,
    /// On a cell reference or range
    Reference,
    /// On an operator (+, -, *, /, &, <, >, =)
    Operator,
    /// On a number literal
    Number,
    /// After a complete expression (e.g., after closing paren)
    Complete,
}

#[derive(Debug, Clone)]
pub struct TokenSpan {
    pub token_type: TokenType,
    pub range: Range<usize>,
    pub text: String,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TokenType {
    Function,
    CellRef,
    Range,
    NamedRange,
    Number,
    String,
    Boolean,
    Operator,
    Comparison,
    Paren,
    Comma,
    Colon,
    Error,
    // Structural tokens (needed for correct parsing, may not be highlighted)
    Whitespace,     // Spaces/tabs between tokens
    Bang,           // '!' for sheet references (Sheet1!A1)
    Percent,        // '%' suffix (10%)
    UnaryMinus,     // Leading minus (distinguished from subtraction)
}
```

### The Analyzer

```rust
/// Analyze formula at cursor position. This is the single source of truth
/// for all formula UI features.
pub fn analyze(formula: &str, cursor: usize) -> FormulaContext {
    // 1. Tokenize entire formula
    // 2. Find token at cursor (or gap between tokens)
    // 3. Determine mode based on token type and surroundings
    // 4. If in ArgList, walk back to find function name and count commas
    // 5. Compute replace_range for autocomplete
}
```

### Invariants

1. **Single source of truth**: All UI features call `analyze()`, never parse independently
2. **Cursor always valid**: If cursor > formula.len(), clamp to end
3. **Mode is unambiguous**: Every cursor position maps to exactly one mode
4. **Nested functions handled**: `current_function` is the innermost function containing cursor
5. **Function names normalized**: Internal names are UPPERCASE, matching is case-insensitive

### Function Name Normalization

To prevent "SUM" vs "sum" inconsistencies:

- **Internal storage**: All function names in `FunctionInfo` are UPPERCASE
- **Display**: Show as UPPERCASE (matches Excel convention)
- **Matching**: Case-insensitive (user types "sum", matches "SUM")
- **Enforcement**: `analyze()` normalizes identifiers before lookup

```rust
fn normalize_function_name(name: &str) -> String {
    name.to_ascii_uppercase()
}
```

---

## Cursor Indexing Policy

> **Rule:** UI cursor positions are char indices. Engine tokenizer uses byte offsets. A conversion layer bridges them.

### The Problem

Rust strings are UTF-8, where `"é".len() == 2` bytes but is 1 char. If you mix byte offsets with char indices:
- Highlighting spans will be wrong
- Cursor positioning will jump unexpectedly
- Mid-character slicing will panic

### The Solution

| Layer | Index Type | Why |
|-------|------------|-----|
| UI (cursor, selection) | Char index | User thinks in characters |
| Engine (tokenizer) | Byte offset | Rust string slicing |
| Conversion | `char_to_byte()` / `byte_to_char()` | Bridge layer |

```rust
/// Convert char index to byte offset
fn char_to_byte(s: &str, char_idx: usize) -> usize {
    s.char_indices()
        .nth(char_idx)
        .map(|(i, _)| i)
        .unwrap_or(s.len())
}

/// Convert byte offset to char index
fn byte_to_char(s: &str, byte_idx: usize) -> usize {
    s[..byte_idx].chars().count()
}
```

### Invariants

1. **All `FormulaContext` positions are char indices**
2. **All `TokenSpan.range` values are char indices** (converted from engine)
3. **Engine tokenizer may use bytes internally**, but `analyze()` converts before returning
4. **Emojis may behave oddly** (multi-codepoint), but won't crash — acceptable for v1

---

## UI State Model

> **Explicit state prevents precedence bugs and ensures predictable Escape behavior.**

### Formula Bar State

```rust
pub struct FormulaBarState {
    /// Is formula bar focused for editing?
    pub focused: bool,

    /// Current formula text
    pub text: String,

    /// Cursor position (char index, not byte offset)
    pub cursor: usize,

    /// Text selection range, if any
    pub selection: Option<Range<usize>>,

    /// Autocomplete state
    pub autocomplete: Option<AutocompleteState>,

    /// Signature help state
    pub signature_help: Option<SignatureHelpState>,

    /// Error banner state
    pub error_banner: Option<ErrorBannerState>,
}

pub struct AutocompleteState {
    pub items: Vec<&'static FunctionInfo>,
    pub selected_index: usize,
    pub replace_range: Range<usize>,
    pub filter_text: String,
}

pub struct SignatureHelpState {
    pub function: &'static FunctionInfo,
    pub current_arg: usize,
}

pub struct ErrorBannerState {
    pub message: String,
    pub span: Option<Range<usize>>,  // For future squiggles
}
```

### State Precedence Rules

**Autocomplete and signature help CAN be open simultaneously.** This matches VS Code behavior where you see both the completion list and parameter hints.

**Escape closes in order:**
1. Autocomplete (if open) → close autocomplete, keep signature help
2. Signature help (if open, autocomplete closed) → close signature help
3. Neither open → cancel edit, restore original value, exit formula bar

**Enter behavior:**
1. If autocomplete open → accept selected item, close autocomplete
2. If autocomplete closed → commit formula, exit edit mode

**Tab behavior:**
1. If autocomplete open → accept selected item, close autocomplete
2. If autocomplete closed → commit formula and move active cell right (Excel behavior)

**Shift+Tab behavior:**
1. If autocomplete open → dismiss autocomplete (no accept)
2. If autocomplete closed → commit formula and move active cell left

> **Rationale:** This is a spreadsheet, not a text editor. Literal tabs in formulas are not useful. Tab = commit + navigate matches Excel muscle memory.

### State Transitions

```
┌─────────────────┐
│   Not Editing   │
└────────┬────────┘
         │ Focus formula bar / type '=' / F2
         ▼
┌─────────────────┐
│    Editing      │◄───────────────────────────┐
│                 │                             │
│ autocomplete: None                            │
│ signature_help: None                          │
│ error_banner: None                            │
└────────┬────────┘                             │
         │                                      │
         │ Type letter in Identifier/Start mode │
         ▼                                      │
┌─────────────────┐                             │
│    Editing      │                             │
│                 │──── Escape ─────────────────┤
│ autocomplete: Some(...)                       │
│ signature_help: maybe                         │
└────────┬────────┘                             │
         │                                      │
         │ Type '(' after function              │
         ▼                                      │
┌─────────────────┐                             │
│    Editing      │                             │
│                 │──── Escape (x2 if both) ────┘
│ autocomplete: maybe
│ signature_help: Some(...)
└─────────────────┘
```

---

## Autocomplete Popup

### Trigger Rules (Context-Based)

Autocomplete appears when `FormulaContext.mode` is:
- `Start` (right after `=`) — show full function list
- `Identifier` (typing what could be a function name) — **only if identifier length ≥ 2**
- `Operator` (after an operator, ready for next operand) — show full function list

Autocomplete does **NOT** appear when mode is:
- `String` (inside quotes)
- `Number` (typing a numeric literal)
- `Reference` (on a cell ref like A1 — future: show named ranges here)
- `Identifier` with length 1 (ambiguous: `A` could be cell ref or AVERAGE)

### Identifier Length Rule

> **Rule:** Don't show autocomplete for single-letter identifiers.

**Problem:** `=A` is ambiguous — could be start of cell ref `A1` or function `AVERAGE`.

**Solution:** Require ≥2 characters before showing function suggestions:
- `=A` → no autocomplete (could be A1, A2, AVERAGE, ABS, AND...)
- `=AV` → show autocomplete (clearly a function prefix)
- `=S` → no autocomplete (could be SUM, STDEV, or cell S1)
- `=SU` → show autocomplete (SUM, SUMIF, SUBSTITUTE)

This avoids noise and prevents accidental acceptance when user intended a cell reference.

### Behavior

```
Formula bar: =SU|
            ┌─────────────────────────────┐
            │ SUM(number1, [number2], ...)│
            │ SUMIF(range, criteria, ...) │
            │ SUBSTITUTE(text, old, new)  │
            └─────────────────────────────┘
```

**Rules:**
- Show top 7 matching functions (prefix match, case-insensitive)
- Highlight matching prefix in bold
- Show signature inline (truncated to ~40 chars)
- Arrow Up/Down to navigate (wraps at ends)
- Tab or Enter to accept
- Escape to dismiss (not exit edit)
- Continue typing to filter
- Click to accept
- Typing non-matching char closes autocomplete

### Acceptance Behavior

When user accepts a suggestion:
1. Replace `replace_range` with function name + `(`
2. Position cursor after `(`
3. Close autocomplete
4. Open signature help for the accepted function

### Data Structure

```rust
pub struct FunctionInfo {
    pub name: &'static str,
    pub signature: &'static str,
    pub description: &'static str,
    pub category: FunctionCategory,
    pub parameters: &'static [ParameterInfo],
}

pub struct ParameterInfo {
    pub name: &'static str,
    pub description: &'static str,
    pub optional: bool,
    pub repeatable: bool,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FunctionCategory {
    Math,
    Logical,
    Text,
    Lookup,
    DateTime,
    Statistical,
    Array,
    Conditional,
    Trigonometry,
}
```

---

## Signature Help

### Trigger Rules

Signature help appears when:
1. User types `(` after a known function name
2. User types `,` inside function arguments
3. Cursor moves into a function's argument list (via arrow keys or click)
4. User accepts autocomplete suggestion (auto-opens)

Signature help closes when:
1. Cursor moves outside the function's parentheses
2. User types `)` that closes the function
3. User presses Escape
4. Formula bar loses focus

### Behavior

```
Formula bar: =VLOOKUP(A1, |
                      ↓
┌─────────────────────────────────────────────────────────┐
│ VLOOKUP(lookup_value, **table_array**, col_index, [exact]) │
│                                                         │
│ table_array: The range containing the lookup table      │
└─────────────────────────────────────────────────────────┘
```

**Rules:**
- Current parameter shown in bold
- Parameter description shown below signature
- Updates as cursor moves or commas are typed
- Tracks nested functions (shows innermost)
- Optional parameters shown in `[brackets]`

### Nested Function Handling

```
Formula: =IF(SUM(A1:A5) > 10, |
```

At cursor position, context is:
- Outer function: IF
- Current arg index in IF: 1 (the "value if true" argument)
- No inner function at cursor

```
Formula: =IF(SUM(|) > 10, "yes", "no")
```

At cursor position:
- Inner function: SUM (this is what we show)
- Current arg index in SUM: 0

---

## Error Highlighting

### Real-time Validation

As user types:
1. Parse formula after each keystroke
2. Classify error as `Hard` or `Transient`
3. Only show banner for `Hard` errors (after 300ms debounce)
4. Clear banner when formula becomes valid

### Error Taxonomy

> **Rule:** Avoid "red screaming while mid-typing." Only show errors the user should fix now.

```rust
pub enum DiagnosticKind {
    /// Always show (after debounce) - user made a definite mistake
    Hard,
    /// Don't show - user is likely still typing
    Transient,
}
```

| Error Type | Kind | Rationale |
|------------|------|-----------|
| Unknown function name | Hard | User typed something wrong |
| Invalid token (e.g., `@@@`) | Hard | Definite syntax error |
| Unmatched `(` when cursor moved away | Hard | User left an incomplete construct |
| Unmatched `(` with cursor at end | Transient | User might be about to type `)` |
| Trailing operator with cursor at end | Transient | User is about to type operand |
| Trailing operator with cursor elsewhere | Hard | User left incomplete expression |

**Implementation rule:**
- If cursor is at end of formula AND error is "incomplete construct" → Transient
- Otherwise → Hard

### Error Banner Display

```
Formula bar: =SUM(A1:A10
             ─────────────────────
             ⚠ Missing closing parenthesis
```

**Hard errors to detect:**
- Unknown function names (always show)
- Invalid tokens (always show)
- Mismatched parentheses (when cursor moved away from the incomplete spot)
- Invalid cell references (e.g., `$$$A1`)

### Squiggles (Future)

When we have span information, underline the error location in red within the formula text itself.

---

## Quick Actions

### F4 Cycle Reference Type

**Trigger:** F4 while cursor is on a cell reference in formula

**Precondition:** `FormulaContext.token_at_cursor` is a `CellRef` or `Range`

**Behavior:**
- Cycle: `A1` → `$A$1` → `A$1` → `$A1` → `A1`
- For ranges, cycle both endpoints: `A1:B2` → `$A$1:$B$2` → ...
- Only affects reference under cursor
- Updates formula bar immediately
- Preserves cursor position within the reference

**Implementation:**
```rust
fn cycle_reference(formula: &str, ref_span: Range<usize>) -> String {
    // Parse ref at span
    // Determine current type (relative, absolute, mixed)
    // Generate next type
    // Replace in formula
}
```

### Alt+= AutoSum

**Trigger:** Alt+= with cell selection (not in edit mode)

**Behavior:**
1. Detect context:
   - Look up: find contiguous numeric cells above active cell
   - Look left: find contiguous numeric cells left of active cell
   - Prefer above if both exist and above is longer
   - Minimum 2 cells to trigger
2. Insert `=SUM(range)` into active cell
3. Enter edit mode with formula selected (so user can modify)
4. If no suitable range found, just enter `=SUM()` with cursor inside parens

---

## Syntax Highlighting

### Token Colors

| Token Type | Example | Theme Key |
|------------|---------|-----------|
| Function name | `SUM`, `VLOOKUP` | `FormulaFunction` |
| Cell reference | `A1`, `$B$2` | `FormulaCellRef` |
| Range | `A1:B10` | `FormulaRange` |
| Named range | `SalesData` | `FormulaNamedRange` |
| Number | `123.45` | `FormulaNumber` |
| String | `"text"` | `FormulaString` |
| Boolean | `TRUE`, `FALSE` | `FormulaBoolean` |
| Constant | `PI()` result | `FormulaConstant` |
| Operator | `+`, `-`, `*`, `/`, `&` | `FormulaOperator` |
| Comparison | `<`, `>`, `=`, `<>` | `FormulaComparison` |
| Parenthesis | `(`, `)` | `FormulaParen` |
| Comma | `,` | `FormulaComma` |
| Error | `#DIV/0!` | `FormulaError` |
| Sheet reference | `Sheet1!A1` | `FormulaSheetRef` |
| Spill reference | `A1#` | `FormulaSpillRef` |
| Whitespace | ` ` | (default/invisible) |
| Percent | `10%` | `FormulaPercent` |

**Note:** Add all theme keys now, even if some (sheet ref, spill ref, percent) are unused in v1. This prevents breaking theme APIs later.

### Theme Integration

Add to `TokenKey` enum in `theme.rs`:

```rust
// Formula syntax highlighting
FormulaFunction,
FormulaCellRef,
FormulaRange,
FormulaNamedRange,
FormulaNumber,
FormulaString,
FormulaBoolean,
FormulaConstant,
FormulaOperator,
FormulaComparison,
FormulaParen,
FormulaComma,
FormulaError,
FormulaSheetRef,    // Future
FormulaSpillRef,    // Future
FormulaPercent,     // Future
// Note: Whitespace and Bang don't need theme keys (invisible/structural)
```

### Implementation

Expose tokenizer from `parser.rs`:

```rust
pub fn tokenize_for_highlight(formula: &str) -> Vec<(Range<usize>, TokenType)> {
    // Reuse existing tokenizer
    // Return spans with semantic types
    // Skip leading '='
}
```

Render formula bar with styled spans instead of plain text.

---

## Hover Documentation

### Trigger

Mouse hovers over a function name in the formula bar for 500ms.

### Behavior

```
Formula bar: =VLOOKUP(A1, B:B, 2, FALSE)
                ▲
            ┌───┴───────────────────────────────────┐
            │ VLOOKUP                               │
            │                                       │
            │ Looks for a value in the leftmost     │
            │ column of a table and returns a       │
            │ value in the same row from a column   │
            │ you specify.                          │
            │                                       │
            │ Syntax: VLOOKUP(lookup_value,         │
            │         table_array, col_index_num,   │
            │         [range_lookup])               │
            └───────────────────────────────────────┘
```

**Rules:**
- Only appears on function names
- Shows full description (not truncated like autocomplete)
- Shows complete signature
- Dismiss on mouse move away or any keystroke
- Does not interfere with autocomplete/signature help

---

## Array Visualization

### Spill Borders

When a cell contains an array formula that spills:

| Cell Type | Visual |
|-----------|--------|
| Spill parent | Blue border (2px solid) |
| Spill receiver | Light blue border (1px dashed) |
| Blocked (#SPILL!) | Red border (2px solid) |

### Selection Behavior

- Click spill receiver → select the spill parent instead
- Edit spill receiver → show message "Cannot edit spill range. Edit [parent cell] instead."
- Delete on spill receiver → no action (must delete parent)

### Implementation

Track in cell rendering:

```rust
if cell.is_spill_parent() {
    border_color = theme.get(TokenKey::FormulaSpillBorder);
    border_width = 2.0;
} else if cell.is_spill_receiver() {
    border_color = theme.get(TokenKey::FormulaSpillReceiverBorder);
    border_style = BorderStyle::Dashed;
    border_width = 1.0;
}
```

---

## Formula View Toggle

### Ctrl+` Toggle

**Behavior:**
- Toggle between showing computed values and raw formulas in cells
- Affects all cells globally
- Formulas shown in monospace font
- **Do NOT auto-widen columns** — clip/ellipsize formula text instead
- Show subtle `…` indicator when formula is clipped
- Rely on formula bar for full formula text

### State

```rust
pub struct SpreadsheetApp {
    pub show_formulas: bool,
    // ...
}
```

### Rendering

```rust
fn cell_display_text(&self, row: usize, col: usize) -> String {
    if self.show_formulas {
        self.sheet.get_raw(row, col)
    } else {
        self.sheet.get_display(row, col)
    }
}
```

---

## Copy Behavior Clarification

> **Rule:** Copy uses raw formula text for formula cells; otherwise uses the cell's displayed value.

| Cell Contains | Copied As |
|---------------|-----------|
| Formula `=A1+B1` | `=A1+B1` (raw text) |
| Plain value `123` | `123` (display) |
| Formatted value `$1,234.56` | `$1,234.56` (display) |
| Spill receiver (from array) | Displayed value (what user sees) |
| Error `#DIV/0!` | `#DIV/0!` (display) |

This matches Excel behavior and user expectations.

---

## Function Inventory (96 Functions)

### Math (23)
`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `INT`, `MOD`, `POWER`, `SQRT`, `CEILING`, `FLOOR`, `PRODUCT`, `MEDIAN`, `LOG`, `LOG10`, `LN`, `EXP`, `RAND`, `RANDBETWEEN`

### Logical (12)
`IF`, `AND`, `OR`, `NOT`, `IFERROR`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR`, `IFS`, `SWITCH`, `CHOOSE`

### Text (14)
`CONCATENATE`, `CONCAT`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `TRIM`, `TEXT`, `VALUE`, `FIND`, `SUBSTITUTE`, `REPT`

### Conditional (3)
`SUMIF`, `COUNTIF`, `COUNTBLANK`

### Lookup (8)
`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`

### Date/Time (13)
`TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `DATEDIF`, `EDATE`, `EOMONTH`, `HOUR`, `MINUTE`, `SECOND`

### Trigonometry (10)
`PI`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `DEGREES`, `RADIANS`

### Statistical (8)
`STDEV`, `STDEV.S`, `STDEV.P`, `STDEVP`, `VAR`, `VAR.S`, `VAR.P`, `VARP`

### Array (5)
`SEQUENCE`, `TRANSPOSE`, `SORT`, `UNIQUE`, `FILTER`

---

## Implementation Priority

### Sprint 1: Foundation + Core IntelliSense
1. **FormulaContext analyzer** — Everything else depends on this
2. **Autocomplete popup** — Most requested feature
3. **Signature help** — Natural follow-on, shares context
4. **Error banner** — Uses same debounce/parse infrastructure
5. **F4 ref cycle** — Depends on token_at_cursor from analyzer

### Sprint 2: Visual Polish
6. **Syntax highlighting** — Makes formulas readable
7. **Hover docs** — High perceived quality, now possible in gpui
8. **Alt+= AutoSum** — Quick win

### Sprint 3: Array Support
9. **Spill visualization** — Shows array behavior
10. **Ctrl+` formula view** — Debugging aid (no column resize)

---

## Files to Modify

| File | Changes |
|------|---------|
| NEW: `src/formula_context.rs` | FormulaContext, analyze(), all context logic |
| `src/app.rs` | FormulaBarState, autocomplete/signature state |
| `src/views/formula_bar.rs` | Render popups, syntax highlighting, hover |
| `src/theme.rs` | Add 15 formula token colors |
| `src/mode.rs` | No changes needed (formula editing is within existing Edit mode) |
| `crates/engine/src/formula/parser.rs` | Expose tokenize_for_highlight() |

---

## Edge Cases

> **These are the subtle decisions that prevent "it feels off" bugs.**

### 1. Autocomplete Never Auto-Accepts

**Invariant:** Autocomplete acceptance is always explicit (Enter/Tab/click). Typing never "auto-corrects."

- User types `=IN` → popup shows INDEX, INT
- User continues typing `=IND` → popup filters to INDEX
- User types `=INDEX` and keeps typing `(` → user intended INDEX, acceptance happens because `(` follows known function
- User types `=IN` then types a number → popup closes, no insertion

**Rule:** If popup is open and user types a character that doesn't match any suggestion, close popup. Never insert "best guess."

### 2. Replace Range Definition

**Exact definition:**

```rust
fn compute_replace_range(formula: &str, cursor: usize, mode: FormulaEditMode) -> Range<usize> {
    match mode {
        // Zero-width at cursor - nothing to replace yet
        Start | Operator => cursor..cursor,

        // Expand to full identifier span
        Identifier => {
            let start = scan_backward_while(formula, cursor, |c| c.is_ascii_alphanumeric() || c == '_');
            let end = scan_forward_while(formula, cursor, |c| c.is_ascii_alphanumeric() || c == '_');
            start..end
        }

        // Other modes: no replacement
        _ => cursor..cursor,
    }
}
```

**Examples:**
- `=SU|M` (cursor between U and M) → replace_range covers "SUM" (0..3 after `=`)
- `=SUM(|` → replace_range is empty (cursor..cursor)
- `=|` → replace_range is empty

### 3. Whitespace Permissiveness

**Rule:** Allow whitespace anywhere that doesn't change meaning. The editor should feel robust.

| Input | Valid? | Parsing |
|-------|--------|---------|
| `=SUM(A1:A3)` | Yes | Normal |
| `=SUM (A1:A3)` | Yes | Whitespace between function and `(` allowed |
| `=SUM( A1:A3 )` | Yes | Whitespace inside parens allowed |
| `=A1 + B1` | Yes | Whitespace around operators allowed |
| `=A1 : B1` | Yes | Whitespace around colon allowed |
| `= SUM(A1)` | Yes | Whitespace after `=` allowed |

**Implementation:**
- Tokenizer emits `Whitespace` tokens but they don't affect parsing
- Spans are calculated excluding whitespace for highlighting purposes
- Signature help arg counting ignores whitespace tokens

### 4. Percent Literals (Unsupported in v1)

**Rule:** If user types `10%` and engine doesn't support percent, show explicit message.

| Input | Error Kind | Message |
|-------|------------|---------|
| `=10%` | Hard | "Percent literals not yet supported. Use `10/100` or `0.1` instead." |
| `=A1*50%` | Hard | Same message |

**Rationale:** Users coming from Excel will try this. A clear message is better than "Unexpected character: %".

### 5. Unary Minus Classification

**Rule:** `-` is `UnaryMinus` when it appears in a position expecting an operand. Otherwise it's binary `Operator`.

| Context | Token Type | Example |
|---------|------------|---------|
| After `=` | UnaryMinus | `=-1` |
| After `(` | UnaryMinus | `=(-1)` |
| After `,` | UnaryMinus | `=SUM(-1, -2)` |
| After operator | UnaryMinus | `=1*-2` |
| After number/ref/`)` | Operator (binary minus) | `=1-2`, `=A1-1` |

**Signature help impact:** `=SUM(-1, -2)` has 2 arguments, not 4. The `-` before each number is part of the number, not a separate token for arg counting.

### 6. Empty Function Calls

**Rule:** `=SUM()` is valid syntax (zero arguments).

- Signature help shows: `SUM(**number1**, [number2], ...)` with first param highlighted
- No error banner (some functions accept zero args)
- Evaluation may return error depending on function semantics

### 7. Incomplete Range References

**Rule:** `=A1:` with cursor at end is Transient error. `=A1:|` (cursor moved away) becomes Hard.

| Input | Cursor | Error Kind |
|-------|--------|------------|
| `=A1:|` | at end | Transient |
| `=A1:| + 1` | after colon | Hard |
| `=SUM(A1:|)` | after colon | Transient (inside construct) |

### 8. Case Preservation in Non-Function Identifiers

**Rule:** Function names are normalized to UPPERCASE. Cell refs preserve user's case for display but normalize for evaluation.

| User Types | Displayed | Evaluated As |
|------------|-----------|--------------|
| `=sum(a1)` | `=SUM(A1)` | SUM of A1 |
| `=Sum(a1)` | `=SUM(A1)` | SUM of A1 |
| `=a1+b1` | `=A1+B1` | A1 + B1 |

**Note:** We normalize on commit, not while typing. User sees their input until Enter.

### 9. Nested Parentheses Without Functions

**Rule:** `=(1+2)*3` is valid. Signature help does not appear for non-function parens.

- `=(|1+2)` → mode is `Number`, no signature help
- `=SUM(|1+2)` → mode is `ArgList`, signature help shows SUM

### 10. Selection Overrides Cursor for Replace Range

**Rule:** If user has text selected in formula bar, `replace_range` = selection, not computed span.

This allows precise replacement when user explicitly selects text.

---

## Testing Checklist

### Behavior Tests

**Autocomplete:**
- [ ] Appears after typing `=S`
- [ ] Filters as you type
- [ ] Arrow keys navigate (wraps)
- [ ] Enter/Tab accepts
- [ ] Escape dismisses (doesn't exit edit)
- [ ] Works after `(` and `,` and operators
- [ ] Accepting inserts function name + `(` and opens signature help

**Signature Help:**
- [ ] Appears after `(`
- [ ] Highlights current parameter
- [ ] Updates on `,`
- [ ] Handles nested functions (shows innermost)
- [ ] Dismisses on `)` or cursor leaving function
- [ ] Persists while navigating within args

**Syntax Highlighting:**
- [ ] Functions colored
- [ ] Cell refs colored
- [ ] Strings colored
- [ ] Numbers colored
- [ ] Errors colored red
- [ ] Named ranges colored (when implemented)

**Quick Actions:**
- [ ] F4 cycles reference type
- [ ] F4 works on ranges (cycles both endpoints)
- [ ] Alt+= inserts SUM with detected range
- [ ] Alt+= prefers above over left

### Correctness Invariants

These are the bugs that ruin "it feels like an IDE":

- [ ] **Autocomplete does NOT appear inside string literals**
- [ ] **Autocomplete does NOT appear for single-letter identifiers** (`=A` → no popup)
- [ ] **Autocomplete never auto-accepts** — only explicit Enter/Tab/click inserts
- [ ] **Signature help tracks nested functions when cursor moves** (not just on comma)
- [ ] **F4 toggles only the ref under cursor** when multiple refs exist
- [ ] **Escape precedence**: autocomplete → signature help → cancel edit
- [ ] **Autocomplete replace_range is correct** for mid-word edits (e.g., `=SU|M` → replacing "SUM")
- [ ] **Context analyzer returns consistent results** regardless of which feature calls it
- [ ] **Transient errors don't show** when cursor is at end of incomplete construct
- [ ] **Hard errors show** after debounce even while typing

---

## Golden Tests

> **Behavioral fixtures that must pass before shipping. These are not unit tests — they're acceptance criteria.**

### Autocomplete Behavior

| Formula | Cursor | Expected |
|---------|--------|----------|
| `=|` | after `=` | Autocomplete shows full function list |
| `=A|` | after A | **No autocomplete** (single letter) |
| `=AV|` | after V | Autocomplete shows AVERAGE |
| `=SU|` | after U | Autocomplete shows SUM, SUMIF, SUBSTITUTE |
| `=SU|M` | between U and M | Autocomplete shows SUM; replace_range covers "SUM" |
| `=SUM|` | after M | Autocomplete shows SUM (exact match first) |
| `="hello|"` | inside string | **No autocomplete** |
| `=SUM(|` | after `(` | Autocomplete shows full list (new operand position) |
| `=SUM(A1,|` | after `,` | Autocomplete shows full list |
| `=SUM(A1)+|` | after `+` | Autocomplete shows full list |

### Signature Help Behavior

| Formula | Cursor | Expected |
|---------|--------|----------|
| `=SUM(|)` | inside parens | Signature: `SUM(**number1**, [number2], ...)` — arg 0 |
| `=SUM(1,|2)` | after comma | Signature: `SUM(number1, **[number2]**, ...)` — arg 1 |
| `=SUM(1,2,|3)` | after second comma | Signature: arg 2 (repeatable param) |
| `=VLOOKUP(A1,|B:B,2)` | after first comma | Signature: `**table_array**` highlighted |
| `=IF(SUM(|),1,2)` | inside nested SUM | Signature shows **SUM**, not IF |
| `=IF(SUM(A1)|,1,2)` | after nested SUM | Signature shows **IF**, arg 0 |
| `=IF(SUM(A1),|1,2)` | after first comma in IF | Signature shows **IF**, arg 1 |
| `=(|1+2)` | inside non-function parens | **No signature help** |
| `=SUM (|A1)` | with space before `(` | Signature help works (whitespace allowed) |

### Error Banner Behavior

| Formula | Cursor | Expected |
|---------|--------|----------|
| `=SUM(A1|` | at end | **No banner** (Transient — user typing) |
| `=SUM(A1| + 1` | not at end | Banner: "Missing closing parenthesis" (Hard) |
| `=SUMM(A1)` | anywhere | Banner: "Unknown function: SUMM" (Hard) |
| `=A1 +|` | at end | **No banner** (Transient — incomplete) |
| `=A1 +| B1` | not at end | Banner: "Expected operand" (Hard) |
| `=10%` | anywhere | Banner: "Percent literals not yet supported" (Hard) |
| `=@@@` | anywhere | Banner: "Invalid character" (Hard) |

### Unary Minus & Arg Counting

| Formula | Cursor | Expected |
|---------|--------|----------|
| `=SUM(-1,|` | after comma | Signature shows arg 1 (not arg 2) |
| `=SUM(-1,-2,|` | after second comma | Signature shows arg 2 |
| `=1-|2` | after minus | Mode is `Operator` (binary minus) |
| `=-|1` | after minus | Mode is `Number` (unary minus) |
| `=(-|1)` | after minus | Mode is `Number` (unary minus) |
| `=1*-|2` | after minus | Mode is `Number` (unary minus after operator) |

### F4 Reference Cycling

| Formula | Cursor | Before F4 | After F4 |
|---------|--------|-----------|----------|
| `=A1|+B1` | on A1 | `A1` | `$A$1` |
| `=$A$1|+B1` | on $A$1 | `$A$1` | `A$1` |
| `=A$1|+B1` | on A$1 | `A$1` | `$A1` |
| `=$A1|+B1` | on $A1 | `$A1` | `A1` |
| `=A1:B2|+C1` | on range | `A1:B2` | `$A$1:$B$2` |
| `=SUM(A1|)` | on A1 inside function | `A1` | `$A$1` |

### Whitespace Tolerance

| Formula | Valid? | Notes |
|---------|--------|-------|
| `=SUM(A1:A3)` | Yes | Normal |
| `=SUM (A1:A3)` | Yes | Space before `(` |
| `= SUM(A1:A3)` | Yes | Space after `=` |
| `=SUM( A1:A3 )` | Yes | Spaces inside |
| `=A1 + B1` | Yes | Spaces around operator |
| `=A1 : B1` | Yes | Spaces around colon |
| `=SUM(A1 , A2)` | Yes | Spaces around comma |

### Mode Detection

| Formula | Cursor | Expected Mode |
|---------|--------|---------------|
| `=|` | after `=` | Start |
| `=SUM|` | after identifier | Identifier |
| `=SUM(|` | after `(` | ArgList |
| `=SUM(A1|` | on cell ref | Reference |
| `=SUM(A1:|B2` | on range | Reference |
| `=SUM(123|` | on number | Number |
| `="text|"` | inside string | String |
| `=A1+|` | after operator | Operator |
| `=SUM(A1)|` | after `)` | Complete |

---

## Version History

- **v1** (2026-01): Initial gpui implementation spec
- **v1.1** (2026-01): Added FormulaContext analyzer, UI state model, expanded token types, copy clarification
- **v1.2** (2026-01): Added cursor indexing policy (char vs byte), Tab/Shift+Tab behavior, identifier length rule for autocomplete, error taxonomy (Hard vs Transient), primary_span, function name normalization, additional token types (Whitespace, Bang, Percent, UnaryMinus)
- **v1.3** (2026-01): Added Edge Cases section (10 decisions), Golden Tests section (50+ behavioral fixtures)
