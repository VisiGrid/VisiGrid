# Formula Roadmap: VisiGrid vs Google Sheets

This document tracks VisiGrid's formula coverage compared to Google Sheets, organized by priority tiers.

---

## Current State

## Structural Primitives (Foundational)

Before expanding formula coverage, VisiGrid needs array semantics:

- Range-in → range-out evaluation
- Broadcasting rules
- Spill behavior
- Scalar ↔ array coercion

These unlock:
- `SUMIFS`, `AVERAGEIFS`, `COUNTIFS`
- `FILTER`, `SORT`, `UNIQUE`
- Compact finance schedules and table-style modeling

### Implemented

**Arithmetic Operators:**
- `+` Addition
- `-` Subtraction
- `*` Multiplication
- `/` Division
- Parentheses for grouping
- Unary minus

**Comparison Operators:**
- `<` Less than
- `>` Greater than
- `=` Equal
- `<=` Less than or equal
- `>=` Greater than or equal
- `<>` Not equal

**String Operators:**
- `&` Concatenation
- `"text"` String literals

**Cell References:**
- Single cell: `A1`, `B2`
- Ranges: `A1:B5`
- Multi-letter columns: `AA1`, `AB5`, `ZZ100`
- Absolute references: `$A$1`, `$A1`, `A$1`

**Data Types:**
- Numbers
- Text/Strings
- Booleans (TRUE/FALSE)
- Errors (#DIV/0!, #VALUE!, #ERR, etc.)

**Functions (96 total):**

| Category | Functions |
|----------|-----------|
| **Math (23)** | SUM, AVERAGE, AVG, MIN, MAX, COUNT, COUNTA, ABS, ROUND, INT, MOD, POWER, SQRT, CEILING, FLOOR, PRODUCT, MEDIAN, LOG, LOG10, LN, EXP, RAND, RANDBETWEEN |
| **Logical (12)** | IF, AND, OR, NOT, IFERROR, ISBLANK, ISNUMBER, ISTEXT, ISERROR, IFS, SWITCH, CHOOSE |
| **Text (14)** | CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, TEXT, VALUE, FIND, SUBSTITUTE, REPT |
| **Conditional (3)** | SUMIF, COUNTIF, COUNTBLANK |
| **Lookup (8)** | VLOOKUP, HLOOKUP, INDEX, MATCH, ROW, COLUMN, ROWS, COLUMNS |
| **Date/Time (13)** | TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DATEDIF, EDATE, EOMONTH, HOUR, MINUTE, SECOND |
| **Trigonometry (10)** | PI, SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, DEGREES, RADIANS |
| **Statistical (8)** | STDEV, STDEV.S, STDEV.P, STDEVP, VAR, VAR.S, VAR.P, VARP |
| **Array (5)** | SEQUENCE, TRANSPOSE, SORT, UNIQUE, FILTER |

### Not Yet Implemented

- Named ranges in formulas (defined but not referenced)
- Cross-sheet references
- ~~Array formulas~~ ✓ Basic array support with spill
- Financial functions

---

## Priority Tiers

### Tier 1: Essential (Ship Blockers) ✓ COMPLETE

These are expected by any spreadsheet user. Missing them feels broken.

#### Parser Improvements
- [x] Multi-letter columns (AA, AB, ... ZZ, AAA)
- [x] Comparison operators: `<`, `>`, `=`, `<=`, `>=`, `<>`
- [x] Text literals in formulas: `"hello"`
- [x] Concatenation operator: `&`
- [x] Absolute/mixed references: `$A$1`, `$A1`, `A$1`

#### Math Functions
| Function | Description | Status |
|----------|-------------|--------|
| `SUMIF` | Conditional sum | ✓ Done |
| `SUMIFS` | Multi-condition sum | Not yet |
| `PRODUCT` | Multiply all values | ✓ Done |
| `POWER` | Exponentiation | ✓ Done |
| `SQRT` | Square root | ✓ Done |
| `MOD` | Modulo/remainder | ✓ Done |
| `INT` | Truncate to integer | ✓ Done |
| `CEILING` | Round up | ✓ Done |
| `FLOOR` | Round down | ✓ Done |

#### Logical Functions
| Function | Description | Status |
|----------|-------------|--------|
| `IF` | Conditional | ✓ Done |
| `AND` | Logical AND | ✓ Done |
| `OR` | Logical OR | ✓ Done |
| `NOT` | Logical NOT | ✓ Done |
| `IFERROR` | Error handling | ✓ Done |
| `ISBLANK` | Check if empty | ✓ Done |
| `ISNUMBER` | Check if numeric | ✓ Done |
| `ISTEXT` | Check if text | ✓ Done |

#### Statistical Functions
| Function | Description | Status |
|----------|-------------|--------|
| `COUNTA` | Count non-empty cells | ✓ Done |
| `COUNTIF` | Conditional count | ✓ Done |
| `COUNTIFS` | Multi-condition count | Not yet |
| `COUNTBLANK` | Count empty cells | ✓ Done |
| `AVERAGEIF` | Conditional average | Not yet |
| `AVERAGEIFS` | Multi-condition average | Not yet |
| `MEDIAN` | Median value | ✓ Done |
| `STDEV` | Standard deviation | ✓ Done |
| `VAR` | Variance | ✓ Done |

#### Text Functions
| Function | Description | Status |
|----------|-------------|--------|
| `CONCATENATE` / `CONCAT` | Join strings | ✓ Done |
| `LEFT` | Left substring | ✓ Done |
| `RIGHT` | Right substring | ✓ Done |
| `MID` | Middle substring | ✓ Done |
| `LEN` | String length | ✓ Done |
| `UPPER` | Uppercase | ✓ Done |
| `LOWER` | Lowercase | ✓ Done |
| `TRIM` | Remove whitespace | ✓ Done |
| `TEXT` | Format number as text | ✓ Done |
| `VALUE` | Parse text to number | ✓ Done |
| `FIND` | Find substring position | ✓ Done |
| `SUBSTITUTE` | Replace text | ✓ Done |

---

### Tier 2: Power User (Pro Differentiator) - PARTIAL

These separate casual users from power users. Good candidates for Pro edition.

**Progress:** Lookup (8/11) and Date/Time (13/15) functions implemented.

#### Lookup Functions
| Function | Description | Status |
|----------|-------------|--------|
| `VLOOKUP` | Vertical lookup | ✓ Done |
| `HLOOKUP` | Horizontal lookup | ✓ Done |
| `INDEX` | Return value at position | ✓ Done |
| `MATCH` | Find position of value | ✓ Done |
| `XLOOKUP` | Modern lookup | Not yet |
| `INDIRECT` | Reference from string | Not yet |
| `OFFSET` | Dynamic range | Not yet |
| `ROW` | Current row number | ✓ Done |
| `COLUMN` | Current column number | ✓ Done |
| `ROWS` | Count rows in range | ✓ Done |
| `COLUMNS` | Count columns in range | ✓ Done |

#### Date/Time Functions
| Function | Description | Status |
|----------|-------------|--------|
| `TODAY` | Current date | ✓ Done |
| `NOW` | Current datetime | ✓ Done |
| `DATE` | Create date | ✓ Done |
| `YEAR` | Extract year | ✓ Done |
| `MONTH` | Extract month | ✓ Done |
| `DAY` | Extract day | ✓ Done |
| `WEEKDAY` | Day of week | ✓ Done |
| `DATEDIF` | Difference between dates | ✓ Done |
| `EDATE` | Add months to date | ✓ Done |
| `EOMONTH` | End of month | ✓ Done |
| `HOUR` | Extract hour from time | ✓ Done |
| `MINUTE` | Extract minute from time | ✓ Done |
| `SECOND` | Extract second from time | ✓ Done |
| `NETWORKDAYS` | Working days between | Not yet |
| `WORKDAY` | Add working days | Not yet |

#### More Math
| Function | Description | Status |
|----------|-------------|--------|
| `RAND` | Random 0-1 | ✓ Done |
| `RANDBETWEEN` | Random in range | ✓ Done |
| `LOG` | Logarithm | ✓ Done |
| `LOG10` | Base-10 log | ✓ Done |
| `LN` | Natural log | ✓ Done |
| `EXP` | e^x | ✓ Done |
| `PI` | Pi constant | ✓ Done |
| `DEGREES` | Radians to degrees | ✓ Done |
| `RADIANS` | Degrees to radians | ✓ Done |
| `SIN`, `COS`, `TAN` | Trigonometry | ✓ Done |
| `ASIN`, `ACOS`, `ATAN` | Inverse trig | ✓ Done |
| `ATAN2` | Two-argument arctangent | ✓ Done |

#### Advanced Logical
| Function | Description | Status |
|----------|-------------|--------|
| `IFS` | Multiple conditions | ✓ Done |
| `SWITCH` | Match value to results | ✓ Done |
| `CHOOSE` | Select by index | ✓ Done |
| `ISERROR` | Check for error | ✓ Done |
| `ISNA` | Check for #N/A | Not yet |

---

### Tier 3: Specialist (Long-term)

Niche but expected by domain experts.

#### Financial Functions
| Function | Description | Use Case |
|----------|-------------|----------|
| `PMT` | Payment amount | Loans |
| `IPMT` | Interest portion of payment | Loans |
| `PPMT` | Principal portion of payment | Loans |
| `FV` | Future value | Investments |
| `PV` | Present value | Investments |
| `NPV` | Net present value | Finance |
| `XNPV` | NPV with irregular dates | Finance |
| `IRR` | Internal rate of return | Finance |
| `XIRR` | IRR with irregular dates | Finance |
| `MIRR` | Modified internal rate of return | Finance |
| `RATE` | Interest rate | Loans |
| `NPER` | Number of periods | Loans |
| `SLN` | Straight-line depreciation | Accounting |

#### Statistical (Advanced)
| Function | Description |
|----------|-------------|
| `PERCENTILE` | Nth percentile |
| `QUARTILE` | Quartile value |
| `RANK` | Rank in list |
| `LARGE` | Nth largest |
| `SMALL` | Nth smallest |
| `CORREL` | Correlation coefficient |
| `FORECAST` | Linear prediction |
| `TREND` | Linear trend array |
| `GROWTH` | Exponential trend |

#### Array/Modern Functions
| Function | Description | Status |
|----------|-------------|--------|
| `SEQUENCE` | Generate sequence | ✓ Done |
| `TRANSPOSE` | Swap rows/cols | ✓ Done |
| `SORT` | Sort range | ✓ Done |
| `UNIQUE` | Unique values | ✓ Done |
| `FILTER` | Filter rows | ✓ Done |
| `FLATTEN` | Flatten to column | Not yet |
| `ARRAYFORMULA` | Apply to range | Google Sheets specific |
| `QUERY` | SQL-like queries | Google Sheets specific |

#### Information Functions
| Function | Description |
|----------|-------------|
| `TYPE` | Value type |
| `N` | Convert to number |
| `NA` | Return #N/A |
| `ERROR.TYPE` | Error type number |
| `CELL` | Cell information |
| `INFO` | Environment info |

---

## Implementation Notes

### Data Types

The formula system now supports multiple data types:

```rust
pub enum EvalResult {
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(String),
}
```

### Parser Features

The parser now supports:

1. **Multi-letter columns:**
   ```
   AA1 -> col 26, row 0
   AZ1 -> col 51, row 0
   BA1 -> col 52, row 0
   ```

2. **Comparison operators:**
   ```
   =A1>10         -> TRUE/FALSE
   =A1=B1         -> TRUE/FALSE
   =A1<>B1        -> TRUE/FALSE (not equal)
   ```

3. **String literals:**
   ```
   ="Hello"
   =A1&" "&B1     -> Concatenation
   =IF(A1>0,"Yes","No")
   ```

4. **Absolute/mixed references:**
   ```
   $A$1  -> Absolute
   $A1   -> Column absolute
   A$1   -> Row absolute
   ```

### Dependency Graph

For recalculation efficiency, need:
- Track which cells depend on which ✓ (basic)
- Topological sort for eval order
- Detect circular references
- Incremental recalc (only dirty cells)

---

## Recommended Implementation Order

### Phase 1: Core Completeness ✓ DONE
1. ~~`IF` function (unlocks conditional logic)~~ ✓
2. ~~Comparison operators~~ ✓
3. ~~String literals and `CONCATENATE`~~ ✓
4. ~~`SUMIF`, `COUNTIF`~~ ✓
5. ~~Text functions: `LEFT`, `RIGHT`, `MID`, `LEN`~~ ✓

### Phase 2: Lookup & Reference ✓ DONE
1. ~~`VLOOKUP`~~ ✓
2. ~~`INDEX` + `MATCH`~~ ✓
3. ~~Date functions: `TODAY`, `DATE`, `YEAR`, `MONTH`, `DAY`~~ ✓
4. `SUMIFS`, `COUNTIFS`, `AVERAGEIF` (deferred to Phase 3)

### Phase 3: Power Features
1. `XLOOKUP` (modern replacement for VLOOKUP)
2. Array formulas
3. `FILTER`, `SORT`, `UNIQUE`
4. Financial functions

---

## Google Sheets Function Count

For reference, Google Sheets has **500+** functions across categories:

| Category | Approx Count |
|----------|--------------|
| Math | 50+ |
| Statistical | 80+ |
| Text | 30+ |
| Logical | 15+ |
| Lookup | 20+ |
| Date/Time | 30+ |
| Financial | 50+ |
| Database | 10+ |
| Information | 20+ |
| Engineering | 50+ |
| Array | 20+ |
| Web | 10+ |
| Google-specific | 50+ |

**VisiGrid Target:** ~100 functions covers 95% of real-world use cases. The long tail (engineering, specialized statistical) can wait.

**Current VisiGrid Count:** 96 functions (96% of target)

---

## Success Metrics

A formula system is "complete enough" when users can:

1. Do basic arithmetic with cell references ✓
2. Use conditional logic (`IF`, `SUMIF`) ✓
3. Look up values (`VLOOKUP` or `INDEX/MATCH`) ✓
4. Work with text (`LEFT`, `RIGHT`, `CONCATENATE`) ✓
5. Handle dates (`TODAY`, `DATE`, `DATEDIF`) ✓
6. Handle errors gracefully (`IFERROR`) ✓

**Current coverage: 6/6 core capabilities. ✓ COMPLETE**

---

## Changelog

### 2026-01-18 - Dynamic Arrays Complete (Track B Phase 2)
- Fixed reference resolution: formulas referencing spill receivers now work correctly
- Added spill range highlight: selecting a spill parent highlights the full spill rectangle
- Implemented high-impact array functions:
  - SORT(range, [sort_col], [is_asc]) - Sort rows by column (numbers < text < empty < errors)
  - UNIQUE(range) - Deduplicate rows (first occurrence wins, case-insensitive)
  - FILTER(range, include) - Filter rows where include column is TRUE (#CALC! if no matches)
- Total functions: 96 (up from 93)
- All core dynamic array functions now complete: SEQUENCE, TRANSPOSE, SORT, UNIQUE, FILTER

### 2026-01-18 - Array Primitives Foundation (Track B Phase 1)
- Implemented `Value` enum (Empty, Number, Text, Boolean, Error) for scalar values
- Implemented `Array2D` type for 2D arrays with dense row-major storage
- Added `EvalResult::Array` variant for array formula results
- Implemented spill mechanics:
  - Spill bookkeeping (spill_parent, spill_info, spill_error on Cell)
  - Spill collision detection with #SPILL! error display
  - Automatic spill clearing when formula changes
  - Edit blocking for spill-receiving cells
- Visual indicators:
  - Blue border for spill parent cells
  - Light blue border for spill receiver cells
  - Red border for cells with #SPILL! errors
- Added Array functions:
  - SEQUENCE(rows, [cols], [start], [step]) - Generate sequences
  - TRANSPOSE(array) - Swap rows and columns
- Total functions: 93 (up from 91)

### 2026-01-18 - Track A Scalar Functions (91% of target)
- Added Random functions: RAND, RANDBETWEEN
- Added Logarithm functions: LOG, LOG10, LN, EXP
- Added Trigonometry functions: PI, SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, DEGREES, RADIANS
- Added Advanced logical functions: IFS, SWITCH, CHOOSE
- Added Statistical functions: STDEV, STDEV.S, STDEV.P, STDEVP, VAR, VAR.S, VAR.P, VARP
- Total functions: 91 (up from 68)

### 2026-01-17 - Phase 2 Complete (6/6 Core Capabilities)
- Added Lookup functions: VLOOKUP, HLOOKUP, INDEX, MATCH, ROW, COLUMN, ROWS, COLUMNS
- Added Date/Time functions: TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DATEDIF, EDATE, EOMONTH, HOUR, MINUTE, SECOND
- Excel-compatible date serial numbers (days since 1899-12-30)
- Total functions: 68 (up from 47)
- **All 6 core capabilities now complete**

### 2026-01-17 - Phase 1 Complete
- Added comparison operators (`<`, `>`, `=`, `<=`, `>=`, `<>`)
- Added string literals and concatenation operator (`&`)
- Added multi-letter column support (AA, AB, etc.)
- Added absolute/mixed cell references ($A$1, $A1, A$1)
- Added Boolean and Text types to formula evaluation
- Added 40 new functions across Math, Logical, Text, and Conditional categories
- Total functions: 47 (up from 7)
