# Formula Support: VisiGrid gpui

The formula engine lives in `crates/engine/` and is shared between iced and gpui versions.

---

## Engine Status: COMPLETE

The formula engine is fully functional with 96 functions. No changes needed for gpui migration.

### Implemented Features

| Feature | Status |
|---------|--------|
| Arithmetic operators (+, -, *, /) | ✅ |
| Comparison operators (<, >, =, <=, >=, <>) | ✅ |
| String concatenation (&) | ✅ |
| Cell references (A1, $A$1, A$1, $A1) | ✅ |
| Range references (A1:B10) | ✅ |
| Multi-letter columns (AA, AB, ZZ) | ✅ |
| Boolean literals (TRUE, FALSE) | ✅ |
| String literals ("text") | ✅ |
| Error types (#DIV/0!, #VALUE!, etc.) | ✅ |
| Array formulas with spill | ✅ |
| Dependency tracking | ✅ |

### Function Count: 96

| Category | Count | Functions |
|----------|-------|-----------|
| Math | 23 | SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ABS, ROUND, INT, MOD, POWER, SQRT, CEILING, FLOOR, PRODUCT, MEDIAN, LOG, LOG10, LN, EXP, RAND, RANDBETWEEN |
| Logical | 12 | IF, AND, OR, NOT, IFERROR, ISBLANK, ISNUMBER, ISTEXT, ISERROR, IFS, SWITCH, CHOOSE |
| Text | 14 | CONCATENATE, CONCAT, LEFT, RIGHT, MID, LEN, UPPER, LOWER, TRIM, TEXT, VALUE, FIND, SUBSTITUTE, REPT |
| Conditional | 3 | SUMIF, COUNTIF, COUNTBLANK |
| Lookup | 8 | VLOOKUP, HLOOKUP, INDEX, MATCH, ROW, COLUMN, ROWS, COLUMNS |
| Date/Time | 13 | TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DATEDIF, EDATE, EOMONTH, HOUR, MINUTE, SECOND |
| Trigonometry | 10 | PI, SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, DEGREES, RADIANS |
| Statistical | 8 | STDEV, STDEV.S, STDEV.P, STDEVP, VAR, VAR.S, VAR.P, VARP |
| Array | 5 | SEQUENCE, TRANSPOSE, SORT, UNIQUE, FILTER |

---

## gpui UI Integration Status

The engine works, but UI features for formulas need work:

### Working

| Feature | Status |
|---------|--------|
| Formula entry (=SUM(A1:A10)) | ✅ |
| Formula evaluation | ✅ |
| Display results | ✅ |
| Error display (#DIV/0!, etc.) | ✅ |
| Formula bar shows formula | ✅ |

### Not Yet Implemented in gpui

| Feature | Priority | Notes |
|---------|----------|-------|
| Formula syntax highlighting | P2 | Color refs, functions |
| Function autocomplete | P1 | Popup while typing |
| Signature help | P1 | Show args in tooltip |
| Spill border visualization | P2 | Blue border for arrays |
| F9 recalculate | P2 | Force recalc |
| Ctrl+` formula view toggle | P2 | Show formulas in cells |
| Alt+= AutoSum | P1 | Quick sum insertion |
| F4 cycle reference type | P2 | A1 → $A$1 → A$1 → $A1 |

---

## Missing Functions (Future)

### Multi-Condition (Not Yet)
- SUMIFS
- COUNTIFS
- AVERAGEIF
- AVERAGEIFS

### Lookup (Not Yet)
- XLOOKUP
- INDIRECT
- OFFSET

### Date (Not Yet)
- NETWORKDAYS
- WORKDAY

### Financial (Not Yet)
- PMT, IPMT, PPMT
- FV, PV, NPV
- IRR, XIRR
- RATE, NPER

### Statistical (Not Yet)
- PERCENTILE, QUARTILE
- RANK, LARGE, SMALL
- CORREL, FORECAST

---

## Array Formula Support

### Implemented

```
=SEQUENCE(5,3)      → 5x3 grid of numbers
=TRANSPOSE(A1:C3)   → Swap rows/cols
=SORT(A1:A10)       → Sort values
=UNIQUE(A1:A10)     → Deduplicate
=FILTER(A1:B10,A1:A10>5) → Filter rows
```

### Spill Behavior
- Parent cell shows result
- Spill cells receive overflow
- #SPILL! error if blocked
- Edit blocked on spill receivers

### gpui Visualization (TODO)
- [ ] Blue border on spill parent
- [ ] Light border on spill receivers
- [ ] Red border on #SPILL! cells
- [ ] Highlight full spill range on select

---

## Implementation Priority

### Sprint 1: Basic Polish
1. Alt+= AutoSum
2. F4 cycle reference type

### Sprint 2: IntelliSense
1. Function autocomplete popup
2. Signature help tooltip
3. Formula syntax highlighting

### Sprint 3: Array Visualization
1. Spill borders
2. Spill range highlighting

### Sprint 4: Advanced
1. Ctrl+` formula view
2. F9 recalculate
3. Formula auditing (trace dependents)
