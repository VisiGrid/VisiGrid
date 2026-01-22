# VisiGrid CLI v1 Specification

Command-line interface for headless spreadsheet operations.

> **Status:** Shipped (v0.1.8). v1 scope locked.
>
> **Changes to this spec require an RFC.** This is the contract.

---

## Principles

1. **Explicit over magic** - User declares format. No guessing.
2. **Deterministic** - Same input = same output. No locale, no timestamps.
3. **Fast** - <50ms cold start. Lazy loading only.
4. **Loud failure** - Non-zero exit on any error. Stderr for diagnostics, stdout for data.

---

## Architecture

```
crates/cli/     → visigrid-cli (headless, no gpui)
gpui-app/       → visigrid (GUI)
```

CLI depends only on `engine`, `io`, `core`. No GUI dependencies.

---

## v1 Commands

Three commands ship in v1:

- `convert` - format conversion
- `calc` - formula evaluation
- `list-functions` - function catalog

---

## Format Resolution

**Rule:**

| Input Source | `--from` | Behavior |
|--------------|----------|----------|
| stdin | required | Use `--from` value |
| file | omitted | Infer from extension |
| file | provided | `--from` overrides extension |

Unknown extension without `--from` → exit 2: `error: cannot infer format, use --from`

**Supported formats:**

| Format | Extensions | Read | Write |
|--------|------------|------|-------|
| csv | .csv | ✓ | ✓ |
| tsv | .tsv | ✓ | ✓ |
| json | .json | ✓ | ✓ |
| lines | — | ✓ | ✓ |
| xlsx | .xlsx, .xls | ✓ | — |
| sheet | .sheet | ✓ | ✓ |

`lines` = one value per line, single column.

---

## JSON Determinism Rules

**Input (array-of-arrays):**
```json
[[1, 2], [3, 4]]
```
Maps directly to grid. Row-major order.

**Input (array-of-objects):**
```json
[{"a": 1, "b": 2}, {"a": 3}]
```

- Column order = lexicographic sort of union of all keys
- Missing keys → empty cell
- Non-scalar values (nested arrays/objects) → exit 4: `error: non-scalar value at row N, key "K"`

**Output (array-of-arrays):**
```json
[[1, 2], [3, 4]]
```
Default. No keys.

**Output (array-of-objects):**
Requires `--headers`. First row values become keys.
```json
[{"col1": 3, "col2": 4}]
```
Keys sanitized: lowercase, spaces → `_`, invalid chars stripped.

---

## Unicode

CLI does not normalize Unicode. Input and output preserve codepoints exactly as provided.

---

## `visigrid-cli convert`

Convert between formats.

```
visigrid-cli convert [INPUT] --to FORMAT [OPTIONS]
```

**Examples:**
```bash
# File to stdout
visigrid-cli convert data.xlsx --to csv

# File to file
visigrid-cli convert data.xlsx --to csv --output data.csv

# Stdin to stdout
cat data.csv | visigrid-cli convert --from csv --to json

# Override inferred format
visigrid-cli convert data.txt --from tsv --to csv
```

**Arguments:**

| Argument | Required | Description |
|----------|----------|-------------|
| `INPUT` | No | Input file. Omit for stdin. |
| `--from FORMAT` | If stdin or override | Input format |
| `--to FORMAT` | Yes | Output format |
| `--output FILE` | No | Output file. Omit for stdout. |
| `--sheet NAME` | No | Sheet name for multi-sheet files. Default: first. |
| `--delimiter CHAR` | No | CSV/TSV delimiter. Default: `,` for csv, `\t` for tsv. |
| `--headers` | No | First row is headers. Affects JSON object keys. |

**No positional output.** Use `--output` explicitly.

**CSV/TSV output:** RFC 4180-compliant quoting. Values containing delimiter, quote, or newline are quoted; embedded quotes are doubled.

---

## `visigrid-cli calc`

Evaluate formula against stdin data.

```
visigrid-cli calc FORMULA --from FORMAT [OPTIONS]
```

**Examples:**
```bash
# Sum a column of numbers
echo -e "10\n20\n30" | visigrid-cli calc "=SUM(A:A)" --from lines
# stdout: 60

# Average from CSV
cat sales.csv | visigrid-cli calc "=AVERAGE(B:B)" --from csv
# stdout: 4500.5

# With spill output
cat matrix.csv | visigrid-cli calc "=TRANSPOSE(A1:C3)" --from csv --spill csv
```

**Arguments:**

| Argument | Required | Description |
|----------|----------|-------------|
| `FORMULA` | Yes | Formula to evaluate. Must start with `=`. |
| `--from FORMAT` | Yes | Stdin format: csv, tsv, lines, json |
| `--into CELL` | No | Load data starting at cell. Default: A1. |
| `--delimiter CHAR` | No | CSV delimiter. Default: `,` |
| `--headers` | No | First row is headers. See Header Behavior. |
| `--spill FORMAT` | No | Output format if result is array: csv, json. |

**Loading rules:**

| Format | Behavior |
|--------|----------|
| `lines` | Column A, starting at `--into` |
| `csv` | Grid starting at `--into` |
| `tsv` | Grid starting at `--into` |
| `json` | Grid starting at `--into` (see JSON rules) |

**Empty stdin:** Exit 4 with `error: empty input`. No silent empty-sheet behavior.

**Output rules:**

| Result | Behavior |
|--------|----------|
| Scalar | Print raw value to stdout, exit 0 |
| Error token | Print token (e.g., `#DIV/0!`) to stdout, diagnostic to stderr, exit 1 |
| Array (spill) | If `--spill` provided, output in that format. Else exit 1: `error: result is array, use --spill` |

**Number output:** Raw numeric value, full precision. No locale formatting.

- Integers print without decimal point: `2` not `2.0`
- Floats use minimal representation to round-trip: `1.5` not `1.50000`
- No scientific notation unless engine produces it

Example: `1234.5678` not `1,234.57` or `$1,234.57`.

---

## Header Behavior

`--headers` affects:

1. **JSON output** - First row values become object keys
2. **Data loading** - First row excluded from formula range (headers don't count as data)

`--headers` does NOT affect:

- Formula identifiers. Use `B:B`, not `revenue`.

**Future consideration:** `--columns-from-headers` flag to enable mapping header names to named ranges. Not in v1.

---

## `visigrid-cli list-functions`

Print supported functions.

```
visigrid-cli list-functions
```

**Output:**
```
ABS
AVERAGE
AVERAGEIF
...
XLOOKUP
YEAR
```

One function per line, sorted alphabetically. Suitable for `grep`, `wc -l`.

**No `--verbose` in v1.** Function signatures ship later.

---

## `visigrid-cli open`

Launch GUI.

```
visigrid-cli open [FILE]
```

Shells out to `visigrid-gui` (Linux/Windows) or `VisiGrid.app` (macOS).

---

## Spill Behavior

**Definition:** Result is a spill if it's a rectangular array with rows × cols > 1.

**1×1 arrays are scalar.** They print the single value and exit 0. No `--spill` required.

**Without `--spill`:**
```
error: result is 3x2 array, use --spill csv or --spill json
```
Exit 1.

**With `--spill csv`:**
Output as CSV to stdout.

**With `--spill json`:**
Output as array-of-arrays. Never objects (regardless of `--headers`).

---

## Exit Codes

| Code | Category | Meaning |
|------|----------|---------|
| 0 | Success | Operation completed |
| 1 | Eval error | Formula error (#VALUE!, #REF!, spill without --spill) |
| 2 | Args error | Invalid arguments, missing required flags |
| 3 | IO error | File not found, permission denied, write failed |
| 4 | Parse error | Malformed input (bad CSV, invalid JSON, etc.) |
| 5 | Format error | Unsupported or unknown format |

---

## stdout / stderr Contract

| Situation | stdout | stderr | Exit |
|-----------|--------|--------|------|
| Success (scalar) | result value | — | 0 |
| Success (convert) | converted data | — | 0 |
| Formula error | error token | `error: ...` | 1 |
| Spill without flag | — | `error: result is NxM array...` | 1 |
| Bad args | — | `error: ...` | 2 |
| IO failure | — | `error: ...` | 3 |
| Parse failure | — | `error: ... at line N` | 4 |
| Bad format | — | `error: unknown format 'X'` | 5 |

**Rule:** stdout is the data stream. stderr is diagnostics. Exit code is truth for pipelines.

---

## Deferred (Post-v1)

| Command | Reason |
|---------|--------|
| `diff` | Needs design: values vs formulas, tolerance, ordering |
| `import` | Mutation risk, needs atomic writes |
| `batch` | Lua runtime adds binary weight |
| `export` | Redundant with `convert` |
| `examples` | Add after core commands are stable |
| `--verbose` | Requires shipping docstrings, impacts binary size |
| `--display` | Locale-sensitive formatting, complex edge cases |
| `--columns-from-headers` | Named range mapping, sanitization rules |

---

## Testing

**Golden tests:** Committed input + expected output.

```
tests/cli/
  convert/
    xlsx-to-csv/
      input.xlsx
      args.txt          # --to csv
      expected.stdout
      expected.exit     # 0
    malformed-csv/
      input.csv
      args.txt
      expected.stderr
      expected.exit     # 4
  calc/
    sum-lines/
      input.txt
      args.txt          # "=SUM(A:A)" --from lines
      expected.stdout   # 60
      expected.exit     # 0
    div-zero/
      input.txt
      args.txt
      expected.stdout   # #DIV/0!
      expected.stderr   # error: division by zero
      expected.exit     # 1
```

**Edge cases to cover:**

- CSV with quotes, commas, newlines in values
- JSON objects with missing keys
- JSON with non-scalar values (expect exit 4)
- Empty input
- Single cell input
- Formula referencing out-of-bounds
- Unicode in values and headers

**Platform parity:** Run all golden tests on Linux, macOS, Windows in CI.

---

## Performance

| Metric | Target | Actual |
|--------|--------|--------|
| Cold start | <50ms | ~3ms (`time visigrid-cli list-functions > /dev/null`) |
| calc 10K rows | <500ms | TBD |
| convert 1MB CSV | <1s | TBD |

**CI enforcement:** Benchmark on every PR. Block merge if regression >10%.

**Implementation constraint:** Function registry loads lazily. No startup-time initialization of all 97 functions.

---

## Distribution

| Platform | CLI | GUI |
|----------|-----|-----|
| macOS | `visigrid-cli` in `VisiGrid.app/Contents/MacOS/` | `VisiGrid.app` |
| Windows | `visigrid-cli.exe` | `visigrid.exe` |
| Linux | `visigrid-cli` | `visigrid` |

Both binaries are included in all release packages.
