# VisiGrid Agent Instructions

Copy this into your CLAUDE.md or AGENTS.md to enable verifiable spreadsheet builds.

---

## VisiGrid CLI

VisiGrid is a deterministic spreadsheet engine. Use it for calculations, data reconciliation, and building verifiable financial models.

### Available Commands

```bash
# Evaluate formulas against CSV data
echo "amount\n100\n200" | vgrid calc "=SUM(A:A)" --from csv --headers

# Reconcile two datasets
vgrid diff expected.csv actual.csv --key id --tolerance 0.01 --out json

# Build a .sheet from Lua script
vgrid sheet apply model.sheet --lua build.lua --json

# Inspect cells/workbook
vgrid sheet inspect model.sheet A1 --json
vgrid sheet inspect model.sheet A1:D10 --json
vgrid sheet inspect model.sheet --json  # workbook metadata

# Get fingerprint
vgrid sheet fingerprint model.sheet --json

# Verify fingerprint
vgrid sheet verify model.sheet --fingerprint v1:42:abc123...
```

### Lua Build API

When writing Lua scripts for `sheet apply`:

```lua
set("A1", "Revenue Model")    -- set value or formula (affects fingerprint)
set("B2", 10000)              -- numbers work directly
set("C3", "=SUM(B:B)")        -- formulas start with =
clear("D4")                   -- clear cell (affects fingerprint)
meta("A1", { role = "title" }) -- semantic metadata (affects fingerprint)
style("A1", { bold = true })   -- presentation only (does NOT affect fingerprint)
```

### Fingerprint Boundary (Critical)

| Function | Affects Fingerprint |
|----------|---------------------|
| `set()` | Yes |
| `clear()` | Yes |
| `meta()` | Yes |
| `style()` | **No** |

This means you can format sheets without breaking verification. Style is presentation, not semantics.

### Workflow Rules

1. **Always use `--json` for tool calls.** Never parse table output.

2. **Never assume results.** After `sheet apply`, always `sheet inspect` to verify values.

3. **Always verify before declaring success.** The workflow is:
   ```
   sheet apply → sheet inspect → sheet verify
   ```

4. **Capture fingerprint before modifications.** If you need to prove what changed:
   ```bash
   BEFORE=$(vgrid sheet fingerprint model.sheet --json | jq -r .fingerprint)
   # ... make changes ...
   AFTER=$(vgrid sheet fingerprint model.sheet --json | jq -r .fingerprint)
   ```

5. **Use `--dry-run` to preview.** Before writing, verify the fingerprint:
   ```bash
   vgrid sheet apply model.sheet --lua build.lua --dry-run --json
   ```

### Error Handling

Errors are JSON on stderr. Check exit codes:
- `0` = success
- `1` = verification failed / diffs found
- `2` = usage error

Example error:
```json
{
  "ok": false,
  "error": "fingerprint_mismatch",
  "expected": "v1:42:abc123...",
  "computed": "v1:42:def456..."
}
```

### Example: Build and Verify a Model

```bash
# Write the Lua script
cat > model.lua << 'EOF'
set("A1", "Q1 Revenue")
set("B1", 100000)
set("A2", "Q2 Revenue")
set("B2", 120000)
set("A3", "Total")
set("B3", "=SUM(B1:B2)")
meta("B1:B2", { type = "input" })
style("A3:B3", { bold = true })
EOF

# Build
vgrid sheet apply model.sheet --lua model.lua --json

# Inspect the total
vgrid sheet inspect model.sheet B3 --json
# → {"cell":"B3","value":"220000","formula":"=SUM(B1:B2)","value_type":"formula"}

# Get fingerprint for future verification
vgrid sheet fingerprint model.sheet --json
# → {"file":"model.sheet","fingerprint":"v1:7:abc123...","ops":7}

# Verify (e.g., in CI)
vgrid sheet verify model.sheet --fingerprint v1:7:abc123...
# → Verification: PASS
```

### Supported Functions

96+ spreadsheet functions including:
SUM, AVERAGE, COUNT, VLOOKUP, HLOOKUP, INDEX, MATCH, IF, SUMIF, COUNTIF,
SUMIFS, AVERAGEIF, LEFT, RIGHT, MID, CONCATENATE, TEXT, DATE, etc.

Run `vgrid list-functions` for the full list.
