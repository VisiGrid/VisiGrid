# Phase 2: Agent-Ready Kit

**Status:** After Phase 1 (Session Server)
**Effort:** 3-5 days (mostly documentation + examples)
**Prerequisite:** Phase 1 (Session Server) for structured errors

## What It Does

Package VisiGrid as a **tool** that LLM agents can call. Provide:

1. Formal tool definition (MCP, OpenAI function calling, Claude tools)
2. Example prompts and guardrails
3. Demo scripts showing the workflow
4. CLAUDE.md snippet for Claude Code users

This is the **launch moment** for "VisiGrid for Agents."

## Tool Definition (MCP Format)

MCP (Model Context Protocol) is becoming the standard for tool definitions.

```json
{
  "name": "visigrid",
  "version": "1.0.0",
  "description": "Deterministic spreadsheet engine. Use for calculations, data reconciliation, and building financial models. Provides immediate structured feedback on formulas.",
  "tools": [
    {
      "name": "calc",
      "description": "Evaluate a spreadsheet formula against data. Use for SUM, AVERAGE, VLOOKUP, conditional aggregations, etc. Returns the computed value or a structured error.",
      "inputSchema": {
        "type": "object",
        "properties": {
          "formula": {
            "type": "string",
            "description": "Spreadsheet formula starting with '='. Examples: '=SUM(A:A)', '=VLOOKUP(\"key\",A:B,2,FALSE)'"
          },
          "data": {
            "type": "string",
            "description": "CSV data to evaluate against (piped to stdin)"
          },
          "headers": {
            "type": "boolean",
            "description": "Whether first row contains headers"
          }
        },
        "required": ["formula", "data"]
      }
    },
    {
      "name": "diff",
      "description": "Reconcile two datasets by key column. Reports matches, missing rows, and value differences. Handles financial formats ($1,234.56) and numeric tolerance for rounding.",
      "inputSchema": {
        "type": "object",
        "properties": {
          "left": { "type": "string", "description": "CSV data for left side" },
          "right": { "type": "string", "description": "CSV data for right side" },
          "key": { "type": "string", "description": "Key column name or letter" },
          "tolerance": { "type": "number", "description": "Numeric tolerance for differences (default: 0)" }
        },
        "required": ["left", "right", "key"]
      }
    },
    {
      "name": "set_cell",
      "description": "Write a value or formula to a cell in a VisiGrid file. Use for building models programmatically.",
      "inputSchema": {
        "type": "object",
        "properties": {
          "file": { "type": "string", "description": "Path to .sheet file" },
          "cell": { "type": "string", "description": "Cell reference (e.g., 'A1', 'B2')" },
          "value": { "type": "string", "description": "Value or formula (formulas start with '=')" }
        },
        "required": ["file", "cell", "value"]
      }
    },
    {
      "name": "inspect_cell",
      "description": "Read cell value, formula, format, and dependencies. Use to understand existing spreadsheet logic.",
      "inputSchema": {
        "type": "object",
        "properties": {
          "file": { "type": "string", "description": "Path to .sheet file" },
          "cell": { "type": "string", "description": "Cell reference" },
          "include_deps": { "type": "boolean", "description": "Include dependency graph" }
        },
        "required": ["file", "cell"]
      }
    },
    {
      "name": "apply_ops",
      "description": "Apply a batch of operations to a running VisiGrid instance. Use for live model building with immediate visual feedback.",
      "inputSchema": {
        "type": "object",
        "properties": {
          "session": { "type": "string", "description": "Session ID from 'vgrid sessions'" },
          "ops": {
            "type": "array",
            "items": {
              "type": "object",
              "properties": {
                "op": { "enum": ["set_value", "set_formula", "clear", "style", "format"] },
                "cell": { "type": "string" },
                "range": { "type": "string" },
                "value": { "type": "string" },
                "formula": { "type": "string" },
                "bold": { "type": "boolean" },
                "format": { "type": "string" }
              }
            }
          }
        },
        "required": ["session", "ops"]
      }
    }
  ]
}
```

## CLAUDE.md Snippet

For Claude Code users, add this to project CLAUDE.md:

```markdown
## VisiGrid CLI

VisiGrid is available for spreadsheet calculations and data reconciliation.

### Quick Reference

```bash
# Calculate formula against CSV data
cat data.csv | vgrid calc "=SUM(A:A)" --from csv --headers

# Reconcile two datasets
vgrid diff expected.csv actual.csv --key id --tolerance 0.01

# Build a model cell by cell
vgrid set model.sheet A1 "Revenue" --create
vgrid set model.sheet B1 "=SUM(B2:B100)"

# Inspect cell dependencies
vgrid inspect model.sheet B1 --deps
```

### When to Use

- **calc**: Evaluate formulas (SUM, VLOOKUP, SUMIF, etc.) against data
- **diff**: Compare datasets with tolerance for rounding differences
- **set/inspect**: Build or analyze spreadsheet models

### Error Handling

Errors are structured JSON on stderr. Read the `message` and `suggestion` fields
to self-correct:

```json
{
  "error": "formula_error",
  "cell": "B1",
  "message": "Circular reference",
  "suggestion": "Cell B1 references itself via B1 → C1 → B1"
}
```

### Supported Functions

96+ spreadsheet functions including:
SUM, AVERAGE, COUNT, VLOOKUP, HLOOKUP, INDEX, MATCH, IF, SUMIF, COUNTIF,
SUMIFS, AVERAGEIF, LEFT, RIGHT, MID, CONCATENATE, TEXT, DATE, TODAY, etc.

Run `vgrid list-functions` for full list.
```

## Example Prompts for Agents

### Prompt 1: Data Reconciliation

```
You have access to vgrid for spreadsheet operations.

Task: Reconcile the vendor invoice (vendor.csv) against our ledger (ledger.csv).
Use the Invoice Number column as the key. Allow $0.01 tolerance for rounding.

Report:
1. How many invoices match exactly?
2. Which invoices are missing from our ledger?
3. Which invoices have amount differences outside tolerance?
```

### Prompt 2: Build a Financial Model

```
You have access to vgrid for spreadsheet operations.

Task: Build a revenue projection model in projection.sheet with:
- Monthly growth rate input in A2 (default 5%)
- Base revenue in B2 (default $100,000)
- 12 months of projected revenue in C2:C13 using compound growth
- Headers in row 1, bold formatting
- Currency format for all revenue cells

Use vgrid set to create the model. Verify calculations with calc.
```

### Prompt 3: Audit Existing Model

```
You have access to vgrid for spreadsheet operations.

Task: Audit model.sheet for issues:
1. Use inspect to trace the dependency graph from the output cell (Z100)
2. Identify any circular references
3. Check if any "input" cells (colored blue per finance convention) contain formulas
4. Report findings
```

## Guardrails / Safety

### Agent Behavior Rules

Include in system prompt:

```
When using VisiGrid:

1. VERIFY before committing: After building a model, use `calc` to verify
   key output cells match expectations before declaring success.

2. INSPECT before modifying: Before changing an existing model, use `inspect`
   to understand the current cell's dependencies.

3. BATCH related changes: Group related cell changes into a single `apply_ops`
   call to maintain consistency and trigger only one recalc.

4. REPORT errors clearly: If VisiGrid returns an error, include the full
   error message in your response and explain what went wrong.

5. NEVER hallucinate values: If you need a calculation result, use `calc`.
   Do not guess or approximate.
```

### Error Recovery Pattern

```
When VisiGrid returns an error:

1. Read the `message` and `suggestion` fields
2. Identify the specific issue (circular ref, bad syntax, missing data)
3. Modify your command to fix the issue
4. Retry
5. If it fails 3 times, report the issue to the user with full context
```

## Demo Scripts

### demo-reconciliation.sh

```bash
#!/bin/bash
# Demo: Agent-style reconciliation workflow

# Step 1: Agent checks the data shape
echo "Inspecting left dataset..."
head -3 vendor.csv

echo "Inspecting right dataset..."
head -3 ledger.csv

# Step 2: Agent runs diff
echo "Running reconciliation..."
vgrid diff vendor.csv ledger.csv --key "Invoice Number" --tolerance 0.01

# Step 3: Agent summarizes (this would be LLM output)
echo "Summary: X matched, Y only in vendor, Z with differences"
```

### demo-live-build.sh

```bash
#!/bin/bash
# Demo: Build a model with live preview

# Terminal 1: Start GUI in watch mode
# vgrid open projection.sheet --watch

# Terminal 2: Build the model
SESSION=$(vgrid sessions --json | jq -r '.[0].id')

vgrid apply --session $SESSION <<'EOF'
{"op":"set_value","cell":"A1","value":"Growth Rate"}
{"op":"set_value","cell":"A2","value":"0.05"}
{"op":"set_value","cell":"B1","value":"Base Revenue"}
{"op":"set_value","cell":"B2","value":"100000"}
{"op":"set_value","cell":"C1","value":"Month 1"}
{"op":"set_formula","cell":"C2","formula":"=B2*(1+$A$2)"}
{"op":"style","range":"A1:C1","bold":true}
{"op":"format","range":"B2:C2","format":"currency_usd"}
EOF

echo "Model built. Check the GUI!"
```

## Marketing Assets

### One-liner

> "VisiGrid: The spreadsheet compiler for LLMs"

### Elevator pitch

> "Most AI agents hallucinate math. VisiGrid gives them a high-speed
> arithmetic co-processor with structured error feedback and a
> verifiable audit trail."

### Technical differentiator

> "We're not claiming AI is smarter at VisiGrid. We're claiming the
> feedback loop is tighter. The agent interacts with a reactive engine
> that handles math and dependencies, leaving the agent to focus on
> logic — which we verify via deterministic fingerprints."

### Comparison hook

> "Why not Python/Pandas? Because you can't hand a Pandas script to a CFO.
> VisiGrid provides database performance with spreadsheet explainability."

## Files to Create

| File | Description |
|------|-------------|
| `docs/agent-tools.json` | MCP tool definition |
| `docs/claude-md-snippet.md` | Copy-paste for CLAUDE.md |
| `examples/agent-reconciliation.sh` | Demo script |
| `examples/agent-model-build.sh` | Demo script |
| `examples/agent-prompts.md` | Example prompts |

## Success Criteria

1. MCP tool definition validates against MCP schema
2. CLAUDE.md snippet works with Claude Code out of the box
3. Demo scripts run successfully
4. At least one "agent builds a spreadsheet" video/gif for launch
5. README includes "Agent-Ready" section with link to docs
