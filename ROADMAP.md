# VisiGrid Roadmap

This document tracks what's built, what's in progress, and what's intentionally not planned.

> This roadmap reflects current functionality and direction, not guarantees.
> Some features may be partial, evolving, or subject to change prior to v1.0.

---

## Shipped

### Core Spreadsheet (Stable)
- GPU-accelerated grid rendering (GPUI)
- Cell editing with formula bar
- Selection, multi-selection, and range operations
- Undo/redo
- Clipboard integration (copy/paste)
- Fill down (Ctrl+D) and fill right (Ctrl+R)

### Navigation & Workflow
- Keyboard-driven navigation (arrow keys, Ctrl+arrows, Home/End)
- Command palette (Ctrl+Shift+P)
- Go To dialog (Ctrl+G)
- Find (with incremental search)
- Session restore (reopen files, scroll position, selection, panels)

### File Formats
- Native `.sheet` format (preserves values, formulas, formatting)
- CSV/TSV import and export
- JSON export
- Excel import (common cases; fidelity varies by file)

### Formula Engine (96 built-in functions)

Broad coverage across:
- Math, logical, text, lookup, date/time, statistical, and array formulas
- Dynamic arrays (spill behavior)
- Dependency tracking and recalculation

Exact function list available in documentation.

### Multi-Sheet & Organization
- Multiple sheets per workbook
- Sheet tabs (rename, delete, reorder)
- Named ranges (create, rename, delete)

### Platform
- macOS (Universal binary, signed and notarized)
- Windows (x64)
- Linux (x86_64, tar.gz and AppImage)

### Pro Features (Advanced)
- Background Excel import with progress reporting
- Import Report (fidelity tracking, unsupported features)
- Inspector panel (dependencies, diagnostics)

### Pro Platform Capabilities (Evolving)
- Lua scripting console (Ctrl+Shift+L)
- Configurable keybindings (`~/.config/visigrid/keybindings.json`)

---

## In Progress

### Polish
- Freeze panes (lock rows/columns while scrolling)

### Formula Coverage
- Multi-condition functions: SUMIFS, COUNTIFS, AVERAGEIF, AVERAGEIFS
- Modern lookup: XLOOKUP
- Financial functions: PMT, FV, PV, NPV, IRR

---

## Planned

### Near-term
- Find and replace
- Data validation (dropdowns, constraints)
- Conditional formatting (basic rules)
- Print to PDF
- Comments/notes on cells

### Long-term: Systems of Record

Connect to authoritative data sources without sync risk.

**Concept:** Import data from ERPs, payment processors, and ledgers as read-only snapshots. Every pull is versioned with a timestamp. User controls when to refresh. No silent syncs. No write-back.

**Target sources:**
- Stripe / payment processors
- QuickBooks / accounting systems
- Bank feeds
- Database exports

### Long-term: AI Reconciliation

AI helps you reason about data without modifying it.

**Concept:** AI proposes operations (compare, flag, explain differences). User reviews and approves. VisiGrid executes deterministically. Every step is visible and auditable.

**AI can:**
- Compare records across sources
- Explain differences
- Flag anomalies
- Summarize discrepancies

**AI cannot:**
- Edit cells without approval
- Fetch data directly
- Hide steps
- Invent values

---

## Non-Goals

These are explicit rejections, not "not yet":

| Feature | Rationale |
|---------|-----------|
| VBA/macro compatibility | Lua scripting instead |
| XLSX export | One-way import; save as .sheet |
| Real-time collaboration | Local-first philosophy |
| Web version | Desktop-native performance |
| Mobile/tablet | Desktop-first workflows |
| Pivot tables (v1) | Complexity; later if demand |
| Charts (v1) | Separate concern; later |
| Perfect Excel formatting | Diminishing returns |

---

## Versioning

VisiGrid uses [CalVer](https://calver.org/) for releases: `YYYY.MM.PATCH`

Current status: Early access (pre-v1.0). Expect breaking changes to the `.sheet` format.
