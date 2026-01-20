# Future Features: VisiGrid gpui

Long-term vision beyond MVP. These are NOT current priorities.

---

## Core Positioning (Unchanged)

VisiGrid is building:
- A local-first computation surface
- A programmable decision engine people trust
- Native, fast, keyboard-driven spreadsheet

VisiGrid is NOT building:
- AI models
- Domain-specific templates
- SaaS collaboration platform

---

## Post-MVP Roadmap

### Phase 1: Feature Parity ✅ COMPLETE

Get gpui version to match what iced had:
- ✅ Command Palette
- ✅ Multi-sheet support
- ✅ Full keyboard shortcuts
- ✅ Dropdown menus

### Phase 2: Polish ✅ MOSTLY COMPLETE

- ✅ Themes (10+ themes, dark/light)
- ✅ Column/row resize
- ❌ Configurable keybindings
- ❌ Freeze panes
- ❌ Print to PDF

### Phase 3: Power Features ✅ COMPLETE

- ✅ Named ranges UI (create, rename, delete, extract)
- ✅ Excel import (xlsx/xls/xlsb/ods)
- ✅ Background import with progress overlay
- ✅ Import Report dialog (fidelity tracking)
- ✅ Inspector panel
- ✅ Formula autocomplete & signature help
- ✅ Syntax highlighting
- ❌ Data validation
- ❌ Conditional formatting (basic)
- ❌ Comments/notes
- ❌ Charts (basic)

### Phase 4: Systems of Record

From future-features.md - the differentiator:

**Concept:** Authoritative data ingress from trusted systems.

| Property | Requirement |
|----------|-------------|
| Read-only | Never write back to source |
| Local | Data lives on user's machine |
| Explicit refresh | User controls when |
| Snapshot-preserving | Every pull is versioned |

**Sources:**
- Stripe / payment processors
- QuickBooks / accounting
- Bank feeds
- Database exports
- CSV uploads

**NOT a sync engine.** Treat every source as "signed statement of facts at time T."

### Phase 5: AI Reconciliation

**Contract:** AI proposes, never mutates.

AI can:
- Compare records across sources
- Explain differences
- Flag anomalies
- Summarize discrepancies

AI cannot:
- Edit cells
- Fetch data
- Hide steps
- Invent values

**Workflow:**
1. User prompts: "Compare Stripe to bank deposits"
2. AI returns: comparison plan
3. User clicks: [Run]
4. VisiGrid executes: deterministically
5. Result: diff sheet + annotations

---

## Technical Debt to Address

### Before Phase 2 ✅ COMPLETE

| Issue | Priority | Status |
|-------|----------|--------|
| Proper dropdown menus | High | ✅ Done |
| Alt accelerators | High | ✅ Done |
| Context menus (right-click) | Medium | ✅ Done (sheets) |
| Window title (filename) | Low | ✅ Done |

### Before Phase 3 ✅ COMPLETE

| Issue | Priority | Status |
|-------|----------|--------|
| Multi-sheet persistence | High | ✅ Done |
| Named range persistence | High | ✅ Done |
| Format persistence (fonts, colors) | Medium | ✅ Done |
| Column width persistence | Medium | ✅ Done |

### Before Phase 4

| Issue | Priority |
|-------|----------|
| Plugin architecture | High |
| Credential storage | High |
| Background refresh | Medium |

---

## Non-Goals (Hard Boundaries)

These remain rejected:

| Non-Goal | Rationale |
|----------|-----------|
| VBA/macro compatibility | Lua scripting instead |
| Perfect Excel formatting | Diminishing returns |
| Real-time collaboration | Complexity; local-first |
| XLSX write/export | One-way import only; save as .sheet |
| Charts (v1) | Separate concern |
| Pivot tables (v1) | Too complex |
| Mobile/tablet | Desktop-first |

**Note:** XLSX *read* is now implemented (one-way import). Excel files open with background import, Import Report shows fidelity info, and unsupported functions are tracked for prioritization.

---

## Success Metrics

### MVP (gpui)
- [ ] 50% shortcut coverage
- [ ] Command palette works
- [ ] Multi-sheet works
- [ ] <300ms cold start
- [ ] 60fps scrolling

### v1.0
- [ ] 70% shortcut coverage
- [ ] Themes work
- [ ] Print to PDF
- [ ] 100+ GitHub stars

### v2.0
- [ ] Systems of Record (1 source)
- [ ] AI reconciliation (1 workflow)
- [ ] Plugin API
