# VisiGrid Roadmap

## In Progress

Features with significant implementation already shipped.

- **Data Validation** — Phases 0-6D complete (v2.5). Dropdown lists, number/date/text constraints, error alerts, XLSX import/export, circle invalid data. Spec: `docs/features/data-validation-spec.md`

- **XLSX Formatting Import** — Phase 1C complete. Cell colors, fonts, borders, alignment, custom number formats all parse and render. Spec: `docs/features/xlsx-import-formatting-spec.md`

- **Explainability** — Phases 0-11 complete. Dependency graph, verified mode, inspector, provenance, history panel, CLI replay. Spec: `docs/features/explainability-roadmap.md`

- **Context Menu** — Phases 1-2 complete (flat menus). Phases 3-5 blocked on gpui nested submenu support (upstream zed#19837). Spec: `docs/features/future/context-menu-spec.md`

- **Ask AI / AI Reconciliation** — Phases 0-3 complete. Ask AI + Explain Differences working, OpenAI provider. Phase 4 deferred. Spec: `docs/features/future/ai-reconciliation-spec.md`

- **Custom Functions** — Partial. Lua scripting exists for Pro features; needs formula integration and sandbox polish. Spec: `docs/features/future/custom-functions-spec.md`

## Near-term

Ready to build, no major blockers.

- **Cell Comments** — Text notes on cells with red triangle indicator, hover preview, edit/delete, bulk operations. Spec: `docs/features/cell-comments-spec.md`

- **Print to PDF** — Paginated PDF export with print preview, page setup, headers/footers, margins, scaling. Spec: `docs/features/print-to-pdf-spec.md`

- **Conditional Formatting** — Highlight rules, color scales, data bars, icon sets, top/bottom ranking, formula-based rules. Spec: `docs/features/conditional-formatting-spec.md`

- **Paste Special Phase 2-3** — Arithmetic paste operations (Add/Subtract/Multiply/Divide) and Transpose with formula reference adjustment. Phase 1 shipped in v0.3.5. Spec: `docs/features/future/paste-special-phase2-3.md`

- **Problems Panel** — Bottom panel aggregating all workbook errors, warnings, and info with filtering, navigation, and error help. Spec: `docs/features/future/problems-panel-spec.md`

- **Merged Cells extensions** — Merge Across (merge each row separately) and Merge/Unmerge in right-click context menu. Basic merge/unmerge shipped.

## Medium-term

Need infrastructure, design decisions, or upstream dependencies.

- **Range Picker** — Excel-style RefEdit control for selecting cell ranges from within modal dialogs. Used by Data Validation, Named Ranges, Ask AI. Spec: `docs/features/future/range-picker-spec.md`

- **Split View** — Horizontal/vertical/four-way panes for comparing distant regions with independent scrolling. Spec: `docs/features/future/split-view-spec.md`

- **Windows Title Bar** — Custom title bar integrating menu bar for native Windows look. Spec: `docs/features/windows-titlebar-spec.md`

- **Panel primitive** — Top-aligned, partial-opacity backdrop container for `preferences.rs` and similar panels. Blocked on 3+ call sites (currently only 1). Spec: `docs/features/done/ui-components.md`

- **Validation dialog** — Nested dropdowns require overlay stacking. Blocked on panel primitive.

## Long-term / Platform

Major infrastructure investment.

- **Plugin Architecture** — WASM-based sandboxed plugins for custom functions, data connectors, UI panels, and file formats. Permission system, signing, developer tools. Spec: `docs/features/future/plugin-architecture-spec.md`

- **Data Connectors** — SQL databases, REST APIs, GraphQL, S3. Query interface, parameterized queries, connected ranges with auto-refresh. Spec: `docs/features/future/data-connectors-spec.md`

- **Systems of Record** — Read-only SaaS integrations (Stripe, QuickBooks, Plaid, Salesforce). OAuth, versioned snapshots, comparison UI. Spec: `docs/features/future/systems-of-record-spec.md`

- **Minimap** — Bird's-eye view sidebar showing data density, viewport position, error/formula highlighting, quick navigation. Spec: `docs/features/future/minimap-spec.md`
