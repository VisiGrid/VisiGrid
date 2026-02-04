# VisiGrid Roadmap

## Future

### UI Components

- **`ui/popup.rs` extraction** — Unify the three existing context menus (sheet tab, history, cell/header) into a single reusable popup primitive. API: `Popup::new(position).child(items)` — handles clamping, outside-click dismissal, shadow, border, padding. Optional `max_height` for scroll. Migrate all three menus, then lock styles. Meets freeze policy (3+ call sites).

### Merged Cells

- **Merge Across** — Merge each row in the selection separately. Selection A1:C3 produces three independent merges (A1:C1, A2:C2, A3:C3). Data loss warning lists affected cells per row. Requires UX decision on menu placement and shortcut. See `docs/features/done/merged-cells-spec.md` for full spec.

- **Context menu entries** — Show "Merge Cells" / "Unmerge Cells" conditionally in the cell right-click menu (Merge only when selection can merge; Unmerge only when selection contains a merge). Context menu infrastructure now exists (`views/context_menu.rs`).
