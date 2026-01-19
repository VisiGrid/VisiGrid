# VisiGrid gpui Design Document

A lightweight, native, GPU-accelerated spreadsheet built with gpui (Zed's UI framework).

---

## Migration Status: iced → gpui

VisiGrid was originally built with iced. This document tracks the gpui rebuild.

### Why gpui?

| Benefit | Details |
|---------|---------|
| GPU-accelerated | Native Wayland/Metal rendering |
| Zed-proven | Battle-tested in production editor |
| Modern Rust | Clean async patterns, good DX |
| Wayland-native | First-class Linux desktop support |

---

## Current Implementation Status

### Core Features

| Feature | Status | Notes |
|---------|--------|-------|
| Basic grid with cell selection | ✅ Done | Dynamic row/col sizing |
| Formula bar | ✅ Done | Shows cell ref + value |
| Keyboard navigation | ✅ Done | Arrows, Tab, Enter, Ctrl+Arrows |
| Cell editing | ✅ Done | F2, typing starts edit |
| Copy/Cut/Paste | ✅ Done | Ctrl+C/X/V |
| Undo/Redo | ✅ Done | Ctrl+Z/Y |
| Multi-cell selection | ✅ Partial | Shift+Arrow only, no Ctrl+Click |
| Go To (Ctrl+G) | ✅ Done | |
| Find (Ctrl+F) | ✅ Done | F3/Shift+F3 for next/prev |
| Bold/Italic/Underline | ✅ Done | Ctrl+B/I/U |
| Native format save/load | ✅ Done | SQLite .sheet files |
| CSV export | ✅ Done | |
| Toolbar menu bar | ✅ Done | Buttons, not dropdowns |
| Dynamic grid sizing | ✅ Done | Fills window, resizes |
| Double-click to edit | ✅ Done | |
| Mouse wheel scroll | ✅ Done | |
| Page Up/Down | ✅ Done | |

### Missing from MVP (Priority Order)

| Feature | Priority | Effort | Notes |
|---------|----------|--------|-------|
| Command Palette | P0 | Medium | Core differentiator |
| Fill Down/Right | P0 | Easy | Ctrl+D/R |
| Multi-edit | P0 | Medium | Typing replaces ALL selected |
| Ctrl+Click selection | P1 | Easy | Discontiguous ranges |
| Dropdown menus | P1 | Medium | Replace toolbar buttons |
| Alt menu accelerators | P1 | Medium | Alt+F, Alt+E, etc. |
| Sheet tabs | P1 | Hard | Multi-sheet support |
| Number format shortcuts | P2 | Easy | Ctrl+Shift+$, %, etc. |
| Format Cells dialog | P2 | Medium | Ctrl+1 |
| Replace (Ctrl+H) | P2 | Easy | |
| CSV import | P2 | Easy | |
| Zoom | P2 | Easy | Ctrl++/- |
| Context menu | P3 | Medium | Right-click |
| Zen Mode | P3 | Easy | F11 |
| Split view | P3 | Hard | |

### Editor-Style Features (VS Code Inspired)

| Feature | Priority | Status |
|---------|----------|--------|
| Command Palette (Ctrl+Shift+P) | P0 | ❌ |
| Quick Open (Ctrl+P) | P2 | ❌ |
| Problems Panel (Ctrl+Shift+M) | P3 | ❌ |
| Cell Inspector (Ctrl+Shift+I) | P3 | ❌ |
| Context Help (F1) | P2 | ❌ |
| Go to Definition (F12) | P3 | ❌ |
| Find All References (Shift+F12) | P3 | ❌ |

---

## Architecture

### Current Structure

```
gpui-app/
├── src/
│   ├── main.rs           # Entry point
│   ├── app.rs            # Spreadsheet state, methods
│   ├── actions.rs        # gpui actions
│   ├── keybindings.rs    # Key bindings
│   ├── mode.rs           # Navigation/Edit/GoTo/Find modes
│   ├── history.rs        # Undo/redo
│   ├── file_ops.rs       # File operations
│   └── views/
│       ├── mod.rs        # Main render + action handlers
│       ├── grid.rs       # Cell grid rendering
│       ├── headers.rs    # Row/column headers
│       ├── formula_bar.rs
│       ├── status_bar.rs
│       ├── menu_bar.rs   # Toolbar
│       ├── goto_dialog.rs
│       └── find_dialog.rs
```

### Engine (Shared with iced version)

```
crates/
├── engine/    # Core spreadsheet logic
│   ├── sheet.rs      # Sheet data structure
│   ├── cell.rs       # Cell types, formats
│   ├── parser.rs     # Formula parser
│   ├── eval.rs       # Formula evaluator
│   └── functions.rs  # 96 built-in functions
├── core/      # Shared types
├── io/        # File format handling
└── config/    # Settings
```

---

## Implementation Plan

### Phase 1: Core Parity (Current Focus)

1. ✅ Basic grid navigation
2. ✅ Cell editing
3. ✅ Copy/paste
4. ✅ Undo/redo
5. ✅ Save/load
6. ✅ Find/GoTo
7. ⏳ Command Palette
8. ⏳ Fill Down/Right
9. ⏳ Multi-edit

### Phase 2: Excel Compatibility

1. Dropdown menus
2. Alt accelerators
3. Number format shortcuts
4. Format Cells dialog (Ctrl+1)
5. Replace (Ctrl+H)
6. Ctrl+Click selection

### Phase 3: Multi-Sheet

1. Sheet tabs UI
2. Ctrl+PageUp/Down navigation
3. Cross-sheet references
4. Sheet management (add/delete/rename)

### Phase 4: Power Features

1. Named ranges UI
2. Array formula visualization (spill borders)
3. Zoom
4. Split view
5. Zen mode

---

## Key Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| UI framework | gpui | GPU-accelerated, Wayland-native |
| Window sizing | Dynamic | Calculate visible rows/cols from window size |
| Menu style | Toolbar (v1) | Simpler; dropdowns in v2 |
| Mode system | Enum | Navigation/Edit/GoTo/Find |
| Selection | Anchor+End | Supports range selection |
| History | Per-action | Single undo step per user action |

---

## Performance Targets

| Metric | Target | Current |
|--------|--------|---------|
| Cold start | <300ms | ~200ms ✅ |
| Scroll (65k rows) | 60fps | ✅ |
| Keystroke latency | <16ms | ✅ |
| Binary size | <30MB | ~11MB ✅ |

---

## Next Steps

1. **Command Palette** - Fuzzy search all commands
2. **Fill Down/Right** - Ctrl+D/R
3. **Multi-edit** - Typing affects all selected cells
4. **Dropdown menus** - Proper File/Edit/View/Format menus
