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
| Multi-cell selection | ✅ Done | Shift+Arrow, Ctrl+Shift+Arrow |
| Go To (Ctrl+G) | ✅ Done | |
| Find (Ctrl+F) | ✅ Done | F3/Shift+F3 for next/prev |
| Bold/Italic/Underline | ✅ Done | Ctrl+B/I/U |
| Native format save/load | ✅ Done | SQLite .sheet files |
| CSV export | ✅ Done | |
| Dynamic grid sizing | ✅ Done | Fills window, resizes |
| Double-click to edit | ✅ Done | |
| Mouse wheel scroll | ✅ Done | |
| Page Up/Down | ✅ Done | |
| Column/Row resize | ✅ Done | Drag headers to resize |
| Command Palette | ✅ Done | Ctrl+Shift+P, fuzzy search |
| Fill Down/Right | ✅ Done | Ctrl+D/R with formula adjustment |
| Multi-edit | ✅ Done | Ctrl+Enter applies to selection |
| Excel 2003 Menu Bar | ✅ Done | File/Edit/View/Insert/Format/Data/Help |
| Alt menu accelerators | ✅ Done | Alt+F, Alt+E, Alt+V, etc. |
| Multi-sheet support | ✅ Done | Workbook with multiple sheets |
| Sheet tabs | ✅ Done | Click to switch, + to add |
| Sheet navigation | ✅ Done | Ctrl+PageUp/Down, Shift+F11 |
| Sheet context menu | ✅ Done | Right-click: Insert/Delete/Rename |

### Remaining Features (Priority Order)

| Feature | Priority | Effort | Notes |
|---------|----------|--------|-------|
| Ctrl+Click selection | P1 | Easy | Discontiguous ranges |
| Number format shortcuts | P2 | Easy | Ctrl+Shift+$, %, etc. |
| Format Cells dialog | P2 | Medium | Ctrl+1 |
| Replace (Ctrl+H) | P2 | Easy | |
| CSV import | P2 | Easy | |
| Zoom | P2 | Easy | Ctrl++/- |
| Cell context menu | P3 | Medium | Right-click on cells |
| Zen Mode | P3 | Easy | F11 |
| Split view | P3 | Hard | |
| Cross-sheet references | P3 | Medium | =Sheet2!A1 |

### Editor-Style Features (VS Code Inspired)

| Feature | Priority | Status |
|---------|----------|--------|
| Command Palette (Ctrl+Shift+P) | P0 | ✅ Done |
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
│   ├── workbook.rs   # Multi-sheet workbook
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

### Phase 1: Core Parity ✅ COMPLETE

1. ✅ Basic grid navigation
2. ✅ Cell editing
3. ✅ Copy/paste
4. ✅ Undo/redo
5. ✅ Save/load
6. ✅ Find/GoTo
7. ✅ Command Palette
8. ✅ Fill Down/Right (with formula reference adjustment)
9. ✅ Multi-edit (Ctrl+Enter)

### Phase 2: Excel Compatibility ✅ COMPLETE

1. ✅ Excel 2003-style dropdown menus
2. ✅ Alt accelerators (Alt+F, Alt+E, etc.)
3. ✅ Ctrl+Shift+Arrow selection
4. ✅ Column/row resize by dragging
5. ❌ Number format shortcuts
6. ❌ Format Cells dialog (Ctrl+1)
7. ❌ Replace (Ctrl+H)
8. ❌ Ctrl+Click selection

### Phase 3: Multi-Sheet ✅ COMPLETE

1. ✅ Workbook engine (multiple sheets)
2. ✅ Sheet tabs UI
3. ✅ Ctrl+PageUp/Down navigation
4. ✅ Shift+F11 to add sheet
5. ✅ Sheet management (add/delete/rename via context menu)
6. ❌ Cross-sheet references (=Sheet2!A1)

### Phase 4: Power Features (Next)

1. Named ranges UI
2. Array formula visualization (spill borders)
3. Zoom (Ctrl++/-)
4. Split view
5. Zen mode (F11)
6. Cell context menu

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

## Known Issues

| Issue | Description | Status |
|-------|-------------|--------|
| Per-cell font not rendering | Font picker UI works, data is stored/persisted correctly, but gpui doesn't render the font change. Tried both style inheritance and explicit TextRun approaches. | Investigating |

---

## Next Steps

1. Ctrl+Click discontiguous selection
2. Format Cells dialog (Ctrl+1)
3. Replace (Ctrl+H)
4. Cross-sheet references
