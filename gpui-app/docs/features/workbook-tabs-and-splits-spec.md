# Workbook Tabs and Split View

Multi-document support with side-by-side comparison.

## Overview

VisiGrid needs to support multiple open workbooks with comparison capability. This spec defines:
1. **Workbook tabs** - Multiple documents open in one window
2. **Split view** - Side-by-side panes for comparison
3. **Focus and navigation** - Keyboard-driven pane/tab switching

## Mental Model

Four distinct concepts, each with a clear responsibility:

| Concept | What it is | Examples |
|---------|-----------|----------|
| **Workbook** | The document (data, formulas, sheets) | Book1, budget.xlsx |
| **WorkbookView** | Workbook + view state (scroll, selection, zoom, active sheet) | "budget.xlsx scrolled to row 500" |
| **Pane** | Tab container holding WorkbookViews | Left pane with 3 tabs |
| **PaneGroup** | Split tree of Panes | Two panes side-by-side |

**Key insight:** A single Workbook can have multiple WorkbookViews (e.g., same file in two panes showing different sheets or scroll positions). View state belongs to the View, not the Workbook.

## Architecture

Based on Zed's proven pane system, adapted for spreadsheets.

### Data Model

```
Workspace (Spreadsheet struct, expanded)
└── PaneGroup (center layout)
    └── Member (recursive enum)
        ├── Member::Pane(Entity<Pane>)        // leaf
        └── Member::Axis(PaneAxis)            // split
            ├── axis: Axis (Horizontal | Vertical)
            ├── members: Vec<Member>
            └── flexes: Vec<f32>              // resize ratios
```

### Core Types

```rust
/// Split direction for pane operations
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SplitDirection {
    Up,
    Down,
    Left,
    Right,
}

impl SplitDirection {
    pub fn axis(&self) -> Axis {
        match self {
            Self::Up | Self::Down => Axis::Vertical,
            Self::Left | Self::Right => Axis::Horizontal,
        }
    }

    pub fn increasing(&self) -> bool {
        matches!(self, Self::Down | Self::Right)
    }
}

/// View state for a workbook in a specific pane
/// This is the per-(pane, workbook) state that allows the same
/// workbook to be viewed differently in multiple panes.
#[derive(Clone)]
pub struct WorkbookViewState {
    pub scroll_row: usize,
    pub scroll_col: usize,
    pub active_cell: (usize, usize),
    pub selection: Selection,
    pub active_sheet: usize,
    pub zoom: f32,
    // ... other view-specific state
}

/// A view of a workbook with independent view state.
/// Multiple WorkbookViews can point to the same Workbook.
pub struct WorkbookView {
    /// The underlying workbook (shared, may have other views)
    pub workbook: Entity<Workbook>,
    /// This view's independent state
    pub state: WorkbookViewState,
}

/// A container for workbook tabs (views)
pub struct Pane {
    /// Open views in this pane (tabs)
    items: Vec<Entity<WorkbookView>>,
    /// Which view is active (0-indexed)
    active_index: usize,
    /// Focus handle for this pane
    focus_handle: FocusHandle,
}

/// Recursive pane layout
pub enum Member {
    Pane(Entity<Pane>),
    Axis(PaneAxis),
}

/// Split container holding 2+ members
pub struct PaneAxis {
    pub axis: Axis,              // Horizontal or Vertical
    pub members: Vec<Member>,    // Children (panes or nested axes)
    pub flexes: Vec<f32>,        // Resize proportions (1.0 each by default)
}

/// Root layout manager
pub struct PaneGroup {
    pub root: Member,
}
```

## Commands and Shortcuts

### Workbook Management

| Command | macOS | Windows/Linux | Description |
|---------|-------|---------------|-------------|
| New Workbook | Cmd+N | Ctrl+N | Create Book2, Book3, etc. in current pane |
| Open | Cmd+O | Ctrl+O | Open file in current pane |
| Close Tab | Cmd+W | Ctrl+W | Close active tab |
| Next Tab | Ctrl+Tab, Cmd+Shift+] | Ctrl+Tab | Cycle to next tab |
| Previous Tab | Ctrl+Shift+Tab, Cmd+Shift+[ | Ctrl+Shift+Tab | Cycle to previous tab |
| Go to Tab N | Cmd+1..9 | Ctrl+1..9 | Jump to Nth tab |

### Split Operations

| Command | macOS | Windows/Linux | Description |
|---------|-------|---------------|-------------|
| Split Right | Cmd+\ | Ctrl+\ | Clone active view into new right pane |
| Split Down | Cmd+Shift+\ | Ctrl+Shift+\ | Clone active view into new bottom pane |
| Close Pane | Cmd+K Cmd+W | Ctrl+K Ctrl+W | Close pane, merge tabs to neighbor |
| Join All | Cmd+K Cmd+J | Ctrl+K Ctrl+J | Merge all panes into one |

### Pane Navigation

| Command | macOS | Windows/Linux | Description |
|---------|-------|---------------|-------------|
| Focus Next Pane | Cmd+K Cmd+Right | Ctrl+K Ctrl+Right | Move focus to next pane |
| Focus Prev Pane | Cmd+K Cmd+Left | Ctrl+K Ctrl+Left | Move focus to previous pane |
| Focus Pane Up | Cmd+K Cmd+Up | Ctrl+K Ctrl+Up | Move focus to pane above |
| Focus Pane Down | Cmd+K Cmd+Down | Ctrl+K Ctrl+Down | Move focus to pane below |

**Note:** Using `Cmd+K` chord prefix avoids conflicts with Mission Control (Cmd+Alt+Arrow) and keeps pane operations grouped. This matches VS Code's pane commands.

## Behavior

### Split Semantics

1. **Split Right/Down clones the active view**
   - Creates new pane with a new WorkbookView pointing to the same Workbook
   - Both views have independent scroll, selection, zoom, active sheet
   - New pane gets focus
   - If current pane is empty, new pane is also empty

2. **Why clone, not empty?**
   - Users split to compare or see two parts of the same document
   - Empty pane feels like "nothing happened"
   - Matches Zed and Excel behavior
   - User can Cmd+O or Cmd+N in the new pane if they want different content

3. **Split when only one pane exists**
   - `root: Member::Pane(p)` becomes `root: Member::Axis { members: [Pane(p), Pane(new)], flexes: [1.0, 1.0] }`

### Close Pane Semantics

1. **Close pane merges tabs to neighbor** (no data loss)
   - All tabs move to the nearest sibling pane
   - Merge target priority: sibling in same axis > nearest leaf in traversal order
   - No save prompts (tabs are preserved, just relocated)

2. **Last pane cannot be closed**
   - Always maintain at least one pane
   - Use Close Tab (Cmd+W) to close individual tabs

3. **Save prompts are per-workbook, not per-pane**
   - Prompt when closing a tab with unsaved changes
   - Prompt when quitting with any dirty workbooks
   - Never prompt when just closing/merging panes

### Tab Behavior

1. **Tab order** - Most recently opened is rightmost
2. **Close tab activates previous** - `new_active = max(closed_index - 1, 0)`
3. **Middle-click** - Closes tab
4. **Double-click tab bar** - New workbook
5. **Drag tabs between panes** - V2 (deferred)

### Focus Rules

1. Only ONE pane is "active" at a time (has keyboard focus)
2. Active pane has visible focus ring
3. Clicking any cell in a pane focuses that pane
4. Tab switching within pane doesn't change pane focus

### Resize Behavior

1. Drag divider between panes to resize
2. Double-click divider to reset to 50/50
3. Minimum pane width: 200px
4. Minimum pane height: 150px

## Visual Design

### Tab Bar (per pane)

```
┌─────────────────────────────────────────────────────┐
│ [Book1] [budget.xlsx ●] [customers.csv]    [+] [⋮] │
├─────────────────────────────────────────────────────┤
│                                                     │
│                  Spreadsheet Grid                   │
│                                                     │
└─────────────────────────────────────────────────────┘
```

- Active tab: filled background, bold text
- Dirty indicator: ● after filename
- Hover: show close button (×)
- [+] button: new workbook
- [⋮] button: tab overflow menu

### Split Divider

```
┌──────────────┬──────────────┐
│              │              │
│   Pane 1     │   Pane 2     │
│  (focused)   │              │
│              │              │
└──────────────┴──────────────┘
        ↑
    4px hitbox, 1px visual line
```

- Divider: 4px hitbox, 1px visual line
- Hover: cursor changes to resize
- Active pane: subtle border highlight

## Edge Cases

### Closing a pane with multiple tabs

Tabs merge into the nearest neighbor pane:
1. Prefer sibling pane in the same axis
2. If no sibling, use nearest leaf in pre-order traversal
3. Merged tabs append to the end of the target pane's tab list
4. Focus moves to the target pane's active tab

### Active tab after close

When closing tab at index `i`:
- New active index = `max(i - 1, 0)`
- Exception: if `i == 0` and tabs remain, new active = 0

### Dirty workbook with multiple views

If a workbook has views in multiple panes:
- Dirty indicator (●) shows on ALL tabs viewing that workbook
- Save prompt appears once per workbook, not per view
- Saving from any view saves the workbook (updates all views)

### Empty pane after closing last tab

- Pane shows "No open files" placeholder
- Cmd+N or Cmd+O to add content
- Pane can still receive focus and be split

## Implementation Phases

### Phase 1: Foundation

1. Create `WorkbookView` struct with view state
2. Create `Pane` struct holding `Vec<Entity<WorkbookView>>`
3. Create `PaneGroup` with single-pane support
4. Migrate current `Spreadsheet` grid logic to use WorkbookView
5. Tab bar rendering (single tab initially)

### Phase 2: Multi-Tab

1. Support multiple WorkbookViews in one pane
2. Tab switching (Ctrl+Tab, Cmd+Shift+]/[, Cmd+1-9)
3. New Workbook creates new tab
4. Close Tab closes active tab
5. Dirty state per workbook (shared across views)

### Phase 3: Split View

1. Implement `Member` enum and `PaneAxis`
2. Split Right / Split Down (clone active view)
3. Pane focus management
4. Pane navigation (Cmd+K chord)
5. Close Pane with tab merging
6. Divider drag to resize

### Phase 4: Polish (V2)

1. Tab drag between panes
2. Linked scrolling toggle (for comparing same workbook)
3. Session restore for splits
4. Detach to window (optional)

## State Persistence

Session should save:
```rust
struct PaneSession {
    /// Recursive layout
    layout: MemberSession,
    /// Which pane was focused (index in pre-order traversal)
    active_pane_index: usize,
}

enum MemberSession {
    Pane {
        views: Vec<WorkbookViewSession>,
        active_index: usize,
    },
    Axis {
        axis: Axis,
        members: Vec<MemberSession>,
        flexes: Vec<f32>,
    },
}

struct WorkbookViewSession {
    /// Path to workbook (None if unsaved)
    path: Option<PathBuf>,
    /// View state
    scroll_row: usize,
    scroll_col: usize,
    active_cell: (usize, usize),
    active_sheet: usize,
    zoom: f32,
    // Note: selection is not persisted (too complex, low value)
}
```

## Migration Path

Current `Spreadsheet` struct holds one workbook. Migration:

1. Keep `Spreadsheet` as the root workspace view
2. Extract view state from `Spreadsheet` into `WorkbookViewState`
3. Create `WorkbookView` wrapping `Entity<Workbook>` + state
4. Create `Pane` with single `WorkbookView`
5. Add `PaneGroup` to `Spreadsheet` with single pane
6. Gradually move grid rendering to read from active WorkbookView

No breaking changes to file format - this is UI-only.

## Decisions Made

| Question | Decision | Rationale |
|----------|----------|-----------|
| Tab position | Top | Modern app convention (Zed, VS Code, browsers) |
| Sheet tabs vs workbook tabs | Workbook top, sheets bottom | Clear hierarchy, matches Excel |
| Split default | Clone active view | Users split to compare; empty pane feels broken |
| Detach to window | V2 | Split view covers 90% of comparison needs |
| Max tabs | Scroll with overflow | No arbitrary limit, like browsers |
| Pane shortcuts | Cmd+K chord | Avoids Mission Control conflicts |

## References

- Zed source: `zed/crates/workspace/src/pane.rs`
- Zed source: `zed/crates/workspace/src/pane_group.rs`
- VS Code pane shortcuts: `Cmd+K` chord pattern
