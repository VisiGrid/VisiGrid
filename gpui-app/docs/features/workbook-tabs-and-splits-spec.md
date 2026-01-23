# Workbook Tabs and Split View

Multi-document support with side-by-side comparison.

## Overview

VisiGrid needs to support multiple open workbooks with comparison capability. This spec defines:
1. **Workbook tabs** - Multiple documents open in one window
2. **Split view** - Side-by-side panes for comparison
3. **Focus and navigation** - Keyboard-driven pane/tab switching

## Architecture

Based on Zed's proven pane system, adapted for spreadsheets.

### Data Model

```
Workspace (Spreadsheet struct, expanded)
└── PaneGroup (center layout)
    └── Member (recursive enum)
        ├── Member::Pane(Entity<WorkbookPane>)   // leaf
        └── Member::Axis(PaneAxis)               // split
            ├── axis: Axis (Horizontal | Vertical)
            ├── members: Vec<Member>
            └── flexes: Vec<f32>                 // resize ratios
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

/// A container for workbook tabs
pub struct WorkbookPane {
    /// Open workbooks in this pane (tabs)
    workbooks: Vec<Entity<Workbook>>,
    /// Which workbook is active (0-indexed)
    active_index: usize,
    /// Focus handle for this pane
    focus_handle: FocusHandle,
    /// Scroll/selection state per workbook
    view_states: HashMap<EntityId, WorkbookViewState>,
}

/// Recursive pane layout
pub enum Member {
    Pane(Entity<WorkbookPane>),
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
| New Workbook | Cmd+N | Ctrl+N | Create Book2, Book3, etc. |
| Open | Cmd+O | Ctrl+O | Open file in current pane |
| Close Workbook | Cmd+W | Ctrl+W | Close active workbook tab |
| Next Workbook | Ctrl+Tab | Ctrl+Tab | Cycle to next tab |
| Previous Workbook | Ctrl+Shift+Tab | Ctrl+Shift+Tab | Cycle to previous tab |
| Go to Workbook N | Cmd+1..9 | Ctrl+1..9 | Jump to Nth tab |

### Split Operations

| Command | macOS | Windows/Linux | Description |
|---------|-------|---------------|-------------|
| Split Right | Cmd+\ | Ctrl+\ | Split pane vertically |
| Split Down | Cmd+Shift+\ | Ctrl+Shift+\ | Split pane horizontally |
| Close Split | Cmd+Shift+W | Ctrl+Shift+W | Close current pane (keeps workbooks) |
| Join All | - | - | Merge all panes into one |

### Pane Navigation

| Command | macOS | Windows/Linux | Description |
|---------|-------|---------------|-------------|
| Focus Pane Left | Cmd+Alt+Left | Ctrl+Alt+Left | Move focus to left pane |
| Focus Pane Right | Cmd+Alt+Right | Ctrl+Alt+Right | Move focus to right pane |
| Focus Pane Up | Cmd+Alt+Up | Ctrl+Alt+Up | Move focus to pane above |
| Focus Pane Down | Cmd+Alt+Down | Ctrl+Alt+Down | Move focus to pane below |
| Swap Pane | Cmd+Alt+S | Ctrl+Alt+S | Swap this pane with adjacent |

## Behavior

### Split Semantics

1. **Split Right/Down** creates a new empty pane
   - New pane gets focus
   - Original pane keeps its workbooks
   - User can Ctrl+N or Ctrl+O to add content

2. **Split cloning** (optional, V2)
   - Cmd+Alt+\ could duplicate the current workbook view
   - Both panes show same workbook, independent scroll/selection

3. **Closing a pane**
   - If pane has dirty workbooks: prompt save
   - If pane is last one: don't close (always have one pane)
   - Workbooks move to adjacent pane OR close with confirmation

### Tab Behavior

1. **Tab order** - Most recently opened is rightmost
2. **Close tab** - Activates previous tab (not next)
3. **Drag tabs** - Between panes to move workbooks (V2)
4. **Middle-click** - Closes tab
5. **Double-click tab bar** - New workbook

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
    4px divider, draggable
```

- Divider: 4px hitbox, 1px visual line
- Hover: cursor changes to resize
- Active pane: subtle border highlight

## Implementation Phases

### Phase 1: Foundation (do first)

1. Create `WorkbookPane` struct
2. Create `PaneGroup` with single-pane support
3. Move current `Spreadsheet` grid logic into pane
4. Workbook tabs (visual only, single workbook)

### Phase 2: Multi-Workbook

1. Support multiple workbooks in one pane
2. Tab switching (Ctrl+Tab, Ctrl+1-9)
3. New Workbook creates tab (not replaces)
4. Close Workbook closes tab
5. Dirty state per workbook

### Phase 3: Split View

1. Implement `Member` enum and `PaneAxis`
2. Split Right / Split Down actions
3. Pane focus management
4. Pane navigation (Cmd+Alt+arrows)
5. Divider drag to resize

### Phase 4: Polish

1. Tab drag between panes
2. Detach to window (optional)
3. Linked scrolling toggle
4. Session restore for splits

## State Persistence

Session should save:
```rust
struct PaneSession {
    /// Recursive layout
    layout: MemberSession,
    /// Which pane was focused
    active_pane_id: usize,
}

enum MemberSession {
    Pane {
        workbooks: Vec<WorkbookSession>,
        active_index: usize,
    },
    Axis {
        axis: Axis,
        members: Vec<MemberSession>,
        flexes: Vec<f32>,
    },
}

struct WorkbookSession {
    path: Option<PathBuf>,
    // ... existing scroll, selection, etc.
}
```

## Migration Path

Current `Spreadsheet` struct holds one workbook. Migration:

1. Keep `Spreadsheet` as the root view
2. Add `PaneGroup` field to `Spreadsheet`
3. `PaneGroup` starts with one `WorkbookPane`
4. `WorkbookPane` starts with one workbook
5. Gradually move grid rendering into pane

No breaking changes to file format - this is UI-only.

## Open Questions

1. **Tab position**: Top (like Zed) or bottom (like Excel)?
   - Recommendation: Top, it's more standard for modern apps

2. **Sheet tabs vs workbook tabs**: Both visible?
   - Recommendation: Workbook tabs top, sheet tabs bottom (status bar area)

3. **Detach to window**: V1 or later?
   - Recommendation: Later. Split view covers 90% of comparison needs.

4. **Max tabs**: Limit or scroll?
   - Recommendation: Scroll with overflow menu, like browsers

## References

- Zed source: `/tmp/zed/crates/workspace/src/pane.rs`
- Zed source: `/tmp/zed/crates/workspace/src/pane_group.rs`
- [Zed Actions Docs](https://zed.dev/docs/all-actions)
