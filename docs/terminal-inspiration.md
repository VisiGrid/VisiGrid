# Terminal-Inspired Features

Design patterns borrowed from premium terminals (Ghostty, Kitty, WezTerm, Alacritty) that could translate to VisiGrid.

---

## Already Implemented

These patterns validate VisiGrid's direction:

| Pattern | Terminal Example | VisiGrid Status |
|---------|------------------|-----------------|
| Config as code | TOML/JSON config files | settings.json, keybindings.json |
| GPU rendering | wgpu, Metal, OpenGL | iced/wgpu |
| Keyboard-first | Everything via shortcuts | Command palette, 55+ shortcuts |
| Splits | tmux-style panes | Ctrl+\\ split view |
| Minimal chrome | No window decorations | Zen mode (F11) |
| Themes | Easy theme switching | JSON themes, Omarchy integration |
| Keyboard hints | Vimium, EasyMotion | 'g' key shows hints, type to jump |
| URL detection | iTerm2 semantic history | Ctrl+click opens URLs/emails/paths |
| Session persistence | tmux-resurrect | Auto-restore, workspaces |

---

## Proposed Features

### 1. Keyboard Hints (Vimium-style) ✓

**Inspiration:** Vimium browser extension, EasyMotion vim plugin

**Status:** ✓ Implemented

Press `g` in navigation mode to show letter hints on all visible cells. Type the hint letters to jump instantly.

```
┌───────┬───────┬───────┬───────┐
│   A   │   B   │   C   │   D   │  ← hints appear when you press 'g'
├───────┼───────┼───────┼───────┤
│   E   │   F   │   G   │   H   │
├───────┼───────┼───────┼───────┤
│   I   │   J   │   K   │   L   │
└───────┴───────┴───────┴───────┘

Type 'H' → cursor jumps to that cell
```

**How it works:**
- Press `g` to enter hint mode
- Hints use a-z, then aa-az, ba-bz for large grids
- Type letters to filter; jumps when unique match found
- Matching hints highlight in yellow on blue
- Press `Backspace` to correct, `Escape` to cancel

---

### 2. Pipe-Friendly CLI

**Inspiration:** Unix philosophy, jq, csvkit

Terminals integrate with shell pipelines. VisiGrid should too.

```bash
# Quick calculation without opening GUI
cat sales.csv | visigrid --headless "=SUM(B:B)"
# Output: 125450

# Export and pipe to other tools
visigrid export budget.sheet --format json | jq '.rows[] | select(.total > 1000)'

# Import from API
curl -s api.example.com/data.json | visigrid import --to A1

# Diff two sheets
visigrid diff old.sheet new.sheet

# Apply formula to stdin
echo -e "10\n20\n30" | visigrid calc "=SUM(A:A)"
# Output: 60

# Batch operations
visigrid batch budget.sheet --script cleanup.lua

# Headless format conversion
visigrid convert data.xlsx --to csv --output data.csv
```

**Subcommands:**
- `visigrid open <file>` - Open in GUI (default)
- `visigrid calc <formula>` - Evaluate formula against stdin
- `visigrid export <file> --format <fmt>` - Export to stdout
- `visigrid import --to <cell>` - Import from stdin
- `visigrid diff <a> <b>` - Text diff of two sheets
- `visigrid convert <file> --to <fmt>` - Format conversion
- `visigrid batch <file> --script <lua>` - Run script headlessly

**Priority:** High - Makes VisiGrid composable, appeals to terminal users

---

### 3. Inline Sparklines

**Inspiration:** Sixel graphics, Kitty image protocol

Mini visualizations rendered directly in cells. Very "terminal aesthetic."

```
┌──────────┬─────────────────┬──────────┐
│ Product  │ Trend           │ Total    │
├──────────┼─────────────────┼──────────┤
│ Revenue  │ ▂▅▇▅▃▁▃▅▇       │ $847,000 │
│ Users    │ ▁▂▃▄▅▆▇█▇       │ 12,450   │
│ Churn    │ ▇▅▃▂▁▁▂▃▂       │ 2.3%     │
└──────────┴─────────────────┴──────────┘
```

**Types:**
- Line sparkline: `▁▂▃▄▅▆▇█`
- Bar sparkline: `|||||`
- Win/loss: `▲▼▲▲▼▲`
- Bullet/progress: `████░░░░`

**Implementation:**
```
=SPARKLINE(B2:M2)                    # Auto line
=SPARKLINE(B2:M2, "bar")             # Bar style
=SPARKLINE(B2:M2, "winloss")         # Up/down indicators
=PROGRESS(0.75)                       # Progress bar
```

**Priority:** Medium - Visual differentiation from Excel, fits the aesthetic

---

### 4. Background Jobs with Notifications

**Inspiration:** Shell job control, desktop notifications

Long operations run async, notify on completion.

```
┌─────────────────────────────────────────────────┐
│ [Import] salesdata.csv → Sheet1 ████████░░ 80% │
│ [Export] Running: Q4-report.xlsx                │
└─────────────────────────────────────────────────┘

🔔 Notification: "Export complete: Q4-report.xlsx"
```

**Use cases:**
- Large file imports (100K+ rows)
- Complex recalculations
- Export to slow formats (Excel, PDF)
- Data refresh from external sources
- Script execution

**Implementation:**
- Status bar shows active jobs
- Jobs panel lists all (like browser downloads)
- Desktop notification on completion
- Keyboard shortcut to view jobs: `Ctrl+Shift+J`

**Priority:** Low - Nice for power users, not essential early

---

### 5. Formula Ligatures

**Inspiration:** Programming font ligatures (Fira Code, JetBrains Mono)

Transform operator sequences into proper symbols in the formula bar.

| Typed | Display |
|-------|---------|
| `>=`  | `≥`     |
| `<=`  | `≤`     |
| `<>`  | `≠`     |
| `!=`  | `≠`     |
| `->`  | `→`     |
| `=>`  | `⇒`     |
| `&&`  | `∧`     |
| `||`  | `∨`     |

**Implementation notes:**
- Display only (source remains ASCII)
- Toggle in settings: `editor.ligatures: true`
- Cursor movement treats ligature as original characters
- Only in formula bar, not in cells

**Priority:** Low - Polish feature, fun but not essential

---

### 6. Semantic Regions

**Inspiration:** Semantic shell prompts (OSC 133), LSP semantic tokens

Auto-detect and treat different regions specially.

**Region types:**

| Region | Detection | Behavior |
|--------|-----------|----------|
| Headers | First row with text, followed by data | Auto-freeze, bold, filter row |
| Data | Contiguous filled cells | Auto-select for operations |
| Formulas | Cells starting with `=` | Highlight dependencies |
| Totals | Row after data with SUM/AVERAGE | Protect from accidental edit |
| Empty | Large unfilled areas | Skip in navigation |

**Features:**
- `Ctrl+Shift+H` - Toggle header detection
- `Ctrl+A` - Select current data region (not entire sheet)
- Visual indicators for region boundaries
- Warnings when editing protected regions

**Priority:** Medium - Improves UX, reduces errors

---

### 7. Session Persistence ✓

**Inspiration:** tmux-resurrect, vim sessions

**Status:** ✓ Implemented (see docs/editor-inspiration.md)

Auto-save everything on quit, restore exactly on reopen.

**Saved state:**
```json
{
  "files": [
    {
      "path": "/home/user/budget.sheet",
      "cursor": "D15",
      "scroll": { "row": 10, "col": 0 },
      "selection": "D15:F20",
      "split": { "enabled": true, "position": 0.5 }
    }
  ],
  "layout": {
    "inspector": true,
    "problems": false,
    "zen": false
  },
  "recent_commands": ["sort desc", "format currency", "freeze row"],
  "undo_history": "budget.sheet.undo"
}
```

**Behavior:**
- Auto-save session every N seconds
- Restore on next launch
- `--no-restore` flag to start fresh
- Named sessions: `visigrid --session work`

**Priority:** Medium - Power user feature, already planned (Workspaces v2)

---

### 8. URL/Path Detection ✓

**Inspiration:** Terminal URL detection, iTerm2 semantic history

**Status:** ✓ Implemented

Auto-detect and open URLs, file paths, and email addresses with Ctrl+Click.

```
┌──────────┬─────────────────────────────────────┐
│ Invoice  │ https://stripe.com/inv_abc123       │ ← Ctrl+Click opens browser
├──────────┼─────────────────────────────────────┤
│ Receipt  │ ~/Documents/receipts/jan.pdf        │ ← Opens in file viewer
├──────────┼─────────────────────────────────────┤
│ Contact  │ billing@example.com                 │ ← Opens mail client
└──────────┴─────────────────────────────────────┘
```

**Supported patterns:**
- URLs: `http://`, `https://`, `ftp://`
- Email: `user@domain.com` (opens as `mailto:`)
- Paths: `/absolute/path` or `~/home/relative` (if file exists)

**Usage:** `Ctrl+Click` on a cell containing a link to open it with your system's default handler. If no link is detected, falls back to adding cell to selection.

---

### 9. Status Line Customization

**Inspiration:** Starship prompt, vim statusline, tmux status

User-configurable status bar with template syntax.

**Default:**
```
[Sheet1] A1 | 3 errors | Selection: 4 cells | SUM=12,450 | Modified
```

**Configuration (settings.json):**
```json
{
  "statusline.left": "{sheet} {cell}",
  "statusline.center": "{errors} | {selection_info}",
  "statusline.right": "{modified} {mode}",
  "statusline.components": {
    "selection_info": "{count} cells | SUM={sum} AVG={avg}"
  }
}
```

**Available variables:**
- `{sheet}` - Current sheet name
- `{cell}` - Current cell reference
- `{mode}` - Edit/Normal/Visual mode
- `{errors}` - Error count with icon
- `{selection}` - Selection range
- `{count}` - Selected cell count
- `{sum}`, `{avg}`, `{min}`, `{max}` - Selection stats
- `{modified}` - Unsaved indicator
- `{file}` - File name
- `{path}` - Full path

**Priority:** Low - Power user customization, v2 feature

---

### 10. Shell Commands in Formulas

**Inspiration:** Unix pipes, org-mode babel

Execute shell commands and use output in cells.

```
=SHELL("date +%Y-%m-%d")           # → 2024-01-15
=SHELL("curl -s api.com/rate")     # → 1.23
=SHELL("wc -l < data.txt")         # → 1542
```

**Security model:**
- Disabled by default
- Opt-in per file: "This file wants to run shell commands. Allow?"
- Sandboxed execution (firejail, bubblewrap)
- Whitelist specific commands
- No network access by default
- Cached results, explicit refresh with `Ctrl+Shift+R`

**Alternative: External data functions:**
```
=HTTP("https://api.example.com/rate")
=FILE("/path/to/data.txt")
=ENV("HOME")
```

**Priority:** Low - Powerful but risky, needs careful design

---

## Implementation Priority

Sorted by impact and feasibility:

| Rank | Feature | Effort | Impact | Status |
|------|---------|--------|--------|--------|
| 1 | Pipe-friendly CLI | Medium | High | Partial (diff done) |
| 2 | Keyboard hints | Low | High | ✓ Done |
| 3 | Inline sparklines | Medium | Medium | Not started |
| 4 | Semantic regions | Medium | Medium | Not started |
| 5 | Session persistence | Low | Medium | ✓ Done |
| 6 | URL detection | Low | Low | ✓ Done |
| 7 | Background jobs | Medium | Low | Not started |
| 8 | Status customization | Low | Low | Not started |
| 9 | Formula ligatures | Low | Low | Not started |
| 10 | Shell commands | High | Low | Not started |

---

## The Terminal Ethos

What makes terminal apps feel premium:

1. **Speed** - Instant response, no lag
2. **Composability** - Works with other tools
3. **Configurability** - Everything customizable
4. **Keyboard-first** - Mouse optional
5. **Text-based** - Version control, diff-able
6. **Focused** - Does one thing well
7. **Transparent** - No hidden state
8. **Respectful** - No telemetry, no cloud requirement

VisiGrid already embodies most of these. The CLI and sparklines would complete the picture.
