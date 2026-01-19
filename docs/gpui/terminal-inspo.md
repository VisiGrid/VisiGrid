# Terminal-Inspired Features: VisiGrid gpui

Design patterns from premium terminals (Ghostty, Kitty, WezTerm, Alacritty) that apply to VisiGrid.

---

## Current Status

| Pattern | Terminal Example | gpui Status |
|---------|------------------|-------------|
| GPU rendering | wgpu, Metal | ✅ gpui uses Metal/Vulkan |
| Keyboard-first | Shortcuts for everything | ✅ 37 shortcuts implemented |
| Minimal chrome | No decorations | ❌ Zen mode not yet |
| Config as code | TOML/JSON files | ❌ Not yet |
| Themes | Easy switching | ❌ Not yet |
| Session persistence | tmux-resurrect | ❌ Not yet |

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

VisiGrid embodies most of these. The features below would complete the picture.

---

## Proposed Features

### 1. Keyboard Hints (Vimium-style)

**Inspiration:** Vimium browser extension, EasyMotion vim plugin

Press a key to show letter hints on visible cells. Type to jump.

```
┌───────┬───────┬───────┬───────┐
│   A   │   B   │   C   │   D   │  ← hints appear
├───────┼───────┼───────┼───────┤
│   E   │   F   │   G   │   H   │
├───────┼───────┼───────┼───────┤
│   I   │   J   │   K   │   L   │
└───────┴───────┴───────┴───────┘

Type 'H' → cursor jumps to that cell
```

**How it works:**
- Press `g` to enter hint mode
- Hints use a-z, then aa-az for large grids
- Type letters to filter; jumps when unique
- `Backspace` to correct, `Escape` to cancel

**Priority:** P2 | **Status:** ❌ Not implemented

---

### 2. Pipe-Friendly CLI

**Inspiration:** Unix philosophy, jq, csvkit

```bash
# Quick calculation without GUI
cat sales.csv | visigrid --headless "=SUM(B:B)"
# Output: 125450

# Export and pipe
visigrid export budget.sheet --format json | jq '.rows[]'

# Diff two sheets
visigrid diff old.sheet new.sheet

# Headless format conversion
visigrid convert data.xlsx --to csv --output data.csv
```

**Subcommands:**
- `visigrid open <file>` - Open in GUI (default)
- `visigrid calc <formula>` - Evaluate formula
- `visigrid export <file> --format <fmt>` - Export
- `visigrid diff <a> <b>` - Text diff
- `visigrid convert <file> --to <fmt>` - Convert

**Priority:** P1 (High) | **Status:** ❌ Not implemented

---

### 3. Inline Sparklines

**Inspiration:** Sixel graphics, Kitty image protocol

Mini visualizations in cells using Unicode:

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
- Bar sparkline
- Win/loss: `▲▼▲▲▼▲`
- Progress: `████░░░░`

**Formula:**
```
=SPARKLINE(B2:M2)
=SPARKLINE(B2:M2, "bar")
=PROGRESS(0.75)
```

**Priority:** P3 | **Status:** ❌ Not implemented

---

### 4. URL/Path Detection

**Inspiration:** Terminal URL detection, iTerm2

Auto-detect and open URLs, file paths, email with Ctrl+Click.

```
┌──────────┬─────────────────────────────────────┐
│ Invoice  │ https://stripe.com/inv_abc123       │ ← Ctrl+Click opens
├──────────┼─────────────────────────────────────┤
│ Receipt  │ ~/Documents/receipts/jan.pdf        │ ← Opens file viewer
├──────────┼─────────────────────────────────────┤
│ Contact  │ billing@example.com                 │ ← Opens mail client
└──────────┴─────────────────────────────────────┘
```

**Patterns:**
- URLs: `http://`, `https://`
- Email: `user@domain.com`
- Paths: `/absolute/path`, `~/relative`

**Priority:** P2 | **Status:** ❌ Not implemented

---

### 5. Session Persistence

**Inspiration:** tmux-resurrect, vim sessions

Auto-save state on quit, restore on reopen.

**Saved state:**
```json
{
  "files": [{
    "path": "/home/user/budget.sheet",
    "cursor": "D15",
    "scroll": { "row": 10, "col": 0 },
    "selection": "D15:F20"
  }],
  "layout": {
    "zen": false
  },
  "recent_commands": ["sort desc", "format currency"]
}
```

**Priority:** P2 | **Status:** ❌ Not implemented

---

### 6. Background Jobs with Notifications

**Inspiration:** Shell job control, desktop notifications

Long operations run async, notify on completion.

```
┌─────────────────────────────────────────────────┐
│ [Import] salesdata.csv → Sheet1 ████████░░ 80% │
│ [Export] Running: Q4-report.xlsx                │
└─────────────────────────────────────────────────┘

🔔 Notification: "Export complete"
```

**Use cases:**
- Large file imports (100K+ rows)
- Complex recalculations
- Export to slow formats

**Priority:** P3 | **Status:** ❌ Not implemented

---

### 7. Formula Ligatures

**Inspiration:** Fira Code, JetBrains Mono

Transform operators into proper symbols in formula bar.

| Typed | Display |
|-------|---------|
| `>=`  | `≥`     |
| `<=`  | `≤`     |
| `<>`  | `≠`     |

**Notes:**
- Display only (source stays ASCII)
- Toggle in settings
- Only in formula bar

**Priority:** P4 | **Status:** ❌ Not implemented

---

### 8. Semantic Regions

**Inspiration:** LSP semantic tokens, org-mode

Auto-detect and treat regions specially.

| Region | Detection | Behavior |
|--------|-----------|----------|
| Headers | First row with text | Auto-freeze, bold |
| Data | Contiguous filled cells | Auto-select |
| Formulas | Cells with `=` | Highlight deps |
| Totals | Row with SUM/AVERAGE | Protect |

**Priority:** P3 | **Status:** ❌ Not implemented

---

### 9. Status Line Customization

**Inspiration:** Starship prompt, vim statusline

User-configurable status bar.

**Default:**
```
[Sheet1] A1 | Selection: 4 cells | SUM=12,450 | Modified
```

**Config (future):**
```json
{
  "statusline.left": "{sheet} {cell}",
  "statusline.right": "{modified} {mode}"
}
```

**Priority:** P4 | **Status:** ❌ Not implemented

---

## Implementation Priority

| Rank | Feature | Effort | Impact | gpui Status |
|------|---------|--------|--------|-------------|
| 1 | Pipe-friendly CLI | Medium | High | ❌ |
| 2 | Keyboard hints | Low | High | ❌ |
| 3 | URL detection | Low | Medium | ❌ |
| 4 | Session persistence | Low | Medium | ❌ |
| 5 | Inline sparklines | Medium | Medium | ❌ |
| 6 | Semantic regions | Medium | Medium | ❌ |
| 7 | Background jobs | Medium | Low | ❌ |
| 8 | Status customization | Low | Low | ❌ |
| 9 | Formula ligatures | Low | Low | ❌ |

---

## Near-Term Focus

For gpui MVP, focus on:

1. **Core spreadsheet functionality** (current)
2. **Command Palette** (editor-style)
3. **Fill Down/Right** (Excel compat)

Terminal-inspired features come after core parity.
