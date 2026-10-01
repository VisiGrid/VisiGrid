# Cell comments

The first version supports one plain-text comment per cell, with an optional author.
These correspond to traditional Excel Notes, not threaded discussions. Comments
persist in `.sheet`/`.vgrid` and import/export with `.xlsx`.

## Excel Notes round trips

- Text, optional author, cell address, blank annotated cells, and multiple sheets
  survive unchanged, edited, and native-format round trips. Deleting a comment
  removes it from the export; deleting all comments removes the note/VML parts.
- Rich-text note bodies import as plain text. Existing author prefixes in the body
  are retained verbatim; exports do not add another prefix. Note font styling,
  popup position/size, and always-visible state are not preserved.
- Threaded Excel discussions are not supported. Import stops with an explicit
  error before a workbook is created, so saving cannot silently erase discussions.
  The original remains untouched. Legacy binary `.xls` Notes are not imported.
- Invalid references, missing comment parts, and invalid author IDs fail import
  rather than silently dropping notes. Export rejects bodies exceeding Excel's
  32,767 UTF-16-unit limit without truncation.

Generate visual QA copies with:

```sh
cargo run -p visigrid-io --example xlsx_comments_probe -- /tmp/visigrid-xlsx-comments-qa
cargo test -p visigrid-io --test xlsx_comments --locked
```

Open `notes-source.xlsx` and `notes-unchanged.xlsx` in Excel and VisiGrid. Check
the B3 multiline note, blank B5, and Other sheet A1. In `notes-edited.xlsx`, B3
has new text/author, B7 has no note, and B9 has a new authorless note. Check
`notes-native-roundtrip.xlsx` against the source. Excel must open without repair.

The XLSX writer supplies note relationships and VML; VisiGrid replaces the generated
comment XML to correct the library's author-ID ordering and literal escape handling.
Other ZIP entries are copied unchanged. Workbooks with comments use a buffer for
this pass; exports without comments retain the direct-to-file path.

## UI checks

- Right-click a cell and choose **New Comment…**. Enter multiline text and an
  optional author. Save with the button or Cmd+Enter (Ctrl+Enter elsewhere).
- Click the colored corner marker or choose **Edit Comment…** to reopen it.
- Hover anywhere in a commented cell for half a second to read its preview. Move
  into the card to keep it open; **Edit comment** opens the editor. Moving away
  dismisses it. Check comments with and without authors, plus multiline text.
- **Preferences → Appearance → Comment hover previews** turns previews on/off
  across workbooks and persists after restart. **Toggle Comment Previews** in the
  command palette uses the same setting. Markers and click-to-edit remain available.
- Long previews wrap and scroll within a bounded card; scrolling over the card
  must not move or zoom the grid. Check the last visible row/column, merged cells
  across frozen panes, switching sheets, and entering the editor with a preview open.
- Check mouse caret placement, drag selection, arrow keys, select all, copy/paste,
  and Unicode text. Tab switches between text and author. Escape cancels drafts.
- Delete from the editor or cell context menu. Undo/redo restores comment metadata
  without changing cell values or formatting.
- Ordinary full cell copy/paste carries comments. **Paste Special → All** performs
  a full paste even when the values-only default is enabled. Values-only,
  formulas-only, and formats-only paste leave destination comments alone.
- Cut moves comments with the copied cells; undo restores source comments.
- Insert/delete rows and columns, undo deletions, and check blank annotated cells.
- Check markers on frozen and merged cells, then save and reopen a native workbook.

Comments are cell metadata, so clearing values keeps them. Blank comments cannot
be saved; use Delete to remove an existing comment. Comment actions appear in
history and support workbook rewind, but do not yet have a Lua replay API.

## Comment navigation

**Comments** in the toolbar (or **Toggle Comments Sidebar** in the command palette)
opens a persistent sidebar alongside the grid. On macOS it is also in View → Comments.
Search matches text, authors, cell addresses, and sheet names, ignoring case. Multiple
words must all match. **This sheet / All sheets** controls scope; sheet order then
cell order controls the result list. Click a result to jump and open its reading card.
Enter in search opens the first result. The × clears search; Escape closes the sidebar.

**Add/Edit Comment** in the palette, **Shift+F2**, and **Add / edit selected** in the
sidebar use the selected cell. Existing comments open for editing; empty cells get
a new draft. macOS also has Edit → Add/Edit Comment. On Macs with media function
keys, Fn+Shift+F2 may be needed. Custom keybindings: `comments.edit`, `comments.sidebar`.

Sidebar QA: search an author and a body word, filter to the current sheet, then
switch to All sheets and click a result on another tab. Check zero results, blank
annotated cells, and comments beyond the viewport. Search must not change cell data
or mark the workbook modified. Click back into the grid and check that typing/shortcuts
return to the grid. Add/edit/delete and undo should update results immediately.

Suggested demo: open Comments → search an author → jump to a match → Shift+F2 →
edit and Cmd+Enter → switch to All sheets → find a comment on another tab.

Use **Next Comment** or **Previous Comment** in the command palette. Navigation
follows workbook sheet order, then row and column order, and wraps at either end.
The reading card includes the cell address, author, position/count, Previous/Next,
and Edit comment. It works with hover previews disabled; Escape dismisses it.

Check blank commented cells, merged cells, distant cells, sheet boundaries and
both wrap directions. Hidden/filtered cells remain hidden and their comment is
shown with an explanation. Navigation must not mark the workbook modified.
An empty workbook reports “No comments in this workbook”; a single comment wraps
to itself. Selecting another cell, editing, or changing cell data dismisses the card.

## Automated checks

```sh
cargo test -p visigrid-engine --test comments --locked
cargo test -p visigrid-io --test native_fidelity --locked
cargo test -p visigrid-gpui --bin visigrid comment --locked
cargo test -p visigrid-gpui --bin visigrid ime:: --locked
```

These cover comment storage, structural movement, clearing values, cloned cells,
Unicode serialization, all native workbook save APIs plus the legacy single-sheet
API, deletion persistence, comment undo/redo and history replay, and IME composition.

## Preferences keyboard checks

- Open Preferences, press Tab to focus the sidebar, then Up/Down to change sections.
- Tab and Shift+Tab traverse only the current section, Done, and Close, wrapping
  within the modal. Focused controls scroll into view; Space/Enter activate them.
- In Appearance, tab to Comment hover previews and press Space twice. Confirm
  the setting changes and returns, without modifying the selected spreadsheet cell.
- Numeric fields accept typing, select-all and paste. Tab validates and applies
  before moving; Enter applies in place. Invalid sizes keep focus with an inline error.
- Escape cancels a pending numeric edit, closes Preferences, and returns keyboard
  focus to the grid. Done applies a valid pending size before closing.
