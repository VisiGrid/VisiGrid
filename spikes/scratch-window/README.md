# Scratch window spike (Phase 0)

Throwaway. Proves the window and lifecycle behaviour scratch mode depends on,
on the pinned gpui fork, before any app work starts. See
`planning/visigrid/scratch-mode-brief.md`, section "Phase 0".

## Build and run

    cargo run --release            # from this directory; own workspace, own lockfile

Environment:

- `SPIKE_KIND=popup|floating|normal` (default `popup`) picks `WindowKind`.

Drive it:

- **Ctrl+Shift+Space** anywhere (the global hotkey under test). A failed
  registration is logged, not fatal.
- `kill -USR1 <pid>` (unix) or type `t` + Enter in the terminal: same toggle.
- `d` + Enter cycles the target display; `q` + Enter quits; Ctrl+Q in the
  window quits.

Everything is logged to stdout with timestamps.

## What to record per platform

1. Hotkey summon while another app is frontmost (macOS: also with a
   full-screen app on screen). Did the window appear? Log line
   `first render Nms after summon`.
2. Type `12+3` then some letters. Do `key:` lines appear with `char=` set, and
   is `active=true`? On Windows, also summon via `kill`-equivalent from a
   second terminal and note whether the window takes foreground or flashes.
3. Esc, then summon again: `first render` time on the second summon (this is
   the close+reopen cost; the fork has no per-window hide/show).
4. `d` then summon: does it land on the other display, and is `scale` right?
5. macOS only: repeat 1–3 in the sandboxed (App Store) build of this binary.

Paste the log and one screenshot per platform into the Phase 0 results page.

## Findings so far (Linux, 2026-09-06)

See `planning/visigrid/scratch-mode-phase0-results.md`.
