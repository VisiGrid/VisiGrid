# VisiGrid Release Checklist

Pre-release verification checklist for VisiGrid releases.

## Session Persistence

### Basic Functionality
- [ ] Open file, quit, relaunch → file reopens at same position
- [ ] Scroll to row 500, col Z, quit, relaunch → same viewport
- [ ] Select A1:G50, quit, relaunch → same selection
- [ ] Open inspector, switch to Names tab, quit → state restored

### Multi-Window
- [ ] Open 3 windows with different files, quit → all restore
- [ ] Maximized window restores as maximized
- [ ] Window positions preserved across monitors

### Edge Cases
- [ ] Missing file on restore → shows status message, doesn't crash
- [ ] Corrupt session.json → backs up to .bad-<timestamp>, starts fresh
- [ ] Delete session.json mid-session → recreates on quit
- [ ] Session from newer version → backs up, starts fresh
- [ ] Empty session.json → starts fresh
- [ ] First run (no session file) → starts normally

### CLI Flags
- [ ] `--no-restore` skips restore but preserves session file
- [ ] `--reset-session` deletes session file and starts fresh
- [ ] `--dump-session` prints session JSON and exits
- [ ] `--help` shows usage and exits
- [ ] Unknown flag shows error and exits

### File Operations
- [ ] Save As updates session with new path
- [ ] New File clears session state for that window
- [ ] Import (.xlsx) updates session after import completes

## Debugging

When troubleshooting session issues:

```bash
# View current session state
visigrid --dump-session

# Enable debug logging
VISIGRID_DEBUG_SESSION=1 visigrid

# Reset to fresh start
visigrid --reset-session
```

Session file location: `~/.config/visigrid/session.json`

## Build Verification

- [ ] `cargo test` passes (all crates)
- [ ] `cargo build --release` succeeds
- [ ] App launches on target platform
- [ ] Session metadata shows correct version, platform

## Version Bump

Before tagging:
1. Update version in `gpui-app/Cargo.toml`
2. Update version in `crates/*/Cargo.toml` if changed
3. Set VISIGRID_GIT_SHA in build (optional)

## Platform-Specific

### Linux
- [ ] App launches from terminal
- [ ] Desktop file works (if installed)
- [ ] XDG paths used correctly

### macOS
- [ ] App launches from Finder
- [ ] Cmd+Q saves session before quit
- [ ] Menu bar integrates correctly

### Windows
- [ ] App launches without console window
- [ ] Atomic save uses backup pattern
- [ ] %APPDATA% paths used correctly
