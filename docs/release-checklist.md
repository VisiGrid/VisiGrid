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

## Releasing

Use the automated release script:

```bash
# Dry run (prints what it would do):
./scripts/release.sh 0.7.0 --dry-run

# Full release:
./scripts/release.sh 0.7.0
```

The script handles all of these in order:
1. Pre-flight checks (branch, clean tree, untracked .rs files, build)
2. Version bump in workspace `Cargo.toml` + `Cargo.lock` update
3. Tag + push, wait for CI to pass
4. Publish the draft release (triggers Homebrew/Winget workflows)
5. Update AUR PKGBUILD with new SHA
6. Verify Homebrew and AUR

### Manual version bump (if needed)

Update `version` in the workspace `Cargo.toml` (the workspace version propagates to all crates):
```bash
# Edit Cargo.toml [workspace.package] version
cargo generate-lockfile
```

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
