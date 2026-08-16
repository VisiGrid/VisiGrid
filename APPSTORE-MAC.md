# Mac App Store Submission Kit — VisiGrid for Mac

Bundle: com.visigrid.mac · Price: $14.99 one-time · Category: Productivity
(secondary: Business)

## Build & upload

```sh
# Requires (from developer.apple.com): Apple Distribution cert,
# Mac Installer Distribution cert, Mac App Store provisioning profile
# for com.visigrid.mac.
export APPLE_SIGNING_IDENTITY="Apple Distribution: RegAtlas, LLC (8KKQG868XP)"
export MAS_INSTALLER_IDENTITY="3rd Party Mac Developer Installer: RegAtlas, LLC (8KKQG868XP)"
export MAS_PROVISIONING_PROFILE=~/Downloads/VisiGrid_Mac_App_Store.provisionprofile

./gpui-app/scripts/bundle-macos.sh --appstore --sign
# -> gpui-app/build/VisiGrid-<version>.pkg
# Upload the .pkg with Transporter.app (Mac App Store, free).
```

MAS build behavior (automatic): session server / CLI pairing disabled
(sandbox denies the bind, app degrades), session restore off
(APP_SANDBOX_CONTAINER_ID detected; bookmarks are the 1.1 follow-up).

## Listing copy

**Name:** VisiGrid — if taken with the iPad record, "VisiGrid for Mac"

**Subtitle (30 chars):** `The keyboard-first spreadsheet`

**Promotional text:**
`A real spreadsheet engine, fully local. Command palette, Lua
scripting, Excel-grade shortcuts. Buy once — no subscription, no
account, no cloud.`

**Description:**
```
VisiGrid is a fast, native spreadsheet built the way modern code
editors are built — keyboard-first, scriptable, and entirely yours.

No subscription. No account. No cloud requirement. Your files live on
your Mac; the engine runs on your Mac.

A REAL ENGINE
• 123 formula functions with live recalculation
• Opens and saves Excel (.xlsx) and CSV/TSV — formulas survive
• Handles 100,000-row files smoothly, opens in a fraction of a second
• Multi-sheet workbooks, conditional formatting, undo across everything

BUILT LIKE A CODE EDITOR
• Command palette (⌘K) — every action, fuzzy-searchable
• A Lua console with live sheet access — scriptable like an editor
• Minimap, problems panel, configurable keybindings
• Excel-grade navigation: ⌘-arrows, ⇧-extension, point mode, AutoSum

MAKE IT YOURS
• Themes, including a VisiCalc green-phosphor homage to the
  spreadsheet's 1979 roots
• Native GPU rendering via Metal — no Electron, no web views
• The whole app is under 50 MB

VisiGrid is open source (AGPL) and remains free to download at
visigrid.app. Buying on the App Store supports development and gets
you automatic updates. Same engine as VisiGrid for iPad and Linux.

Buy once, keep forever.
```

**Keywords:**
`spreadsheet,excel,csv,xlsx,formulas,keyboard,lua,scripting,grid,finance,budget,editor,command palette`

**URLs:** support/marketing https://visigrid.app · privacy https://visigrid.app/privacy

## Review notes (paste into App Review Information)

```
Fully local spreadsheet app. No account or login required. To test:
open the app -> New Workbook; ⌘K opens the command palette.

About the Lua console (⌘K -> "Show Lua Console"): this is user-authored
automation over the open document, in the spirit of spreadsheet
formulas or AppleScript. The interpreter is embedded and sandboxed
inside the app process; it operates only on the open workbook, and the
app provides no mechanism to download or install code. Try:
sheet:set("A1", 42).

The terminal panel from the open-source build is disabled in this App
Store build (the panel explains this if opened); no shell is spawned.
```

## Screenshots (Mac sizes: 2880×1800 or 2560×1600, up to 10)

1. HERO — light theme, small model, =SUM in the formula bar
2. VISICALC THEME — green phosphor (the scroll-stopper)
3. COMMAND PALETTE — ⌘K open over the grid
4. LUA CONSOLE — script running, cells filled above
5. 100K ROWS — large file open, minimap visible

## Questionnaires

- App Privacy: Data Not Collected
- Age rating: all None/No -> 4+
- Export compliance: uses HTTPS only -> exempt (answer the standard
  encryption question "Yes, exempt")

## Post-approval

- Manual release. Launch beat shares assets with the iPad launch.
- v1.1: security-scoped bookmarks (session restore + recents in MAS).
