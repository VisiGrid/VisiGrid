# gpui Font Rendering Issue on Linux

## Summary

Bold, italic, and per-cell font changes do not render on Linux despite being correctly implemented in VisiGrid. This is a **gpui framework limitation**, not a VisiGrid bug.

## Symptoms

| Feature | Status | Notes |
|---------|--------|-------|
| Underline | ✅ Works | Drawn as separate decoration |
| Bold | ❌ Doesn't render | Font weight ignored |
| Italic | ❌ Doesn't render | Font style ignored |
| Per-cell font family | ❌ Doesn't render | Always uses default font |
| Number formats | ✅ Works | Not font-dependent |
| Alignment | ✅ Works | Not font-dependent |

## Root Cause

gpui's Linux text system uses [cosmic-text](https://github.com/pop-os/cosmic-text) for text rendering. The font selection pipeline has issues finding and applying font variants (bold, italic, different families).

### Code Flow Analysis

1. **VisiGrid correctly sets font properties**:
   ```rust
   let mut text_style = window.text_style();
   text_style.font_weight = FontWeight::BOLD;
   text_style.font_style = FontStyle::Italic;
   text_style.font_family = "Liberation Sans".into();
   ```

2. **gpui creates TextRun with Font**:
   ```rust
   // In gpui/src/style.rs - TextStyle::to_run()
   TextRun {
       font: Font {
           family: self.font_family.clone(),
           weight: self.font_weight,    // BOLD is set here
           style: self.font_style,      // Italic is set here
           ...
       },
       ...
   }
   ```

3. **Font resolution attempts to find matching face**:
   ```rust
   // In gpui/src/platform/linux/text_system.rs
   let ix = font_kit::matching::find_best_match(
       &candidate_properties,
       &font_into_properties(font)  // Contains weight=BOLD, style=Italic
   )?;
   ```

4. **But rendering uses the loaded font's properties directly**:
   ```rust
   // In layout_line() - uses font face info, not requested style
   attrs_list.add_span(
       offs..(offs + run.len),
       &Attrs::new()
           .family(Family::Name(&font.families.first().unwrap().0))
           .weight(font.weight)   // From loaded font, not request
           .style(font.style)     // From loaded font, not request
   );
   ```

## TODOs in gpui's Linux Text System

These are from `/tmp/zed-source/crates/gpui/src/platform/linux/text_system.rs`:

| Line | TODO | Impact |
|------|------|--------|
| 60 | `todo(linux) make font loading non-blocking` | Performance |
| 99 | `todo(linux): Do we need to use CosmicText's Font APIs? Can we consolidate this to use font_kit?` | **Font selection confusion** |
| 110 | `todo(linux) ideally we would make fontdb's `find_best_match` pub instead of using font-kit here` | **Using wrong matching API** |
| 229 | `TODO: Determine the proper system UI font.` | Hardcoded to IBM Plex Sans |
| 247 | `HACK: To let the storybook run...` | Font fallback issues |

And from `/tmp/zed-source/crates/gpui/src/text_system.rs`:

| Line | TODO | Impact |
|------|------|--------|
| 74 | `TODO: Remove this when Linux have implemented setting fallbacks.` | Fallback stack incomplete |

## Potential Fixes (in gpui)

### Option 1: Fix font_id resolution
The `font_id()` function uses `font_kit::matching::find_best_match` but then `layout_line()` uses the font face's actual properties rather than confirming the match was successful. Need to verify the returned font actually matches requested weight/style.

### Option 2: Use cosmic-text's font APIs consistently
The TODO at line 99 suggests uncertainty about whether to use cosmic-text or font-kit APIs. Cosmic-text has its own font matching that might work better.

### Option 3: Explicit font face loading
Instead of loading a font family and trying to match, explicitly load font faces by their full names (e.g., "Liberation Sans Bold" instead of "Liberation Sans" + weight=BOLD).

## Workaround Attempts (Failed)

1. **Explicit font family** - Set `font_family = "Liberation Sans"` which has bold/italic variants installed. Result: Still uses regular weight.

2. **StyledText with explicit TextStyle** - Used `StyledText::new(text).with_default_highlights(&text_style, [])`. Result: Correct TextRuns created but not rendered differently.

3. **div style inheritance** - Used `div().font_weight(FontWeight::BOLD).child(text)`. Result: No effect.

## Verification

To verify Liberation Sans has all variants:
```bash
$ fc-list : family style | grep "Liberation Sans"
Liberation Sans:style=Regular
Liberation Sans:style=Bold
Liberation Sans:style=Bold Italic
Liberation Sans:style=Italic
```

## Related Issues

- This may affect Zed editor's syntax highlighting on Linux (if they use bold for keywords)
- The font picker feature in VisiGrid is also affected
- Any gpui application needing font weight/style variants on Linux

## Upstream

This should be reported to Zed/gpui:
- Repository: https://github.com/zed-industries/zed
- Relevant file: `crates/gpui/src/platform/linux/text_system.rs`

## Current VisiGrid Status

The code is correctly implemented and will work when gpui fixes this. The formatting data is stored and persisted correctly - only the visual rendering is affected on Linux.

```rust
// gpui-app/src/views/grid.rs - Current implementation
if format.bold {
    text_style.font_weight = FontWeight::BOLD;
}
if format.italic {
    text_style.font_style = FontStyle::Italic;
}
// This creates correct TextRuns, but cosmic-text doesn't render them differently
cell = cell.child(StyledText::new(text_content).with_default_highlights(&text_style, []));
```
