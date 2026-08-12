#!/usr/bin/env python3
"""Rebuild the bundled symbol fallback font from the glyphs the UI actually uses.

The UI font (IBM Plex) carries Latin and a few symbols but not ✕ ▾ ⌥ ⏱ ⚙ ↵ ─ ⚠.
Without a bundled fallback those depend on whatever fonts the machine has, which
works on macOS and is a coin toss on Linux — ↵ is in a dozen fonts locally and
renders, ⏱ is in four and shows a missing-glyph box.

Run this whenever the UI gains a character the font does not have. The test in
gpui-app/src/main.rs fails in that situation and names the missing character,
so you should not have to remember to.

    python3 scripts/build-symbol-font.py [--source PATH]

Source font: Adwaita Mono, SIL OFL 1.1 with no reserved font name, so subsetting
and redistribution are permitted provided the licence travels with it (it sits
beside the output as LICENSE-OFL.txt). Any OFL font with the coverage would do;
pass --source to use a different one.
"""

import argparse
import pathlib
import subprocess
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
UI_SRC = REPO / "gpui-app" / "src"
OUT_DIR = REPO / "gpui-app" / "assets" / "fonts" / "visigrid-symbols"
OUT_FONT = OUT_DIR / "VisiGridSymbols-Regular.ttf"
MANIFEST = OUT_DIR / "COVERED-GLYPHS.txt"
FAMILY = "VisiGrid Symbols"
DEFAULT_SOURCE = "/usr/share/fonts/Adwaita/AdwaitaMono-Regular.ttf"


def in_scope(cp: int) -> bool:
    """Is this a character the bundled font should carry?

    Kept identical to `is_in_scope` in gpui-app/src/main.rs — that test is the
    authority on what must be covered, and this decides what gets built. If the
    two ever disagree the test fails, which is the safe direction.
    """
    if cp < 0x80:
        return False  # ASCII: every font has it
    if 0x1F000 <= cp:
        return False  # emoji live in a colour font, not this one
    if cp in (0x2601, 0x2615):
        # ☁ (sync indicator) and ☕ — excluded because the source font does not
        # have them, not because they do not matter. They still depend on the
        # machine having a font that does; around eighteen local fonts carry ☁,
        # so it renders here, but that is luck rather than a guarantee. A source
        # font covering them would let both be dropped from this list.
        return False
    if 0x3000 <= cp <= 0x9FFF or 0xFE00 <= cp <= 0xFE0F:
        return False  # CJK and variation selectors need their own fonts
    return True


def sweep() -> set[int]:
    """Every in-scope character appearing anywhere in the UI source."""
    found: set[int] = set()
    for path in sorted(UI_SRC.rglob("*.rs")):
        for ch in path.read_text(encoding="utf-8", errors="replace"):
            if in_scope(ord(ch)):
                found.add(ord(ch))
    return found


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=DEFAULT_SOURCE)
    args = ap.parse_args()

    source = pathlib.Path(args.source)
    if not source.exists():
        print(f"source font not found: {source}", file=sys.stderr)
        print("Install adwaita-fonts, or pass --source with any OFL font that", file=sys.stderr)
        print("covers the characters listed in", MANIFEST, file=sys.stderr)
        return 1

    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        print("fonttools is required: pip install fonttools", file=sys.stderr)
        return 1

    wanted = sweep()
    print(f"{len(wanted)} in-scope characters used by the UI")

    src = TTFont(source, lazy=True)
    available: set[int] = set()
    for table in src["cmap"].tables:
        available |= set(table.cmap.keys())

    missing = sorted(wanted - available)
    if missing:
        # Loud, because the alternative is a font that silently lacks a glyph
        # and a UI that silently falls back to the machine's fonts.
        print("source font cannot supply:", file=sys.stderr)
        for cp in missing:
            print(f"  U+{cp:04X} {chr(cp)}", file=sys.stderr)
        print("Pick a source font that covers these, or exclude them in in_scope().", file=sys.stderr)
        return 1

    covered = sorted(wanted)
    unicodes = ",".join(f"U+{cp:04X}" for cp in covered)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    result = subprocess.run(
        ["pyftsubset", str(source), f"--unicodes={unicodes}",
         f"--output-file={OUT_FONT}", "--no-hinting", "--desubroutinize",
         "--name-IDs=*", "--drop-tables+=DSIG"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(result.stderr, file=sys.stderr)
        return 1

    # Rename so the subset is not passed off as the original, and so the
    # fallback chain has a stable name to ask for.
    font = TTFont(OUT_FONT)
    for rec in font["name"].names:
        if rec.nameID in (1, 4, 16):
            rec.string = FAMILY
        elif rec.nameID == 2:
            rec.string = "Regular"
        elif rec.nameID == 6:
            rec.string = "VisiGridSymbols-Regular"
        elif rec.nameID == 3:
            rec.string = "VisiGridSymbols-Regular-subset"
    font.save(OUT_FONT)

    MANIFEST.write_text(
        "# Regenerate with: python3 scripts/build-symbol-font.py\n"
        "# Every character here is one the UI uses and the UI font lacks.\n"
        + "".join(f"U+{cp:04X}\t{chr(cp)}\n" for cp in covered)
    )

    print(f"wrote {OUT_FONT.relative_to(REPO)} ({OUT_FONT.stat().st_size:,} bytes)")
    print(f"wrote {MANIFEST.relative_to(REPO)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
