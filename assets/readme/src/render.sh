#!/bin/bash
# Render README art from these HTML templates into assets/readme/*.png.
# Usage: ./render.sh [name...]   (no args renders everything)
set -euo pipefail
cd "$(dirname "$0")"
# Set CHROME to override, e.g. CHROME=chromium ./render.sh
if [ -z "${CHROME:-}" ]; then
  for c in "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" google-chrome-stable google-chrome chromium; do
    if command -v "$c" >/dev/null 2>&1 || [ -x "$c" ]; then CHROME="$c"; break; fi
  done
fi
: "${CHROME:?no Chrome found; set CHROME}"
height() { case $1 in hero) echo 1080;; *) echo 820;; esac; }
[ $# -eq 0 ] && set -- hero card-numbers card-palette card-engine card-verify card-formats
for name in "$@"; do
  "$CHROME" --headless=new --disable-gpu --hide-scrollbars --force-device-scale-factor=1 \
    --virtual-time-budget=4000 --window-size=2000,$(height $name) \
    --screenshot="$PWD/../$name.png" "file://$PWD/$name.html" 2>/dev/null
  echo "rendered $name.png"
done
