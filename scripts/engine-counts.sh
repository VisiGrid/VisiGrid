#!/usr/bin/env bash
# Count CPU instructions for the engine's hot paths with callgrind.
#
# Usage: scripts/engine-counts.sh [out-dir]
# Writes <out-dir>/<scenario>.callgrind and <out-dir>/results.json, then
# prints the counts. Requires valgrind (Linux). The gate that compares the
# results against the checked-in ceilings is scripts/engine-counts-gate.py.
#
# Why instructions: wall-clock time is what users feel but it is noisy, and
# a CI runner's milliseconds cannot gate a pull request. Instruction counts
# under callgrind are deterministic to well under a percent, so one run is a
# measurement. See docs/engine-counts.md.
set -euo pipefail

out="${1:-target/counts}"
mkdir -p "$out"

if ! command -v valgrind >/dev/null; then
  echo "valgrind not found; this script needs Linux with valgrind installed" >&2
  exit 3
fi

# The `counts` profile is release with symbols kept, so callgrind can see
# the measured_* boundary. Without the engine's default features the io
# crate skips DuckDB, which the scenarios do not use.
cargo build --quiet --locked -p visigrid-io --no-default-features --profile counts --example engine_counts
bin="target/counts/examples/engine_counts"

arch="$(uname -m)"
results="$out/results.json"
echo "{" > "$results"
first=1
for scenario in $("$bin" list); do
  valgrind --tool=callgrind --quiet \
    --toggle-collect='*measured_*' \
    --callgrind-out-file="$out/$scenario.callgrind" \
    "$bin" "$scenario"
  # callgrind writes the collected total as "summary:" (or "totals:" on
  # older versions). The first number is Ir, instructions executed.
  ir="$(grep -m1 -E '^(summary|totals):' "$out/$scenario.callgrind" | awk '{print $2}')"
  if [ -z "$ir" ]; then
    echo "no instruction total in $out/$scenario.callgrind" >&2
    exit 1
  fi
  [ $first -eq 1 ] || echo "," >> "$results"
  first=0
  printf '  "%s": %s' "$scenario" "$ir" >> "$results"
  printf '%-18s %14s Ir\n' "$scenario" "$ir"
done
printf '\n}\n' >> "$results"
echo "arch: $arch; results: $results"
