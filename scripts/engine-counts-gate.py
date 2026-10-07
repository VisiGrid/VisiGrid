#!/usr/bin/env python3
"""Compare engine instruction counts against the checked-in ceilings.

    scripts/engine-counts-gate.py target/engine-counts/results.json [--ratchet]

The ceilings live in benches/engine-counts.ceilings.<arch>.json, one number
per scenario. A count above its ceiling by more than the tolerance fails.
A count below it does not change anything on a pull request; the nightly
run passes --ratchet, which lowers each ceiling to the count it saw, so the
ceiling only ever goes down. A scenario with no ceiling yet is reported and
passes, and --ratchet writes it, which is how a new scenario is seeded.

Counts differ between CPU architectures, so the file is per arch and the
gate is skipped when there is no file for the one it runs on.
"""
import json
import os
import platform
import sys

TOLERANCE = 0.01  # 1%: HashMap seeding and allocator placement move counts a little

def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    results_path = sys.argv[1]
    ratchet = "--ratchet" in sys.argv
    arch = platform.machine()
    ceilings_path = os.path.join("benches", f"engine-counts.ceilings.{arch}.json")

    with open(results_path) as f:
        results = {k: int(v) for k, v in json.load(f).items()}
    ceilings = {}
    if os.path.exists(ceilings_path):
        with open(ceilings_path) as f:
            ceilings = {k: int(v) for k, v in json.load(f).items()}
    elif not ratchet:
        print(f"no ceilings for {arch} ({ceilings_path}); nothing to gate")
        return 0

    failed = []
    lowered = {}
    print(f"{'scenario':<18} {'count':>14} {'ceiling':>14} {'change':>9}")
    for name, count in sorted(results.items()):
        ceiling = ceilings.get(name)
        if ceiling is None:
            print(f"{name:<18} {count:>14,} {'(none)':>14} {'seed':>9}")
            lowered[name] = count
            continue
        change = (count - ceiling) / ceiling
        flag = ""
        if count > ceiling * (1 + TOLERANCE):
            flag = "  FAIL"
            failed.append((name, count, ceiling, change))
        elif count < ceiling:
            lowered[name] = count
        print(f"{name:<18} {count:>14,} {ceiling:>14,} {change:>+8.2%}{flag}")

    if failed:
        print()
        for name, count, ceiling, change in failed:
            print(f"{name}: {count:,} instructions is {change:+.2%} over the ceiling of {ceiling:,}.")
        print("Either the change made this path slower, or it is legitimately more work;")
        print("if the latter, raise the ceiling in the same PR and say why in the commit.")
        return 1

    if ratchet and lowered:
        ceilings.update(lowered)
        os.makedirs(os.path.dirname(ceilings_path), exist_ok=True)
        with open(ceilings_path, "w") as f:
            json.dump(dict(sorted(ceilings.items())), f, indent=2)
            f.write("\n")
        print(f"\nlowered {len(lowered)} ceiling(s) in {ceilings_path}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
