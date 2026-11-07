#!/usr/bin/env python3
"""Exit 0 if the release build's cache inputs changed between two commits, else 1.

Used by release-cache.yml to skip a rebuild on pushes that cannot change the
cached dependencies, above all the release version bump, which rewrites
Cargo.lock and every Cargo.toml but only for VisiGrid's own crates.

Changed means: rust-toolchain.toml differs, or the set of external packages in
Cargo.lock (those with a `source`, i.e. from crates.io or git, not path
crates) differs.

Usage: deps-changed.py BASE_SHA HEAD_SHA
"""
import subprocess
import sys
import tomllib


def show(sha, path):
    r = subprocess.run(["git", "show", f"{sha}:{path}"], capture_output=True, text=True)
    return r.stdout if r.returncode == 0 else None


def external_packages(lock_text):
    if lock_text is None:
        return None
    lock = tomllib.loads(lock_text)
    return {
        (p["name"], p["version"], p["source"])
        for p in lock.get("package", [])
        if "source" in p
    }


base, head = sys.argv[1], sys.argv[2]
if show(base, "rust-toolchain.toml") != show(head, "rust-toolchain.toml"):
    print("rust-toolchain.toml changed")
    sys.exit(0)
before, after = external_packages(show(base, "Cargo.lock")), external_packages(show(head, "Cargo.lock"))
if before != after:
    changed = sorted((before or set()) ^ (after or set()))
    print(f"{len(changed)} external package entries changed, e.g. {changed[:5]}")
    sys.exit(0)
print("no dependency or toolchain change")
sys.exit(1)
