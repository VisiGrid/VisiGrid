#!/usr/bin/env bash
# The release cargo build, shared by release.yml (tag builds) and
# release-cache.yml (main, which saves the cache tag builds restore).
#
# Both workflows MUST build through this script. Cargo reuses a cached
# dependency only when the profile, target triple, features and build
# environment match exactly; if the two drifted, the cache would restore
# fine and DuckDB's C++ (~19 min on Linux) would quietly recompile anyway.
#
# Usage: scripts/ci/release-build.sh linux|windows|macos
set -euo pipefail

# Bound the memory used by bundled DuckDB C++ compilation.
export CARGO_BUILD_JOBS="${CARGO_BUILD_JOBS:-2}"

case "${1:-}" in
  linux)
    cargo build --release -p visigrid-gpui -p visigrid-cli
    ;;
  windows)
    # Two invocations, as before: one build per package.
    cargo build --release -p visigrid-gpui
    cargo build --release -p visigrid-cli
    ;;
  macos)
    # aarch64 (native) first. ssh2's vendored-openssl feature builds a static
    # target-architecture OpenSSL, so the Intel cross-build needs no Homebrew.
    cargo build --release --target aarch64-apple-darwin -p visigrid-gpui -p visigrid-cli
    cargo build --release --target x86_64-apple-darwin -p visigrid-gpui -p visigrid-cli
    ;;
  *)
    echo "usage: $0 linux|windows|macos" >&2
    exit 2
    ;;
esac
