#!/usr/bin/env bash
# Fail if a Linux binary needs a newer glibc than the one we promise.
#
# The requirement comes from the build machine, not the code: linking on a
# newer distro binds libm/libc symbols to that distro's versions. When the
# x86_64 runner label moved from Ubuntu 24.04 to 26.04, the desktop binary
# began to require glibc 2.43 and stopped starting on Debian 13 and Ubuntu
# 24.04 (#183). The build image is pinned; this catches the next drift.
#
# Usage: scripts/ci/check-glibc-floor.sh <max-version> <binary>...
set -euo pipefail

max="${1:?usage: $0 <max-version> <binary>...}"
shift
[ "$#" -gt 0 ] || { echo "usage: $0 <max-version> <binary>..." >&2; exit 2; }

status=0
for bin in "$@"; do
  [ -f "$bin" ] || { echo "not found: $bin" >&2; exit 2; }
  needed=$(objdump -T "$bin" | grep -o 'GLIBC_[0-9][0-9.]*' | sed 's/^GLIBC_//' | sort -uV | tail -1)
  [ -n "$needed" ] || { echo "no glibc symbol versions found in $bin" >&2; exit 2; }
  newest=$(printf '%s\n%s\n' "$max" "$needed" | sort -V | tail -1)
  if [ "$newest" != "$max" ]; then
    echo "FAIL $bin needs glibc $needed (limit $max):" >&2
    objdump -T "$bin" | grep -E "GLIBC_${needed//./\\.}\)" >&2 || true
    status=1
  else
    echo "ok   $bin needs glibc $needed (limit $max)"
  fi
done
exit "$status"
