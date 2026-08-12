#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-}"
DRY_RUN=false
[[ "${2:-}" == "--dry-run" ]] && DRY_RUN=true

REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
AUR_DIR="$HOME/Code/visigrid-bin"
GITHUB_REPO="VisiGrid/VisiGrid"
HOMEBREW_REPO="VisiGrid/homebrew-visigrid"

# --- Platform detection ---

IS_MACOS=false
IS_LINUX=false
case "$(uname -s)" in
    Darwin) IS_MACOS=true ;;
    Linux)  IS_LINUX=true ;;
    *)      echo "Warning: unknown platform $(uname -s), assuming Linux-like." ; IS_LINUX=true ;;
esac

# --- Helpers ---

bold() { printf '\033[1m%s\033[0m\n' "$*"; }
green() { printf '\033[32m%s\033[0m\n' "$*"; }
red() { printf '\033[31m%s\033[0m\n' "$*"; }
yellow() { printf '\033[33m%s\033[0m\n' "$*"; }

die() { red "ERROR: $*" >&2; exit 1; }

# Set when a post-publish step fails. The run continues so the summary can
# report what landed, then exits non-zero.
AUR_FAILED=false
# Computed in phase 5. Declared here so the summary can quote it without
# depending on how far the run got — under `set -u` an unset variable there
# would abort the very report that explains what went wrong.
SHA256=""


# Update the AUR package. Returns non-zero on any failure rather than exiting,
# because this phase runs after the release is published and the caller needs
# to report the partial state rather than stop mid-way.
aur_update() {
    git pull --rebase || return 1
    sed_i "s/^pkgver=.*/pkgver=$VERSION/" PKGBUILD || return 1
    sed_i "s/^sha256sums=.*/sha256sums=('$SHA256')/" PKGBUILD || return 1
    bold "Generating .SRCINFO..."
    makepkg --printsrcinfo > .SRCINFO || return 1
    git add PKGBUILD .SRCINFO || return 1
    git commit -m "Bump to v$VERSION" || return 1
    git push || return 1
}

run() {
    if $DRY_RUN; then
        yellow "[dry-run] $*"
    else
        "$@"
    fi
}

# Portable SHA-256: prefer sha256sum, fall back to shasum -a 256 (macOS)
sha256() {
    if command -v sha256sum &>/dev/null; then
        sha256sum "$1" | awk '{print $1}'
    elif command -v shasum &>/dev/null; then
        shasum -a 256 "$1" | awk '{print $1}'
    else
        die "No sha256sum or shasum found."
    fi
}

# Portable sed -i: macOS sed requires '' after -i, GNU sed does not.
sed_i() {
    if $IS_MACOS; then
        sed -i '' "$@"
    else
        sed -i "$@"
    fi
}

# --- Phase 1: Pre-flight checks ---

bold "=== Phase 1: Pre-flight checks ==="

# Version argument
[[ -z "$VERSION" ]] && die "Usage: $0 <version> [--dry-run]"
[[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] || die "Version must be semver (e.g. 0.6.6), got: $VERSION"

# Required tools (all platforms)
for cmd in gh cargo git sed curl jq; do
    command -v "$cmd" &>/dev/null || die "Required tool not found: $cmd"
done

# Linux-only tools (AUR)
if $IS_LINUX; then
    for cmd in makepkg; do
        command -v "$cmd" &>/dev/null || die "Required tool not found: $cmd (needed for AUR on Linux)"
    done
fi

cd "$REPO_ROOT"

# Branch check
BRANCH="$(git branch --show-current)"
[[ "$BRANCH" == "main" ]] || die "Must be on main branch (currently on: $BRANCH)"

# Clean working tree (ignore submodule changes with --ignore-submodules)
git diff --exit-code --quiet --ignore-submodules || die "Unstaged changes exist. Commit or stash them first."
git diff --cached --exit-code --quiet --ignore-submodules || die "Staged uncommitted changes exist. Commit or stash them first."

# Check for untracked .rs files (catches forgotten module files)
UNTRACKED_RS="$(git ls-files --others --exclude-standard -- '*.rs')"
if [[ -n "$UNTRACKED_RS" ]]; then
    red "Untracked .rs files found:"
    echo "$UNTRACKED_RS"
    die "Commit or remove these files before releasing."
fi

# Up to date with remote
git fetch origin main --quiet
LOCAL="$(git rev-parse HEAD)"
REMOTE="$(git rev-parse origin/main)"
[[ "$LOCAL" == "$REMOTE" ]] || die "Local main is not up to date with origin/main. Pull or push first."

# Tag doesn't already exist
if git rev-parse "v$VERSION" &>/dev/null 2>&1; then
    die "Tag v$VERSION already exists."
fi

# Build check
bold "Running cargo build..."
cargo build --release -p visigrid-gpui -p visigrid-cli || die "cargo build failed"

green "Pre-flight checks passed."

# --- Phase 2: Version bump ---

bold "=== Phase 2: Version bump ==="

CURRENT_VERSION="$(grep '^version' "$REPO_ROOT/Cargo.toml" | head -1 | sed 's/.*"\(.*\)".*/\1/')"
if [[ "$CURRENT_VERSION" == "$VERSION" ]]; then
    yellow "Cargo.toml already at version $VERSION, skipping bump."
else
    bold "Bumping version: $CURRENT_VERSION -> $VERSION"
    run sed_i "s/^version = \"$CURRENT_VERSION\"/version = \"$VERSION\"/" "$REPO_ROOT/Cargo.toml"
    bold "Updating Cargo.lock..."
    run cargo check --workspace
    run git add Cargo.toml Cargo.lock
    run git commit -m "Bump version to $VERSION"
    run git push origin main
fi

green "Version bump complete."

# --- Phase 3: Tag and wait for CI ---

bold "=== Phase 3: Tag and wait for CI ==="

# Re-check the branch here, not just in pre-flight.
#
# Several people and agents share this working directory, and a checkout
# anywhere in it moves HEAD for everyone. During the v0.21.0 release something
# checked out the release commit while CI was running, leaving the tree
# detached — harmless that time, because the commit already matched main, but
# only by luck. Pre-flight passed twenty minutes earlier and had nothing left
# to say about it.
#
# The tag is the point of no return, so the assertion belongs immediately
# before it rather than at the start.
# `branch --show-current` is empty when detached, which is the case that
# actually happened; `rev-parse --abbrev-ref` reports the literal "HEAD" there
# and reads like a branch name in an error message.
BRANCH_NOW="$(git branch --show-current)"
if [[ "$BRANCH_NOW" != "main" ]]; then
    WHERE="${BRANCH_NOW:-a detached HEAD}"
    die "HEAD is on $WHERE, not main — something checked out this working
     directory since pre-flight passed. Nothing has been tagged. Restore with
     'git checkout main', confirm it still points at what you meant to ship,
     and re-run. Use a git worktree for parallel checkouts."
fi

HEAD_NOW="$(git rev-parse HEAD)"
REMOTE_NOW="$(git rev-parse origin/main 2>/dev/null || echo "")"
if [[ -n "$REMOTE_NOW" && "$HEAD_NOW" != "$REMOTE_NOW" ]]; then
    die "main and origin/main diverged since pre-flight (local ${HEAD_NOW:0:7},
     remote ${REMOTE_NOW:0:7}) — someone pushed while this release was running.
     Nothing has been tagged. Reconcile, then re-run."
fi

run git tag "v$VERSION"
run git push origin "v$VERSION"

if $DRY_RUN; then
    yellow "[dry-run] Would wait for Release workflow to complete."
else
    bold "Waiting for Release workflow to start..."
    sleep 10

    # Find the workflow run for this tag
    TIMEOUT=1800  # 30 minutes
    INTERVAL=30
    ELAPSED=0

    while true; do
        STATUS="$(gh run list --workflow=release.yml --branch="v$VERSION" --limit=1 --json status,conclusion --jq '.[0]' 2>/dev/null || echo "")"

        if [[ -z "$STATUS" ]]; then
            if (( ELAPSED > 60 )); then
                die "No Release workflow run found for v$VERSION after 60s."
            fi
            echo "Waiting for workflow to appear..."
            sleep "$INTERVAL"
            ELAPSED=$((ELAPSED + INTERVAL))
            continue
        fi

        RUN_STATUS="$(echo "$STATUS" | jq -r '.status')"
        RUN_CONCLUSION="$(echo "$STATUS" | jq -r '.conclusion')"

        if [[ "$RUN_STATUS" == "completed" ]]; then
            if [[ "$RUN_CONCLUSION" == "success" ]]; then
                green "Release workflow completed successfully."
                break
            else
                die "Release workflow failed with conclusion: $RUN_CONCLUSION"
            fi
        fi

        if (( ELAPSED >= TIMEOUT )); then
            die "Timed out waiting for Release workflow (${TIMEOUT}s)."
        fi

        echo "Workflow status: $RUN_STATUS (${ELAPSED}s elapsed)..."
        sleep "$INTERVAL"
        ELAPSED=$((ELAPSED + INTERVAL))
    done
fi

green "CI complete."

# --- Phase 4: Publish release ---

bold "=== Phase 4: Publish release ==="

run gh release edit "v$VERSION" --draft=false

green "Release v$VERSION published. Homebrew and Winget workflows triggered."

# --- Phase 5: Update AUR (Linux only) ---

if $IS_LINUX; then
    bold "=== Phase 5: Update AUR ==="

    if [[ ! -d "$AUR_DIR" ]]; then
        die "AUR directory not found: $AUR_DIR"
    fi

    if $DRY_RUN; then
        yellow "[dry-run] Would download tarball, compute SHA, update PKGBUILD, push to AUR."
    else
        bold "Downloading Linux tarball for SHA256..."
        TARBALL_URL="https://github.com/$GITHUB_REPO/releases/download/v$VERSION/VisiGrid-linux-x86_64.tar.gz"

        # Wait for CDN propagation before downloading.
        # GitHub's CDN can serve stale/incomplete assets for up to 60s after
        # a release is published. We download twice with a gap and compare
        # checksums to ensure we have the final, stable asset.
        bold "Waiting 30s for CDN propagation..."
        sleep 30

        TMPFILE="$(mktemp)"
        TMPFILE2="$(mktemp)"
        trap "rm -f '$TMPFILE' '$TMPFILE2'" EXIT

        download_tarball() {
            local dest="$1"
            for attempt in 1 2 3 4 5; do
                if curl -sL -o "$dest" -w '%{http_code}' "$TARBALL_URL" | grep -q '^200$'; then
                    return 0
                fi
                if (( attempt == 5 )); then
                    return 1
                fi
                echo "Download attempt $attempt failed, retrying in 10s..."
                sleep 10
            done
        }

        download_tarball "$TMPFILE" || die "Failed to download tarball after 5 attempts: $TARBALL_URL"
        SHA_FIRST="$(sha256 "$TMPFILE")"

        # Second download after a gap to confirm CDN consistency
        bold "Verifying CDN consistency (second download in 15s)..."
        sleep 15
        download_tarball "$TMPFILE2" || die "Failed to download tarball (verification): $TARBALL_URL"
        SHA_SECOND="$(sha256 "$TMPFILE2")"

        if [[ "$SHA_FIRST" != "$SHA_SECOND" ]]; then
            yellow "CDN returned different checksums — waiting 60s and retrying..."
            sleep 60
            download_tarball "$TMPFILE" || die "Failed to download tarball (final): $TARBALL_URL"
            SHA_FIRST="$(sha256 "$TMPFILE")"
        fi

        SHA256="$SHA_FIRST"
        bold "SHA256: $SHA256"

        cd "$AUR_DIR"

        # Everything from here on is recoverable and happens AFTER the release
        # is published, so a failure must not abort the run: the operator still
        # needs the summary telling them what did and did not land. The failure
        # is carried to the end and the script exits non-zero there.
        if aur_update; then
            green "AUR updated."
        else
            AUR_FAILED=true
            red "AUR was NOT updated. The release itself is published and unaffected."
        fi

        cd "$REPO_ROOT"
    fi
else
    bold "=== Phase 5: Update AUR (skipped — not on Linux) ==="
    yellow "Run this script on Linux to update AUR, or update manually."
fi

# --- Phase 6: Verify ---

bold "=== Phase 6: Verify ==="

if $DRY_RUN; then
    yellow "[dry-run] Would verify Homebrew and AUR versions."
else
    bold "Checking Homebrew..."
    sleep 30  # Give the workflow time to run
    BREW_STATUS="$(gh run list --repo "$HOMEBREW_REPO" --limit=1 --json status,conclusion --jq '.[0].conclusion' 2>/dev/null || echo "unknown")"
    echo "Homebrew workflow conclusion: $BREW_STATUS"

    if $IS_LINUX && ! $AUR_FAILED; then
        # The AUR's RPC lags a successful push by a few minutes, so a single
        # read straight afterwards reports the previous version — it did so on
        # both 0.25.0 and 0.25.1, each time saying "not done" about something
        # that was. A check that is wrong in a predictable direction is worse
        # than no check: it stays in the summary looking authoritative while
        # everyone learns to disregard it. Poll, and say plainly when the wait
        # expires rather than printing whatever the last read happened to be.
        bold "Checking AUR (its RPC lags a push by a few minutes)..."
        AUR_VER="unknown"
        AUR_CONFIRMED=false
        for _ in $(seq 1 10); do
            AUR_VER="$(curl -s "https://aur.archlinux.org/rpc/v5/info?arg[]=visigrid-bin" | jq -r '.results[0].Version' 2>/dev/null || echo "unknown")"
            if [[ "$AUR_VER" == "$VERSION"* ]]; then
                AUR_CONFIRMED=true
                break
            fi
            sleep 30
        done
        if $AUR_CONFIRMED; then
            echo "AUR version: $AUR_VER"
        else
            yellow "AUR still reports $AUR_VER after 5 minutes."
            yellow "The push succeeded; this is the RPC lagging, or it needs checking by hand:"
            yellow "  https://aur.archlinux.org/packages/visigrid-bin"
        fi
    fi
fi

echo ""
bold "=== Release Summary ==="
green "Version:     $VERSION"
green "Tag:         v$VERSION"
green "Release:     https://github.com/$GITHUB_REPO/releases/tag/v$VERSION"
if $IS_MACOS; then
    yellow "AUR:         skipped (run on Linux to update)"
fi

if $AUR_FAILED; then
    echo ""
    red "=== AUR NOT UPDATED ==="
    red "The release is published and every other channel is current; only the"
    red "AUR package is still on its previous version. To finish it later:"
    echo ""
    echo "  cd $AUR_DIR"
    echo "  git pull --rebase"
    if [[ -n "$SHA256" ]]; then
        echo "  # set pkgver=$VERSION and sha256sums=('$SHA256') in PKGBUILD"
    else
        echo "  # set pkgver=$VERSION in PKGBUILD, and sha256sums to the sha256 of"
        echo "  #   https://github.com/$GITHUB_REPO/releases/download/v$VERSION/VisiGrid-linux-x86_64.tar.gz"
    fi
    echo "  makepkg --printsrcinfo > .SRCINFO"
    echo "  git add PKGBUILD .SRCINFO && git commit -m 'Bump to v$VERSION' && git push"
    echo ""
    red "Exiting non-zero: the release succeeded, this run did not fully complete."
    exit 1
fi

echo ""
bold "Done!"
