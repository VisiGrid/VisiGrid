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

        git pull --rebase || die "Failed to pull AUR repo. Resolve conflicts manually."

        sed_i "s/^pkgver=.*/pkgver=$VERSION/" PKGBUILD
        sed_i "s/^sha256sums=.*/sha256sums=('$SHA256')/" PKGBUILD

        bold "Generating .SRCINFO..."
        makepkg --printsrcinfo > .SRCINFO

        git add PKGBUILD .SRCINFO
        git commit -m "Bump to v$VERSION"
        git push

        cd "$REPO_ROOT"
    fi

    green "AUR updated."
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

    if $IS_LINUX; then
        bold "Checking AUR..."
        AUR_VER="$(curl -s "https://aur.archlinux.org/rpc/v5/info?arg[]=visigrid-bin" | jq -r '.results[0].Version' 2>/dev/null || echo "unknown")"
        echo "AUR version: $AUR_VER"
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
echo ""
bold "Done!"
