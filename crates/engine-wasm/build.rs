//! Stamp the commit the wasm bundle was built from.
//!
//! The crate version cannot distinguish a build made at a release tag from one
//! made a few commits later — both report the same number. That is exactly the
//! difference that matters to anything vendoring this artifact and claiming a
//! version alongside it, so the commit is recorded too.
//!
//! When the commit cannot be determined the value is empty rather than
//! something plausible like "unknown". A consumer checking it should treat
//! empty as a failure to identify the build; a placeholder would let the check
//! pass while verifying nothing.

use std::process::Command;

fn main() {
    let commit = Command::new("git")
        .args(["rev-parse", "HEAD"])
        .output()
        .ok()
        .filter(|out| out.status.success())
        .and_then(|out| String::from_utf8(out.stdout).ok())
        .map(|s| s.trim().to_string())
        .unwrap_or_default();

    let dirty = Command::new("git")
        .args(["status", "--porcelain", "--untracked-files=no"])
        .output()
        .ok()
        .filter(|out| out.status.success())
        .map(|out| !out.stdout.is_empty())
        .unwrap_or(false);

    // A build from a modified tree is not the commit it names, and saying so is
    // the whole point of recording it.
    let stamp = match (commit.is_empty(), dirty) {
        (true, _) => String::new(),
        (false, true) => format!("{commit}-modified"),
        (false, false) => commit,
    };

    println!("cargo:rustc-env=VISIGRID_ENGINE_COMMIT={stamp}");
    println!("cargo:rerun-if-changed=../../.git/HEAD");
}
