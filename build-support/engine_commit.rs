// Shared by the build scripts of `visigrid-engine-wasm` and `visigrid-cli`
// (`include!`d, not a crate), so the browser engine's `engine_commit()` and
// `vgrid collab-host`'s hello report the identical string for the same source.
// The collaboration server refuses a client whose engine differs from its host
// (replicas recalculating with different engines would diverge), so the two
// must agree exactly.
//
// The stamp is, in order:
//   1. `VISIGRID_ENGINE_COMMIT` from the build environment, when set: for
//      builds without a usable `.git` (pass the commit being built);
//   2. the full `git rev-parse HEAD`, with `-modified` when tracked files
//      differ from it (a dirty build is not the commit it names);
//   3. empty: unidentified. Never a plausible placeholder, so a check against
//      it fails instead of passing while verifying nothing.

fn engine_commit_stamp() -> String {
    use std::process::Command;

    println!("cargo:rerun-if-env-changed=VISIGRID_ENGINE_COMMIT");
    if let Ok(v) = std::env::var("VISIGRID_ENGINE_COMMIT") {
        let v = v.trim();
        if !v.is_empty() {
            return v.to_string();
        }
    }

    let git = |args: &[&str]| {
        Command::new("git")
            .args(args)
            .output()
            .ok()
            .filter(|out| out.status.success())
            .and_then(|out| String::from_utf8(out.stdout).ok())
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
    };

    // Re-run when the commit moves. In a worktree `.git` is a pointer file,
    // so watch the real paths git reports, not `../../.git/HEAD`.
    if let Some(dir) = git(&["rev-parse", "--absolute-git-dir"]) {
        println!("cargo:rerun-if-changed={dir}/HEAD");
        println!("cargo:rerun-if-changed={dir}/index");
    }
    if let Some(common) = git(&["rev-parse", "--path-format=absolute", "--git-common-dir"]) {
        println!("cargo:rerun-if-changed={common}/packed-refs");
        if let Some(head_ref) = git(&["symbolic-ref", "-q", "HEAD"]) {
            println!("cargo:rerun-if-changed={common}/{head_ref}");
        }
    }

    let Some(commit) = git(&["rev-parse", "HEAD"]) else {
        return String::new();
    };
    let dirty = Command::new("git")
        .args(["status", "--porcelain", "--untracked-files=no"])
        .output()
        .ok()
        .filter(|out| out.status.success())
        .map(|out| !out.stdout.is_empty())
        .unwrap_or(false);
    if dirty {
        format!("{commit}-modified")
    } else {
        commit
    }
}
