use std::process::Command;

fn main() {
    // Embed git commit hash for version info
    println!("cargo:rerun-if-changed=../../.git/HEAD");
    println!("cargo:rerun-if-changed=../../.git/refs/heads");
    // Tags matter too: creating one turns an unreleased build into a released
    // one, and without this the stale marker would survive the rebuild.
    println!("cargo:rerun-if-changed=../../.git/refs/tags");

    let git_hash = Command::new("git")
        .args(["rev-parse", "--short=7", "HEAD"])
        .output()
        .ok()
        .and_then(|output| {
            if output.status.success() {
                String::from_utf8(output.stdout).ok()
            } else {
                None
            }
        })
        .map(|s| s.trim().to_string())
        .unwrap_or_else(|| "unknown".to_string());

    println!("cargo:rustc-env=GIT_COMMIT_HASH={}", git_hash);

    // Say so when this build is not a released version.
    //
    // The version number comes from Cargo.toml, which is bumped as part of
    // making a release — so every build between one release and the next
    // reports the version it is working towards, not one that exists. A
    // documentation page verified against such a build cites an engine nobody
    // can install, and the version string alone gives no hint of it.
    //
    // The release workflow sets VGRID_RELEASE_TAG, which is authoritative: its
    // shallow checkout may not carry the tag object, so asking git there could
    // brand a genuine release "unreleased". Locally there is no such variable
    // and `--exact-match` answers correctly.
    println!("cargo:rerun-if-env-changed=VGRID_RELEASE_TAG");
    let is_tagged = std::env::var("VGRID_RELEASE_TAG").is_ok_and(|v| !v.is_empty())
        || Command::new("git")
            .args(["describe", "--tags", "--exact-match", "HEAD"])
            .output()
            .map(|o| o.status.success())
            .unwrap_or(false);
    println!(
        "cargo:rustc-env=VGRID_RELEASE_STATE={}",
        if is_tagged { "" } else { ", unreleased" }
    );

    // Pass build target triple to source code
    let target = std::env::var("TARGET").unwrap_or_else(|_| "unknown".to_string());
    println!("cargo:rustc-env=TARGET={}", target);
}
