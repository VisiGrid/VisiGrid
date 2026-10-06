use std::process::Command;

include!("../build-support/engine_commit.rs");

fn main() {
    println!("cargo:rustc-env=VISIGRID_ENGINE_COMMIT={}", engine_commit_stamp());
    // Embed git commit hash for version info
    println!("cargo:rerun-if-changed=../.git/HEAD");
    println!("cargo:rerun-if-changed=../.git/refs/heads");

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

    // Only run Windows resource compilation on Windows
    #[cfg(target_os = "windows")]
    {
        let mut res = winresource::WindowsResource::new();
        res.set_icon("windows/visigrid.ico");
        // Don't set manifest - gpui already provides one
        res.set("FileDescription", "VisiGrid - A fast, keyboard-driven spreadsheet");
        res.set("ProductName", "VisiGrid");
        res.set("CompanyName", "RegAtlas, LLC");
        res.set("LegalCopyright", "Copyright © 2025 RegAtlas, LLC");
        res.compile().expect("Failed to compile Windows resources");
    }
}
