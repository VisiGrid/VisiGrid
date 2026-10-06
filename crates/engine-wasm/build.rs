//! Stamp the commit the wasm bundle was built from.
//!
//! The crate version cannot distinguish a build made at a release tag from one
//! made a few commits later — both report the same number. That is exactly the
//! difference that matters to anything vendoring this artifact and claiming a
//! version alongside it, so the commit is recorded too. The collaboration
//! server also compares it with its engine host's (see the shared helper).

include!("../../build-support/engine_commit.rs");

fn main() {
    println!("cargo:rustc-env=VISIGRID_ENGINE_COMMIT={}", engine_commit_stamp());
}
