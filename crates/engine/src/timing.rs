//! `Instant` that compiles on wasm32.
//!
//! `std::time::Instant::now()` panics ("time not implemented") on
//! wasm32-unknown-unknown, which turns every recompute into an
//! `unreachable` trap in the browser. Recalc timing is diagnostics-only
//! (RecalcReport durations), so on wasm it reports zero. Native targets
//! get the real `std::time::Instant`, re-exported — behavior unchanged.

#[cfg(not(target_arch = "wasm32"))]
pub(crate) use std::time::Instant;

#[cfg(target_arch = "wasm32")]
#[derive(Clone, Copy, Debug)]
pub(crate) struct Instant;

#[cfg(target_arch = "wasm32")]
impl Instant {
    pub(crate) fn now() -> Self {
        Instant
    }

    pub(crate) fn elapsed(&self) -> std::time::Duration {
        std::time::Duration::ZERO
    }
}

/// Wall-clock for RecalcReport metadata: real on native, UNIX_EPOCH on wasm
/// (SystemTime::now panics there; the field is diagnostics-only).
pub(crate) fn system_now() -> std::time::SystemTime {
    #[cfg(not(target_arch = "wasm32"))]
    {
        std::time::SystemTime::now()
    }
    #[cfg(target_arch = "wasm32")]
    {
        std::time::SystemTime::UNIX_EPOCH
    }
}

/// Duration since the Unix epoch, for volatile functions (NOW/TODAY/RAND).
/// Zero on wasm — the web verify layer never verifies volatile formulas,
/// so the stub value is never user-visible.
pub(crate) fn now_since_epoch() -> std::time::Duration {
    #[cfg(not(target_arch = "wasm32"))]
    {
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or(std::time::Duration::ZERO)
    }
    #[cfg(target_arch = "wasm32")]
    {
        std::time::Duration::ZERO
    }
}
