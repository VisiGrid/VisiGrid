//! `Instant` and wall-clock that compile on wasm32.
//!
//! `std::time::Instant::now()` and `SystemTime::now()` both panic ("time not
//! implemented") on wasm32-unknown-unknown, so this module supplies them. On
//! native it re-exports the real thing; on wasm it reads the JS clock.
//!
//! These used to return zero on wasm, on the reasoning that the only caller
//! was the web verify layer and it never verified volatile formulas. That
//! stopped being true — conditional formatting and validation evaluate through
//! the same bundle, so a rule like `=A1>TODAY()` compared against 1970 and
//! flagged every row, with no error and nothing in the divergence report.
//!
//! A stub is only invisible while its callers stay the ones you had in mind,
//! and nothing tells you when that changes.

#[cfg(not(target_arch = "wasm32"))]
pub(crate) use std::time::Instant;

/// Milliseconds since the Unix epoch, from the JS clock.
#[cfg(target_arch = "wasm32")]
fn js_epoch_millis() -> f64 {
    let ms = js_sys::Date::now();
    if ms.is_finite() && ms > 0.0 { ms } else { 0.0 }
}

#[cfg(target_arch = "wasm32")]
#[derive(Clone, Copy, Debug)]
pub(crate) struct Instant(f64);

#[cfg(target_arch = "wasm32")]
impl Instant {
    pub(crate) fn now() -> Self {
        Instant(js_epoch_millis())
    }

    /// Resolution is the browser's, typically a millisecond and sometimes
    /// coarsened further for fingerprinting reasons. Fine for the phase
    /// timings in RecalcReport, which are tens of milliseconds and up, and
    /// far better than the zero this used to report — a recompute in the
    /// browser previously claimed every phase took no time at all.
    pub(crate) fn elapsed(&self) -> std::time::Duration {
        let ms = (js_epoch_millis() - self.0).max(0.0);
        std::time::Duration::from_secs_f64(ms / 1000.0)
    }
}

/// Duration since the Unix epoch, for volatile functions (NOW/TODAY/RAND).
pub(crate) fn now_since_epoch() -> std::time::Duration {
    #[cfg(not(target_arch = "wasm32"))]
    {
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or(std::time::Duration::ZERO)
    }
    #[cfg(target_arch = "wasm32")]
    {
        std::time::Duration::from_secs_f64(js_epoch_millis() / 1000.0)
    }
}

/// Wall-clock for RecalcReport metadata.
///
/// Derived from `now_since_epoch` rather than read separately, so the two can
/// never disagree about what time it is on one platform and not the other.
pub(crate) fn system_now() -> std::time::SystemTime {
    std::time::UNIX_EPOCH + now_since_epoch()
}

/// Seconds to add to UTC to reach local wall-clock time.
///
/// Excel's TODAY and NOW are local, not UTC. Without this the serial rolls
/// over at midnight UTC, so anyone west of it gets tomorrow's date for the
/// last hours of their evening — five hours a day in US Central, eight in
/// Pacific — and a rule like `=A1<TODAY()` marks work overdue a day early.
///
/// chrono reads the system zone on native and the browser's on wasm, so both
/// targets agree with the machine the user is looking at.
pub(crate) fn local_utc_offset_seconds() -> i64 {
    use chrono::Offset;
    chrono::Local::now().offset().fix().local_minus_utc() as i64
}
