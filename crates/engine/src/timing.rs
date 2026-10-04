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

/// Overrides for what volatile functions read, so that several replicas of one
/// workbook (browsers, desktops and a server in a collaboration session) can
/// recalculate to identical values. Unset fields fall back to the machine:
/// the system clock, its time zone, and the process-wide random generator.
///
/// See the Collaborative Workbook Model spec, "Determinism": the server stamps
/// each operation with its clock and a seed, and every replica recalculates
/// with them.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RecalcClock {
    /// Milliseconds since the Unix epoch that NOW and TODAY report.
    pub now_ms: Option<i64>,
    /// Seconds east of UTC used for NOW and TODAY's local day.
    pub utc_offset_seconds: Option<i64>,
    /// When set, RAND/RANDBETWEEN/RANDARRAY are a pure function of this seed,
    /// the evaluating cell and the call's position in its formula.
    pub seed: Option<u64>,
}

thread_local! {
    static CLOCK: std::cell::Cell<Option<RecalcClock>> = const { std::cell::Cell::new(None) };
    /// The cell being evaluated and how many random draws it has made, for
    /// seeded RAND. Set by the workbook around each evaluation.
    static RANDOM_CELL: std::cell::Cell<Option<(u64, usize, usize, u64)>> = const { std::cell::Cell::new(None) };
}

/// Installs a clock for the current thread until dropped, restoring the
/// previous one (installs nest).
pub(crate) struct ClockGuard(Option<RecalcClock>);

impl ClockGuard {
    pub(crate) fn install(clock: Option<RecalcClock>) -> Self {
        ClockGuard(CLOCK.with(|c| c.replace(clock)))
    }
}

impl Drop for ClockGuard {
    fn drop(&mut self) {
        let previous = self.0;
        CLOCK.with(|c| c.set(previous));
    }
}

pub(crate) fn current_clock() -> Option<RecalcClock> {
    CLOCK.with(|c| c.get())
}

/// Marks the cell whose formula is being evaluated, so seeded RAND draws are
/// per cell. Cleared when dropped.
pub(crate) struct RandomCellGuard(Option<(u64, usize, usize, u64)>);

impl RandomCellGuard {
    pub(crate) fn enter(sheet: u64, row: usize, col: usize) -> Self {
        RandomCellGuard(RANDOM_CELL.with(|c| c.replace(Some((sheet, row, col, 0)))))
    }
}

impl Drop for RandomCellGuard {
    fn drop(&mut self) {
        let previous = self.0;
        RANDOM_CELL.with(|c| c.set(previous));
    }
}

/// A seeded draw, if a seed is installed: SplitMix64 over (seed, cell, n),
/// where n counts the draws this cell has made in this evaluation. The same
/// seed and document give the same numbers on every replica, whatever order
/// cells are evaluated in.
pub(crate) fn seeded_random_u64() -> Option<u64> {
    let seed = current_clock()?.seed?;
    let (sheet, row, col, n) = RANDOM_CELL.with(|c| {
        let cur = c.get().unwrap_or((0, 0, 0, 0));
        c.set(Some((cur.0, cur.1, cur.2, cur.3 + 1)));
        cur
    });
    let mut z = seed;
    for part in [sheet, row as u64, col as u64, n] {
        z = splitmix64(z ^ part.wrapping_mul(0x9E37_79B9_7F4A_7C15));
    }
    Some(z)
}

fn splitmix64(x: u64) -> u64 {
    let mut z = x.wrapping_add(0x9E37_79B9_7F4A_7C15);
    z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
    z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
    z ^ (z >> 31)
}

/// Duration since the Unix epoch, for volatile functions (NOW/TODAY/RAND).
/// An installed `RecalcClock` with `now_ms` wins over the machine clock.
pub(crate) fn now_since_epoch() -> std::time::Duration {
    if let Some(ms) = current_clock().and_then(|c| c.now_ms) {
        return std::time::Duration::from_millis(ms.max(0) as u64);
    }
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
    if let Some(offset) = current_clock().and_then(|c| c.utc_offset_seconds) {
        return offset;
    }
    use chrono::Offset;
    chrono::Local::now().offset().fix().local_minus_utc() as i64
}
