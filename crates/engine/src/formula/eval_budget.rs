//! Temporary array limits for interactive validation evaluation. Ordinary
//! formula evaluation retains its existing policy; nested validation helpers
//! share the outer budget instead of resetting it.
use std::cell::Cell;

#[derive(Clone, Copy)]
struct Budget {
    remaining: usize,
    per_array: usize,
}
thread_local! { static BUDGET: Cell<Option<Budget>> = const { Cell::new(None) }; }

pub(crate) fn validation<T>(f: impl FnOnce() -> T) -> T {
    if BUDGET.with(|b| b.get().is_some()) {
        return f();
    }
    struct Reset;
    impl Drop for Reset {
        fn drop(&mut self) {
            BUDGET.with(|b| b.set(None));
        }
    }
    BUDGET.with(|b| {
        b.set(Some(Budget {
            remaining: 1_000_000,
            per_array: 100_000,
        }))
    });
    let _reset = Reset;
    f()
}

pub(crate) fn array(rows: usize, cols: usize) -> Result<(), String> {
    let cells = rows.checked_mul(cols).ok_or("#NUM! Array size overflow")?;
    BUDGET.with(|state| {
        if let Some(mut budget) = state.get() {
            if cells > budget.per_array || cells > budget.remaining {
                return Err("#NUM! Validation formula array limit exceeded".into());
            }
            budget.remaining -= cells;
            state.set(Some(budget));
        }
        Ok(())
    })
}

pub(crate) fn max_array_cells() -> usize {
    BUDGET.with(|state| state.get().map_or(usize::MAX, |budget| budget.per_array))
}
