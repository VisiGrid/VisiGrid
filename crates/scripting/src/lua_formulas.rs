//! The Lua-aware formula adapter: what the desktop app and the CLI install so
//! that custom functions from `functions.lua` and `=LUA` cells compute
//! everywhere a workbook is recalculated, including after every edit and
//! every agent op, in `vgrid serve`, and in headless CLI commands.
//!
//! Installed once per process with [`install`], which registers [`dispatch`]
//! as the engine's default custom-function handler. The engine consults it
//! from every evaluation path that has no explicit handler. The state behind
//! it, a sandboxed Lua runtime, the loaded registry, the compiled-chunk cache
//! and the memo, lives in a thread-local, built lazily on the first call on
//! each thread: the desktop loads files on background threads and the engine
//! must stay `Send`, and a Lua state is neither.
//!
//! The session host never sees any of this. It applies ops to a workbook and
//! the workbook does the rest, which is the point: it stays Lua-agnostic.

use std::cell::{Cell, RefCell};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;

use visigrid_engine::formula::eval::{EvalArg, EvalResult};
use visigrid_engine::recalc::RecalcReport;
use visigrid_engine::workbook::Workbook;

use crate::custom_functions::{
    call_custom_function, load_custom_functions, CustomFunctionRegistry, MemoCache,
    RESERVED_LUA_CELL_NAME,
};
use crate::lua_cell::{call_lua_cell, ChunkCache};
use crate::runtime::{Limits, LuaRuntime};

/// The name of the cell-chunk function.
pub const LUA_CELL_FUNCTION: &str = RESERVED_LUA_CELL_NAME;

/// Bumped by every reload. A thread whose host is older than this rebuilds
/// it on its next use, so "Reload Custom Functions" reaches the background
/// threads that load files, not only the thread that ran the command.
static GENERATION: AtomicU64 = AtomicU64::new(1);

struct Host {
    runtime: LuaRuntime,
    registry: CustomFunctionRegistry,
    load_error: Option<String>,
    chunks: RefCell<ChunkCache>,
    /// Memo for one full recalc: filled between the start and end of
    /// [`recompute`], empty otherwise. That is the contract the cache was
    /// written to (a function may capture mutable state, and an error may be
    /// a timeout), and the incremental path evaluates too few cells to need
    /// one.
    memo: RefCell<MemoCache>,
    in_recalc: Cell<bool>,
    lua_time_us: Cell<u64>,
    generation: u64,
}

impl Host {
    fn load() -> Self {
        let runtime = LuaRuntime::with_limits(Limits::batch()).unwrap_or_else(|e| {
            // A Lua state that cannot be created leaves every custom function
            // unknown, which the engine reports per cell. Not fatal.
            eprintln!("custom functions: Lua runtime unavailable: {}", e);
            LuaRuntime::default()
        });
        let (registry, load_error) = match load_custom_functions(runtime.lua()) {
            Ok(registry) => (registry, None),
            Err(e) => (CustomFunctionRegistry::empty(), Some(e)),
        };
        Host {
            runtime,
            registry,
            load_error,
            chunks: RefCell::new(ChunkCache::new()),
            memo: RefCell::new(MemoCache::new()),
            in_recalc: Cell::new(false),
            lua_time_us: Cell::new(0),
            generation: GENERATION.load(Ordering::SeqCst),
        }
    }
}

thread_local! {
    static HOST: RefCell<Option<Host>> = const { RefCell::new(None) };
}

/// Build or refresh this thread's host, then run `f` against it.
///
/// A reload that fails keeps the last working registry: a typo saved into
/// functions.lua must not turn every custom function in an open sheet into
/// "Unknown function" on the next edit. The error is kept for the status
/// line, and the failed generation is adopted so the load is not retried on
/// every call.
fn with_host<R>(f: impl FnOnce(&Host) -> R) -> R {
    HOST.with(|slot| {
        let current = GENERATION.load(Ordering::SeqCst);
        let needs_load = match slot.borrow().as_ref() {
            None => true,
            Some(host) => host.generation != current,
        };
        if needs_load {
            let fresh = Host::load();
            let mut slot_mut = slot.borrow_mut();
            match (slot_mut.as_mut(), fresh.load_error.clone()) {
                (Some(old), Some(error)) => {
                    old.load_error = Some(error);
                    old.generation = current;
                }
                _ => *slot_mut = Some(fresh),
            }
        }
        let host = slot.borrow();
        f(host.as_ref().expect("host loaded above"))
    })
}

/// Register the adapter as the engine's default handler. Idempotent.
pub fn install() {
    visigrid_engine::custom_fns::set_default_custom_fn_handler(Some(dispatch));
}

/// The handler the engine calls for any function name it does not know.
pub fn dispatch(name: &str, args: &[EvalArg]) -> Option<EvalResult> {
    with_host(|host| {
        let start = Instant::now();
        // Outside a full recalc there is no memo: a fresh, discarded one.
        let scratch = RefCell::new(MemoCache::new());
        let memo = if host.in_recalc.get() { &host.memo } else { &scratch };
        let result = if name == LUA_CELL_FUNCTION {
            Some(call_lua_cell(host.runtime.lua(), args, &host.chunks, memo))
        } else if host.registry.functions.contains_key(name) {
            Some(call_custom_function(host.runtime.lua(), name, args, memo))
        } else {
            None
        };
        if result.is_some() {
            host.lua_time_us
                .set(host.lua_time_us.get() + start.elapsed().as_micros() as u64);
        }
        result
    })
}

/// What `functions.lua` gave this thread's host.
pub struct LoadStatus {
    pub function_count: usize,
    pub warnings: Vec<String>,
    pub error: Option<String>,
}

pub fn load_status() -> LoadStatus {
    with_host(|host| LoadStatus {
        function_count: host.registry.functions.len(),
        warnings: host.registry.warnings.clone(),
        error: host.load_error.clone(),
    })
}

/// Names the editor should treat as callable: the registered custom
/// functions plus `LUA` itself, which is not a registry entry but answers.
pub fn function_names() -> Vec<String> {
    with_host(|host| {
        let mut names: Vec<String> = vec![LUA_CELL_FUNCTION.to_string()];
        names.extend(host.registry.functions.keys().cloned());
        names
    })
}

/// Re-read `functions.lua`. This thread reloads now; every other thread
/// reloads on its next use. A failed load keeps the previous registry and
/// reports the error.
pub fn reload() -> LoadStatus {
    GENERATION.fetch_add(1, Ordering::SeqCst);
    load_status()
}

/// Full ordered recalc through the adapter, with the memo alive for exactly
/// its duration. The workbook consults the registered handler itself; this
/// exists so callers have one name for "the Lua-aware recalc", so the memo
/// has a boundary, and so the Lua time for the report is accounted.
pub fn recompute(wb: &mut Workbook) -> RecalcReport {
    let _ = take_lua_time_us();
    with_host(|host| {
        host.memo.borrow_mut().clear();
        host.in_recalc.set(true);
    });
    let mut report = wb.recompute_full_ordered();
    with_host(|host| {
        host.in_recalc.set(false);
        host.memo.borrow_mut().clear();
    });
    report.phase_lua_total_us = take_lua_time_us();
    report
}

#[cfg(test)]
fn memo_len() -> usize {
    with_host(|host| host.memo.borrow().len())
}

#[cfg(test)]
fn host_generation() -> u64 {
    with_host(|host| host.generation)
}

/// Microseconds spent inside Lua since the last call, on this thread.
pub fn take_lua_time_us() -> u64 {
    with_host(|host| host.lua_time_us.replace(0))
}

#[cfg(test)]
mod tests {
    use super::*;
    use visigrid_engine::formula::eval::Value;

    #[test]
    fn a_lua_cell_updates_when_its_input_is_edited() {
        install();
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "21");
        wb.set_cell_value_tracked(0, 0, 1, "=LUA(\"return args[1] * 2\", A1)");
        assert_eq!(wb.active_sheet().get_display(0, 1), "42", "computed on entry");

        // The incremental path after an edit, outside any batch.
        wb.set_cell_value_tracked(0, 0, 0, "5");
        assert_eq!(wb.active_sheet().get_display(0, 1), "10", "dependent followed the edit");

        // The batch path the session host uses for agent ops.
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 0, "8");
        }
        assert_eq!(wb.active_sheet().get_display(0, 1), "16", "batch close recalculated");

        // A dependent that spills follows too, and re-places on the edit.
        wb.set_cell_value_tracked(0, 0, 2, "=LUA(\"return {args[1], args[1] + 1}\", B1)");
        assert_eq!(wb.active_sheet().get_display(1, 2), "17");
        wb.set_cell_value_tracked(0, 0, 0, "1");
        assert_eq!(wb.active_sheet().get_display(0, 1), "2");
        assert_eq!(wb.active_sheet().get_display(1, 2), "3", "spilled receiver followed the edit");

        // And the full recalc every CLI command runs after a load.
        let report = recompute(&mut wb);
        assert_eq!(wb.active_sheet().get_display(0, 1), "2");
        assert!(report.errors.is_empty(), "{:?}", report.errors);
    }

    #[test]
    fn a_spilling_lua_cell_settles_through_the_incremental_path() {
        install();
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "3");
        wb.set_cell_value_tracked(0, 0, 1, "=LUA(\"local t = {} for i = 1, args[1] do t[i] = i end return t\", A1)");
        wb.recompute_full_ordered();
        assert_eq!(wb.active_sheet().get_display(2, 1), "3", "spilled three rows");
    }

    #[test]
    fn the_memo_lives_only_inside_a_full_recalc() {
        install();
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "=LUA(\"return 1\")");
        assert_eq!(memo_len(), 0, "incremental evaluation leaves nothing behind");
        recompute(&mut wb);
        assert_eq!(memo_len(), 0, "cleared when the recalc ends");
    }

    #[test]
    fn lua_is_a_known_name_for_the_editor() {
        install();
        assert!(function_names().iter().any(|n| n == "LUA"));
    }

    #[test]
    fn reload_reaches_other_threads_on_their_next_use() {
        install();
        let before = host_generation();
        let worker = std::thread::spawn(|| {
            let _ = dispatch("LUA", &[EvalArg::Scalar(Value::Text("return 1".into()))]);
            let seen_first = host_generation();
            // Park until the main thread has reloaded, then use the host again.
            let (tx, rx) = std::sync::mpsc::channel::<()>();
            (seen_first, tx, rx)
        });
        let (seen_first, _tx, _rx) = worker.join().unwrap();
        assert_eq!(seen_first, before);
        let status = reload();
        assert!(status.error.is_none(), "{:?}", status.error);
        assert_eq!(host_generation(), before + 1, "this thread reloaded");
        let other = std::thread::spawn(|| {
            let _ = dispatch("LUA", &[EvalArg::Scalar(Value::Text("return 1".into()))]);
            host_generation()
        })
        .join()
        .unwrap();
        assert_eq!(other, before + 1, "a fresh thread builds at the new generation");
    }

    #[test]
    fn unknown_names_are_not_claimed() {
        install();
        assert!(dispatch("NOT_A_FUNCTION_ANYWHERE", &[]).is_none());
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "=NOT_A_FUNCTION_ANYWHERE()");
        assert!(wb.active_sheet().get_display(0, 0).contains("Unknown function"));
    }
}
