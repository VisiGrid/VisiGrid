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
use std::time::Instant;

use visigrid_engine::formula::eval::{EvalArg, EvalResult};
use visigrid_engine::recalc::RecalcReport;
use visigrid_engine::workbook::Workbook;

use crate::custom_functions::{
    call_custom_function, load_custom_functions, CustomFunctionRegistry, MemoCache,
};
use crate::lua_cell::{call_lua_cell, ChunkCache};
use crate::runtime::{Limits, LuaRuntime};

/// Memo entries kept before the memo is emptied. Results are pure in their
/// inputs, so the memo never goes stale; it only grows, and a long editing
/// session or a large workbook of distinct inputs must not grow it forever.
pub const MAX_MEMO_ENTRIES: usize = 50_000;

/// The name of the cell-chunk function.
pub const LUA_CELL_FUNCTION: &str = "LUA";

struct Host {
    runtime: LuaRuntime,
    registry: CustomFunctionRegistry,
    load_error: Option<String>,
    chunks: RefCell<ChunkCache>,
    memo: RefCell<MemoCache>,
    lua_time_us: Cell<u64>,
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
            lua_time_us: Cell::new(0),
        }
    }
}

thread_local! {
    static HOST: RefCell<Option<Host>> = const { RefCell::new(None) };
}

fn with_host<R>(f: impl FnOnce(&Host) -> R) -> R {
    HOST.with(|slot| {
        if slot.borrow().is_none() {
            *slot.borrow_mut() = Some(Host::load());
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
        let result = if name == LUA_CELL_FUNCTION {
            Some(call_lua_cell(host.runtime.lua(), args, &host.chunks, &host.memo))
        } else if host.registry.functions.contains_key(name) {
            Some(call_custom_function(host.runtime.lua(), name, args, &host.memo))
        } else {
            None
        };
        if result.is_some() {
            host.lua_time_us
                .set(host.lua_time_us.get() + start.elapsed().as_micros() as u64);
            if host.memo.borrow().len() > MAX_MEMO_ENTRIES {
                host.memo.borrow_mut().clear();
            }
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

/// Registered custom-function names, for autocomplete and diagnostics.
pub fn function_names() -> Vec<String> {
    with_host(|host| host.registry.functions.keys().cloned().collect())
}

/// Re-read `functions.lua` on this thread, dropping the memo and compiled
/// chunks with it. Other threads keep what they loaded until they reload.
pub fn reload() -> LoadStatus {
    HOST.with(|slot| {
        *slot.borrow_mut() = Some(Host::load());
    });
    load_status()
}

/// Full ordered recalc through the adapter. The workbook consults the
/// registered handler itself; this exists so callers have one name for "the
/// Lua-aware recalc" and so the Lua time for the report is accounted.
pub fn recompute(wb: &mut Workbook) -> RecalcReport {
    let _ = take_lua_time_us();
    let mut report = wb.recompute_full_ordered();
    report.phase_lua_total_us = take_lua_time_us();
    report
}

/// Microseconds spent inside Lua since the last call, on this thread.
pub fn take_lua_time_us() -> u64 {
    with_host(|host| host.lua_time_us.replace(0))
}

#[cfg(test)]
mod tests {
    use super::*;

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
    fn unknown_names_are_not_claimed() {
        install();
        assert!(dispatch("NOT_A_FUNCTION_ANYWHERE", &[]).is_none());
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "=NOT_A_FUNCTION_ANYWHERE()");
        assert!(wb.active_sheet().get_display(0, 0).contains("Unknown function"));
    }
}
