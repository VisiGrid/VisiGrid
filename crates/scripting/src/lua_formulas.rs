//! The Lua-aware formula adapter: what the desktop app and the CLI install so
//! that custom functions from `functions.lua` and `=LUA` cells compute
//! everywhere a workbook is recalculated, including after every edit and
//! every agent op, in `vgrid serve`, and in headless CLI commands.
//!
//! Installed once per process with [`install`], which registers the engine's
//! custom-function hooks: [`dispatch`] answers names, and the recalc hooks
//! bracket every recalculation the engine runs so the memo lives for exactly
//! one of them, full or incremental, however it was reached.
//!
//! The source of `functions.lua` is published process-wide, one version at a
//! time, and every thread compiles that exact version into its own Lua state
//! (a Lua state is neither `Send` nor shareable, and the desktop loads files
//! on background threads). A reload that fails to compile keeps the previous
//! version published and reports the error, so a typo saved into the file
//! cannot turn every custom function into "Unknown function" anywhere.
//!
//! The session host never sees any of this. It applies ops to a workbook and
//! the workbook does the rest, which is the point: it stays Lua-agnostic.

use std::cell::{Cell, RefCell};
use std::path::PathBuf;
use std::sync::{Arc, RwLock};
use std::time::Instant;

use visigrid_engine::custom_fns::CustomFnHooks;
use visigrid_engine::formula::eval::{EvalArg, EvalResult};
use visigrid_engine::recalc::RecalcReport;
use visigrid_engine::workbook::Workbook;

use crate::custom_functions::{
    call_custom_function, custom_functions_path, load_custom_functions_from_source,
    CustomFunctionRegistry, MemoCache, RESERVED_LUA_CELL_NAME,
};
use crate::lua_cell::{call_lua_cell, ChunkCache};
use crate::runtime::{Limits, LuaRuntime};

/// The name of the cell-chunk function.
pub const LUA_CELL_FUNCTION: &str = RESERVED_LUA_CELL_NAME;

/// The published `functions.lua`: what every thread compiles.
struct Published {
    /// Bumped only when a version that compiles is published.
    generation: u64,
    /// The source text and where it came from; `None` when there is no file.
    source: Option<(Arc<str>, PathBuf)>,
    /// The last load that failed, kept for the status line. The published
    /// source is the last one that compiled.
    error: Option<String>,
}

static PUBLISHED: RwLock<Option<Published>> = RwLock::new(None);

/// Read `functions.lua` from disk and prove it compiles, on this thread, in
/// a throwaway runtime. Only then does it become the published version.
fn publish_from_disk() -> Result<(), String> {
    let path = custom_functions_path()?;
    let candidate = if path.exists() {
        Some(std::fs::read_to_string(&path).map_err(|e| format!("Failed to read {}: {}", path.display(), e))?)
    } else {
        None
    };
    publish(candidate, path)
}

fn publish(candidate: Option<String>, path: PathBuf) -> Result<(), String> {
    // Checked in a closure so a failure is a value to record, not an early
    // return that would skip recording it.
    let checked: Result<Option<(Arc<str>, PathBuf)>, String> = (|| match &candidate {
        None => Ok(None),
        Some(text) => {
            let probe = LuaRuntime::with_limits(Limits::batch()).map_err(|e| e.to_string())?;
            load_custom_functions_from_source(probe.lua(), text, &path)?;
            Ok(Some((Arc::from(text.as_str()), path.clone())))
        }
    })();
    let mut slot = PUBLISHED.write().unwrap_or_else(|e| e.into_inner());
    match checked {
        Ok(source) => {
            let generation = slot.as_ref().map(|p| p.generation + 1).unwrap_or(1);
            *slot = Some(Published { generation, source, error: None });
            Ok(())
        }
        Err(e) => {
            match slot.as_mut() {
                Some(p) => p.error = Some(e.clone()),
                None => *slot = Some(Published { generation: 1, source: None, error: Some(e.clone()) }),
            }
            Err(e)
        }
    }
}

/// The published version, publishing from disk first if nothing is yet.
/// The first publish happens under the write lock, so two threads starting
/// at once cannot both read the disk and the slower one overwrite whatever
/// was published in between.
fn published() -> (u64, Option<(Arc<str>, PathBuf)>, Option<String>) {
    {
        let guard = PUBLISHED.read().unwrap_or_else(|e| e.into_inner());
        if let Some(p) = guard.as_ref() {
            return (p.generation, p.source.clone(), p.error.clone());
        }
    }
    let mut slot = PUBLISHED.write().unwrap_or_else(|e| e.into_inner());
    if slot.is_none() {
        *slot = Some(initial_from_disk());
    }
    let p = slot.as_ref().expect("published above");
    (p.generation, p.source.clone(), p.error.clone())
}

/// Read and compile-check functions.lua for the first publish. A file that
/// fails to compile publishes as "no functions" with the error kept.
fn initial_from_disk() -> Published {
    let read = || -> Result<Option<(Arc<str>, PathBuf)>, String> {
        let path = custom_functions_path()?;
        if !path.exists() {
            return Ok(None);
        }
        let text = std::fs::read_to_string(&path).map_err(|e| format!("Failed to read {}: {}", path.display(), e))?;
        let probe = LuaRuntime::with_limits(Limits::batch()).map_err(|e| e.to_string())?;
        load_custom_functions_from_source(probe.lua(), &text, &path)?;
        Ok(Some((Arc::from(text.as_str()), path)))
    };
    match read() {
        Ok(source) => Published { generation: 1, source, error: None },
        Err(e) => Published { generation: 1, source: None, error: Some(e) },
    }
}

struct Host {
    runtime: LuaRuntime,
    registry: CustomFunctionRegistry,
    chunks: RefCell<ChunkCache>,
    /// Filled between the engine's begin and end of one recalc; empty
    /// otherwise. That is the contract the cache was written to: a function
    /// may capture mutable state and an error may be a timeout, so nothing
    /// answers from a previous recalc.
    memo: RefCell<MemoCache>,
    recalc_depth: Cell<u32>,
    lua_time_us: Cell<u64>,
    generation: u64,
}

impl Host {
    /// Compile the published version. It compiled once already when it was
    /// published, so this does not fail in practice; if it does, the registry
    /// is empty and the reason is reported through `load_status`.
    fn build(generation: u64, source: Option<(Arc<str>, PathBuf)>) -> Self {
        let runtime = LuaRuntime::with_limits(Limits::batch()).unwrap_or_else(|e| {
            eprintln!("custom functions: Lua runtime unavailable: {}", e);
            LuaRuntime::default()
        });
        let registry = match source {
            None => CustomFunctionRegistry::empty(),
            Some((text, path)) => match load_custom_functions_from_source(runtime.lua(), &text, &path) {
                Ok(r) => r,
                Err(e) => {
                    eprintln!("custom functions: {}", e);
                    CustomFunctionRegistry::empty()
                }
            },
        };
        Host {
            runtime,
            registry,
            chunks: RefCell::new(ChunkCache::new()),
            memo: RefCell::new(MemoCache::new()),
            recalc_depth: Cell::new(0),
            lua_time_us: Cell::new(0),
            generation,
        }
    }
}

thread_local! {
    static HOST: RefCell<Option<Host>> = const { RefCell::new(None) };
}

/// This thread's host at the published generation, building or rebuilding
/// it first when it is missing or behind.
fn with_host<R>(f: impl FnOnce(&Host) -> R) -> R {
    let (generation, source, _) = published();
    HOST.with(|slot| {
        let behind = match slot.borrow().as_ref() {
            None => true,
            Some(host) => host.generation != generation,
        };
        if behind {
            *slot.borrow_mut() = Some(Host::build(generation, source));
        }
        let host = slot.borrow();
        f(host.as_ref().expect("host built above"))
    })
}

/// Register the adapter with the engine. Idempotent.
pub fn install() {
    visigrid_engine::custom_fns::set_default_custom_fns(Some(CustomFnHooks {
        call: dispatch,
        begin_recalc,
        end_recalc,
    }));
}

fn begin_recalc() {
    with_host(|host| {
        if host.recalc_depth.get() == 0 {
            host.memo.borrow_mut().clear();
        }
        host.recalc_depth.set(host.recalc_depth.get() + 1);
    });
}

fn end_recalc() {
    with_host(|host| {
        let depth = host.recalc_depth.get().saturating_sub(1);
        host.recalc_depth.set(depth);
        if depth == 0 {
            host.memo.borrow_mut().clear();
        }
    });
}

/// The handler the engine calls for any function name it does not know.
pub fn dispatch(name: &str, args: &[EvalArg]) -> Option<EvalResult> {
    with_host(|host| {
        let start = Instant::now();
        // Outside a recalc (a direct call from a test or a tool) there is no
        // memo: a fresh, discarded one.
        let scratch = RefCell::new(MemoCache::new());
        let memo = if host.recalc_depth.get() > 0 { &host.memo } else { &scratch };
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

/// What `functions.lua` gave the process.
pub struct LoadStatus {
    pub function_count: usize,
    pub warnings: Vec<String>,
    /// The last load that failed. The published version is still the last
    /// one that compiled, and `function_count` describes that one.
    pub error: Option<String>,
}

pub fn load_status() -> LoadStatus {
    let (_, _, error) = published();
    with_host(|host| LoadStatus {
        function_count: host.registry.functions.len(),
        warnings: host.registry.warnings.clone(),
        error,
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

/// Re-read `functions.lua` and publish it if it compiles. This thread
/// rebuilds now; every other thread rebuilds on its next use. A failed load
/// leaves the previous version published and reports the error.
pub fn reload() -> LoadStatus {
    let _ = publish_from_disk();
    load_status()
}

/// Full ordered recalc through the adapter. The workbook consults the
/// registered hooks itself; this exists so callers have one name for "the
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
fn memo_len() -> usize {
    with_host(|host| host.memo.borrow().len())
}

#[cfg(test)]
fn host_generation() -> u64 {
    with_host(|host| host.generation)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Mutex;
    use visigrid_engine::formula::eval::Value;

    /// The publish tests rewrite the process-wide source; they take turns.
    static PUBLISH_LOCK: Mutex<()> = Mutex::new(());

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

    /// The memo is bracketed by the engine's own recalc hooks, so it covers
    /// the incremental path and the full path alike and outlives neither.
    #[test]
    fn the_memo_lives_only_inside_a_recalc_whichever_path_ran_it() {
        install();
        let mut wb = Workbook::new();
        wb.set_cell_value_tracked(0, 0, 0, "=LUA(\"return 1\")");
        assert_eq!(memo_len(), 0, "incremental recalc cleared on end");
        wb.recompute_full_ordered();
        assert_eq!(memo_len(), 0, "full recalc cleared on end");
        {
            let mut guard = wb.batch_guard();
            guard.set_cell_value_tracked(0, 0, 1, "=LUA(\"return 2\")");
        }
        assert_eq!(memo_len(), 0, "batch close cleared on end");
        assert_eq!(with_host(|h| h.recalc_depth.get()), 0, "depth returns to zero");
        // A direct call outside any recalc leaves nothing behind either.
        let _ = dispatch("LUA", &[EvalArg::Scalar(Value::Text("return 3".into()))]);
        assert_eq!(memo_len(), 0);
    }

    #[test]
    fn lua_is_a_known_name_for_the_editor() {
        install();
        assert!(function_names().iter().any(|n| n == "LUA"));
    }

    /// A worker thread that already built its host sees the new generation
    /// on its next use after a reload on another thread.
    #[test]
    fn an_already_initialised_thread_observes_a_reload() {
        let _serial = PUBLISH_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        install();
        let path = std::path::PathBuf::from("functions.lua");
        publish(Some("function GEN_A() return 1 end".into()), path.clone()).unwrap();
        let (to_worker, from_main) = std::sync::mpsc::channel::<()>();
        let (to_main, from_worker) = std::sync::mpsc::channel::<(u64, Vec<String>)>();
        let worker = std::thread::spawn(move || {
            // Build the host now, at the current generation.
            let first = host_generation();
            let names_first = function_names();
            to_main.send((first, names_first)).unwrap();
            // Wait for the main thread to publish a new version, then use
            // the host again without doing anything else.
            from_main.recv().unwrap();
            let second = host_generation();
            let names_second = function_names();
            to_main.send((second, names_second)).unwrap();
        });
        let (first, names_first) = from_worker.recv().unwrap();
        assert!(names_first.iter().any(|n| n == "GEN_A"));
        publish(Some("function GEN_B() return 2 end".into()), path).unwrap();
        to_worker.send(()).unwrap();
        let (second, names_second) = from_worker.recv().unwrap();
        worker.join().unwrap();
        assert_eq!(second, first + 1, "the worker rebuilt at the new generation");
        assert!(names_second.iter().any(|n| n == "GEN_B") && !names_second.iter().any(|n| n == "GEN_A"));
    }

    /// A version that fails to compile is never published: the thread that
    /// tried keeps the last good registry, and so does a thread that starts
    /// afterwards, because both compile the same published text.
    #[test]
    fn a_failed_reload_keeps_the_last_good_version_on_every_thread() {
        let _serial = PUBLISH_LOCK.lock().unwrap_or_else(|e| e.into_inner());
        install();
        let path = std::path::PathBuf::from("functions.lua");
        publish(Some("function GOOD() return 1 end".into()), path.clone()).unwrap();
        let good_gen = host_generation();
        assert!(function_names().iter().any(|n| n == "GOOD"));

        let err = publish(Some("function BAD( return".into()), path).unwrap_err();
        assert!(err.contains("syntax") || err.contains("expected"), "{}", err);
        assert_eq!(host_generation(), good_gen, "nothing new was published");
        assert!(function_names().iter().any(|n| n == "GOOD"), "this thread kept the last good version");
        let status = load_status();
        assert!(status.error.is_some(), "the failure is reported: count={} warnings={:?}", status.function_count, status.warnings);

        let fresh = std::thread::spawn(|| (host_generation(), function_names())).join().unwrap();
        assert_eq!(fresh.0, good_gen);
        assert!(fresh.1.iter().any(|n| n == "GOOD"), "a new thread compiled the last good version too");
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
