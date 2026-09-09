//! `=LUA(code, args…)`: a cell whose formula is a Lua chunk.
//!
//! The code is the first argument, a formula string; the rest are ordinary
//! formula arguments, which is how the engine knows what the cell depends on.
//! The chunk sees them as `args[1]`, `args[2]`, … with `args.n` for the count
//! (so an empty argument does not shorten the table), returns a scalar for the
//! cell or a table that spills, and runs in the same read-only sandbox as a
//! custom function: no sheet access, no globals, no randomness, an
//! instruction limit. That is what keeps a cell pure, and purity is what lets
//! the engine order and memoise it like any other formula.
//!
//! Compiled chunks are cached by code hash. A workbook with a thousand copies
//! of one formula compiles it once; the environment is set per call.

use std::cell::RefCell;
use std::collections::HashMap;

use mlua::Lua;
use visigrid_engine::formula::eval::{EvalArg, EvalResult, Value};

use crate::custom_functions::{
    build_memo_key, eval_arg_to_lua, freeze_env, lua_return_to_eval_result, sandbox_globals,
    scrub_lua_runtime_error, FormulaLimitGuard, MemoCache,
};

/// Largest chunk a cell may hold. A cell is a formula, not a module; anything
/// past this belongs in an attached script.
pub const MAX_CELL_CODE_BYTES: usize = 64 * 1024;

/// Compiled chunks by code hash, so repeated formulas compile once.
#[derive(Default)]
pub struct ChunkCache {
    chunks: HashMap<u64, mlua::Function>,
}

impl ChunkCache {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn len(&self) -> usize {
        self.chunks.len()
    }

    pub fn is_empty(&self) -> bool {
        self.chunks.is_empty()
    }

    fn get_or_compile(&mut self, lua: &Lua, code: &str) -> Result<mlua::Function, String> {
        let hash = hash_code(code);
        if let Some(f) = self.chunks.get(&hash) {
            return Ok(f.clone());
        }
        let f = lua
            .load(code)
            .set_name("=LUA")
            .into_function()
            .map_err(|e| format!("#LUA! {}", scrub_lua_runtime_error(&e, "LUA")))?;
        self.chunks.insert(hash, f.clone());
        Ok(f)
    }
}

fn hash_code(code: &str) -> u64 {
    // FNV-1a: cheap, deterministic, and this is a cache key, not a fingerprint.
    let mut h: u64 = 0xcbf29ce484222325;
    for b in code.bytes() {
        h ^= b as u64;
        h = h.wrapping_mul(0x100000001b3);
    }
    h
}

/// Evaluate `=LUA(...)` from its formula arguments: `args[0]` is the code,
/// the rest are the chunk's inputs.
pub fn call_lua_cell(
    lua: &Lua,
    args: &[EvalArg],
    chunks: &RefCell<ChunkCache>,
    memo: &RefCell<MemoCache>,
) -> EvalResult {
    let Some(first) = args.first() else {
        return EvalResult::Error("#LUA! LUA needs the code as its first argument".to_string());
    };
    let code = match first {
        EvalArg::Scalar(Value::Text(code)) => code.as_str(),
        EvalArg::Scalar(Value::Error(e)) => return EvalResult::Error(e.clone()),
        _ => return EvalResult::Error("#LUA! the first argument must be a string of Lua code".to_string()),
    };
    eval_cell_chunk(lua, code, &args[1..], chunks, memo)
}

/// Run a chunk over its inputs. Errors in the inputs propagate untouched, as
/// they do for custom functions, so a `#REF!` upstream stays a `#REF!` here.
pub fn eval_cell_chunk(
    lua: &Lua,
    code: &str,
    inputs: &[EvalArg],
    chunks: &RefCell<ChunkCache>,
    memo: &RefCell<MemoCache>,
) -> EvalResult {
    if code.len() > MAX_CELL_CODE_BYTES {
        return EvalResult::Error(format!(
            "#LUA! cell code larger than {} KiB; put it in an attached script",
            MAX_CELL_CODE_BYTES / 1024
        ));
    }
    for arg in inputs {
        match arg {
            EvalArg::Scalar(Value::Error(e)) => return EvalResult::Error(e.clone()),
            EvalArg::Range { values, .. } => {
                if let Some(Value::Error(e)) = values.iter().find(|v| matches!(v, Value::Error(_))) {
                    return EvalResult::Error(e.clone());
                }
            }
            _ => {}
        }
    }

    // Memoised on code plus inputs, within one recalc.
    let mut key_args: Vec<EvalArg> = Vec::with_capacity(inputs.len() + 1);
    key_args.push(EvalArg::Scalar(Value::Text(code.to_string())));
    key_args.extend(inputs.iter().cloned());
    let memo_key = build_memo_key("LUA", &key_args);
    if let Some(cached) = memo.borrow().lookup(&memo_key) {
        return cached.clone();
    }

    let result = {
        let _guard = FormulaLimitGuard::new(lua);
        run_chunk(lua, code, inputs, chunks)
    };
    memo.borrow_mut().store(memo_key, result.clone());
    result
}

fn run_chunk(
    lua: &Lua,
    code: &str,
    inputs: &[EvalArg],
    chunks: &RefCell<ChunkCache>,
) -> EvalResult {
    let chunk = match chunks.borrow_mut().get_or_compile(lua, code) {
        Ok(f) => f,
        Err(e) => return EvalResult::Error(e),
    };
    let built = (|| -> mlua::Result<mlua::Function> {
        let args_table = lua.create_table()?;
        for (i, arg) in inputs.iter().enumerate() {
            args_table.raw_set(i + 1, eval_arg_to_lua(lua, arg)?)?;
        }
        args_table.raw_set("n", inputs.len())?;
        let env = sandbox_globals(lua)?;
        env.set("args", args_table)?;
        let wrapper = freeze_env(lua, env)?;
        chunk.set_environment(wrapper)?;
        Ok(chunk)
    })();
    let chunk = match built {
        Ok(f) => f,
        Err(e) => return EvalResult::Error(format!("#LUA! {}", e)),
    };
    match chunk.call::<mlua::Value>(()) {
        Ok(ret) => lua_return_to_eval_result(&ret),
        Err(e) => EvalResult::Error(format!("#LUA! {}", scrub_lua_runtime_error(&e, "LUA"))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn scalar(n: f64) -> EvalArg {
        EvalArg::Scalar(Value::Number(n))
    }

    fn run(code: &str, inputs: &[EvalArg]) -> EvalResult {
        let lua = Lua::new();
        let chunks = RefCell::new(ChunkCache::new());
        let memo = RefCell::new(MemoCache::new());
        eval_cell_chunk(&lua, code, inputs, &chunks, &memo)
    }

    #[test]
    fn doubles_its_input() {
        assert_eq!(run("return args[1] * 2", &[scalar(21.0)]), EvalResult::Number(42.0));
    }

    #[test]
    fn args_n_counts_empty_arguments() {
        let inputs = [scalar(1.0), EvalArg::Scalar(Value::Empty), scalar(3.0)];
        assert_eq!(run("return args.n", &inputs), EvalResult::Number(3.0));
        assert_eq!(run("return args[2] == nil", &inputs), EvalResult::Boolean(true));
    }

    #[test]
    fn ranges_arrive_as_range_objects_and_tables_spill() {
        let range = EvalArg::Range {
            values: vec![Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)],
            num_cells: 3,
        };
        let r = run(
            "local out = {} for i = 1, args[1].n do out[i] = args[1]:get(i) * 10 end return out",
            &[range],
        );
        match r {
            EvalResult::Array(a) => {
                assert_eq!((a.rows(), a.cols()), (3, 1));
                assert_eq!(a.get(2, 0), Some(&Value::Number(30.0)));
            }
            other => panic!("expected array, got {:?}", other),
        }
    }

    #[test]
    fn errors_are_reported_not_raised() {
        match run("return (", &[]) {
            EvalResult::Error(e) => assert!(e.starts_with("#LUA!"), "{}", e),
            other => panic!("{:?}", other),
        }
        match run("error('nope')", &[]) {
            EvalResult::Error(e) => assert!(e.contains("nope"), "{}", e),
            other => panic!("{:?}", other),
        }
        match run("x = 1 return x", &[]) {
            EvalResult::Error(e) => assert!(e.contains("Global state"), "{}", e),
            other => panic!("{:?}", other),
        }
        assert_eq!(
            run("return args[1]", &[EvalArg::Scalar(Value::Error("#REF!".into()))]),
            EvalResult::Error("#REF!".into())
        );
    }

    #[test]
    fn oversized_code_is_refused() {
        let code = format!("return 1 --{}", "x".repeat(MAX_CELL_CODE_BYTES));
        match run(&code, &[]) {
            EvalResult::Error(e) => assert!(e.contains("larger than"), "{}", e),
            other => panic!("{:?}", other),
        }
    }

    #[test]
    fn chunks_compile_once_per_code() {
        let lua = Lua::new();
        let chunks = RefCell::new(ChunkCache::new());
        let memo = RefCell::new(MemoCache::new());
        for n in 0..5 {
            let r = eval_cell_chunk(&lua, "return args[1] + 1", &[scalar(n as f64)], &chunks, &memo);
            assert_eq!(r, EvalResult::Number(n as f64 + 1.0));
        }
        assert_eq!(chunks.borrow().len(), 1);
    }

    #[test]
    fn call_lua_cell_takes_the_code_from_the_first_argument() {
        let lua = Lua::new();
        let chunks = RefCell::new(ChunkCache::new());
        let memo = RefCell::new(MemoCache::new());
        let args = [EvalArg::Scalar(Value::Text("return args[1] * 2".into())), scalar(21.0)];
        assert_eq!(call_lua_cell(&lua, &args, &chunks, &memo), EvalResult::Number(42.0));
        match call_lua_cell(&lua, &[scalar(1.0)], &chunks, &memo) {
            EvalResult::Error(e) => assert!(e.contains("string of Lua code"), "{}", e),
            other => panic!("{:?}", other),
        }
    }

    /// The headless end-to-end: a workbook cell =LUA("return args[1] * 2", A1)
    /// evaluated through the engine's custom-function handler boundary.
    #[test]
    fn a_lua_cell_evaluates_inside_a_workbook_recalc() {
        use visigrid_engine::workbook::Workbook;
        let lua = Lua::new();
        let chunks = RefCell::new(ChunkCache::new());
        let memo = RefCell::new(MemoCache::new());
        let handler = |name: &str, args: &[EvalArg]| -> Option<EvalResult> {
            (name == "LUA").then(|| call_lua_cell(&lua, args, &chunks, &memo))
        };
        let mut wb = Workbook::new();
        wb.active_sheet_mut().set_value(0, 0, "21");
        wb.active_sheet_mut().set_value(0, 1, "=LUA(\"return args[1] * 2\", A1)");
        wb.active_sheet_mut().set_value(0, 2, "=LUA(\"return {args[1], args[1] + 1}\", B1)");
        wb.rebuild_dep_graph();
        let report = wb.recompute_full_ordered_with_custom_fns(&handler);
        assert!(report.errors.is_empty(), "{:?}", report.errors);
        let sheet = wb.active_sheet();
        assert_eq!(sheet.get_display(0, 1), "42");
        assert_eq!(sheet.get_display(0, 2), "42");
        assert_eq!(sheet.get_display(1, 2), "43", "table return spilled");
    }
}
