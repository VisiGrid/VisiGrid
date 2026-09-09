//! Custom (user-defined) Lua formula functions for VisiGrid.
//!
//! Users write `functions.lua` in the config directory, VisiGrid loads and
//! sandboxes them, and they become callable in formulas (`=DOUBLE(A1)`).
//!
//! # Architecture
//!
//! The formula engine lives in `crates/engine/` (no Lua dependency). The bridge
//! is the `CellLookup::try_custom_function` method: the engine evaluates
//! arguments into typed `EvalArg` values, then calls back here to execute the
//! Lua function.

use std::cell::RefCell;
use std::collections::HashMap;
use std::path::PathBuf;
use std::time::Instant;

use mlua::{self, Lua, HookTriggers, VmState};

use visigrid_engine::formula::eval::{Array2D, EvalArg, EvalResult, Value};
use visigrid_engine::formula::functions::{is_known_function, is_valid_custom_function_name};

// =============================================================================
// Registry
// =============================================================================

/// Registry of user-defined custom functions loaded from `functions.lua`.
pub struct CustomFunctionRegistry {
    pub functions: HashMap<String, CustomFunction>,
    pub source_path: PathBuf,
    pub last_loaded: Option<Instant>,
    pub warnings: Vec<String>,
}

/// A single registered custom function.
pub struct CustomFunction {
    pub name: String,
}

impl CustomFunctionRegistry {
    pub fn empty() -> Self {
        Self {
            functions: HashMap::new(),
            source_path: PathBuf::new(),
            last_loaded: None,
            warnings: Vec::new(),
        }
    }
}

// =============================================================================
// Loading
// =============================================================================

/// Load custom functions from `~/.config/visigrid/functions.lua`.
///
/// Returns an empty registry (not an error) if the file doesn't exist.
/// Returns Err only for parse/compile failures.
pub fn load_custom_functions(lua: &Lua) -> Result<CustomFunctionRegistry, String> {
    let config_dir = dirs::config_dir()
        .ok_or_else(|| "Could not determine config directory".to_string())?;
    let source_path = config_dir.join("visigrid").join("functions.lua");

    if !source_path.exists() {
        return Ok(CustomFunctionRegistry {
            functions: HashMap::new(),
            source_path,
            last_loaded: Some(Instant::now()),
            warnings: Vec::new(),
        });
    }

    let source = std::fs::read_to_string(&source_path)
        .map_err(|e| format!("Failed to read {}: {}", source_path.display(), e))?;

    // Execute the file in the Lua runtime to populate globals
    lua.load(&source)
        .set_name(source_path.to_string_lossy().into_owned())
        .exec()
        .map_err(|e| scrub_lua_load_error(&e, &source_path))?;

    // Scan globals for uppercase function names
    let mut functions = HashMap::new();
    let mut warnings = Vec::new();
    let globals = lua.globals();

    // Iterate all globals
    for pair in globals.pairs::<String, mlua::Value>() {
        let (name, value) = match pair {
            Ok(p) => p,
            Err(_) => continue,
        };

        // Only consider functions with valid custom function names
        if !is_valid_custom_function_name(&name) {
            continue;
        }

        // Must be a Lua function
        if !matches!(value, mlua::Value::Function(_)) {
            continue;
        }

        // Check for built-in collision
        if is_known_function(&name) {
            warnings.push(format!("{} shadows built-in", name));
            continue;
        }

        functions.insert(name.clone(), CustomFunction { name });
    }

    Ok(CustomFunctionRegistry {
        functions,
        source_path,
        last_loaded: Some(Instant::now()),
        warnings,
    })
}

// =============================================================================
// RAII instruction limit guard
// =============================================================================

/// RAII guard that sets a tighter instruction limit for formula evaluation
/// and restores the previous state on drop.
pub(crate) struct FormulaLimitGuard<'a> {
    lua: &'a Lua,
}

/// Instruction budget for custom formula functions (10 million).
const FORMULA_INSTRUCTION_LIMIT: i64 = 10_000_000;

/// Hook check interval for formula functions.
const FORMULA_HOOK_INTERVAL: u32 = 1_000;

/// Wall-clock timeout for formula functions (1 second).
const FORMULA_TIMEOUT: std::time::Duration = std::time::Duration::from_secs(1);

impl<'a> FormulaLimitGuard<'a> {
    pub(crate) fn new(lua: &'a Lua) -> Self {
        use std::sync::atomic::{AtomicI64, Ordering};

        let counter = std::sync::Arc::new(AtomicI64::new(0));
        let start = Instant::now();

        let counter_clone = counter.clone();
        lua.set_hook(
            HookTriggers::new().every_nth_instruction(FORMULA_HOOK_INTERVAL),
            move |_lua, _debug| {
                let count = counter_clone.fetch_add(FORMULA_HOOK_INTERVAL as i64, Ordering::Relaxed);
                if count >= FORMULA_INSTRUCTION_LIMIT {
                    return Err(mlua::Error::RuntimeError(
                        "instruction limit exceeded".to_string(),
                    ));
                }
                if start.elapsed() > FORMULA_TIMEOUT {
                    return Err(mlua::Error::RuntimeError(
                        "execution timeout (1s limit)".to_string(),
                    ));
                }
                Ok(VmState::Continue)
            },
        );

        Self { lua }
    }
}

impl<'a> Drop for FormulaLimitGuard<'a> {
    fn drop(&mut self) {
        self.lua.remove_hook();
    }
}

// =============================================================================
// VisiGridRange userdata
// =============================================================================

/// Lua userdata for range arguments passed to custom functions.
/// Provides `.n` field and `:get(i)` method (1-based indexing).
struct VisiGridRange {
    values: Vec<Value>,
}

impl mlua::UserData for VisiGridRange {
    fn add_fields<F: mlua::UserDataFields<Self>>(fields: &mut F) {
        fields.add_field_method_get("n", |_, this| Ok(this.values.len()));
    }

    fn add_methods<M: mlua::UserDataMethods<Self>>(methods: &mut M) {
        methods.add_method("get", |lua, this, i: usize| {
            if i == 0 || i > this.values.len() {
                return Err(mlua::Error::RuntimeError(
                    format!("index {} out of range (1..{})", i, this.values.len()),
                ));
            }
            value_to_lua(lua, &this.values[i - 1])
        });
    }
}

/// Convert an engine Value to a Lua value.
fn value_to_lua(lua: &Lua, v: &Value) -> mlua::Result<mlua::Value> {
    match v {
        Value::Number(n) => Ok(mlua::Value::Number(*n)),
        Value::Text(s) => Ok(mlua::Value::String(lua.create_string(s)?)),
        Value::Boolean(b) => Ok(mlua::Value::Boolean(*b)),
        Value::Empty => Ok(mlua::Value::Nil),
        Value::Error(_) => Ok(mlua::Value::Nil), // errors pre-propagated
    }
}

// =============================================================================
// Memo cache
// =============================================================================

/// In-cycle memoization cache for custom function calls.
/// Lifetime: single recalc cycle.
pub struct MemoCache {
    cache: HashMap<MemoKey, EvalResult>,
}

#[derive(Hash, Eq, PartialEq)]
pub(crate) struct MemoKey {
    name: String,
    args: Vec<MemoArg>,
}

#[derive(Hash, Eq, PartialEq)]
enum MemoArg {
    Number(u64),   // f64 bits after -0.0 → 0.0 canonicalization
    Text(String),
    Boolean(bool),
    Nil,
    Error(String),
    Range(u64),    // blake3 fingerprint
}

impl Default for MemoCache {
    fn default() -> Self {
        Self::new()
    }
}

impl MemoCache {
    pub(crate) fn lookup(&self, key: &MemoKey) -> Option<&EvalResult> {
        self.get(key)
    }
    pub(crate) fn store(&mut self, key: MemoKey, result: EvalResult) {
        self.insert(key, result)
    }
    pub fn new() -> Self {
        Self { cache: HashMap::new() }
    }

    fn get(&self, key: &MemoKey) -> Option<&EvalResult> {
        self.cache.get(key)
    }

    fn insert(&mut self, key: MemoKey, result: EvalResult) {
        self.cache.insert(key, result);
    }
}

/// Canonicalize f64 for hashing: -0.0 → 0.0.
fn canon_f64_bits(n: f64) -> u64 {
    let canonical = if n == 0.0 { 0.0f64 } else { n };
    canonical.to_bits()
}

/// Build a memo key from function name and evaluated args.
pub(crate) fn build_memo_key(name: &str, args: &[EvalArg]) -> MemoKey {
    let memo_args: Vec<MemoArg> = args.iter().map(|arg| {
        match arg {
            EvalArg::Scalar(v) => match v {
                Value::Number(n) => MemoArg::Number(canon_f64_bits(*n)),
                Value::Text(s) => MemoArg::Text(s.clone()),
                Value::Boolean(b) => MemoArg::Boolean(*b),
                Value::Empty => MemoArg::Nil,
                Value::Error(e) => MemoArg::Error(e.clone()),
            },
            EvalArg::Range { values, .. } => {
                MemoArg::Range(compute_range_fingerprint(values))
            }
        }
    }).collect();

    MemoKey { name: name.to_string(), args: memo_args }
}

/// Compute a deterministic fingerprint for a range of values using blake3.
fn compute_range_fingerprint(values: &[Value]) -> u64 {
    let mut hasher = blake3::Hasher::new();
    for v in values {
        match v {
            Value::Number(n) => {
                hasher.update(&[0]);
                let canonical = if *n == 0.0 { 0.0f64 } else { *n };
                hasher.update(&canonical.to_le_bytes());
            }
            Value::Text(s) => {
                hasher.update(&[1]);
                hasher.update(s.as_bytes());
            }
            Value::Boolean(b) => {
                hasher.update(&[2]);
                hasher.update(&[*b as u8]);
            }
            Value::Empty => {
                hasher.update(&[3]);
            }
            Value::Error(e) => {
                hasher.update(&[4]);
                hasher.update(e.as_bytes());
            }
        }
    }
    let hash = hasher.finalize();
    u64::from_le_bytes(hash.as_bytes()[..8].try_into().unwrap())
}

// =============================================================================
// Calling custom functions
// =============================================================================

/// Call a custom Lua function with already-evaluated arguments.
///
/// Error args are propagated without calling Lua. Results are memoized
/// per-cycle via the provided cache.
pub fn call_custom_function(
    lua: &Lua,
    name: &str,
    args: &[EvalArg],
    memo_cache: &RefCell<MemoCache>,
) -> EvalResult {
    // 1. Propagate error args
    for arg in args {
        match arg {
            EvalArg::Scalar(Value::Error(e)) => return EvalResult::Error(e.clone()),
            EvalArg::Range { values, .. } => {
                for v in values {
                    if let Value::Error(e) = v {
                        return EvalResult::Error(e.clone());
                    }
                }
            }
            _ => {}
        }
    }

    // 2. Check memo cache
    let memo_key = build_memo_key(name, args);
    if let Some(cached) = memo_cache.borrow().get(&memo_key) {
        return cached.clone();
    }

    // 3. Build Lua args
    let lua_args: Vec<mlua::Value> = match args.iter().map(|arg| eval_arg_to_lua(lua, arg)).collect() {
        Ok(a) => a,
        Err(e) => return EvalResult::Error(format!("#LUA! {}", e)),
    };

    // 4. Get the function from globals
    let func: mlua::Function = match lua.globals().get(name) {
        Ok(f) => f,
        Err(_) => return EvalResult::Error(format!("#NAME? '{}'", name)),
    };

    // 5. Call with RAII instruction limit guard
    let result = {
        let _guard = FormulaLimitGuard::new(lua);

        // Set up a read-only environment for the function call
        match setup_formula_env(lua, &func) {
            Ok(sandboxed_func) => {
                match sandboxed_func.call::<mlua::Value>(mlua::MultiValue::from_iter(lua_args)) {
                    Ok(ret) => lua_return_to_eval_result(&ret),
                    Err(e) => EvalResult::Error(format!("#LUA! {}", scrub_lua_runtime_error(&e, name))),
                }
            }
            Err(e) => EvalResult::Error(format!("#LUA! {}", e)),
        }
    };

    // 6. Cache and return
    memo_cache.borrow_mut().insert(memo_key, result.clone());
    result
}

/// Convert an EvalArg to a Lua value.
pub(crate) fn eval_arg_to_lua(lua: &Lua, arg: &EvalArg) -> mlua::Result<mlua::Value> {
    match arg {
        EvalArg::Scalar(v) => value_to_lua(lua, v),
        EvalArg::Range { values, .. } => {
            let range = VisiGridRange { values: values.clone() };
            Ok(mlua::Value::UserData(lua.create_userdata(range)?))
        }
    }
}

/// Convert a Lua return value to EvalResult.
pub(crate) fn lua_return_to_eval_result(val: &mlua::Value) -> EvalResult {
    match val {
        mlua::Value::Number(n) => EvalResult::Number(*n),
        mlua::Value::Integer(i) => EvalResult::Number(*i as f64),
        mlua::Value::String(s) => {
            match s.to_str() {
                Ok(s) => EvalResult::Text(s.to_string()),
                Err(_) => EvalResult::Error("#LUA! non-UTF8 string".to_string()),
            }
        }
        mlua::Value::Boolean(b) => EvalResult::Boolean(*b),
        mlua::Value::Nil => {
            // Lua nil → EvalResult::Text("") → cached as Value::Text("").
            //
            EvalResult::Empty
        }
        mlua::Value::Table(t) => lua_table_to_eval_result(t),
        _ => EvalResult::Error("#LUA! unsupported return type".to_string()),
    }
}

/// The most cells one function call may spill. Generous for real results
/// (a 1000-row × 100-column table) and small enough that a runaway loop
/// building a table cannot take the process down before the instruction
/// limit stops it.
pub const MAX_ARRAY_CELLS: usize = 100_000;

/// A returned table becomes an array that spills from the calling cell.
///
/// Two shapes are accepted, and they mirror how the sheet reads:
///
/// - a sequence of scalars, `{1, 2, 3}`, is a column: one value per row;
/// - a sequence of sequences, `{{"a", 1}, {"b", 2}}`, is rows of columns.
///
/// A ragged inner sequence is padded with empty cells to the widest row, so a
/// missing trailing value never shifts its neighbours. Shapes are not mixed:
/// a row that is a scalar beside a row that is a table is an error, since it
/// would have to be guessed at. An empty table is an empty cell. Only the
/// sequence part (1..n) is read; string keys are ignored.
fn lua_table_to_eval_result(t: &mlua::Table) -> EvalResult {
    fn scalar(v: &mlua::Value) -> Result<Value, String> {
        match v {
            mlua::Value::Number(n) => Ok(Value::Number(*n)),
            mlua::Value::Integer(i) => Ok(Value::Number(*i as f64)),
            mlua::Value::String(s) => s
                .to_str()
                .map(|s| Value::Text(s.to_string()))
                .map_err(|_| "#LUA! non-UTF8 string in table".to_string()),
            mlua::Value::Boolean(b) => Ok(Value::Boolean(*b)),
            mlua::Value::Nil => Ok(Value::Empty),
            other => Err(format!("#LUA! unsupported value in table: {}", other.type_name())),
        }
    }

    let len = t.raw_len();
    if len == 0 {
        return EvalResult::Empty;
    }
    // Refuse before copying anything: the guard exists to keep a runaway
    // result from costing memory, so it has to run on the table's shape, not
    // on the copy of it.
    if len > MAX_ARRAY_CELLS {
        return EvalResult::Error(format!("#LUA! array larger than {} cells", MAX_ARRAY_CELLS));
    }

    let mut rows: Vec<Vec<Value>> = Vec::with_capacity(len);
    let mut nested: Option<bool> = None;
    let mut width = 0usize;

    for i in 1..=len {
        let item: mlua::Value = match t.raw_get(i) {
            Ok(v) => v,
            Err(e) => return EvalResult::Error(format!("#LUA! {}", e)),
        };
        let is_table = matches!(item, mlua::Value::Table(_));
        match nested {
            None => nested = Some(is_table),
            Some(expected) if expected != is_table => {
                return EvalResult::Error(
                    "#LUA! table mixes scalars and rows; return either {a, b, c} or {{a, b}, {c, d}}".to_string(),
                )
            }
            _ => {}
        }
        let row: Vec<Value> = if let mlua::Value::Table(inner) = &item {
            let n = inner.raw_len();
            // Widest row seen so far times the number of rows the table has:
            // checked on lengths alone, before this row's cells are copied.
            if len.saturating_mul(width.max(n)) > MAX_ARRAY_CELLS {
                return EvalResult::Error(format!("#LUA! array larger than {} cells", MAX_ARRAY_CELLS));
            }
            let mut row = Vec::with_capacity(n);
            for j in 1..=n {
                let cell: mlua::Value = match inner.raw_get(j) {
                    Ok(v) => v,
                    Err(e) => return EvalResult::Error(format!("#LUA! {}", e)),
                };
                match scalar(&cell) {
                    Ok(v) => row.push(v),
                    Err(e) => return EvalResult::Error(e),
                }
            }
            row
        } else {
            match scalar(&item) {
                Ok(v) => vec![v],
                Err(e) => return EvalResult::Error(e),
            }
        };
        width = width.max(row.len());
        if rows.len().saturating_mul(width.max(1)) > MAX_ARRAY_CELLS {
            return EvalResult::Error(format!("#LUA! array larger than {} cells", MAX_ARRAY_CELLS));
        }
        rows.push(row);
    }

    if width == 0 {
        // Every row was an empty table: nothing to place.
        return EvalResult::Empty;
    }
    if rows.len() * width > MAX_ARRAY_CELLS {
        return EvalResult::Error(format!("#LUA! array larger than {} cells", MAX_ARRAY_CELLS));
    }
    for row in &mut rows {
        row.resize(width, Value::Empty);
    }
    EvalResult::Array(Array2D::from_vec(rows))
}

// =============================================================================
// Sandbox: read-only environment for formula function calls
// =============================================================================

/// Set up a sandboxed environment for calling a formula function.
///
/// Creates a wrapper function with a read-only `_ENV` that includes:
/// - Safe stdlib: math (without random), string, table
/// - Safe globals: type, tonumber, tostring, pairs, ipairs, select, error, pcall
/// - The function itself (via upvalue capture)
/// - Read-only __newindex that blocks global mutation
fn setup_formula_env(lua: &Lua, func: &mlua::Function) -> mlua::Result<mlua::Function> {
    let env = sandbox_globals(lua)?;
    let wrapper = freeze_env(lua, env)?;
    // Set the function's environment
    func.set_environment(wrapper)?;
    Ok(func.clone())
}

/// The globals a formula may see: read-only stdlib proxies (math without
/// random) and a handful of pure functions. Shared by custom functions and by
/// `=LUA` cells, which add their `args` before the table is frozen.
pub(crate) fn sandbox_globals(lua: &Lua) -> mlua::Result<mlua::Table> {
    // Build environment table with allowed globals
    let env = lua.create_table()?;

    // Copy safe stdlib tables through read-only proxies
    for lib_name in &["math", "string", "table"] {
        if let Ok(original) = lua.globals().get::<mlua::Table>(*lib_name) {
            // Strip math.random and math.randomseed
            if *lib_name == "math" {
                let filtered = lua.create_table()?;
                for (key, val) in original.pairs::<String, mlua::Value>().flatten() {
                    if key != "random" && key != "randomseed" {
                        filtered.set(key, val)?;
                    }
                }
                let frozen_filtered = freeze_table(lua, &filtered)?;
                env.set(*lib_name, frozen_filtered)?;
            } else {
                let frozen = freeze_table(lua, &original)?;
                env.set(*lib_name, frozen)?;
            }
        }
    }

    // Copy safe global functions
    for name in &["type", "tonumber", "tostring", "pairs", "ipairs", "select", "error", "pcall"] {
        if let Ok(val) = lua.globals().get::<mlua::Value>(*name) {
            env.set(*name, val)?;
        }
    }

    Ok(env)
}

/// Wrap a globals table so nothing can be assigned through it.
pub(crate) fn freeze_env(lua: &Lua, env: mlua::Table) -> mlua::Result<mlua::Table> {
    let mt = lua.create_table()?;
    mt.set("__index", env)?;
    mt.set("__newindex", lua.create_function(|_, (_t, key, _val): (mlua::Value, String, mlua::Value)| {
        Err::<(), _>(mlua::Error::RuntimeError(
            format!("Global state mutation is not allowed (attempted to set '{}')", key),
        ))
    })?)?;
    mt.set("__metatable", false)?;
    let wrapper = lua.create_table()?;
    wrapper.set_metatable(Some(mt));
    Ok(wrapper)
}

/// Create a read-only proxy table via metatable.
fn freeze_table(lua: &Lua, original: &mlua::Table) -> mlua::Result<mlua::Table> {
    let mt = lua.create_table()?;
    mt.set("__index", original.clone())?;
    mt.set("__newindex", lua.create_function(|_, _: mlua::MultiValue| {
        Err::<(), _>(mlua::Error::RuntimeError("Cannot modify standard library".to_string()))
    })?)?;
    mt.set("__metatable", false)?;

    let proxy = lua.create_table()?;
    proxy.set_metatable(Some(mt));
    Ok(proxy)
}

// =============================================================================
// Error scrubbing
// =============================================================================

/// Scrub a Lua load/compile error for display.
fn scrub_lua_load_error(err: &mlua::Error, source_path: &std::path::Path) -> String {
    let raw = err.to_string();
    let scrubbed = raw.replace(&source_path.to_string_lossy().to_string(), "functions.lua");
    // Truncate long messages
    if scrubbed.len() > 120 {
        format!("{}...", &scrubbed[..117])
    } else {
        scrubbed
    }
}

/// Scrub a Lua runtime error for display as a cell error.
pub(crate) fn scrub_lua_runtime_error(err: &mlua::Error, func_name: &str) -> String {
    let raw = err.to_string();

    // Extract just the message part (after last colon in first line)
    let first_line = raw.lines().next().unwrap_or(&raw);
    let core_msg = first_line.rsplit(": ").next().unwrap_or(first_line).trim();

    let msg = format!("{}: {}", func_name, core_msg);
    if msg.len() > 100 {
        format!("{}...", &msg[..97])
    } else {
        msg
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval_return(src: &str) -> EvalResult {
        let lua = Lua::new();
        let v: mlua::Value = lua.load(src).eval().expect("lua eval");
        lua_return_to_eval_result(&v)
    }

    fn array(r: &EvalResult) -> &Array2D {
        match r {
            EvalResult::Array(a) => a,
            other => panic!("expected array, got {:?}", other),
        }
    }

    #[test]
    fn sequence_of_scalars_spills_as_a_column() {
        let r = eval_return("return {1, 'two', true}");
        let a = array(&r);
        assert_eq!((a.rows(), a.cols()), (3, 1));
        assert_eq!(a.get(0, 0), Some(&Value::Number(1.0)));
        assert_eq!(a.get(1, 0), Some(&Value::Text("two".into())));
        assert_eq!(a.get(2, 0), Some(&Value::Boolean(true)));
    }

    #[test]
    fn sequence_of_sequences_is_rows_of_columns() {
        let r = eval_return("return {{'a', 1}, {'b', 2}}");
        let a = array(&r);
        assert_eq!((a.rows(), a.cols()), (2, 2));
        assert_eq!(a.get(1, 0), Some(&Value::Text("b".into())));
        assert_eq!(a.get(1, 1), Some(&Value::Number(2.0)));
    }

    #[test]
    fn ragged_rows_are_padded_with_empty_cells() {
        let r = eval_return("return {{1, 2, 3}, {4}}");
        let a = array(&r);
        assert_eq!((a.rows(), a.cols()), (2, 3));
        assert_eq!(a.get(1, 1), Some(&Value::Empty));
        assert_eq!(a.get(1, 2), Some(&Value::Empty));
    }

    #[test]
    fn empty_table_is_an_empty_cell_and_mixed_shapes_are_refused() {
        assert!(matches!(eval_return("return {}"), EvalResult::Empty));
        assert!(matches!(eval_return("return {{}, {}}"), EvalResult::Empty));
        match eval_return("return {1, {2, 3}}") {
            EvalResult::Error(e) => assert!(e.contains("mixes"), "{}", e),
            other => panic!("expected error, got {:?}", other),
        }
        match eval_return("return {function() end}") {
            EvalResult::Error(e) => assert!(e.contains("unsupported value"), "{}", e),
            other => panic!("expected error, got {:?}", other),
        }
    }

    #[test]
    fn oversized_arrays_are_refused_before_they_are_built() {
        let src = format!(
            "local t = {{}} for i = 1, {} do t[i] = i end return t",
            MAX_ARRAY_CELLS + 1
        );
        match eval_return(&src) {
            EvalResult::Error(e) => assert!(e.contains("larger than"), "{}", e),
            other => panic!("expected error, got {:?}", other),
        }
        // A single very wide row is refused on its length, before any of its
        // cells are converted: the values are functions, which would fail
        // conversion with a different message if they were ever visited.
        let src = format!(
            "local row = {{}} for i = 1, {} do row[i] = function() end end return {{row, row}}",
            MAX_ARRAY_CELLS
        );
        match eval_return(&src) {
            EvalResult::Error(e) => assert!(e.contains("larger than"), "{}", e),
            other => panic!("expected error, got {:?}", other),
        }
    }

    #[test]
    fn a_custom_function_returning_a_table_reaches_the_engine_as_an_array() {
        let lua = Lua::new();
        lua.load("function SEQ3() return {10, 20, 30} end").exec().unwrap();
        let memo = RefCell::new(MemoCache::new());
        let r = call_custom_function(&lua, "SEQ3", &[], &memo);
        let a = array(&r);
        assert_eq!((a.rows(), a.cols()), (3, 1));
        assert_eq!(a.get(2, 0), Some(&Value::Number(30.0)));
        // And the memo hands back the same array on a repeat call.
        let again = call_custom_function(&lua, "SEQ3", &[], &memo);
        assert_eq!(array(&again).rows(), 3);
    }
}
