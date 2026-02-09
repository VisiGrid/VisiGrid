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

use visigrid_engine::formula::eval::{EvalArg, EvalResult, Value};
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
struct FormulaLimitGuard<'a> {
    lua: &'a Lua,
}

/// Instruction budget for custom formula functions (10 million).
const FORMULA_INSTRUCTION_LIMIT: i64 = 10_000_000;

/// Hook check interval for formula functions.
const FORMULA_HOOK_INTERVAL: u32 = 1_000;

/// Wall-clock timeout for formula functions (1 second).
const FORMULA_TIMEOUT: std::time::Duration = std::time::Duration::from_secs(1);

impl<'a> FormulaLimitGuard<'a> {
    fn new(lua: &'a Lua) -> Self {
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
struct MemoKey {
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

impl MemoCache {
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
fn build_memo_key(name: &str, args: &[EvalArg]) -> MemoKey {
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
fn eval_arg_to_lua(lua: &Lua, arg: &EvalArg) -> mlua::Result<mlua::Value> {
    match arg {
        EvalArg::Scalar(v) => value_to_lua(lua, v),
        EvalArg::Range { values, .. } => {
            let range = VisiGridRange { values: values.clone() };
            Ok(mlua::Value::UserData(lua.create_userdata(range)?))
        }
    }
}

/// Convert a Lua return value to EvalResult.
fn lua_return_to_eval_result(val: &mlua::Value) -> EvalResult {
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
        _ => EvalResult::Error("#LUA! unsupported return type".to_string()),
    }
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
    // Build environment table with allowed globals
    let env = lua.create_table()?;

    // Copy safe stdlib tables through read-only proxies
    for lib_name in &["math", "string", "table"] {
        if let Ok(original) = lua.globals().get::<mlua::Table>(*lib_name) {
            // Strip math.random and math.randomseed
            if *lib_name == "math" {
                let filtered = lua.create_table()?;
                for pair in original.pairs::<String, mlua::Value>() {
                    if let Ok((key, val)) = pair {
                        if key != "random" && key != "randomseed" {
                            filtered.set(key, val)?;
                        }
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

    // Wrap env in read-only metatable
    let mt = lua.create_table()?;
    mt.set("__index", env.clone())?;
    mt.set("__newindex", lua.create_function(|_, (_t, key, _val): (mlua::Value, String, mlua::Value)| {
        Err::<(), _>(mlua::Error::RuntimeError(
            format!("Global state mutation is not allowed (attempted to set '{}')", key),
        ))
    })?)?;
    mt.set("__metatable", false)?;

    let wrapper = lua.create_table()?;
    wrapper.set_metatable(Some(mt));

    // Set the function's environment
    func.set_environment(wrapper)?;

    Ok(func.clone())
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
fn scrub_lua_runtime_error(err: &mlua::Error, func_name: &str) -> String {
    let raw = err.to_string();

    // Extract just the message part (after last colon in first line)
    let first_line = raw.lines().next().unwrap_or(&raw);
    let core_msg = first_line.rsplitn(2, ": ").next().unwrap_or(first_line).trim();

    let msg = format!("{}: {}", func_name, core_msg);
    if msg.len() > 100 {
        format!("{}...", &msg[..97])
    } else {
        msg
    }
}
