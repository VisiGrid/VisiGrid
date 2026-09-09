//! The process-wide custom-function handler.
//!
//! The engine has no Lua. Custom functions and `=LUA` cells are answered by a
//! handler the host process registers once at startup: the desktop app and
//! the CLI install the scripting crate's adapter, the browser installs
//! nothing. Every evaluation path that does not receive an explicit handler,
//! which is the incremental recalc after an edit or an agent op and every
//! `recompute_full_ordered`, consults this one. That is what makes a custom
//! function update when its input changes, rather than only on a full recalc.
//!
//! A plain `fn` pointer, not a closure, so `Workbook` stays `Send` and
//! serialisable; whatever state the handler needs (a Lua runtime, a registry,
//! caches) lives on the handler's side, per thread if it must.

use std::sync::RwLock;

use crate::formula::eval::{EvalArg, EvalResult};

/// `None` means "not mine", and the engine reports the name as unknown.
pub type CustomFnHandler = fn(&str, &[EvalArg]) -> Option<EvalResult>;

static DEFAULT_HANDLER: RwLock<Option<CustomFnHandler>> = RwLock::new(None);

/// Register (or clear) the handler every workbook uses unless given another.
pub fn set_default_custom_fn_handler(handler: Option<CustomFnHandler>) {
    *DEFAULT_HANDLER.write().unwrap_or_else(|e| e.into_inner()) = handler;
}

pub fn default_custom_fn_handler() -> Option<CustomFnHandler> {
    *DEFAULT_HANDLER.read().unwrap_or_else(|e| e.into_inner())
}
