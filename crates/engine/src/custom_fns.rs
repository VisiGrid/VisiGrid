//! The process-wide custom-function hooks.
//!
//! The engine has no Lua. Custom functions and `=LUA` cells are answered by
//! hooks the host process registers once at startup: the desktop app and the
//! CLI install the scripting crate's adapter, the browser installs nothing.
//! Every evaluation path that does not receive an explicit handler, which is
//! the incremental recalc after an edit or an agent op and every
//! `recompute_full_ordered`, consults `call`. That is what makes a custom
//! function update when its input changes, rather than only on a full recalc.
//!
//! `begin_recalc` and `end_recalc` bracket every recalculation the engine
//! runs, full or incremental, however it was reached. They exist so the
//! handler can scope a memo to exactly one logical recalc: identical calls
//! within it answer once, and nothing survives it. A recalc that starts
//! another (the incremental path's cycle fallback) nests, so the hooks are
//! called in matched pairs and the handler counts depth.
//!
//! Plain `fn` pointers, not closures, so `Workbook` stays `Send` and
//! serialisable; whatever state the hooks need lives on their side.

use std::sync::RwLock;

use crate::formula::eval::{EvalArg, EvalResult};

/// `None` means "not mine", and the engine reports the name as unknown.
pub type CustomFnHandler = fn(&str, &[EvalArg]) -> Option<EvalResult>;

#[derive(Clone, Copy)]
pub struct CustomFnHooks {
    pub call: CustomFnHandler,
    pub begin_recalc: fn(),
    pub end_recalc: fn(),
}

static DEFAULT_HOOKS: RwLock<Option<CustomFnHooks>> = RwLock::new(None);

/// Register (or clear) the hooks every workbook uses unless given a handler.
pub fn set_default_custom_fns(hooks: Option<CustomFnHooks>) {
    *DEFAULT_HOOKS.write().unwrap_or_else(|e| e.into_inner()) = hooks;
}

/// Register a bare handler with no recalc bracketing. For hosts and tests
/// that have nothing to scope.
pub fn set_default_custom_fn_handler(handler: Option<CustomFnHandler>) {
    set_default_custom_fns(handler.map(|call| CustomFnHooks {
        call,
        begin_recalc: || {},
        end_recalc: || {},
    }));
}

pub fn default_custom_fn_handler() -> Option<CustomFnHandler> {
    DEFAULT_HOOKS
        .read()
        .unwrap_or_else(|e| e.into_inner())
        .map(|h| h.call)
}

/// Brackets one recalculation: `begin_recalc` now, `end_recalc` on drop.
pub struct RecalcScope {
    end: Option<fn()>,
}

pub fn recalc_scope() -> RecalcScope {
    let hooks = *DEFAULT_HOOKS.read().unwrap_or_else(|e| e.into_inner());
    match hooks {
        Some(h) => {
            (h.begin_recalc)();
            RecalcScope { end: Some(h.end_recalc) }
        }
        None => RecalcScope { end: None },
    }
}

impl Drop for RecalcScope {
    fn drop(&mut self) {
        if let Some(end) = self.end.take() {
            end();
        }
    }
}
