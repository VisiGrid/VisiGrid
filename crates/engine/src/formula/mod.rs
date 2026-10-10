// Formula parsing and evaluation

pub mod parser;
pub mod names;
pub mod sheets;
pub(crate) mod source_refs;
pub mod extract;
pub mod structured;
pub mod eval;
pub mod functions;
pub mod analyze;
pub mod refs;
/// Formula-editing help shared by the desktop and the web: the function
/// table with signatures, caret context, diagnostics, highlight tokens.
pub mod help;
pub mod whole_range;

pub(crate) mod eval_helpers;
mod eval_let;
pub(crate) mod eval_math;
pub(crate) mod eval_financial;
pub(crate) mod eval_regex;
pub(crate) mod eval_text;
pub(crate) mod eval_logical;
pub(crate) mod eval_conditional;
pub(crate) mod eval_lookup;
pub(crate) mod eval_datetime;
pub(crate) mod eval_statistical;
pub(crate) mod eval_trig;
pub(crate) mod eval_array;
mod eval_lambda;
pub(crate) mod lift;

pub(crate) mod eval_subtotal;

pub(crate) mod eval_budget;
pub mod excel_namespaces;

pub(crate) mod reference;
