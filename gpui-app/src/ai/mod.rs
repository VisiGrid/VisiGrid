// AI integration module
//
// AI assists reasoning about data. It does not compute, edit, or decide.
// All computation, mutation, and verification remains deterministic and owned by VisiGrid.

mod context;
mod client;

pub use context::{
    build_ai_context, cell_ref, range_ref,
    find_current_region, find_used_range,
};
pub use client::{ask_ai, analyze, AskError, AnalyzeResponse, ANALYZE_CONTRACT, INSERT_FORMULA_CONTRACT, DIFF_EXPLAIN_CONTRACT};
