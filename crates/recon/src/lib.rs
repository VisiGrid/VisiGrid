//! `visigrid-recon` — Generic multi-source reconciliation engine.
//!
//! Pure engine crate: receives pre-loaded records, returns classified results.
//! No CLI or IO dependencies.

pub mod aggregate;
pub mod classify;
pub mod config;
pub mod engine;
pub mod error;
pub mod evidence;
pub mod matcher;
pub mod model;

pub use config::ReconConfig;
pub use engine::run;
pub use error::ReconError;
pub use model::{ReconInput, ReconResult, ReconRow};
