//! Collaborative workbook model, Phase 1 (see the vault spec
//! "VisiGrid Collaborative Workbook Model").
//!
//! - [`op`]: the operation vocabulary and the protocol v2 envelope.
//! - [`transform`]: `T(a, b)`, the pair table, and list transforms.
//! - [`apply`]: applying ops to the real engine, and convergence fingerprints.
//! - [`server`]: the in-memory per-workbook sequencer.
//! - [`client`]: a replica with optimistic local edits and rebase.
//! - [`sim`]: the headless convergence simulator; [`gen`]: its op generator.

pub mod apply;
pub mod client;
pub mod gen;
pub mod op;
pub mod server;
pub mod sim;
pub mod transform;
