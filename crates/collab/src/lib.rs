//! Collaborative workbook model, Phase 1 (see the vault spec
//! "VisiGrid Collaborative Workbook Model").
//!
//! - [`op`]: the operation vocabulary and the protocol v2 envelope.
//! - [`transform`]: `T(a, b)`, the pair table, and list transforms.
//! - [`apply`]: applying ops to the real engine, and convergence fingerprints.
//! - [`server`]: the in-memory per-workbook sequencer.
//! - [`client`]: a replica with optimistic local edits and rebase.
//! - [`undo`]: per-user undo as inverse operations.
//! - [`sort`]: a sort's row order, computed by the writer.
//! - [`copy`]: a whole sheet as operations (import into a live workbook, duplicate a tab).
//! - [`sim`]: the headless convergence simulator; [`gen`]: its op generator.

pub mod apply;
pub mod client;
pub mod clock;
pub mod copy;
#[cfg(feature = "sim")]
pub mod gen;
pub mod op;
pub mod server;
#[cfg(feature = "sim")]
pub mod sim;
pub mod sort;
pub mod transform;
pub mod undo;

pub mod wire;

#[cfg(feature = "socket")]
pub mod socket;
