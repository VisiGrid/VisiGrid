//! Settings module for VisiGrid
//!
//! This module provides a centralized, strongly-typed settings system with:
//! - `UserSettings`: Personal preferences (global, persistent)
//! - `DocumentSettings`: Per-file settings (stored with document)
//! - `ResolvedSettings`: Runtime truth (merged from user + document + defaults)
//!
//! Key design decisions:
//! - All settings use `Setting<T>` for explicit three-state semantics (Inherit/Value)
//! - Settings are defined in one place - no distributed registration
//! - Files are the organizing principle (user.json, embedded in .vgrid)

mod types;
mod user;
mod document;
mod resolved;
mod persistence;
mod store;

pub use types::*;
pub use user::*;
pub use document::*;
pub use resolved::*;
pub use persistence::*;
pub use store::*;
