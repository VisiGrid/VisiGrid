//! Terminal panel: PTY terminal emulator embedded in VisiGrid.
//!
//! Uses `alacritty_terminal` for ANSI parsing and grid state, with a thin
//! wrapper for PTY lifecycle and GPUI rendering.

pub mod pty;
pub mod state;

pub use state::{TerminalState, resolve_workspace_root};
