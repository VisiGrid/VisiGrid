//! Reusable UI components for VisiGrid dialogs and overlays.
//!
//! This module provides common UI primitives to eliminate duplication
//! across dialog implementations.

mod button;
mod dialog_frame;
mod locked_feature_panel;
mod modal;
mod popup;
pub mod text_input;

pub use button::Button;
pub use dialog_frame::{DialogFrame, DialogSize, dialog_header_simple, dialog_header_with_subtitle};
pub use locked_feature_panel::render_locked_feature_panel;
pub use modal::{modal_overlay, modal_backdrop};
pub use popup::{popup, clamp_to_viewport};
