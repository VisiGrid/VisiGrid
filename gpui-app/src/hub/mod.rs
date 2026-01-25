// VisiHub sync module
//
// Provides sync functionality between VisiGrid and VisiHub.
// Phase 1: Pull-only (download updates from VisiHub)

pub mod auth;
pub mod client;
pub mod types;

pub use auth::{AuthCredentials, load_auth, save_auth, delete_auth, is_authenticated};
pub use client::{HubClient, HubError, UserInfo, hash_file, hash_bytes};
pub use types::{HubStatus, RemoteStatus, compute_status};

// Re-export HubLink from io crate for convenience
pub use visigrid_io::native::{HubLink, load_hub_link, save_hub_link, delete_hub_link};
