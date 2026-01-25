// VisiHub sync module
//
// Provides sync functionality between VisiGrid and VisiHub.
// Phase 1: Pull-only (download updates from VisiHub)

pub mod auth;
pub mod client;
pub mod types;

pub use auth::{AuthCredentials, load_auth, save_auth, delete_auth};
pub use client::{HubClient, RepoInfo, DatasetInfo, hash_file, hash_bytes, hashes_match};
pub use types::{HubStatus, HubActivity, compute_status};

// Re-export HubLink from io crate for convenience
pub use visigrid_io::native::{HubLink, load_hub_link, save_hub_link, delete_hub_link};
