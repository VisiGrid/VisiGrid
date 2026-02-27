// Cloud sync module
//
// Provides cloud-backed sheet sync with VisiHub Sheets API.
// Cloud sync is automatic (post-save), while hub sync is manual (publish/pull).

pub mod identity;
pub mod sheets_client;
pub mod sync;
pub mod operations;

pub use identity::CloudSyncState;
pub use sheets_client::{SheetsClient, SheetInfo, SaveResponse};

// Re-export CloudIdentity and persistence functions from io crate
pub use visigrid_io::native::{CloudIdentity, load_cloud_identity, save_cloud_identity, delete_cloud_identity};
