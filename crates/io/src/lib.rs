// File I/O operations

pub mod csv;
pub mod json;
pub mod native;
pub mod xlsx;
pub mod xlsx_validation;

/// Native .sheet format version
/// Increment when schema changes in a way that old versions can't read
pub const NATIVE_FORMAT_VERSION: u32 = 1;
