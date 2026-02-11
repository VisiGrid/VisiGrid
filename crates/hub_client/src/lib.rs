//! VisiHub API client — shared between desktop and CLI.
//!
//! This crate is the single source of truth for the VisiHub wire contract:
//! auth, create revision, upload, complete, poll run status, proof URL.
//!
//! No GUI concepts. No retries beyond basic backoff. No progress bars.

mod auth;
mod client;

pub use auth::{AuthCredentials, auth_file_path, load_auth, save_auth, delete_auth};
pub use client::{
    HubClient, HubError, UserInfo, RepoInfo, DatasetInfo,
    CreateRevisionOptions, RunResult,
    AssertionInput, AssertionResult, EngineMetadata,
    hash_file, hash_bytes,
};
