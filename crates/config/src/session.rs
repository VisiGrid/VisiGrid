use serde::{Deserialize, Serialize};
use std::collections::hash_map::DefaultHasher;
use std::fs;
use std::hash::{Hash, Hasher};
use std::path::PathBuf;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RangeState {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct SplitState {
    pub enabled: bool,
    pub direction: String,      // "horizontal" | "vertical"
    pub active_pane: String,    // "primary" | "secondary"
    pub scroll_row: usize,
    pub scroll_col: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct Session {
    pub version: u32,
    pub current_file: Option<PathBuf>,
    pub selection: Vec<RangeState>,
    pub active_cell: (usize, usize),
    pub scroll_row: usize,
    pub scroll_col: usize,
    pub split: SplitState,
    pub dark_mode: bool,
    pub zen_mode: bool,
    pub show_inspector: bool,
    pub show_problems: bool,
    pub show_formulas: bool,
    pub show_repl: bool,
}

impl Session {
    pub fn path() -> PathBuf {
        dirs::config_dir()
            .unwrap_or_else(|| PathBuf::from("."))
            .join("visigrid")
            .join("session.json")
    }

    pub fn load() -> Option<Self> {
        let path = Self::path();
        fs::read_to_string(&path).ok()
            .and_then(|s| serde_json::from_str(&s).ok())
    }

    pub fn save(&self) -> Result<(), String> {
        let path = Self::path();
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent).map_err(|e| e.to_string())?;
        }
        let json = serde_json::to_string_pretty(self).map_err(|e| e.to_string())?;
        fs::write(&path, json).map_err(|e| e.to_string())
    }
}

/// Workspace: per-project session state
/// Workspaces are identified by project directory and stored separately
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workspace {
    pub project_path: PathBuf,
    pub session: Session,
}

impl Workspace {
    /// Get the workspaces directory
    fn workspaces_dir() -> PathBuf {
        dirs::config_dir()
            .unwrap_or_else(|| PathBuf::from("."))
            .join("visigrid")
            .join("workspaces")
    }

    /// Hash a path to create a unique filename
    fn hash_path(path: &PathBuf) -> String {
        let mut hasher = DefaultHasher::new();
        path.hash(&mut hasher);
        format!("{:016x}", hasher.finish())
    }

    /// Get workspace file path for a project
    fn workspace_path(project: &PathBuf) -> PathBuf {
        Self::workspaces_dir().join(format!("{}.json", Self::hash_path(project)))
    }

    /// Detect project root by looking for .visigrid marker file
    /// Falls back to current working directory
    pub fn detect_project() -> Option<PathBuf> {
        let cwd = std::env::current_dir().ok()?;

        // Walk up looking for .visigrid marker
        let mut dir = cwd.as_path();
        loop {
            let marker = dir.join(".visigrid");
            if marker.exists() {
                return Some(dir.to_path_buf());
            }
            match dir.parent() {
                Some(parent) => dir = parent,
                None => break,
            }
        }

        // No marker found, use cwd
        Some(cwd)
    }

    /// Load workspace for a project directory
    pub fn load(project: &PathBuf) -> Option<Self> {
        let path = Self::workspace_path(project);
        fs::read_to_string(&path).ok()
            .and_then(|s| serde_json::from_str(&s).ok())
    }

    /// Load workspace for current project (auto-detected)
    pub fn load_current() -> Option<Self> {
        let project = Self::detect_project()?;
        Self::load(&project)
    }

    /// Save workspace
    pub fn save(&self) -> Result<(), String> {
        let dir = Self::workspaces_dir();
        fs::create_dir_all(&dir).map_err(|e| e.to_string())?;

        let path = Self::workspace_path(&self.project_path);
        let json = serde_json::to_string_pretty(self).map_err(|e| e.to_string())?;
        fs::write(&path, json).map_err(|e| e.to_string())
    }

    /// Create a new workspace from a session
    pub fn new(project: PathBuf, session: Session) -> Self {
        Self { project_path: project, session }
    }

    /// List all saved workspaces with their project paths
    pub fn list_all() -> Vec<(PathBuf, String)> {
        let dir = Self::workspaces_dir();
        let mut result = Vec::new();

        if let Ok(entries) = fs::read_dir(&dir) {
            for entry in entries.flatten() {
                if let Ok(contents) = fs::read_to_string(entry.path()) {
                    if let Ok(ws) = serde_json::from_str::<Workspace>(&contents) {
                        let name = ws.project_path
                            .file_name()
                            .and_then(|n| n.to_str())
                            .unwrap_or("Unknown")
                            .to_string();
                        result.push((ws.project_path, name));
                    }
                }
            }
        }

        result
    }
}
