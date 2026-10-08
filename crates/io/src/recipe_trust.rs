//! Which recipe sources the user has approved, shared by the desktop app
//! and the CLI's MCP server.
//!
//! A `.recipe.toml` (or a workbook linked to one) can come from anyone and
//! can name any file. Before a recipe first reads its source the app asks;
//! the answer is remembered here per recipe file and stated source path.
//! An agent working through MCP may only run recipes the user approved.

use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

use crate::recipe::Recipe;

fn store_path() -> PathBuf {
    config_dir().join("recipe_sources.json")
}

/// VisiGrid's settings folder, where the approvals live. Nothing a recipe
/// produces may be written here.
pub fn config_dir() -> PathBuf {
    dirs::config_dir().unwrap_or_else(|| PathBuf::from(".")).join("visigrid")
}

/// The recipe file and the source path as the recipe states it (a pattern
/// stays a pattern), so next month's file under an approved pattern is
/// covered and a recipe edited to read elsewhere is not.
pub fn approval_key(recipe_path: &Path, recipe: &Recipe) -> String {
    let recipe_path = std::path::absolute(recipe_path).unwrap_or_else(|_| recipe_path.to_path_buf());
    let mut text = format!("{}\0{}", recipe_path.display(), recipe.source.identity());
    // A recipe that merges others reads their sources too: the approval
    // covers each one, so editing a merged recipe to read elsewhere (another
    // file, entity or server) asks again
    chain(&recipe_path, recipe, &mut Vec::new(), &mut text);
    blake3::hash(text.as_bytes()).to_hex()[..32].to_string()
}

/// Each merged recipe and what it reads (its source's identity: a file, or
/// a VisiBooks server, entity and report), recursively. `seen` is the chain
/// being followed, so a loop ends at once; the run refuses it separately.
fn chain(recipe_path: &Path, recipe: &Recipe, seen: &mut Vec<PathBuf>, text: &mut String) {
    if seen.len() >= crate::recipe::MAX_MERGE_DEPTH {
        return;
    }
    let dir = recipe_path.parent().unwrap_or(Path::new("."));
    for (_, path) in recipe.merged_recipes(dir) {
        let path = std::path::absolute(&path).unwrap_or(path);
        let key = path.canonicalize().unwrap_or_else(|_| path.clone());
        if seen.contains(&key) {
            text.push_str(&format!("\0loop\0{}", path.display()));
            continue;
        }
        match Recipe::load(&path) {
            Ok(merged) => {
                text.push_str(&format!("\0merge\0{}\0{}", path.display(), merged.source.identity()));
                seen.push(key);
                chain(&path, &merged, seen, text);
                seen.pop();
            }
            Err(_) => text.push_str(&format!("\0merge\0{}", path.display())),
        }
    }
}

fn load() -> BTreeSet<String> {
    std::fs::read_to_string(store_path())
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

pub fn is_approved(recipe_path: &Path, recipe: &Recipe) -> bool {
    load().contains(&approval_key(recipe_path, recipe))
}

pub fn approve(recipe_path: &Path, recipe: &Recipe) -> Result<(), String> {
    let mut all = load();
    if !all.insert(approval_key(recipe_path, recipe)) {
        return Ok(());
    }
    let file = store_path();
    if let Some(dir) = file.parent() {
        std::fs::create_dir_all(dir).map_err(|e| e.to_string())?;
    }
    let json = serde_json::to_string_pretty(&all).map_err(|e| e.to_string())?;
    std::fs::write(&file, json).map_err(|e| e.to_string())
}

