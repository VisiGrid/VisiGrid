//! Which recipe sources the user has approved.
//!
//! A `.recipe.toml` (or a workbook linked to one) can come from anyone and
//! can name any file. Before a recipe first reads its source, the app shows
//! which file that is and asks; the answer is remembered per recipe file and
//! source path, so a monthly refresh does not ask again. Recipes the user
//! builds, or points at a file themselves, are approved as they are saved.
//!
//! The engine separately refuses anything but local regular files within a
//! size cap (`visigrid_io::recipe::read_regular_file`).

use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

use visigrid_io::recipe::{Recipe, Source};

fn store_path() -> PathBuf {
    dirs::config_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("visigrid")
        .join("recipe_sources.json")
}

/// The recipe file and the source path as the recipe states it (a pattern
/// stays a pattern), so next month's file under an approved pattern is
/// covered and a recipe edited to read elsewhere is not.
pub fn approval_key(recipe_path: &Path, recipe: &Recipe) -> String {
    let Source::Csv(src) = &recipe.source;
    let recipe_path = std::path::absolute(recipe_path).unwrap_or_else(|_| recipe_path.to_path_buf());
    let text = format!("{}\0{}", recipe_path.display(), src.path);
    blake3::hash(text.as_bytes()).to_hex()[..32].to_string()
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

/// What the confirmation shows about the file a recipe would read: the
/// file itself, or for a pattern the folder, the pattern and the file it
/// matches now. Err when it cannot be read at all (then only Cancel).
pub fn describe_source(recipe_path: &Path, recipe: &Recipe) -> Result<(String, String), String> {
    let dir = recipe_path.parent().unwrap_or(Path::new("."));
    let resolved = recipe.resolve_source(dir, None)?;
    let meta = std::fs::metadata(&resolved).map_err(|e| format!("{}: {e}", resolved.display()))?;
    if !meta.is_file() {
        return Err(format!("{} is not a regular file", resolved.display()));
    }
    let size = match meta.len() {
        n if n >= 1024 * 1024 => format!("{:.1} MB", n as f64 / (1024.0 * 1024.0)),
        n if n >= 1024 => format!("{} KB", n / 1024),
        n => format!("{n} bytes"),
    };
    let modified = meta
        .modified()
        .ok()
        .map(|t| chrono::DateTime::<chrono::Local>::from(t).format("%b %-d, %Y %H:%M").to_string())
        .unwrap_or_default();
    let detail = format!("{size} · modified {modified}");
    if recipe.source_is_pattern() {
        let pattern = recipe.source_path(dir, None);
        Ok((
            format!(
                "the newest file matching {} in {}; now {}",
                pattern.file_name().and_then(|n| n.to_str()).unwrap_or(""),
                pattern.parent().map(|p| p.display().to_string()).unwrap_or_default(),
                resolved.file_name().and_then(|n| n.to_str()).unwrap_or("")
            ),
            detail,
        ))
    } else {
        Ok((resolved.display().to_string(), detail))
    }
}

#[cfg(test)]
mod tests {
    use super::approval_key;
    use std::path::Path;
    use visigrid_io::recipe::Recipe;

    fn recipe(path: &str) -> Recipe {
        Recipe::from_toml(&format!("version = 1\n[source]\nkind = \"csv\"\npath = \"{path}\"\n")).unwrap()
    }

    #[test]
    fn approval_covers_one_recipe_and_one_stated_source() {
        let a = approval_key(Path::new("/d/orders.recipe.toml"), &recipe("export-*.csv"));
        // Same recipe, same stated source (the steps may change): same key
        assert_eq!(a, approval_key(Path::new("/d/orders.recipe.toml"), &recipe("export-*.csv")));
        // Pointed elsewhere, or another recipe file: asks again
        assert_ne!(a, approval_key(Path::new("/d/orders.recipe.toml"), &recipe("/home/u/.ssh/id_ed25519")));
        assert_ne!(a, approval_key(Path::new("/e/orders.recipe.toml"), &recipe("export-*.csv")));
    }
}
