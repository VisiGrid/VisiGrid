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

use std::path::Path;

use visigrid_io::recipe::Recipe;
// The approval store itself is shared with the CLI's MCP server
pub use visigrid_io::recipe_trust::{approve, is_approved};

/// What the confirmation shows about the file a recipe would read: the
/// file itself, or for a pattern the folder, the pattern and the file it
/// matches now. Err when it cannot be read at all (then only Cancel).
pub fn describe_source(recipe_path: &Path, recipe: &Recipe) -> Result<(String, String), String> {
    let (what, mut detail) = describe_own(recipe_path, recipe)?;
    // Merge steps read other recipes' sources too: name each one
    let dir = recipe_path.parent().unwrap_or(Path::new("."));
    for (with, path) in recipe.merged_recipes(dir) {
        let reads = Recipe::load(&path).map(|r| r.source.path().to_string()).unwrap_or_else(|e| format!("can't be read: {e}"));
        detail.push_str(&format!(" · also merges {with}, which reads {reads}"));
    }
    Ok((what, detail))
}

fn describe_own(recipe_path: &Path, recipe: &Recipe) -> Result<(String, String), String> {
    let dir = recipe_path.parent().unwrap_or(Path::new("."));
    // A VisiBooks report: name the server, entity and report, and whether a
    // key for that server is saved (the key itself never leaves the keychain)
    if let visigrid_io::recipe::Source::Visibooks(src) = &recipe.source {
        let origin = visigrid_io::recipe::visibooks::origin(&src.server)?;
        let key = match visigrid_io::recipe::visibooks::api_key(&origin) {
            Ok(_) => "an API key for this server is saved".to_string(),
            Err(_) => "no API key for this server is saved yet; run `vgrid visibooks key`".to_string(),
        };
        let transport = if origin.starts_with("https://") { "over https" } else { "on this computer" };
        return Ok((src.describe(), format!("read-only, {transport} · {key}")));
    }
    // A PostgreSQL read: name the server, database, role and table, and
    // whether a password is saved (it never leaves the keychain)
    if let visigrid_io::recipe::Source::Postgres(src) = &recipe.source {
        let password = match visigrid_io::recipe::postgres::password(src) {
            Ok(_) => "a password for this role is saved",
            Err(_) => "no password for this role is saved yet; run `vgrid postgres password`",
        };
        let transport = if src.tls { "over verified TLS" } else { "on this computer" };
        return Ok((src.describe(), format!("read-only role, checked before every read, {transport} · {password}")));
    }
    if recipe.source.combine() && recipe.source_is_pattern() {
        let files = recipe.resolve_sources(dir, None)?;
        let pattern = recipe.source_path(dir, None);
        let total: u64 = files.iter().filter_map(|f| std::fs::metadata(f).ok()).map(|m| m.len()).sum();
        return Ok((
            format!(
                "every file matching {} in {}; now {} file{}",
                pattern.file_name().and_then(|n| n.to_str()).unwrap_or(""),
                pattern.parent().map(|p| p.display().to_string()).unwrap_or_default(),
                files.len(),
                if files.len() == 1 { "" } else { "s" }
            ),
            format!("{} KB in all", total / 1024),
        ));
    }
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
    use visigrid_io::recipe_trust::approval_key;
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
