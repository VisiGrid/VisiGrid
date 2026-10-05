//! `vgrid recipe run`: run a saved import recipe without the app.
//!
//! The same executor as the desktop app (`visigrid_io::recipe`). Diagnostics
//! go to stderr; the result goes to `-o` or, as CSV, to stdout. A failed run
//! writes nothing and exits with EXIT_RECIPE_FAILED, so a previous output
//! file stays as it was.

use std::io::Write;
use std::path::{Path, PathBuf};

use visigrid_io::recipe::{self, Recipe, Snapshot};

use crate::convert::{infer_format, write_format};
use crate::exit_codes::{EXIT_RECIPE_FAILED, EXIT_RECIPE_INVALID};
use crate::{CliError, Format};

pub(crate) fn cmd_recipe_run(
    recipe_path: PathBuf,
    source: Option<PathBuf>,
    output: Option<PathBuf>,
    report_path: Option<PathBuf>,
    quiet: bool,
) -> Result<(), CliError> {
    let recipe = Recipe::load(&recipe_path).map_err(|e| CliError { code: EXIT_RECIPE_INVALID, message: e, hint: None })?;
    let recipe_dir = recipe_path.parent().filter(|p| !p.as_os_str().is_empty()).unwrap_or(Path::new("."));
    // One file, or every file an appending recipe matches
    let sources = recipe.resolve_sources(recipe_dir, source.as_deref()).map_err(|e| {
        CliError::io(e).with_hint("the pattern is matched in the recipe's folder; pass --source to read a given file")
    })?;
    for source_path in &sources {
        check_paths(&recipe_path, source_path, output.as_deref(), report_path.as_deref())?;
    }
    let snapshot = if recipe.source.is_remote() {
        // A VisiBooks report: read over the network, with the key from the
        // keychain (or VISIBOOKS_API_KEY in CI)
        check_paths(&recipe_path, &recipe.source_path(recipe_dir, None), output.as_deref(), report_path.as_deref())?;
        recipe.read_snapshot(recipe_dir, None).map_err(|e| {
            CliError::io(e).with_hint("save the key with `vgrid visibooks key`, or set VISIBOOKS_API_KEY")
        })?
    } else {
        Snapshot::read_all(&sources).map_err(|e| {
            CliError::io(e).with_hint("the path is relative to the recipe file; pass --source to read another file")
        })?
    };

    let result = recipe::run(&recipe, &snapshot);
    if let Some(path) = &report_path {
        let json = serde_json::to_string_pretty(&result.report).map_err(|e| CliError::io(e.to_string()))?;
        write_replacing(path, |tmp| std::fs::write(tmp, &json).map_err(|e| CliError::io(e.to_string())))?;
    }
    if !quiet || !result.report.ok {
        eprint!("{}", result.report.summary());
    }
    if !result.report.ok {
        return Err(CliError {
            code: EXIT_RECIPE_FAILED,
            message: "the recipe's checks failed; the result was not written".into(),
            hint: Some("fix the source or the recipe, then run again".into()),
        });
    }

    let sheet = result.output.to_sheet();
    match output {
        None => {
            let bytes = write_format(&sheet, Format::Csv, ',', false, 0, None, None)?;
            std::io::stdout().write_all(&bytes).map_err(|e| CliError::io(e.to_string()))?;
        }
        Some(path) => write_output(&sheet, &path)?,
    }
    Ok(())
}

/// Write a recipe's result to `path` in the format its extension names,
/// replacing it atomically. Shared with the MCP `run_recipe` tool.
pub(crate) fn write_output(sheet: &visigrid_engine::sheet::Sheet, path: &Path) -> Result<(), CliError> {
    let format = infer_format(&path.to_path_buf())?;
    write_replacing(path, |tmp| match format {
        Format::Xlsx => {
            let wb = visigrid_engine::workbook::Workbook::from_sheets(vec![sheet.clone()], 0);
            crate::convert::write_xlsx(&wb, Some(tmp), None)
        }
        Format::Sheet => visigrid_io::native::save(sheet, tmp).map_err(CliError::io),
        Format::Parquet => Err(CliError::format("parquet output is not supported by recipes yet")),
        other => {
            let bytes = write_format(sheet, other, ',', false, 0, None, None)?;
            std::fs::write(tmp, bytes).map_err(|e| CliError::io(e.to_string()))
        }
    })
}

/// The output and report must not land on each other, the source or the
/// recipe: a failed run would otherwise replace the previous output (or the
/// source) with a JSON report. Checked before anything is read or written.
pub(crate) fn check_paths(recipe: &Path, source: &Path, output: Option<&Path>, report: Option<&Path>) -> Result<(), CliError> {
    let clash = |what: &str, a: &Path, other: &str, b: &Path| -> Result<(), CliError> {
        if same_file(a, b) {
            return Err(CliError::args(format!("the {what} {} is also the {other}", a.display()))
                .with_hint("choose a different path"));
        }
        Ok(())
    };
    if let Some(out) = output {
        clash("output", out, "source", source)?;
        clash("output", out, "recipe", recipe)?;
    }
    if let Some(rep) = report {
        clash("report", rep, "source", source)?;
        clash("report", rep, "recipe", recipe)?;
        if let Some(out) = output {
            clash("report", rep, "output", out)?;
        }
    }
    Ok(())
}

/// The same file, whether or not either exists yet.
fn same_file(a: &Path, b: &Path) -> bool {
    let canon = |p: &Path| {
        std::fs::canonicalize(p).ok().or_else(|| {
            let parent = p.parent().filter(|d| !d.as_os_str().is_empty()).unwrap_or(Path::new("."));
            Some(std::fs::canonicalize(parent).ok()?.join(p.file_name()?))
        })
    };
    match (canon(a), canon(b)) {
        (Some(x), Some(y)) => x == y,
        _ => a == b,
    }
}

/// Write to a temporary file beside `path`, then rename it into place, so
/// a reader never sees a half-written output and a failure leaves the old
/// file intact.
fn write_replacing(path: &Path, write: impl FnOnce(&Path) -> Result<(), CliError>) -> Result<(), CliError> {
    let dir = path.parent().filter(|p| !p.as_os_str().is_empty()).unwrap_or(Path::new("."));
    let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("output");
    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("tmp");
    // Keep the extension last: some writers choose behaviour by it
    let tmp = dir.join(format!(".{name}.{}.partial.{ext}", std::process::id()));
    let result = write(&tmp).and_then(|_| {
        std::fs::rename(&tmp, path).map_err(|e| CliError::io(format!("{}: {e}", path.display())))
    });
    if result.is_err() {
        let _ = std::fs::remove_file(&tmp);
    }
    result
}
