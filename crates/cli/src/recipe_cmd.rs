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
    let source_path = recipe.source_path(recipe_dir, source.as_deref());
    let snapshot = Snapshot::read(&source_path).map_err(|e| {
        CliError::io(e).with_hint("the path is relative to the recipe file; pass --source to read another file")
    })?;

    let result = recipe::run(&recipe, &snapshot);
    if let Some(path) = &report_path {
        let json = serde_json::to_string_pretty(&result.report).map_err(|e| CliError::io(e.to_string()))?;
        std::fs::write(path, json).map_err(|e| CliError::io(format!("{}: {e}", path.display())))?;
    }
    if !quiet || !result.report.ok {
        eprint!("{}", result.report.summary());
    }
    if !result.report.ok {
        return Err(CliError {
            code: EXIT_RECIPE_FAILED,
            message: "the recipe's checks failed; nothing was written".into(),
            hint: Some("fix the source or the recipe, then run again".into()),
        });
    }

    let sheet = result.output.to_sheet();
    match output {
        None => {
            let bytes = write_format(&sheet, Format::Csv, ',', false, 0, None, None)?;
            std::io::stdout().write_all(&bytes).map_err(|e| CliError::io(e.to_string()))?;
        }
        Some(path) => {
            let format = infer_format(&path)?;
            write_replacing(&path, |tmp| match format {
                Format::Xlsx => {
                    let wb = visigrid_engine::workbook::Workbook::from_sheets(vec![sheet.clone()], 0);
                    visigrid_io::xlsx::export(&wb, tmp, None).map(|_| ()).map_err(CliError::io)
                }
                Format::Sheet => visigrid_io::native::save(&sheet, tmp).map_err(CliError::io),
                Format::Parquet => Err(CliError::format("parquet output is not supported by recipes yet")),
                other => {
                    let bytes = write_format(&sheet, other, ',', false, 0, None, None)?;
                    std::fs::write(tmp, bytes).map_err(|e| CliError::io(e.to_string()))
                }
            })?;
        }
    }
    Ok(())
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
