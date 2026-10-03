//! Import recipes in the desktop: open a `.recipe.toml` as a linked Table,
//! refresh a linked Table, and the banner shown when a refresh is blocked.
//!
//! The executor and the Table refresh live in `visigrid_io::recipe` and
//! `visigrid_io::recipe_table`. Refresh runs on a candidate copy of the
//! workbook and publishes through the guarded Table batch, so a refused
//! refresh changes nothing and a published one undoes in one step.

use std::path::{Path, PathBuf};
use std::time::Instant;

use gpui::*;
use visigrid_engine::table::{DataTable, RefreshStamp, TableId, TableSource};
use visigrid_io::recipe::{self, OnError, Recipe, RecipeOutput, RunReport, Snapshot};
use visigrid_io::recipe_table;

use crate::app::Spreadsheet;
use crate::history::MutationSource;

/// Height of the strip above a linked Table's sheet.
pub(crate) const RECIPE_STRIP_HEIGHT: f32 = 30.0;

/// Where a run's result goes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum RecipeTarget {
    /// Opening a recipe: a new workbook holding one linked Table.
    NewWorkbook,
    /// Refreshing this Table in place.
    Table(TableId),
}

/// A run that did not publish, kept so a fix can re-run against the same
/// snapshot: what was checked is what gets loaded.
pub struct RecipeBlocked {
    pub target: RecipeTarget,
    pub table_name: String,
    pub recipe_path: PathBuf,
    pub recipe: Recipe,
    /// None when the source could not be read at all.
    pub snapshot: Option<Snapshot>,
    pub report: RunReport,
    /// Why a successful run could not be placed (cells in the way, ...).
    pub refused: Option<String>,
    /// The Table's last good refresh, which it still shows.
    pub last_good: Option<RefreshStamp>,
    /// A fix could not be saved to the recipe file.
    pub save_error: Option<String>,
}

/// What a background run hands back to the window.
struct RunOutcome {
    recipe: Recipe,
    snapshot: Option<Snapshot>,
    output: RecipeOutput,
    report: RunReport,
    source_path: PathBuf,
}

fn run_job(recipe_path: &Path, recipe: Option<Recipe>, snapshot: Option<Snapshot>) -> Result<RunOutcome, String> {
    let recipe = match recipe {
        Some(r) => r,
        None => Recipe::load(recipe_path)?,
    };
    let dir = recipe_path.parent().unwrap_or(Path::new("."));
    let source_path = match &snapshot {
        Some(s) => s.path.clone(),
        None => match recipe.resolve_source(dir, None) {
            Ok(p) => p,
            Err(e) => {
                let report = RunReport::unreadable(&recipe.source_path(dir, None), e);
                return Ok(RunOutcome { recipe, snapshot: None, output: RecipeOutput::empty(), report, source_path: dir.to_path_buf() });
            }
        },
    };
    let snapshot = match snapshot {
        Some(s) => Some(s),
        None => match Snapshot::read(&source_path) {
            Ok(s) => Some(s),
            Err(e) => {
                // Unreadable source: a failed run with nothing to retry against
                let report = RunReport::unreadable(&source_path, e);
                return Ok(RunOutcome { recipe, snapshot: None, output: RecipeOutput::empty(), report, source_path });
            }
        },
    };
    let result = recipe::run(&recipe, snapshot.as_ref().unwrap());
    Ok(RunOutcome { recipe, snapshot, output: result.output, report: result.report, source_path })
}

/// The rename the banner suggests first: a missing column, in the earliest
/// failed step, that the drift check matched to a new column.
pub(crate) fn primary_rename(b: &RecipeBlocked) -> Option<(String, String)> {
    b.report.steps.iter().filter(|s| s.failed).flat_map(|s| s.missing.iter()).find_map(|name| {
        b.report
            .drift
            .possibly_renamed
            .iter()
            .find(|(old, _)| old.eq_ignore_ascii_case(name))
            .map(|(old, new)| (old.clone(), new.clone()))
    })
}

/// "3 minutes ago", "yesterday", "Sep 30".
pub(crate) fn when(stamp: &str) -> String {
    let Ok(at) = chrono::DateTime::parse_from_rfc3339(stamp) else { return stamp.to_string() };
    let at = at.with_timezone(&chrono::Local);
    let secs = (chrono::Local::now() - at).num_seconds();
    match secs {
        s if s < 60 => "just now".into(),
        s if s < 3600 => format!("{} min ago", s / 60),
        s if s < 86_400 && at.date_naive() == chrono::Local::now().date_naive() => at.format("today %H:%M").to_string(),
        _ => at.format("%b %-d").to_string(),
    }
}

pub(crate) fn file_name(path: &str) -> String {
    Path::new(path).file_name().and_then(|n| n.to_str()).unwrap_or(path).to_string()
}

impl Spreadsheet {
    /// The linked Table the strip describes: the one under the cursor, else
    /// the first linked Table on the active sheet.
    pub(crate) fn recipe_strip_table(&self, cx: &App) -> Option<DataTable> {
        if self.zen_mode || self.is_previewing() || self.review_mode.is_some() {
            return None;
        }
        if let Some(t) = self.table_under_cursor(cx).filter(|t| t.source.is_some()) {
            return Some(t);
        }
        self.sheet(cx).tables().iter().find(|t| t.source.is_some()).cloned()
    }

    pub(crate) fn show_recipe_strip(&self, cx: &App) -> bool {
        self.recipe_strip_table(cx).is_some()
    }

    /// Open a `.recipe.toml`: run it, and open the result as a new workbook
    /// with one linked Table. A failed run shows what blocked it instead.
    pub fn open_recipe(&mut self, path: &Path, cx: &mut Context<Self>) {
        let path = std::path::absolute(path).unwrap_or_else(|_| path.to_path_buf());
        self.start_recipe_run(RecipeTarget::NewWorkbook, path, None, None, cx);
    }

    /// Alt+F5 on a recipe-backed Table, the strip's Refresh button, and the
    /// palette command. Returns false when there is no linked Table here.
    pub fn refresh_recipe_table(&mut self, cx: &mut Context<Self>) -> bool {
        let Some(table) = self.recipe_strip_table(cx) else { return false };
        if self.block_if_previewing(cx) || self.block_read_only_recovery(cx) {
            return true;
        }
        let source = table.source.clone().unwrap();
        self.start_recipe_run(RecipeTarget::Table(table.id), PathBuf::from(&source.recipe), None, None, cx);
        true
    }

    fn start_recipe_run(
        &mut self,
        target: RecipeTarget,
        recipe_path: PathBuf,
        recipe: Option<Recipe>,
        snapshot: Option<Snapshot>,
        cx: &mut Context<Self>,
    ) {
        if self.recipe_run_in_progress {
            self.status_message = Some("A recipe is already running".into());
            cx.notify();
            return;
        }
        self.recipe_run_in_progress = true;
        self.status_message = Some(format!("Running {}…", file_name(&recipe_path.display().to_string())));
        cx.notify();
        let started = Instant::now();
        let job_path = recipe_path.clone();
        cx.spawn(async move |this, cx| {
            let outcome = cx.background_executor().spawn(async move { run_job(&job_path, recipe, snapshot) }).await;
            let _ = this.update(cx, |this, cx| {
                this.recipe_run_in_progress = false;
                match outcome {
                    Ok(outcome) => this.finish_recipe_run(target, recipe_path, outcome, started, cx),
                    Err(e) => {
                        this.status_message = Some(format!("Couldn't run the recipe: {e}"));
                        cx.notify();
                    }
                }
            });
        })
        .detach();
    }

    fn finish_recipe_run(&mut self, target: RecipeTarget, recipe_path: PathBuf, outcome: RunOutcome, started: Instant, cx: &mut Context<Self>) {
        let RunOutcome { recipe, snapshot, output, report, source_path } = outcome;
        let table = match target {
            RecipeTarget::Table(id) => match self.wb(cx).table(id) {
                Some((_, t)) => Some(t.clone()),
                None => {
                    self.status_message = Some("The Table was removed while its recipe ran; nothing changed".into());
                    cx.notify();
                    return;
                }
            },
            RecipeTarget::NewWorkbook => None,
        };
        let table_name = table.as_ref().map_or_else(|| recipe_table::table_name_for(&recipe_path), |t| t.name.clone());
        let last_good = table.as_ref().and_then(|t| t.source.as_ref()?.refreshed.clone());
        let blocked = |report: RunReport, refused: Option<String>, recipe: Recipe, snapshot: Option<Snapshot>| RecipeBlocked {
            target,
            table_name: table_name.clone(),
            recipe_path: recipe_path.clone(),
            recipe,
            snapshot,
            report,
            refused,
            last_good: last_good.clone(),
            save_error: None,
        };
        if !report.ok {
            self.recipe_blocked = Some(blocked(report, None, recipe, snapshot));
            self.status_message = Some(match target {
                RecipeTarget::Table(_) => format!("Refresh didn't publish; {table_name} is unchanged"),
                RecipeTarget::NewWorkbook => "The recipe didn't run cleanly; nothing was opened".into(),
            });
            cx.notify();
            return;
        }
        let link = TableSource {
            recipe: recipe_path.display().to_string(),
            refreshed: Some(recipe_table::stamp(&report, &source_path)),
        };
        let ms = started.elapsed().as_millis();
        match target {
            RecipeTarget::NewWorkbook => match recipe_table::new_workbook(&output, &table_name, link) {
                Ok(wb) => {
                    self.recipe_blocked = None;
                    self.csv_doc = None;
                    self.finish_table_import(wb, &recipe_path, cx);
                    self.status_message = Some(format!(
                        "Loaded {} rows into Table {} from {} in {ms} ms · Alt+F5 refreshes it",
                        report.rows,
                        table_name,
                        file_name(&report.source)
                    ));
                }
                Err(e) => self.recipe_blocked = Some(blocked(report, Some(e), recipe, snapshot)),
            },
            RecipeTarget::Table(id) => {
                let mut candidate = self.wb(cx).clone();
                let refreshed = recipe_table::refresh_table(&mut candidate, id, &output, link)
                    .and_then(|r| Ok((r, self.wb(cx).capture_guarded_batch(&candidate)?)));
                match refreshed {
                    Ok((r, commit)) => {
                        let description = format!("Refresh {table_name}");
                        match self.publish_table_batch(candidate, commit, description, MutationSource::Human, cx) {
                            Ok(()) => {
                                self.recipe_blocked = None;
                                let delta = r.rows_after as i64 - r.rows_before as i64;
                                let since = match delta {
                                    0 => "same count as before".to_string(),
                                    d if d > 0 => format!("+{d}"),
                                    d => format!("{d}"),
                                };
                                self.status_message = Some(format!(
                                    "Refreshed {table_name}: {} rows ({since}) from {} · all {} steps ran{}",
                                    r.rows_after,
                                    file_name(&report.source),
                                    report.steps.len(),
                                    if r.columns_changed { " · columns changed" } else { "" }
                                ));
                            }
                            Err(e) => self.recipe_blocked = Some(blocked(report, Some(e), recipe, snapshot)),
                        }
                    }
                    Err(e) => self.recipe_blocked = Some(blocked(report, Some(e), recipe, snapshot)),
                }
            }
        }
        cx.notify();
    }

    // ------------------------------------------------------------------
    // Fixes offered by the blocked banner
    // ------------------------------------------------------------------

    /// "Use Order Number": the source renamed a column. Saved into the
    /// recipe, then re-run against the same snapshot.
    pub fn recipe_fix_rename(&mut self, old: String, new: String, cx: &mut Context<Self>) {
        self.apply_recipe_fix(cx, move |r| {
            if r.rename_source_column(&old, &new) {
                Ok(())
            } else {
                Err(format!("The recipe doesn't use {old}"))
            }
        });
    }

    /// The banner's suggested fix: the first missing column the source
    /// appears to have renamed. Enter takes it; returns whether there was one.
    pub fn recipe_primary_fix(&mut self, cx: &mut Context<Self>) -> bool {
        let Some((old, new)) = self.recipe_blocked.as_ref().and_then(primary_rename) else { return false };
        self.recipe_fix_rename(old, new, cx);
        true
    }

    /// "Keep as text" / "Leave blank" for values that don't fit a type.
    pub fn recipe_fix_on_error(&mut self, step: usize, on_error: OnError, cx: &mut Context<Self>) {
        self.apply_recipe_fix(cx, move |r| r.set_on_error(step, on_error));
    }

    fn apply_recipe_fix(&mut self, cx: &mut Context<Self>, fix: impl FnOnce(&mut Recipe) -> Result<(), String>) {
        let Some(blocked) = self.recipe_blocked.as_mut() else { return };
        let mut recipe = blocked.recipe.clone();
        if let Err(e) = fix(&mut recipe).and_then(|()| recipe.save(&blocked.recipe_path)) {
            blocked.save_error = Some(e);
            cx.notify();
            return;
        }
        let (target, path, snapshot) = (blocked.target, blocked.recipe_path.clone(), blocked.snapshot.clone());
        blocked.recipe = recipe.clone();
        self.start_recipe_run(target, path, Some(recipe), snapshot, cx);
    }

    /// Run the blocked recipe again, reading the source afresh (it may have
    /// been replaced, or edited by hand).
    pub fn recipe_retry(&mut self, cx: &mut Context<Self>) {
        let Some(b) = self.recipe_blocked.as_ref() else { return };
        let (target, path) = (b.target, b.recipe_path.clone());
        self.start_recipe_run(target, path, None, None, cx);
    }

    /// "Choose file…": read another file with this recipe. A file its
    /// pattern already matches is read once without changing the recipe;
    /// any other file becomes the recipe's source (saved), then it runs.
    pub fn recipe_choose_source(&mut self, recipe_path: PathBuf, target: RecipeTarget, cx: &mut Context<Self>) {
        let future = cx.prompt_for_paths(PathPromptOptions {
            files: true,
            directories: false,
            multiple: false,
            prompt: Some("Choose the file to read".into()),
        });
        cx.spawn(async move |this, cx| {
            let Ok(Ok(Some(paths))) = future.await else { return };
            let Some(chosen) = paths.first().cloned() else { return };
            let _ = this.update(cx, |this, cx| {
                let mut recipe = match Recipe::load(&recipe_path) {
                    Ok(r) => r,
                    Err(e) => {
                        this.status_message = Some(format!("Couldn't open the recipe: {e}"));
                        cx.notify();
                        return;
                    }
                };
                let dir = recipe_path.parent().unwrap_or(Path::new(".")).to_path_buf();
                let pattern_path = recipe.source_path(&dir, None);
                let matches_pattern = recipe.source_is_pattern()
                    && chosen.parent() == pattern_path.parent()
                    && chosen.file_name().and_then(|n| n.to_str()).zip(pattern_path.file_name().and_then(|n| n.to_str()))
                        .is_some_and(|(n, p)| recipe::wildcard_match(p, n));
                if !matches_pattern {
                    let visigrid_io::recipe::Source::Csv(src) = &mut recipe.source;
                    src.path = match chosen.parent() {
                        Some(p) if p == dir => chosen.file_name().unwrap().to_string_lossy().into_owned(),
                        _ => chosen.display().to_string(),
                    };
                    if let Err(e) = recipe.save(&recipe_path) {
                        this.status_message = Some(format!("Couldn't save the recipe: {e}"));
                        cx.notify();
                        return;
                    }
                }
                let snapshot = match Snapshot::read(&chosen) {
                    Ok(s) => s,
                    Err(e) => {
                        this.status_message = Some(format!("Couldn't read {}: {e}", chosen.display()));
                        cx.notify();
                        return;
                    }
                };
                this.recipe_blocked = None;
                this.start_recipe_run(target, recipe_path.clone(), Some(recipe), Some(snapshot), cx);
            });
        })
        .detach();
    }

    /// Open the recipe file in the system's editor for TOML.
    pub fn edit_recipe_file(&mut self, path: &Path, cx: &mut Context<Self>) {
        if path.exists() {
            cx.open_with_system(path);
            self.status_message = Some(format!("Opened {} · Alt+F5 refreshes after you save it", file_name(&path.display().to_string())));
        } else {
            self.status_message = Some(format!("{} no longer exists", path.display()));
        }
        cx.notify();
    }

    pub fn dismiss_recipe_blocked(&mut self, cx: &mut Context<Self>) {
        if self.recipe_blocked.take().is_some() {
            cx.notify();
        }
    }
}
