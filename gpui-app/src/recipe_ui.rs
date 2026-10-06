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
use crate::mode::Mode;
use crate::history::MutationSource;

/// Largest refresh (cells before plus after) kept as sparse cell history;
/// the guarded batch refuses past 100,000 changed cells.
const SPARSE_REFRESH_CELLS: usize = 90_000;

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

/// What the user is asked to approve: a recipe about to read a source it
/// has not read before on this computer.
pub struct RecipeConfirm {
    pub then: ConfirmThen,
    pub recipe_path: PathBuf,
    pub recipe: Recipe,
    /// (the file, size and modified), or why it cannot be read.
    pub source: Result<(String, String), String>,
}

/// What approving does next.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ConfirmThen {
    Run(RecipeTarget),
    /// Open the builder, which previews the source.
    Edit(Option<TableId>),
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
    // One file, or every file an appending recipe matches, read once
    let snapshot = match snapshot {
        Some(s) => s,
        None => match recipe.read_snapshot(dir, None) {
            Ok(s) => s,
            Err(e) => {
                // Unreadable source: a failed run with nothing to retry against
                let report = RunReport::unreadable(&recipe.source_path(dir, None), e);
                return Ok(RunOutcome { recipe, snapshot: None, output: RecipeOutput::empty(), report, source_path: dir.to_path_buf() });
            }
        },
    };
    // What the stamp says it read: the file, or the pattern it appended
    let source_path = if snapshot.more.is_empty() { snapshot.path.clone() } else { recipe.source_path(dir, None) };
    let result = recipe::run(&recipe, &snapshot);
    Ok(RunOutcome { recipe, snapshot: Some(snapshot), output: result.output, report: result.report, source_path })
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
    // A VisiBooks source is named by server/entity/report, not a file:
    // "visibooks:https://api.visiapi.com/42/trial_balance" reads as
    // "VisiBooks 42 trial balance"
    if let Some(rest) = path.strip_prefix("visibooks:") {
        let mut parts = rest.rsplitn(3, '/');
        let (report, entity) = (parts.next().unwrap_or(""), parts.next().unwrap_or(""));
        return format!("VisiBooks {entity} {}", report.replace('_', " "));
    }
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

    /// The recipe a Table links to, as a path: a relative link (how a saved
    /// workbook stores a recipe beside it) resolves against the workbook's
    /// folder.
    pub(crate) fn recipe_link_path(&self, recipe: &str) -> PathBuf {
        let path = Path::new(recipe);
        match self.current_file.as_ref().and_then(|f| f.parent()) {
            // A rooted path without a drive (Windows `/data/x`) is not relative
            Some(dir) if path.is_relative() && !path.has_root() => dir.join(path),
            _ => path.to_path_buf(),
        }
    }

    /// Alt+F5 on a recipe-backed Table, the strip's Refresh button, and the
    /// palette command. Returns false when there is no linked Table here.
    pub fn refresh_recipe_table(&mut self, cx: &mut Context<Self>) -> bool {
        let Some(table) = self.recipe_strip_table(cx) else { return false };
        if self.block_if_previewing(cx) || self.block_read_only_recovery(cx) {
            return true;
        }
        let source = table.source.clone().unwrap();
        let recipe_path = self.recipe_link_path(&source.recipe);
        self.start_recipe_run(RecipeTarget::Table(table.id), recipe_path, None, None, cx);
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
        // A fresh read of the source: ask first unless this recipe and source
        // were approved before. (A run with a snapshot re-uses a source the
        // user already approved or chose.)
        let recipe = match (recipe, &snapshot) {
            (recipe, Some(_)) => recipe,
            (recipe, None) => {
                let recipe = match recipe.map(Ok).unwrap_or_else(|| Recipe::load(&recipe_path)) {
                    Ok(r) => r,
                    Err(e) => {
                        self.status_message = Some(format!("Couldn't open the recipe: {e}"));
                        cx.notify();
                        return;
                    }
                };
                if !crate::recipe_trust::is_approved(&recipe_path, &recipe) {
                    self.ask_to_approve(ConfirmThen::Run(target), recipe_path, recipe, cx);
                    return;
                }
                Some(recipe)
            }
        };
        self.recipe_run_in_progress = true;
        // Opening replaces the window's workbook: note its revision, so edits
        // made while the recipe runs are not thrown away
        self.recipe_open_revision = (target == RecipeTarget::NewWorkbook).then(|| self.wb(cx).revision());
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
            RecipeTarget::NewWorkbook if self.recipe_open_revision.take().is_some_and(|r| r != self.wb(cx).revision()) && self.is_dirty() => {
                self.status_message = Some(format!(
                    "You edited this workbook while {} ran, so it was not replaced. Save your work, then open the recipe again.",
                    file_name(&recipe_path.display().to_string())
                ));
            }
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
                let width = output.columns.len().max(table.as_ref().map_or(0, |t| t.columns.len()));
                let refreshed = recipe_table::refresh_table(&mut candidate, id, &output, link);
                match refreshed {
                    Ok(r) => {
                        let description = format!("Refresh {table_name}");
                        // Sparse cell history while it stays small (most monthly
                        // files); past its limit, one whole-workbook undo step
                        // rather than refusing the refresh
                        let changed = (r.rows_before + r.rows_after + 1) * width;
                        let published = if changed <= SPARSE_REFRESH_CELLS {
                            self.wb(cx)
                                .capture_guarded_batch(&candidate)
                                .and_then(|commit| self.publish_table_batch(candidate, commit, description, MutationSource::Human, cx))
                        } else {
                            self.publish_workbook_snapshot(candidate, description, cx)
                        };
                        match published {
                            Ok(()) => {
                                self.recipe_blocked = None;
                                let delta = r.rows_after as i64 - r.rows_before as i64;
                                let since = match delta {
                                    0 => "same count as before".to_string(),
                                    d if d > 0 => format!("+{d}"),
                                    d => format!("{d}"),
                                };
                                self.status_message = Some(format!(
                                    "Refreshed {table_name}: {} rows ({since}) from {} · {}{}",
                                    r.rows_after,
                                    file_name(&report.source),
                                    match report.steps.len() {
                                        0 => "no steps".to_string(),
                                        1 => "its step ran".to_string(),
                                        n => format!("all {n} steps ran"),
                                    },
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
        // A VisiBooks report isn't a file: what it reads changes in the recipe
        if Recipe::load(&recipe_path).is_ok_and(|r| r.source.is_remote()) {
            self.status_message = Some("This recipe reads a VisiBooks report, not a file; Edit recipe changes which report or entity".into());
            cx.notify();
            return;
        }
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
                    recipe.source.set_path(match chosen.parent() {
                        Some(p) if p == dir => chosen.file_name().unwrap().to_string_lossy().into_owned(),
                        _ => chosen.display().to_string(),
                    });
                    if let Err(e) = recipe.save(&recipe_path) {
                        this.status_message = Some(format!("Couldn't save the recipe: {e}"));
                        cx.notify();
                        return;
                    }
                }
                // The user picked this file: that is the approval
                if let Err(e) = crate::recipe_trust::approve(&recipe_path, &recipe) {
                    this.status_message = Some(format!("Couldn't remember the approval: {e}"));
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

    pub(crate) fn ask_to_approve(&mut self, then: ConfirmThen, recipe_path: PathBuf, recipe: Recipe, cx: &mut Context<Self>) {
        let source = crate::recipe_trust::describe_source(&recipe_path, &recipe);
        self.recipe_blocked = None;
        self.recipe_confirm = Some(RecipeConfirm { then, recipe_path, recipe, source });
        cx.notify();
    }

    /// "Load" on the confirmation (Enter): remember the approval, carry on.
    pub fn approve_recipe_source(&mut self, cx: &mut Context<Self>) {
        let Some(c) = self.recipe_confirm.take() else { return };
        if c.source.is_err() {
            self.recipe_confirm = Some(c);
            return;
        }
        if let Err(e) = crate::recipe_trust::approve(&c.recipe_path, &c.recipe) {
            self.status_message = Some(format!("Couldn't remember the approval: {e}"));
        }
        match c.then {
            ConfirmThen::Run(target) => self.start_recipe_run(target, c.recipe_path, Some(c.recipe), None, cx),
            ConfirmThen::Edit(link) => self.open_recipe_builder(&c.recipe_path, link, cx),
        }
        cx.notify();
    }

    /// While the confirmation is open it takes every key: Ctrl+Enter loads,
    /// Esc cancels, nothing else reaches the grid. Plain Enter does not load,
    /// so a keypress meant for the sheet cannot approve reading a file.
    /// With the blocked banner showing, Ctrl+Enter takes its suggested
    /// rename; other keys pass through.
    pub(crate) fn intercept_recipe_confirm_keys(window: &mut Window, cx: &mut Context<Self>) -> Subscription {
        let this = cx.entity().downgrade();
        let handle = window.window_handle();
        cx.intercept_keystrokes(move |event, window, cx| {
            if window.window_handle() != handle {
                return;
            }
            let Some(this) = this.upgrade() else { return };
            let handled = this.update(cx, |this, cx| {
                let k = &event.keystroke;
                let ctrl_enter = k.key == "enter" && (k.modifiers.control || k.modifiers.platform);
                // The blocked banner's suggested rename is a guess: it takes
                // Ctrl+Enter, never the Enter that moves down the sheet
                if this.recipe_confirm.is_none() {
                    return ctrl_enter && this.mode == Mode::Navigation && this.recipe_primary_fix(cx);
                }
                let k = &event.keystroke;
                if k.key == "escape" {
                    this.cancel_recipe_confirm(cx);
                } else if k.key == "enter" && (k.modifiers.control || k.modifiers.platform) {
                    this.approve_recipe_source(cx);
                }
                true
            });
            if handled {
                cx.stop_propagation();
            }
        })
    }

    pub fn cancel_recipe_confirm(&mut self, cx: &mut Context<Self>) {
        if self.recipe_confirm.take().is_some() {
            self.status_message = Some("Nothing was loaded".into());
            cx.notify();
        }
    }

    /// Why an agent may not change the workbook right now, as (code, message).
    fn agent_refresh_blocker(&self, cx: &App) -> Option<(String, String)> {
        if self.review_mode.is_some() {
            return Some(crate::session_adapter::plan_under_review_error());
        }
        if self.recovery_warning.is_some() {
            return Some(("read_only_recovery".into(), "this window is a read-only recovery; it can't be changed".into()));
        }
        if self.is_previewing() || self.mode.is_editing() {
            return Some(("busy".into(), "the user is previewing history or editing a cell; try again when they're done".into()));
        }
        if crate::table_filter_ui::has_table_criteria(self.workbook.read(cx)) {
            return Some(("table_view_active".into(), crate::table_filter_ui::TABLE_VIEW_EDIT_MESSAGE.into()));
        }
        None
    }

    /// A refresh asked for over the session protocol (MCP `refresh_table`):
    /// the same run and publish as Alt+F5, with the outcome sent to `reply`
    /// when it is done. Never asks: a recipe source the user hasn't approved
    /// in the app is refused, so an agent can't make the app read a file the
    /// user never agreed to. What blocks Alt+F5 blocks an agent too: never
    /// under a plan the user is reviewing, a read-only recovery, a rewind
    /// preview or a filtered Table view. The recipe runs in the background,
    /// like Alt+F5, so the window stays responsive while a big file is read.
    pub(crate) fn start_agent_refresh(
        &mut self,
        table: Option<String>,
        client: Option<String>,
        reply: crate::session_server::bridge::oneshot::Sender<crate::session_server::StructureOutcome>,
        cx: &mut Context<Self>,
    ) {
        let (target, recipe_path, recipe) = match self.prepare_agent_refresh(table.as_deref(), cx) {
            Ok(prepared) => prepared,
            Err(error) => {
                let _ = reply.send(self.structure_outcome(Err(error), cx));
                return;
            }
        };
        self.recipe_run_in_progress = true;
        self.status_message = Some(format!("Refreshing {} for {}…", target.name, client.as_deref().unwrap_or("an agent")));
        cx.notify();
        let started = Instant::now();
        let job_path = recipe_path.clone();
        cx.spawn(async move |this, cx| {
            let outcome = cx.background_executor().spawn(async move { run_job(&job_path, Some(recipe), None) }).await;
            let _ = this.update(cx, |this, cx| {
                this.recipe_run_in_progress = false;
                let result = this.finish_agent_refresh(&target, recipe_path, outcome, started, client, cx);
                let _ = reply.send(this.structure_outcome(result, cx));
                cx.notify();
            });
        })
        .detach();
    }

    /// Everything checked before an agent's refresh reads anything: the
    /// window allows edits, the Table exists and is linked, nothing else is
    /// running, and the user approved what the recipe reads.
    fn prepare_agent_refresh(&mut self, table: Option<&str>, cx: &mut Context<Self>) -> Result<(DataTable, PathBuf, Recipe), (String, String)> {
        if let Some(blocked) = self.agent_refresh_blocker(cx) {
            return Err(blocked);
        }
        let linked: Vec<DataTable> = self.wb(cx).tables().filter(|(_, t)| t.source.is_some()).map(|(_, t)| t.clone()).collect();
        let names = || linked.iter().map(|t| t.name.clone()).collect::<Vec<_>>().join(", ");
        let target = match table {
            Some(name) => linked.iter().find(|t| t.name.eq_ignore_ascii_case(name)).cloned().ok_or_else(|| {
                ("table_not_found".to_string(), format!("no recipe-linked Table named {name}; linked Tables: {}", if linked.is_empty() { "none".into() } else { names() }))
            })?,
            None => match linked.len() {
                1 => linked[0].clone(),
                0 => return Err(("table_not_found".into(), "this workbook has no recipe-linked Table".into())),
                _ => return Err(("ambiguous".into(), format!("name the Table to refresh: {}", names()))),
            },
        };
        if self.recipe_run_in_progress {
            return Err(("busy".into(), "a recipe is already running in this window".into()));
        }
        let recipe_path = self.recipe_link_path(&target.source.as_ref().unwrap().recipe);
        let recipe = Recipe::load(&recipe_path).map_err(|e| ("recipe_invalid".to_string(), e))?;
        if !crate::recipe_trust::is_approved(&recipe_path, &recipe) {
            return Err((
                "needs_approval".into(),
                format!(
                    "the user hasn't approved what {} reads; ask them to refresh {} once in VisiGrid (Alt+F5) and confirm the file",
                    file_name(&recipe_path.display().to_string()),
                    target.name
                ),
            ));
        }
        Ok((target, recipe_path, recipe))
    }

    /// Publish an agent's finished run, unless the window changed while it
    /// ran: the user may have opened a plan to review, the Table may be gone.
    fn finish_agent_refresh(
        &mut self,
        target: &DataTable,
        recipe_path: PathBuf,
        outcome: Result<RunOutcome, String>,
        started: Instant,
        client: Option<String>,
        cx: &mut Context<Self>,
    ) -> Result<String, (String, String)> {
        let outcome = outcome.map_err(|e| ("recipe_failed".to_string(), e))?;
        if let Some((code, message)) = self.agent_refresh_blocker(cx) {
            self.status_message = Some(format!("{} was not refreshed: the window changed while its recipe ran", target.name));
            return Err((code, format!("{message} (the recipe ran, but nothing was changed)")));
        }
        if self.wb(cx).table(target.id).is_none() {
            return Err(("table_not_found".into(), format!("{} was removed while its recipe ran; nothing changed", target.name)));
        }
        let summary = outcome.report.summary();
        let revision = self.wb(cx).revision();
        self.finish_recipe_run(RecipeTarget::Table(target.id), recipe_path, outcome, started, cx);
        // A refresh that didn't publish leaves the banner up for the user,
        // and the Table as it was
        if let Some(b) = self.recipe_blocked.as_ref().filter(|b| b.target == RecipeTarget::Table(target.id)) {
            return Err((
                "recipe_blocked".into(),
                match &b.refused {
                    Some(reason) => format!("{} was not refreshed: {reason}", target.name),
                    None => format!("{} was not refreshed; it keeps its last good result.\n{summary}", target.name),
                },
            ));
        }
        if let (Some(client), true) = (client, self.wb(cx).revision() != revision) {
            self.history.retag_last_source(crate::history::MutationSource::Agent { client });
        }
        Ok(self.status_message.clone().unwrap_or_else(|| format!("Refreshed {}", target.name)))
    }

    /// A session reply for an agent's refresh, with the workbook as it is now.
    fn structure_outcome(&self, result: Result<String, (String, String)>, cx: &App) -> crate::session_server::StructureOutcome {
        let wb = self.workbook.read(cx);
        let (description, error) = match result {
            Ok(d) => (d, None),
            Err(e) => (String::new(), Some(e)),
        };
        crate::session_server::StructureOutcome {
            description,
            revision: wb.revision(),
            sheet_count: wb.sheets().len(),
            active_sheet: wb.active_sheet_index(),
            error,
        }
    }

    /// Palette "Unlink Table from Recipe": the Table keeps its records and
    /// becomes an ordinary Table. One undo step relinks it.
    pub fn unlink_recipe_table(&mut self, cx: &mut Context<Self>) {
        let Some(table) = self.recipe_strip_table(cx) else {
            self.status_message = Some("No recipe-backed Table on this sheet".into());
            cx.notify();
            return;
        };
        if self.block_if_previewing(cx) || self.block_read_only_recovery(cx) {
            return;
        }
        let recipe = table.source.as_ref().map(|s| file_name(&s.recipe)).unwrap_or_default();
        // Metadata only: a Table commit, like changing its banding, not a
        // cell-by-cell comparison of the whole sheet
        let result = self.workbook.update(cx, |wb, _| wb.set_table_source(table.id, None));
        self.status_message = Some(match result {
            Ok(commit) => {
                self.record_table_commit(commit, format!("Unlink {} from its recipe", table.name), cx);
                if self.recipe_blocked.as_ref().is_some_and(|b| b.target == RecipeTarget::Table(table.id)) {
                    self.recipe_blocked = None;
                }
                format!("{} is no longer linked to {recipe}; its records stay. Ctrl+Z relinks it.", table.name)
            }
            Err(e) => format!("Couldn't unlink {}: {e}", table.name),
        });
        cx.notify();
    }

    /// Open the recipe file in the system's editor for TOML.
    pub fn edit_recipe_file(&mut self, path: &Path, cx: &mut Context<Self>) {
        // Even checking that a network path exists sends credentials on Windows
        if let Err(e) = recipe::check_local(path, "recipe") {
            self.status_message = Some(e);
        } else if path.exists() {
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
