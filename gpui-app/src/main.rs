// Hide console window on Windows
#![windows_subsystem = "windows"]

mod actions;
mod ai;
mod app;
mod autocomplete;
mod clipboard;
mod color_palette;
mod command_palette;
mod default_app;
mod default_app_prompt;
mod dialogs;
mod docs_links;
mod diff;
mod editing;
mod file_ops;
mod fill;
mod find_replace;
mod formatting;
mod formula_context;
mod formula_refs;
mod grid_ops;
mod hints;
mod history;
mod hub;
mod impact_preview;
mod keyboard_hints;
mod hub_sync;
mod keybindings;
mod text_editing;
mod links;
#[cfg(target_os = "macos")]
mod menus;
mod menu_model;
mod minimap;
mod user_keybindings;
mod mode;
mod navigation;
#[cfg(feature = "pro")]
mod scripting;
#[cfg(not(feature = "pro"))]
mod scripting_stub;
#[cfg(not(feature = "pro"))]
use scripting_stub as scripting;
mod named_ranges;
mod perf;
mod provenance;
mod ref_target;
mod role_styles;
mod search;
mod series_fill;
mod session;
mod session_server;
mod settings;
mod sheet_ops;
mod sort_filter;
mod split_view;
mod theme;
mod trace;
mod ui;
mod undo_redo;
mod validation_dropdown;
mod views;
mod window_registry;
mod window_switcher;
pub mod workbook_view;

#[cfg(test)]
mod tests;

use std::env;
use std::sync::{Arc, Mutex};

use gpui::*;
use app::Spreadsheet;
use session::{SessionManager, dump_session, reset_session};
use settings::init_settings_store;

// =============================================================================
// Window Geometry - Zed/Figma-style window placement
// =============================================================================

/// Minimum window size to ensure UI remains usable
const MIN_WINDOW_SIZE: Size<Pixels> = Size {
    width: px(1000.0),
    height: px(700.0),
};

/// Fallback window size when work area cannot be determined
const FALLBACK_WINDOW_SIZE: Size<Pixels> = Size {
    width: px(1400.0),
    height: px(900.0),
};

/// Margin from screen edges for default window placement
const WINDOW_MARGIN: f32 = 24.0;

/// Get the work area (visible bounds excluding dock/menubar) for the primary display.
/// Returns None if no displays are available.
fn get_work_area(cx: &App) -> Option<Bounds<Pixels>> {
    let displays = cx.displays();
    displays.first().map(|d| d.visible_bounds())
}

/// Create a near-maximized window rect within the work area with margin.
fn default_large_rect(work_area: Bounds<Pixels>) -> Bounds<Pixels> {
    let margin = px(WINDOW_MARGIN);
    let margin2 = px(WINDOW_MARGIN * 2.0);

    Bounds {
        origin: Point::new(
            work_area.origin.x + margin,
            work_area.origin.y + margin,
        ),
        size: Size {
            width: (work_area.size.width - margin2).max(MIN_WINDOW_SIZE.width),
            height: (work_area.size.height - margin2).max(MIN_WINDOW_SIZE.height),
        },
    }
}

/// Clamp a window rect to fit within the work area.
/// - Shrinks if larger than work area (respecting min size)
/// - Moves if partially off-screen
/// Returns the adjusted rect.
fn clamp_rect_to_work_area(rect: Bounds<Pixels>, work_area: Bounds<Pixels>) -> Bounds<Pixels> {
    let margin = px(WINDOW_MARGIN);
    let margin2 = px(WINDOW_MARGIN * 2.0);

    // Effective work area with margin
    let effective_area = Bounds {
        origin: Point::new(
            work_area.origin.x + margin,
            work_area.origin.y + margin,
        ),
        size: Size {
            width: (work_area.size.width - margin2).max(MIN_WINDOW_SIZE.width),
            height: (work_area.size.height - margin2).max(MIN_WINDOW_SIZE.height),
        },
    };

    // Clamp size to fit within effective area (but respect min size)
    let clamped_width = rect.size.width
        .min(effective_area.size.width)
        .max(MIN_WINDOW_SIZE.width);
    let clamped_height = rect.size.height
        .min(effective_area.size.height)
        .max(MIN_WINDOW_SIZE.height);

    // Clamp position to keep window fully visible
    let max_x = effective_area.origin.x + effective_area.size.width - clamped_width;
    let max_y = effective_area.origin.y + effective_area.size.height - clamped_height;

    let clamped_x = rect.origin.x
        .max(effective_area.origin.x)
        .min(max_x);
    let clamped_y = rect.origin.y
        .max(effective_area.origin.y)
        .min(max_y);

    Bounds {
        origin: Point::new(clamped_x, clamped_y),
        size: Size {
            width: clamped_width,
            height: clamped_height,
        },
    }
}

/// Check if a rect is valid (fully contained within work area with margin).
fn is_rect_valid(rect: &Bounds<Pixels>, work_area: &Bounds<Pixels>) -> bool {
    let margin = px(WINDOW_MARGIN);

    let min_x = work_area.origin.x + margin;
    let min_y = work_area.origin.y + margin;
    let max_x = work_area.origin.x + work_area.size.width - margin;
    let max_y = work_area.origin.y + work_area.size.height - margin;

    rect.origin.x >= min_x
        && rect.origin.y >= min_y
        && rect.origin.x + rect.size.width <= max_x
        && rect.origin.y + rect.size.height <= max_y
        && rect.size.width >= MIN_WINDOW_SIZE.width
        && rect.size.height >= MIN_WINDOW_SIZE.height
}

/// Compute the final window bounds for launch.
/// - If restored bounds are valid, use them exactly
/// - If restored bounds are invalid (off-screen), clamp them
/// - If no restored bounds, create a large default window
fn compute_window_bounds(
    restored: Option<Bounds<Pixels>>,
    work_area: Option<Bounds<Pixels>>,
) -> Bounds<Pixels> {
    match (restored, work_area) {
        (Some(restored), Some(work_area)) => {
            if is_rect_valid(&restored, &work_area) {
                // Valid restored bounds - use exactly
                restored
            } else {
                // Invalid - clamp to work area
                clamp_rect_to_work_area(restored, work_area)
            }
        }
        (None, Some(work_area)) => {
            // Fresh start - large window with margin
            default_large_rect(work_area)
        }
        (Some(restored), None) => {
            // No work area info - use restored but ensure min size
            Bounds {
                origin: restored.origin,
                size: Size {
                    width: restored.size.width.max(MIN_WINDOW_SIZE.width),
                    height: restored.size.height.max(MIN_WINDOW_SIZE.height),
                },
            }
        }
        (None, None) => {
            // Fallback - centered default size
            Bounds {
                origin: Point::new(px(100.0), px(100.0)),
                size: FALLBACK_WINDOW_SIZE,
            }
        }
    }
}

/// Build window options with platform-appropriate titlebar styling.
///
/// On macOS: Transparent titlebar that blends with app content, traffic lights
/// positioned inward. The app renders its own draggable title bar area.
/// Never enters macOS fullscreen Space on launch.
///
/// On Windows/Linux: Standard OS chrome.
fn build_window_options(bounds: WindowBounds) -> WindowOptions {
    #[cfg(target_os = "macos")]
    {
        WindowOptions {
            window_bounds: Some(bounds),
            titlebar: Some(TitlebarOptions {
                title: None,
                appears_transparent: true,
                traffic_light_position: Some(point(px(9.0), px(9.0))),
            }),
            window_min_size: Some(MIN_WINDOW_SIZE),
            ..Default::default()
        }
    }

    #[cfg(not(target_os = "macos"))]
    {
        WindowOptions {
            window_bounds: Some(bounds),
            window_min_size: Some(MIN_WINDOW_SIZE),
            ..Default::default()
        }
    }
}

/// CLI flags
struct CliArgs {
    /// Skip session restore this launch (session file preserved)
    no_restore: bool,
    /// Files to open (if specified). Supports multiple: `visigrid a.xlsx b.csv`
    files: Vec<String>,
    /// Enable session server on startup (for CI/automation)
    session_server: bool,
}

fn print_help() {
    eprintln!("VisiGrid - A spreadsheet for power users");
    eprintln!();
    eprintln!("USAGE:");
    eprintln!("    visigrid [OPTIONS] [FILE]");
    eprintln!();
    eprintln!("OPTIONS:");
    eprintln!("    -n, --no-restore     Skip session restore (start fresh, session preserved)");
    eprintln!("    --reset-session      Delete session file and start fresh");
    eprintln!("    --dump-session       Print session JSON to stdout and exit");
    eprintln!("    --session-server     Enable session server on startup (for CI/automation)");
    eprintln!("    -h, --help           Print this help message");
    eprintln!();
    eprintln!("ARGS:");
    eprintln!("    <FILE>...            File(s) to open (.vgrid, .sheet, .xlsx, .xls, .csv, .tsv)");
}

fn parse_args() -> CliArgs {
    let args: Vec<String> = env::args().collect();
    let mut cli = CliArgs {
        no_restore: false,
        files: Vec::new(),
        session_server: false,
    };

    let mut i = 1;
    while i < args.len() {
        match args[i].as_str() {
            "-h" | "--help" => {
                print_help();
                std::process::exit(0);
            }
            "--dump-session" => {
                dump_session();
                std::process::exit(0);
            }
            "--reset-session" => {
                // Explicit deletion - user knows what they're doing
                reset_session();
                eprintln!("Session reset. Starting fresh.");
                cli.no_restore = true;  // Also skip restore (session is gone)
            }
            "--no-restore" | "-n" => {
                // Skip restore this launch, but preserve session file
                cli.no_restore = true;
            }
            "--session-server" => {
                // Enable session server on startup (for CI/automation)
                cli.session_server = true;
            }
            arg if !arg.starts_with('-') => {
                cli.files.push(arg.to_string());
            }
            unknown => {
                eprintln!("Unknown option: {}", unknown);
                eprintln!("Use --help for usage information.");
                std::process::exit(1);
            }
        }
        i += 1;
    }

    cli
}

/// Parse a file:// URL to a PathBuf. Returns None for non-file URLs.
pub(crate) fn url_to_path(url_str: &str) -> Option<std::path::PathBuf> {
    // macOS sends file:///path/to/file.xlsx
    if let Some(path_str) = url_str.strip_prefix("file://") {
        // URL-decode percent-encoded characters (e.g., %20 → space)
        let decoded = percent_decode(path_str);
        Some(std::path::PathBuf::from(decoded))
    } else if !url_str.contains("://") {
        // Plain path (shouldn't happen from macOS, but handle anyway)
        Some(std::path::PathBuf::from(url_str))
    } else {
        None
    }
}

/// Decode percent-encoded URL characters (minimal, no external deps)
pub(crate) fn percent_decode(input: &str) -> String {
    let mut result = Vec::with_capacity(input.len());
    let bytes = input.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let (Some(hi), Some(lo)) = (hex_val(bytes[i + 1]), hex_val(bytes[i + 2])) {
                result.push(hi << 4 | lo);
                i += 3;
                continue;
            }
        }
        result.push(bytes[i]);
        i += 1;
    }
    String::from_utf8(result).unwrap_or_else(|_| input.to_string())
}

fn hex_val(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'a'..=b'f' => Some(b - b'a' + 10),
        b'A'..=b'F' => Some(b - b'A' + 10),
        _ => None,
    }
}

/// Normalize raw URLs into deduplicated canonical paths, preserving order.
#[allow(unused)] // used by tests
/// Handles file:// URLs, percent-encoding, symlinks, and duplicate entries
/// (Finder can send the same file as both file:// and path variants).
pub(crate) fn normalize_and_dedup_urls(urls: Vec<String>) -> Vec<std::path::PathBuf> {
    let mut seen = std::collections::HashSet::new();
    let mut paths = Vec::new();

    for url in urls {
        let Some(path) = url_to_path(&url) else { continue };
        // canonicalize resolves symlinks + normalizes /./foo, //foo, etc.
        let canonical = path.canonicalize().unwrap_or(path);
        if seen.insert(canonical.clone()) {
            paths.push(canonical);
        }
    }

    paths
}

/// Open file URLs in new windows (one window per file).
/// Called from the startup drain and the polling task.
fn open_file_urls(urls: Vec<String>, cx: &mut App) {
    let paths = normalize_and_dedup_urls(urls);
    if paths.is_empty() {
        return;
    }

    let work_area = get_work_area(cx);

    for path in paths {
        if !path.exists() {
            eprintln!("[open-urls] File not found: {}", path.display());
            continue;
        }

        let bounds = compute_window_bounds(None, work_area);

        cx.open_window(
            build_window_options(WindowBounds::Windowed(bounds)),
            move |window, cx| {
                let window_id = cx.update_global::<SessionManager, _>(|mgr, _| mgr.next_window_id());
                let entity = cx.new(|cx| {
                    let mut app = Spreadsheet::new(window, cx);
                    app.session_window_id = window_id;
                    app.load_file(&path, cx);
                    app
                });
                entity.update(cx, |spreadsheet, cx| {
                    spreadsheet.register_with_window_registry(cx);
                });
                entity
            },
        )
        .ok();
    }
}

fn main() {
    let startup_instant = std::time::Instant::now();

    let cli = parse_args();

    // Move cli values out for use in closure
    let no_restore = cli.no_restore;
    let cli_files = cli.files;
    let start_session_server = cli.session_server;

    // Track if session server has been started (only one per app instance)
    let session_server_started: Arc<Mutex<bool>> = Arc::new(Mutex::new(false));

    // Buffer for file URLs from macOS "Open With" / Finder double-click.
    // on_open_urls callback pushes here; polling task inside run() consumes.
    let pending_open_urls: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));

    let app = Application::new();

    // Register handler for files opened from Finder / "Open With" menu
    app.on_open_urls({
        let pending = pending_open_urls.clone();
        move |urls| {
            pending.lock().unwrap().extend(urls);
        }
    });

    app.run(move |cx: &mut App| {
        // Initialize app-level settings store (must be first)
        init_settings_store(cx);

        // Initialize license system
        visigrid_license::init();

        // Initialize session manager
        cx.set_global(SessionManager::new());

        // Initialize window registry for multi-window management
        cx.set_global(window_registry::WindowRegistry::new());

        // Get modifier style preference for keybindings
        let modifier_style = settings::user_settings(cx)
            .navigation
            .modifier_style
            .as_value()
            .copied()
            .unwrap_or_default();

        keybindings::register(cx, modifier_style);
        user_keybindings::register_user_keybindings(cx);
        window_switcher::register_keybindings(cx);

        // Register Alt+letter accelerators based on setting
        // - Enabled (macOS only): Alt opens scoped Command Palette
        // - Disabled: Alt opens dropdown menus (Excel 2003 style)
        #[cfg(target_os = "macos")]
        {
            use settings::AltAccelerators;
            let alt_accel = settings::user_settings(cx)
                .navigation
                .alt_accelerators
                .as_value()
                .copied()
                .unwrap_or_default();

            if alt_accel == AltAccelerators::Enabled {
                keybindings::register_alt_accelerators(cx);
            } else {
                keybindings::register_menu_accelerators(cx);
            }
        }

        // On Windows/Linux, always use menu accelerators (dropdown menus)
        #[cfg(not(target_os = "macos"))]
        keybindings::register_menu_accelerators(cx);

        // Verify no duplicate accelerator keys within menus (debug builds only)
        #[cfg(debug_assertions)]
        crate::menu_model::debug_assert_all_accels();

        // Set up quit handler (save session before quit)
        cx.on_action(|_: &actions::Quit, cx| {
            // Save session before quitting
            cx.update_global::<SessionManager, _>(|mgr, _| {
                mgr.save_now();
            });
            cx.quit();
        });

        // Set up NewWindow handler (Ctrl+N opens a new window, not replace in-place)
        // This must be at App level because we need cx.open_window()
        cx.on_action(move |_: &actions::NewWindow, cx| {
            // Check if this will be the second window (for tip)
            let is_second_window = cx
                .try_global::<window_registry::WindowRegistry>()
                .map(|r| r.count() == 1)
                .unwrap_or(false);

            let work_area = get_work_area(cx);
            let bounds = compute_window_bounds(None, work_area);

            cx.open_window(
                build_window_options(WindowBounds::Windowed(bounds)),
                |window, cx| {
                    let window_id = cx.update_global::<SessionManager, _>(|mgr, _| mgr.next_window_id());
                    let entity = cx.new(|cx| {
                        let mut app = Spreadsheet::new(window, cx);
                        app.session_window_id = window_id;
                        app
                    });
                    entity.update(cx, |spreadsheet, cx| {
                        spreadsheet.register_with_window_registry(cx);

                        // Show tip for window switcher when 2nd window opens
                        if is_second_window {
                            use crate::settings::{user_settings, TipId};
                            if !user_settings(cx).is_tip_dismissed(TipId::WindowSwitcher) {
                                #[cfg(target_os = "macos")]
                                {
                                    spreadsheet.status_message = Some("Tip: Press Cmd+` to switch between windows".into());
                                }
                                #[cfg(not(target_os = "macos"))]
                                {
                                    spreadsheet.status_message = Some("Tip: Press Ctrl+` to switch between windows".into());
                                }
                                // Dismiss tip after showing once
                                settings::update_user_settings(cx, |settings| {
                                    settings.dismiss_tip(TipId::WindowSwitcher);
                                });
                            }
                        }
                    });
                    entity
                },
            )
            .ok(); // Ignore error if window creation fails
        });

        // Set up OpenFile handler at App level so it works even when no windows are open
        // If there's an active window, dispatch to it. Otherwise create a new window first.
        cx.on_action(move |_: &actions::OpenFile, cx| {
            // Try to find an active spreadsheet window to delegate to
            if let Some(window) = cx.active_window() {
                // Dispatch OpenFile to the active window - it will handle the file picker
                window.update(cx, |_, window, cx| {
                    cx.dispatch_action(&actions::OpenFile);
                    let _ = window; // silence unused warning
                }).ok();
            } else {
                // No windows open - create a new window first, then trigger open
                let work_area = get_work_area(cx);
                let bounds = compute_window_bounds(None, work_area);

                if let Ok(window_handle) = cx.open_window(
                    build_window_options(WindowBounds::Windowed(bounds)),
                    |window, cx| {
                        let window_id = cx.update_global::<SessionManager, _>(|mgr, _| mgr.next_window_id());
                        let entity = cx.new(|cx| {
                            let mut app = Spreadsheet::new(window, cx);
                            app.session_window_id = window_id;
                            app
                        });
                        entity.update(cx, |spreadsheet, cx| {
                            spreadsheet.register_with_window_registry(cx);
                        });
                        entity
                    },
                ) {
                    // Now dispatch OpenFile to the newly created window
                    window_handle.update(cx, |_, window, cx| {
                        cx.dispatch_action(&actions::OpenFile);
                        let _ = window; // silence unused warning
                    }).ok();
                }
            }
        });

        // Set up SwitchWindow handler (Cmd+` / Ctrl+` opens window switcher)
        cx.on_action(|_: &actions::SwitchWindow, cx| {
            // Get the currently focused window to pass to the switcher
            if let Some(window) = cx.active_window() {
                window_switcher::open_switcher(cx, window);
            }
        });

        // Set up native macOS menu bar
        #[cfg(target_os = "macos")]
        menus::set_app_menus(cx);

        // Restore session or open fresh window
        let session = SessionManager::global(cx).session().clone();
        let should_restore = !no_restore && !session.windows.is_empty() && cli_files.is_empty();

        // Get work area for smart window placement
        let work_area = get_work_area(cx);

        if should_restore {
            let restored_count = session.windows.len();

            // Restore each window from session with smart bounds clamping
            for window_session in &session.windows {
                let window_session = window_session.clone();
                let session_server_started = session_server_started.clone();

                // Get restored bounds (if any)
                let restored_bounds = window_session.bounds.as_ref().map(|b| b.to_gpui());

                // Compute final bounds with validation and clamping
                let final_bounds = compute_window_bounds(restored_bounds, work_area);

                // Determine window state - never restore to macOS fullscreen Space
                // (user can manually enter fullscreen after launch if desired)
                let window_bounds = if window_session.maximized {
                    WindowBounds::Maximized(final_bounds)
                } else {
                    // Note: We intentionally don't restore fullscreen state to avoid
                    // automatically entering a separate macOS Space on launch
                    WindowBounds::Windowed(final_bounds)
                };

                let _ = cx.open_window(
                    build_window_options(window_bounds),
                    move |window, cx| {
                        // CRITICAL: Use the SAME window_id from session, not a new one.
                        // This ensures close removes the correct entry from session.
                        let window_id = window_session.window_id;

                        let entity = cx.new(|cx| {
                            let mut app = Spreadsheet::new(window, cx);
                            app.startup_instant = Some(startup_instant);
                            app.session_window_id = window_id;

                            // Load file if present and exists
                            if let Some(ref path) = window_session.file {
                                if path.exists() {
                                    app.load_file(path, cx);
                                } else {
                                    // File missing - show status message but don't crash
                                    app.status_message = Some(format!(
                                        "Session file not found: {}",
                                        path.display()
                                    ));
                                }
                            }

                            // Apply session state (scroll, selection, panels)
                            // This is safe even if file didn't load - clamping handles it
                            app.apply(&window_session, cx);

                            // Set "Session restored" status (only if we actually restored windows)
                            if restored_count > 0 && app.status_message.is_none() {
                                app.status_message = Some("Session restored".to_string());
                            }

                            app
                        });
                        // Register with window registry for window switcher
                        entity.update(cx, |spreadsheet, cx| {
                            spreadsheet.register_with_window_registry(cx);

                            // Start session server if requested via CLI (only once per app)
                            if start_session_server {
                                let mut started = session_server_started.lock().unwrap();
                                if !*started {
                                    *started = true;
                                    // Token from env var (test harness passes this)
                                    let token_override = env::var("VISIGRID_SESSION_TOKEN").ok();
                                    if let Err(e) = spreadsheet.start_session_server(
                                        session_server::ServerMode::Apply,
                                        token_override,
                                        cx,
                                    ) {
                                        eprintln!("Failed to start session server: {}", e);
                                    } else if let Some((session_id, port, discovery)) = spreadsheet.session_server_ready_info() {
                                        // Structured READY line for CI parsing
                                        eprintln!("READY session_id={} port={} discovery={}", session_id, port, discovery.display());
                                    }
                                }
                            }
                        });
                        entity
                    },
                );
            }
        } else if !cli_files.is_empty() {
            // Open file(s) from CLI with smart window placement.
            // Supports multi-file: `visigrid a.xlsx b.csv` opens each in its own window.
            for file_path in cli_files {
                let path = std::path::PathBuf::from(&file_path);
                let bounds = compute_window_bounds(None, work_area);
                let session_server_started = session_server_started.clone();

                cx.open_window(
                    build_window_options(WindowBounds::Windowed(bounds)),
                    move |window, cx| {
                        let window_id = cx.update_global::<SessionManager, _>(|mgr, _| mgr.next_window_id());
                        let entity = cx.new(|cx| {
                            let mut app = Spreadsheet::new(window, cx);
                            app.startup_instant = Some(startup_instant);
                            app.session_window_id = window_id;
                            if path.exists() {
                                app.load_file(&path, cx);
                            } else {
                                app.status_message = Some(format!(
                                    "File not found: {}",
                                    path.display()
                                ));
                            }
                            app
                        });
                        // Register with window registry for window switcher
                        entity.update(cx, |spreadsheet, cx| {
                            spreadsheet.register_with_window_registry(cx);

                            // Start session server if requested via CLI (only once per app)
                            if start_session_server {
                                let mut started = session_server_started.lock().unwrap();
                                if !*started {
                                    *started = true;
                                    let token_override = env::var("VISIGRID_SESSION_TOKEN").ok();
                                    if let Err(e) = spreadsheet.start_session_server(
                                        session_server::ServerMode::Apply,
                                        token_override,
                                        cx,
                                    ) {
                                        eprintln!("Failed to start session server: {}", e);
                                    } else if let Some((session_id, port, discovery)) = spreadsheet.session_server_ready_info() {
                                        eprintln!("READY session_id={} port={} discovery={}", session_id, port, discovery.display());
                                    }
                                }
                            }
                        });
                        entity
                    },
                )
                .ok();
            }
        } else {
            // Fresh start - large near-maximized window with margin
            let bounds = compute_window_bounds(None, work_area);
            let session_server_started = session_server_started.clone();

            cx.open_window(
                build_window_options(WindowBounds::Windowed(bounds)),
                move |window, cx| {
                    let window_id = cx.update_global::<SessionManager, _>(|mgr, _| mgr.next_window_id());
                    let entity = cx.new(|cx| {
                        let mut app = Spreadsheet::new(window, cx);
                        app.startup_instant = Some(startup_instant);
                        app.session_window_id = window_id;
                        app
                    });
                    // Register with window registry for window switcher
                    entity.update(cx, |spreadsheet, cx| {
                        spreadsheet.register_with_window_registry(cx);

                        // Start session server if requested via CLI (only once per app)
                        if start_session_server {
                            let mut started = session_server_started.lock().unwrap();
                            if !*started {
                                *started = true;
                                let token_override = env::var("VISIGRID_SESSION_TOKEN").ok();
                                if let Err(e) = spreadsheet.start_session_server(
                                    session_server::ServerMode::Apply,
                                    token_override,
                                    cx,
                                ) {
                                    eprintln!("Failed to start session server: {}", e);
                                } else if let Some((session_id, port, discovery)) = spreadsheet.session_server_ready_info() {
                                    eprintln!("READY session_id={} port={} discovery={}", session_id, port, discovery.display());
                                }
                            }
                        }
                    });
                    entity
                },
            )
            .unwrap();
        }

        // Process any file URLs that arrived during startup
        // (e.g., user double-clicked a file in Finder to launch VisiGrid)
        {
            let urls = std::mem::take(&mut *pending_open_urls.lock().unwrap());
            if !urls.is_empty() {
                open_file_urls(urls, cx);
            }
        }

        // Spawn polling task to handle file URLs from macOS "Open With"
        // when the app is already running. on_open_urls callback pushes
        // URLs into the shared buffer; this task drains and opens them.
        //
        // Adaptive sleep: fast (100ms) right after processing URLs for
        // responsiveness, then backs off to 1s when idle to avoid wasting
        // cycles. Buffer is drained in one shot via swap (minimal lock hold).
        {
            let pending = pending_open_urls;
            cx.spawn(async move |async_cx| {
                let mut idle_count: u32 = 0;
                loop {
                    let sleep_ms = if idle_count > 10 {
                        1000  // idle: check once per second
                    } else if idle_count > 3 {
                        500   // cooling down
                    } else {
                        100   // recently active: stay responsive
                    };
                    async_cx.background_executor()
                        .timer(std::time::Duration::from_millis(sleep_ms))
                        .await;

                    let urls = {
                        let mut buf = pending.lock().unwrap();
                        if buf.is_empty() {
                            idle_count = idle_count.saturating_add(1);
                            continue;
                        }
                        std::mem::take(&mut *buf)
                    };
                    idle_count = 0;

                    let _ = async_cx.update(|cx| {
                        open_file_urls(urls, cx);
                    });
                }
            }).detach();
        }
    });
}
