// Hide console window on Windows
#![windows_subsystem = "windows"]

mod actions;
mod app;
mod autocomplete;
mod clipboard;
mod command_palette;
mod default_app;
mod default_app_prompt;
mod dialogs;
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
mod provenance;
mod ref_target;
mod search;
mod session;
mod settings;
mod sheet_ops;
mod sort_filter;
mod theme;
mod ui;
mod undo_redo;
mod validation_dropdown;
mod views;
pub mod workbook_view;

#[cfg(test)]
mod tests;

use std::env;

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
    /// File to open (if specified)
    file: Option<String>,
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
    eprintln!("    -h, --help           Print this help message");
    eprintln!();
    eprintln!("ARGS:");
    eprintln!("    <FILE>               File to open (.sheet, .csv, .xlsx)");
}

fn parse_args() -> CliArgs {
    let args: Vec<String> = env::args().collect();
    let mut cli = CliArgs {
        no_restore: false,
        file: None,
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
            arg if !arg.starts_with('-') => {
                cli.file = Some(arg.to_string());
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

fn main() {
    let cli = parse_args();

    // Move cli values out for use in closure
    let no_restore = cli.no_restore;
    let cli_file = cli.file;

    Application::new().run(move |cx: &mut App| {
        // Initialize app-level settings store (must be first)
        init_settings_store(cx);

        // Initialize license system
        visigrid_license::init();

        // Initialize session manager
        cx.set_global(SessionManager::new());

        // Get modifier style preference for keybindings
        let modifier_style = settings::user_settings(cx)
            .navigation
            .modifier_style
            .as_value()
            .copied()
            .unwrap_or_default();

        keybindings::register(cx, modifier_style);
        user_keybindings::register_user_keybindings(cx);

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

        // Set up quit handler (save session before quit)
        cx.on_action(|_: &actions::Quit, cx| {
            // Save session before quitting
            cx.update_global::<SessionManager, _>(|mgr, _| {
                mgr.save_now();
            });
            cx.quit();
        });

        // Set up native macOS menu bar
        #[cfg(target_os = "macos")]
        menus::set_app_menus(cx);

        // Restore session or open fresh window
        let session = SessionManager::global(cx).session().clone();
        let should_restore = !no_restore && !session.windows.is_empty() && cli_file.is_none();

        // Get work area for smart window placement
        let work_area = get_work_area(cx);

        if should_restore {
            // Restore each window from session with smart bounds clamping
            for window_session in &session.windows {
                let window_session = window_session.clone();

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
                        cx.new(|cx| {
                            let mut app = Spreadsheet::new(window, cx);

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
                            app.apply(&window_session);

                            app
                        })
                    },
                );
            }
        } else if let Some(file_path) = cli_file {
            // Open specific file from CLI with smart window placement
            let path = std::path::PathBuf::from(&file_path);
            let bounds = compute_window_bounds(None, work_area);

            cx.open_window(
                build_window_options(WindowBounds::Windowed(bounds)),
                move |window, cx| {
                    cx.new(|cx| {
                        let mut app = Spreadsheet::new(window, cx);
                        if path.exists() {
                            app.load_file(&path, cx);
                        } else {
                            app.status_message = Some(format!(
                                "File not found: {}",
                                path.display()
                            ));
                        }
                        app
                    })
                },
            )
            .unwrap();
        } else {
            // Fresh start - large near-maximized window with margin
            let bounds = compute_window_bounds(None, work_area);

            cx.open_window(
                build_window_options(WindowBounds::Windowed(bounds)),
                |window, cx| cx.new(|cx| Spreadsheet::new(window, cx)),
            )
            .unwrap();
        }
    });
}
