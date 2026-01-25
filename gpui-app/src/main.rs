// Hide console window on Windows
#![windows_subsystem = "windows"]

mod actions;
mod app;
mod autocomplete;
mod clipboard;
mod command_palette;
mod default_app;
mod file_ops;
mod fill;
mod find_replace;
mod formatting;
mod formula_context;
mod hints;
mod history;
mod hub;
mod hub_sync;
mod keybindings;
mod text_editing;
mod links;
#[cfg(target_os = "macos")]
mod menus;
mod user_keybindings;
mod mode;
#[cfg(feature = "pro")]
mod scripting;
#[cfg(not(feature = "pro"))]
mod scripting_stub;
#[cfg(not(feature = "pro"))]
use scripting_stub as scripting;
mod ref_target;
mod search;
mod session;
mod settings;
mod theme;
mod views;
pub mod workbook_view;

#[cfg(test)]
mod tests;

use std::env;

use gpui::*;
use app::Spreadsheet;
use session::{SessionManager, dump_session, reset_session};
use settings::init_settings_store;

/// Build window options with platform-appropriate titlebar styling.
///
/// On macOS: Transparent titlebar that blends with app content, traffic lights
/// positioned inward. The app renders its own draggable title bar area.
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
            ..Default::default()
        }
    }

    #[cfg(not(target_os = "macos"))]
    {
        WindowOptions {
            window_bounds: Some(bounds),
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

        if should_restore {
            // Restore each window from session
            for window_session in &session.windows {
                let window_session = window_session.clone();
                let bounds = window_session.bounds
                    .as_ref()
                    .map(|b| b.to_gpui())
                    .unwrap_or_else(|| Bounds {
                        origin: Point::new(px(100.0), px(100.0)),
                        size: Size {
                            width: px(1200.0),
                            height: px(800.0),
                        },
                    });

                let window_bounds = if window_session.maximized {
                    WindowBounds::Maximized(bounds)
                } else if window_session.fullscreen {
                    WindowBounds::Fullscreen(bounds)
                } else {
                    WindowBounds::Windowed(bounds)
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
            // Open specific file from CLI
            let path = std::path::PathBuf::from(&file_path);
            let bounds = Bounds {
                origin: Point::new(px(100.0), px(100.0)),
                size: Size {
                    width: px(1200.0),
                    height: px(800.0),
                },
            };

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
            // Fresh start - open new empty window
            let bounds = Bounds {
                origin: Point::new(px(100.0), px(100.0)),
                size: Size {
                    width: px(1200.0),
                    height: px(800.0),
                },
            };

            cx.open_window(
                build_window_options(WindowBounds::Windowed(bounds)),
                |window, cx| cx.new(|cx| Spreadsheet::new(window, cx)),
            )
            .unwrap();
        }
    });
}
