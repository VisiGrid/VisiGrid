// Hide console window on Windows
#![windows_subsystem = "windows"]

mod actions;
mod app;
mod autocomplete;
mod file_ops;
mod fill;
mod formatting;
mod formula_context;
mod history;
mod keybindings;
#[cfg(target_os = "macos")]
mod menus;
mod mode;
mod search;
mod settings;
mod theme;
mod views;

#[cfg(test)]
mod tests;

use gpui::*;
use app::Spreadsheet;
use settings::init_settings_store;

fn main() {
    Application::new().run(|cx: &mut App| {
        // Initialize app-level settings store (must be first)
        init_settings_store(cx);

        keybindings::register(cx);

        // Set up native macOS menu bar
        #[cfg(target_os = "macos")]
        {
            cx.on_action(|_: &actions::Quit, cx| cx.quit());
            menus::set_app_menus(cx);
        }

        let bounds = Bounds {
            origin: Point::new(px(100.0), px(100.0)),
            size: Size {
                width: px(1200.0),
                height: px(800.0),
            },
        };

        cx.open_window(
            WindowOptions {
                window_bounds: Some(WindowBounds::Windowed(bounds)),
                ..Default::default()
            },
            |window, cx| cx.new(|cx| Spreadsheet::new(window, cx)),
        )
        .unwrap();
    });
}
