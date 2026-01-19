mod actions;
mod app;
mod file_ops;
mod history;
mod keybindings;
mod mode;
mod views;

use gpui::*;
use app::Spreadsheet;

fn main() {
    Application::new().run(|cx: &mut App| {
        keybindings::register(cx);

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
