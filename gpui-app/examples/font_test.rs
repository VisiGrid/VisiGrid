use gpui::*;

struct FontTest;

impl Render for FontTest {
    fn render(&mut self, _window: &mut Window, _cx: &mut Context<Self>) -> impl IntoElement {
        div()
            .size_full()
            .bg(rgb(0x1e1e1e))
            .flex()
            .items_center()
            .justify_center()
            .text_color(rgb(0xffffff))
            .text_size(px(32.0))
            .child("Hello from gpui! If you can read this, fonts work.")
    }
}

fn main() {
    Application::new().run(|cx: &mut App| {
        let bounds = Bounds::centered(None, size(px(600.), px(400.)), cx);
        cx.open_window(
            WindowOptions {
                window_bounds: Some(WindowBounds::Windowed(bounds)),
                ..Default::default()
            },
            |_, cx| cx.new(|_| FontTest),
        )
        .unwrap();
    });
}
