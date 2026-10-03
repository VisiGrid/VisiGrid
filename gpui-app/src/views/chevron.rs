//! Font-independent chevrons shared by compact controls.
use gpui::*;

pub(super) fn render(color: Hsla, zoom: f32, up: bool) -> impl IntoElement {
    canvas(
        |_, _, _| (),
        move |bounds, _, window, _| {
            let p = |x, y| bounds.origin + point(px(x * zoom), px(y * zoom));
            let (shoulder, tip) = if up { (7.5, 4.0) } else { (4.5, 8.0) };
            let mut path = PathBuilder::stroke(px(1.4 * zoom));
            path.move_to(p(2.5, shoulder));
            path.line_to(p(6.0, tip));
            path.line_to(p(9.5, shoulder));
            if let Ok(path) = path.build() {
                window.paint_path(path, color);
            }
        },
    )
    .size(px(12.0 * zoom))
    .flex_shrink_0()
}
