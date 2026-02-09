//! Window switcher palette for navigating between open windows
//!
//! A lightweight modal overlay that shows all open VisiGrid windows,
//! allowing keyboard-first navigation with type-to-filter.

use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::window_registry::{WindowInfo, WindowRegistry};

actions!(window_switcher, [
    SwitcherUp,
    SwitcherDown,
    SwitcherSelect,
    SwitcherCancel,
]);

/// State for the window switcher overlay
pub struct WindowSwitcher {
    /// Filter query for type-to-filter
    query: String,
    /// Currently selected index in the filtered list
    selected_index: usize,
    /// Focus handle for keyboard input
    focus_handle: FocusHandle,
    /// Cached list of windows (snapshot when opened)
    windows: Vec<WindowInfo>,
    /// Current window handle (to exclude from list or highlight)
    current_window: AnyWindowHandle,
}

impl WindowSwitcher {
    pub fn new(window: &mut Window, cx: &mut Context<Self>, current_window: AnyWindowHandle) -> Self {
        let focus_handle = cx.focus_handle();
        window.focus(&focus_handle, cx);

        // Snapshot the window list from global registry
        let windows = if let Some(registry) = cx.try_global::<WindowRegistry>() {
            registry.windows().to_vec()
        } else {
            vec![]
        };

        Self {
            query: String::new(),
            selected_index: 0,
            focus_handle,
            windows,
            current_window,
        }
    }

    /// Get filtered windows based on current query
    fn filtered_windows(&self) -> Vec<&WindowInfo> {
        let query_lower = self.query.to_lowercase();
        self.windows
            .iter()
            .filter(|w| {
                if query_lower.is_empty() {
                    true
                } else {
                    w.title.to_lowercase().contains(&query_lower)
                        || w.path
                            .as_ref()
                            .and_then(|p| p.to_str())
                            .map(|s| s.to_lowercase().contains(&query_lower))
                            .unwrap_or(false)
                }
            })
            .collect()
    }

    fn move_up(&mut self, cx: &mut Context<Self>) {
        if self.selected_index > 0 {
            self.selected_index -= 1;
            cx.notify();
        }
    }

    fn move_down(&mut self, cx: &mut Context<Self>) {
        let count = self.filtered_windows().len();
        if count > 0 && self.selected_index < count - 1 {
            self.selected_index += 1;
            cx.notify();
        }
    }

    fn select(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let filtered = self.filtered_windows();
        if let Some(info) = filtered.get(self.selected_index) {
            let handle = info.handle;
            let title = info.title.clone();

            // Close the switcher first
            window.remove_window();

            // Activate the selected window and focus the grid
            let _ = handle.update(cx, |root, target_window, cx| {
                // Activate the window (bring to front)
                target_window.activate_window();

                // Focus the main spreadsheet grid
                if let Ok(spreadsheet) = root.downcast::<crate::app::Spreadsheet>() {
                    spreadsheet.update(cx, |app, cx| {
                        // Focus the grid handle
                        target_window.focus(&app.focus_handle, cx);

                        // Show confirmation message
                        app.status_message = Some(format!("Switched to {}", title));
                        cx.notify();
                    });
                }
            });
        }
    }

    fn cancel(&mut self, window: &mut Window, _cx: &mut Context<Self>) {
        window.remove_window();
    }

    fn insert_char(&mut self, ch: &str, cx: &mut Context<Self>) {
        self.query.push_str(ch);
        // Reset selection when query changes
        self.selected_index = 0;
        cx.notify();
    }

    fn backspace(&mut self, cx: &mut Context<Self>) {
        self.query.pop();
        self.selected_index = 0;
        cx.notify();
    }
}

impl Render for WindowSwitcher {
    fn render(&mut self, _window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let filtered = self.filtered_windows();
        let selected_idx = self.selected_index;
        let current_handle = self.current_window;
        let query_str = self.query.clone();
        let has_query = !query_str.is_empty();
        let is_empty = self.windows.is_empty();

        // Theme colors
        let bg = rgb(0x252526);
        let border_color = rgb(0x3d3d3d);
        let selected_bg = rgb(0x094771);
        let text_color = rgb(0xcccccc);
        let dim_color = rgb(0x808080);
        let dirty_color = rgb(0xdcdcaa);

        div()
            .id("window-switcher-backdrop")
            .size_full()
            .flex()
            .items_center()
            .justify_center()
            .bg(rgba(0x00000080))
            .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
                this.cancel(window, cx);
            }))
            .child(
                div()
                    .id("window-switcher-panel")
                    .key_context("WindowSwitcher")
                    .track_focus(&self.focus_handle)
                    .on_action(cx.listener(|this, _: &SwitcherUp, _, cx| {
                        this.move_up(cx);
                    }))
                    .on_action(cx.listener(|this, _: &SwitcherDown, _, cx| {
                        this.move_down(cx);
                    }))
                    .on_action(cx.listener(|this, _: &SwitcherSelect, window, cx| {
                        this.select(window, cx);
                    }))
                    .on_action(cx.listener(|this, _: &SwitcherCancel, window, cx| {
                        this.cancel(window, cx);
                    }))
                    .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                        if event.keystroke.key == "backspace" {
                            this.backspace(cx);
                        } else if let Some(ch) = &event.keystroke.key_char {
                            if !event.keystroke.modifiers.control
                                && !event.keystroke.modifiers.platform
                            {
                                this.insert_char(ch, cx);
                            }
                        }
                    }))
                    .on_mouse_down(MouseButton::Left, |_, _, _| {
                        // Prevent clicks inside panel from closing
                    })
                    .w(px(400.0))
                    .max_h(px(300.0))
                    .bg(bg)
                    .border_1()
                    .border_color(border_color)
                    .rounded_md()
                    .shadow_lg()
                    .overflow_hidden()
                    .flex()
                    .flex_col()
                    .child(
                        // Header with search input
                        div()
                            .px_3()
                            .py_2()
                            .border_b_1()
                            .border_color(border_color)
                            .child(
                                div()
                                    .flex()
                                    .items_center()
                                    .gap_2()
                                    .child(
                                        div()
                                            .text_sm()
                                            .text_color(dim_color)
                                            .child("Switch Window")
                                    )
                                    .when(has_query, |d| {
                                        d.child(
                                            div()
                                                .px_2()
                                                .py_1()
                                                .bg(rgb(0x3c3c3c))
                                                .rounded_sm()
                                                .text_sm()
                                                .text_color(text_color)
                                                .child(query_str.clone())
                                        )
                                    })
                            )
                    )
                    .child(
                        // Window list
                        div()
                            .flex_1()
                            .overflow_hidden()
                            .children(
                                filtered.iter().enumerate().map(|(idx, info)| {
                                    let is_selected = idx == selected_idx;
                                    let is_current = info.handle == current_handle;
                                    let title = info.title.clone();
                                    let is_dirty = info.is_dirty;
                                    let subtitle = info.subtitle();

                                    div()
                                        .id(ElementId::NamedInteger("window-item".into(), idx as u64))
                                        .px_3()
                                        .py_2()
                                        .cursor_pointer()
                                        .when(is_selected, |d| d.bg(selected_bg))
                                        .hover(|style| style.bg(rgb(0x2a2d2e)))
                                        .on_mouse_down(MouseButton::Left, {
                                            let handle = info.handle;
                                            cx.listener(move |this, _, window, cx| {
                                                // Find and select this window
                                                let filtered = this.filtered_windows();
                                                if let Some(pos) = filtered.iter().position(|w| w.handle == handle) {
                                                    this.selected_index = pos;
                                                    this.select(window, cx);
                                                }
                                            })
                                        })
                                        .child(
                                            div()
                                                .flex()
                                                .items_center()
                                                .gap_2()
                                                .child(
                                                    div()
                                                        .text_sm()
                                                        .text_color(text_color)
                                                        .when(is_current, |d| d.font_weight(FontWeight::BOLD))
                                                        .child(title)
                                                )
                                                .when(is_dirty, |d| {
                                                    d.child(
                                                        div()
                                                            .text_sm()
                                                            .text_color(dirty_color)
                                                            .child("●")
                                                    )
                                                })
                                        )
                                        .when_some(subtitle, |d, sub| {
                                            d.child(
                                                div()
                                                    .text_xs()
                                                    .text_color(dim_color)
                                                    .child(sub)
                                            )
                                        })
                                })
                            )
                            .when(filtered.is_empty(), |d: Div| {
                                d.child(
                                    div()
                                        .px_3()
                                        .py_4()
                                        .text_sm()
                                        .text_color(dim_color)
                                        .child(if is_empty {
                                            "No other windows open"
                                        } else {
                                            "No matching windows"
                                        })
                                )
                            })
                    )
                    .child(
                        // Footer with hints
                        div()
                            .px_3()
                            .py_1()
                            .border_t_1()
                            .border_color(border_color)
                            .flex()
                            .gap_3()
                            .child(
                                div()
                                    .text_xs()
                                    .text_color(dim_color)
                                    .child("↑↓ navigate")
                            )
                            .child(
                                div()
                                    .text_xs()
                                    .text_color(dim_color)
                                    .child("↵ select")
                            )
                            .child(
                                div()
                                    .text_xs()
                                    .text_color(dim_color)
                                    .child("esc cancel")
                            )
                    )
            )
    }
}

/// Register keybindings for the window switcher
pub fn register_keybindings(cx: &mut App) {
    cx.bind_keys([
        KeyBinding::new("up", SwitcherUp, Some("WindowSwitcher")),
        KeyBinding::new("down", SwitcherDown, Some("WindowSwitcher")),
        KeyBinding::new("enter", SwitcherSelect, Some("WindowSwitcher")),
        KeyBinding::new("escape", SwitcherCancel, Some("WindowSwitcher")),
    ]);
}

/// Cycle directly to the next or previous window without opening a modal.
///
/// `direction`: 1 for next, -1 for previous.
pub fn cycle_window(cx: &mut App, current: AnyWindowHandle, direction: i32) {
    let (handle, title) = {
        let Some(registry) = cx.try_global::<WindowRegistry>() else { return };
        let windows = registry.windows();
        if windows.len() <= 1 {
            return;
        }
        let current_idx = windows.iter().position(|w| w.handle == current).unwrap_or(0);
        let next_idx =
            ((current_idx as i32 + direction).rem_euclid(windows.len() as i32)) as usize;
        (windows[next_idx].handle, windows[next_idx].title.clone())
    };

    let _ = handle.update(cx, |root, target_window, cx| {
        target_window.activate_window();
        if let Ok(spreadsheet) = root.downcast::<crate::app::Spreadsheet>() {
            spreadsheet.update(cx, |app, cx| {
                target_window.focus(&app.focus_handle, cx);
                app.status_message = Some(format!("Switched to {}", title));
                cx.notify();
            });
        }
    });
}

/// Open the window switcher as a new window
pub fn open_switcher(cx: &mut App, current_window: AnyWindowHandle) {
    // Only open if there are windows to switch to
    if let Some(registry) = cx.try_global::<WindowRegistry>() {
        if registry.count() <= 1 {
            // No other windows to switch to
            return;
        }
    }

    let bounds = Bounds::centered(None, size(px(400.0), px(300.0)), cx);

    cx.open_window(
        WindowOptions {
            window_bounds: Some(WindowBounds::Windowed(bounds)),
            titlebar: None,
            window_background: WindowBackgroundAppearance::Transparent,
            focus: true,
            show: true,
            kind: WindowKind::PopUp,
            is_movable: false,
            ..Default::default()
        },
        move |window, cx| cx.new(|cx| WindowSwitcher::new(window, cx, current_window)),
    )
    .ok();
}
