//! Shared layout geometry and behavior for the compact toolbar and ribbon.
use crate::app::Spreadsheet;
use crate::settings::{user_settings, Setting, SettingsStore, ToolbarLayout};
use gpui::{App, Context, Window};

pub const COMPACT_HEIGHT: f32 = 44.0;
pub const RIBBON_TABS_HEIGHT: f32 = 32.0;
pub const RIBBON_BODY_HEIGHT: f32 = 88.0;

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum RibbonTab {
    #[default]
    Home,
    Insert,
    Formulas,
    Data,
    View,
}

impl RibbonTab {
    pub const ALL: [Self; 5] = [
        Self::Home,
        Self::Insert,
        Self::Formulas,
        Self::Data,
        Self::View,
    ];
    pub fn label(self) -> &'static str {
        match self {
            Self::Home => "Home",
            Self::Insert => "Insert",
            Self::Formulas => "Formulas",
            Self::Data => "Data",
            Self::View => "View",
        }
    }
}

pub struct RibbonState {
    pub focus: gpui::FocusHandle,
    pub temporary_focus: gpui::FocusHandle,
    pub group_focus: gpui::FocusHandle,
    pub tab: RibbonTab,
    pub temporary: bool,
    pub group_menu: Option<usize>,
    pub group_menu_x: f32,
    /// Close obsolete popups when another window changes global layout.
    pub observed: Option<(ToolbarLayout, bool, bool)>,
}

impl RibbonState {
    pub fn new(cx: &mut App) -> Self {
        Self {
            focus: cx.focus_handle(),
            temporary_focus: cx.focus_handle(),
            group_focus: cx.focus_handle(),
            tab: RibbonTab::Home,
            temporary: false,
            group_menu: None,
            group_menu_x: 8.,
            observed: None,
        }
    }

    pub fn contains_focus(&self, window: &Window, cx: &App) -> bool {
        self.focus.contains_focused(window, cx)
            || self.temporary_focus.contains_focused(window, cx)
            || self.group_focus.contains_focused(window, cx)
    }
}

/// Logical coordinates above the grid, independent of worksheet zoom.
#[derive(Debug, Clone, Copy)]
pub struct ToolbarGeometry {
    pub toolbar_top: f32,
    pub toolbar_height: f32,
    pub formula_top: f32,
    pub bottom: f32,
    pub popup_top: f32,
}

impl ToolbarGeometry {
    pub fn new(
        chrome: f32,
        formula: f32,
        layout: ToolbarLayout,
        visible: bool,
        collapsed: bool,
    ) -> Self {
        let height = if !visible {
            0.0
        } else {
            match layout {
                ToolbarLayout::Compact => COMPACT_HEIGHT,
                ToolbarLayout::Ribbon => {
                    RIBBON_TABS_HEIGHT + if collapsed { 0.0 } else { RIBBON_BODY_HEIGHT }
                }
            }
        };
        let toolbar_top = chrome
            + if layout == ToolbarLayout::Compact {
                formula
            } else {
                0.0
            };
        Self {
            toolbar_top,
            toolbar_height: height,
            formula_top: chrome
                + if layout == ToolbarLayout::Ribbon {
                    height
                } else {
                    0.0
                },
            bottom: chrome + formula + height,
            // A temporary collapsed panel is an overlay, not extra grid chrome.
            popup_top: toolbar_top
                + if visible && layout == ToolbarLayout::Ribbon {
                    RIBBON_TABS_HEIGHT + RIBBON_BODY_HEIGHT
                } else {
                    height
                },
        }
    }
}

impl Spreadsheet {
    pub fn toolbar_layout(&self, cx: &App) -> ToolbarLayout {
        user_settings(cx).appearance.toolbar.layout()
    }
    pub fn toolbar_visible(&self, cx: &App) -> bool {
        !self.zen_mode && user_settings(cx).appearance.show_format_bar.resolve(true)
    }
    pub fn ribbon_collapsed(&self, cx: &App) -> bool {
        user_settings(cx).appearance.toolbar.collapsed()
    }

    pub fn toolbar_geometry(&self, cx: &App) -> ToolbarGeometry {
        let chrome = if self.zen_mode {
            0.0
        } else if cfg!(target_os = "macos") {
            crate::app::MACOS_TITLEBAR_HEIGHT
        } else {
            crate::app::MENU_BAR_HEIGHT
        };
        ToolbarGeometry::new(
            chrome,
            if self.zen_mode {
                0.0
            } else {
                self.formula_bar_height()
            },
            self.toolbar_layout(cx),
            self.toolbar_visible(cx),
            self.ribbon_collapsed(cx),
        )
    }

    pub fn close_toolbar_popups(&mut self) {
        self.ui.ribbon.temporary = false;
        self.ui.ribbon.group_menu = None;
        self.ui.format_menu_open = false;
        self.ui.format_bar.size_dropdown = false;
        self.ui.format_bar.number_format_menu_open = false;
        self.ui.format_bar.cell_style_menu_open = false;
    }

    pub fn sync_toolbar_preferences(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let current = (
            self.toolbar_layout(cx),
            self.toolbar_visible(cx),
            self.ribbon_collapsed(cx),
        );
        if self.ui.ribbon.observed.is_some_and(|old| old != current) {
            let had_focus = self.ui.ribbon.contains_focus(window, cx)
                || self.ui.format_bar.size_focus.is_focused(window);
            if self.ui.format_bar.size_editing {
                // Another window may have removed this control. Resolve its draft
                // without ever committing or discarding the cell editor buffer.
                crate::views::format_bar::commit_font_size(self, cx);
            }
            self.dismiss_desktop_keytips(cx);
            self.close_toolbar_popups();
            if had_focus {
                window.focus(&self.focus_handle, cx);
            }
        }
        self.ui.ribbon.observed = Some(current);
    }

    /// Preserve a cell edit; only validate the toolbar's own font-size editor.
    fn prepare_toolbar_change(&mut self, window: &mut Window, cx: &mut Context<Self>) -> bool {
        if self.ui.format_bar.size_editing {
            if crate::views::format_bar::parse_font_size_input(&self.ui.format_bar.size_input)
                .is_none()
            {
                self.status_message =
                    Some("Enter a font size from 1 to 400, or press Escape.".into());
                window.focus(&self.ui.format_bar.size_focus, cx);
                cx.notify();
                return false;
            }
            crate::views::format_bar::commit_font_size(self, cx);
            window.focus(&self.focus_handle, cx);
        }
        self.close_toolbar_popups();
        true
    }

    pub fn set_toolbar_layout(
        &mut self,
        layout: ToolbarLayout,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        if !self.prepare_toolbar_change(window, cx) {
            return;
        }
        let changed = self.toolbar_layout(cx) != layout;
        self.status_message = None;
        self.update_toolbar_preferences(cx, |s| s.appearance.toolbar.set_layout(layout));
        if changed && self.status_message.is_none() {
            self.status_message = Some(match layout {
                ToolbarLayout::Ribbon => "Ribbon toolbar on.".into(),
                ToolbarLayout::Compact => "Compact toolbar on.".into(),
            });
        }
    }

    pub fn toggle_ribbon_collapsed(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if self.toolbar_layout(cx) != ToolbarLayout::Ribbon {
            self.status_message = Some(
                "Collapse/Expand applies to the Ribbon toolbar. Use View \u{2192} Ribbon Toolbar to switch."
                    .into(),
            );
            cx.notify();
            return;
        }
        if !self.prepare_toolbar_change(window, cx) {
            return;
        }
        let collapsed = !self.ribbon_collapsed(cx);
        self.status_message = None;
        self.update_toolbar_preferences(cx, |s| s.appearance.toolbar.set_collapsed(collapsed));
        if self.status_message.is_none() {
            self.status_message = Some(if collapsed {
                "Ribbon collapsed. Click a tab to show it until you pick a command.".into()
            } else {
                "Ribbon expanded.".into()
            });
        }
    }

    pub fn toggle_toolbar_visibility(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if !self.prepare_toolbar_change(window, cx) {
            return;
        }
        let visible = user_settings(cx).appearance.show_format_bar.resolve(true);
        self.update_toolbar_preferences(cx, |s| {
            s.appearance.show_format_bar = Setting::Value(!visible);
            true
        });
    }

    fn update_toolbar_preferences(
        &mut self,
        cx: &mut Context<Self>,
        change: impl FnOnce(&mut crate::settings::UserSettings) -> bool,
    ) {
        use gpui::BorrowAppContext;
        let result = cx.update_global::<SettingsStore, _>(|store, _| {
            if !change(store.user_settings_mut()) {
                return Err(
                    "Toolbar settings were created by a newer VisiGrid version.".to_string()
                );
            }
            store.try_save().map_err(|_| {
                "Couldn't save toolbar preferences. This session's choice is still active.".into()
            })
        });
        if let Err(message) = result {
            self.status_message = Some(message);
        }
        cx.notify();
    }

    pub fn select_ribbon_tab(
        &mut self,
        tab: RibbonTab,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        if !self.prepare_toolbar_change(window, cx) {
            return;
        }
        self.ui.ribbon.tab = tab;
        self.ui.ribbon.temporary = self.ribbon_collapsed(cx);
        cx.notify();
    }
}

#[cfg(test)]
mod tests {
    use super::{ToolbarGeometry, ToolbarLayout};
    #[test]
    fn layouts_keep_formula_grid_and_popup_coordinates_consistent() {
        for chrome in [32.0, 34.0] {
            for formula in [40.0, 100.0] {
                let compact =
                    ToolbarGeometry::new(chrome, formula, ToolbarLayout::Compact, true, false);
                assert_eq!(compact.formula_top, chrome);
                assert_eq!(compact.toolbar_top, chrome + formula);
                assert_eq!(compact.bottom, chrome + formula + 44.0);
                assert_eq!(compact.popup_top, compact.bottom);
                let ribbon =
                    ToolbarGeometry::new(chrome, formula, ToolbarLayout::Ribbon, true, false);
                assert_eq!(ribbon.toolbar_top, chrome);
                assert_eq!(ribbon.formula_top, chrome + 120.0);
                assert_eq!(ribbon.bottom, chrome + formula + 120.0);
                let collapsed =
                    ToolbarGeometry::new(chrome, formula, ToolbarLayout::Ribbon, true, true);
                assert_eq!(collapsed.bottom, chrome + formula + 32.0);
                assert_eq!(collapsed.popup_top, ribbon.popup_top);
                for layout in [ToolbarLayout::Compact, ToolbarLayout::Ribbon] {
                    assert_eq!(
                        ToolbarGeometry::new(chrome, formula, layout, false, false).bottom,
                        chrome + formula
                    );
                }
            }
        }
        assert_eq!(
            ToolbarGeometry::new(0.0, 0.0, ToolbarLayout::Ribbon, false, true).bottom,
            0.0
        );
    }
}
