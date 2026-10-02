//! Linux/Windows Alt-tap navigation. macOS keeps its existing Option behavior.
use crate::{app::Spreadsheet, mode::Menu, settings::ToolbarLayout, toolbar::RibbonTab};
use gpui::{App, Context, Keystroke, Window};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Level {
    Root,
    Ribbon(RibbonTab),
}

#[derive(Default)]
pub struct DesktopKeyTips {
    pub level: Option<Level>,
    pub prefix: String,
    // A mouse gesture or loss of activation must not turn Alt release into a tap.
    pub suppress_alt_tap: bool,
    pub alt_down: bool,
}

impl DesktopKeyTips {
    pub fn active(&self) -> bool {
        !cfg!(target_os = "macos") && self.level.is_some()
    }
    pub fn root(&self) -> bool {
        self.active() && self.level == Some(Level::Root)
    }
    pub fn ribbon(&self, tab: RibbonTab) -> bool {
        self.active() && self.level == Some(Level::Ribbon(tab))
    }
    pub fn clear(&mut self) {
        self.level = None;
        self.prefix.clear();
    }
    pub fn start(&mut self) {
        self.level = Some(Level::Root);
        self.prefix.clear();
    }
    pub fn back(&mut self) {
        if !self.prefix.is_empty() {
            self.prefix.clear();
        } else if matches!(self.level, Some(Level::Ribbon(_))) {
            self.start();
        } else {
            self.clear();
        }
    }
}

pub const MENU_KEYS: [(char, Menu); 7] = [
    ('f', Menu::File),
    ('e', Menu::Edit),
    ('v', Menu::View),
    ('i', Menu::Insert),
    ('o', Menu::Format),
    ('d', Menu::Data),
    ('h', Menu::Help),
];
pub const TAB_KEYS: [(char, RibbonTab); 5] = [
    ('b', RibbonTab::Home),
    ('n', RibbonTab::Insert),
    ('m', RibbonTab::Formulas),
    ('a', RibbonTab::Data),
    ('w', RibbonTab::View),
];

/// Alt-tap is part of the opt-in ribbon, never a mode switch while typing.
fn scope_available(layout: ToolbarLayout, mode: crate::mode::Mode, visible: bool) -> bool {
    layout == ToolbarLayout::Ribbon && mode.is_navigation() && visible
}

impl Spreadsheet {
    pub fn desktop_keytips_available(&self, window: &Window, cx: &App) -> bool {
        window.is_window_active()
            && scope_available(self.toolbar_layout(cx), self.mode, self.toolbar_visible(cx))
            && self.edit_marked_range.is_none()
            && (self.focus_handle.is_focused(window) || self.ui.ribbon.contains_focus(window, cx))
            && !self.ui.format_bar.is_active(window)
            && !self.script.open
            && !self.lua_console.visible
            && self.pivot_panel.is_none()
            && self.table_dialog.is_none()
            && self.pending_table_recovery.is_none()
            && !self.close_confirm_visible
            && self.comment_editor.is_none()
            && self.renaming_sheet.is_none()
            && self.context_menu.is_none()
            && self.sheet_context_menu.is_none()
            && self.filter_dropdown_col.is_none()
            && !self.is_validation_dropdown_open()
    }

    pub fn dismiss_desktop_keytips(&mut self, cx: &mut Context<Self>) {
        if self.ui.desktop_keytips.active() {
            self.ui.desktop_keytips.clear();
            self.ui.ribbon.temporary = false;
            self.ui.ribbon.group_menu = None;
            cx.notify();
        }
    }

    #[cfg(not(target_os = "macos"))]
    pub fn toggle_desktop_keytips(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if self.ui.desktop_keytips.suppress_alt_tap || !self.desktop_keytips_available(window, cx) {
            return;
        }
        if self.ui.desktop_keytips.active() {
            self.dismiss_desktop_keytips(cx);
        } else {
            self.close_toolbar_popups();
            self.close_menu(cx);
            self.ui.desktop_keytips.start();
            cx.notify();
        }
    }

    /// Intercept before bindings: letters must not edit cells, and Escape must
    /// only back out of hints. Ordinary Alt chords continue to the normal keymap.
    #[cfg(not(target_os = "macos"))]
    pub fn intercept_desktop_keytips(
        window: &mut Window,
        cx: &mut Context<Self>,
    ) -> gpui::Subscription {
        let this = cx.entity().downgrade();
        let handle = window.window_handle();
        cx.intercept_keystrokes(move |event, window, cx| {
            if window.window_handle() != handle {
                return;
            }
            let Some(this) = this.upgrade() else { return };
            let handled = this.update(cx, |this, cx| {
                this.desktop_keytip_key(&event.keystroke, window, cx)
            });
            if handled {
                cx.stop_propagation();
            }
        })
    }

    pub fn desktop_keytip_key(
        &mut self,
        key: &Keystroke,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) -> bool {
        if !self.ui.desktop_keytips.active() {
            return false;
        }
        if !self.desktop_keytips_available(window, cx) {
            self.dismiss_desktop_keytips(cx);
            return false;
        }
        if key.modifiers.alt
            || key.modifiers.control
            || key.modifiers.platform
            || key.modifiers.function
        {
            self.dismiss_desktop_keytips(cx);
            return false;
        }
        // Modifier-only Alt is dispatched by GPUI on release to ShowKeyTips.
        if matches!(key.key.as_str(), "alt" | "shift" | "control" | "platform") {
            return false;
        }
        if matches!(key.key.as_str(), "escape" | "backspace") {
            self.ui.desktop_keytips.back();
            self.ui.ribbon.group_menu = None;
            if !matches!(self.ui.desktop_keytips.level, Some(Level::Ribbon(_))) {
                self.ui.ribbon.temporary = false;
            }
            cx.notify();
            return true;
        }
        let input = key.key.to_ascii_lowercase();
        match self.ui.desktop_keytips.level {
            Some(Level::Root) => {
                if let Some((_, menu)) = MENU_KEYS.iter().find(|(c, _)| input == c.to_string()) {
                    self.dismiss_desktop_keytips(cx);
                    self.toggle_menu(*menu, cx);
                } else if self.toolbar_visible(cx)
                    && self.toolbar_layout(cx) == ToolbarLayout::Ribbon
                {
                    if let Some((_, tab)) = TAB_KEYS.iter().find(|(c, _)| input == c.to_string()) {
                        self.select_ribbon_tab(*tab, window, cx);
                        self.ui.desktop_keytips.level = Some(Level::Ribbon(*tab));
                        self.ui.desktop_keytips.prefix.clear();
                        cx.notify();
                    }
                }
            }
            Some(Level::Ribbon(tab)) => {
                crate::views::ribbon::handle_keytip(self, tab, &input, window, cx)
            }
            None => {}
        }
        // Unknown keys are consumed, without changing the workbook or dismissing hints.
        true
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn alt_tap_is_only_available_in_visible_ribbon_navigation() {
        use crate::mode::Mode;
        for mode in [
            Mode::Navigation,
            Mode::Edit,
            Mode::Formula,
            Mode::Find,
            Mode::Command,
            Mode::FontPicker,
        ] {
            assert!(!scope_available(ToolbarLayout::Compact, mode, true));
            assert!(!scope_available(ToolbarLayout::Ribbon, mode, false));
            assert_eq!(
                scope_available(ToolbarLayout::Ribbon, mode, true),
                mode == Mode::Navigation
            );
        }
    }

    #[test]
    fn root_letters_do_not_collide_and_help_stays_h() {
        let mut keys: Vec<_> = MENU_KEYS
            .iter()
            .map(|(k, _)| *k)
            .chain(TAB_KEYS.iter().map(|(k, _)| *k))
            .collect();
        let len = keys.len();
        keys.sort();
        keys.dedup();
        assert_eq!(len, keys.len());
        assert!(matches!(
            MENU_KEYS.iter().find(|(k, _)| *k == 'h'),
            Some((_, Menu::Help))
        ));
        assert_eq!(TAB_KEYS[0], ('b', RibbonTab::Home));
    }
    #[test]
    fn escape_backs_out_of_prefix_then_tab_then_root() {
        let mut state = DesktopKeyTips::default();
        state.level = Some(Level::Ribbon(RibbonTab::Home));
        state.prefix = "f".into();
        state.back();
        assert!(state.prefix.is_empty());
        assert_eq!(state.level, Some(Level::Ribbon(RibbonTab::Home)));
        state.back();
        assert_eq!(state.level, Some(Level::Root));
        state.back();
        assert_eq!(state.level, None);
    }
}
