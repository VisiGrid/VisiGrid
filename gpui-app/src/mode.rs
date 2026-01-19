/// Application modes determine how keyboard input is handled
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Mode {
    #[default]
    Navigation,  // Grid focus: keystrokes move selection
    Edit,        // Cell editor focus: keystrokes edit text
    Command,     // Command palette open
    GoTo,        // Go to cell dialog
    QuickOpen,   // Quick file open (Ctrl+P)
    Find,        // Find in cells (Ctrl+F)
    FontPicker,  // Font selection dialog
}

/// Which menu dropdown is currently open (Excel 2003 style)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Menu {
    File,
    Edit,
    View,
    Insert,
    Format,
    Data,
    Help,
}

impl Mode {
    pub fn is_editing(&self) -> bool {
        matches!(self, Mode::Edit)
    }

    pub fn is_navigation(&self) -> bool {
        matches!(self, Mode::Navigation)
    }

    pub fn is_overlay(&self) -> bool {
        matches!(self, Mode::Command | Mode::GoTo | Mode::QuickOpen | Mode::Find | Mode::FontPicker)
    }
}
