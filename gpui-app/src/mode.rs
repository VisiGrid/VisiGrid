/// Application modes determine how keyboard input is handled
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Mode {
    #[default]
    Navigation,    // Grid focus: keystrokes move selection
    Edit,          // Cell editor focus: keystrokes edit text (non-formula)
    Formula,       // Formula entry: grid navigation inserts references
    Command,       // Command palette open
    GoTo,          // Go to cell dialog
    QuickOpen,     // Quick file open (Ctrl+P)
    Find,          // Find in cells (Ctrl+F)
    FontPicker,    // Font selection dialog
    ThemePicker,   // Theme selection dialog
    About,         // About VisiGrid dialog
    RenameSymbol,  // Rename named range (Ctrl+Shift+R)
    CreateNamedRange,  // Create named range from selection (Ctrl+Shift+N)
    EditDescription,   // Edit named range description
    Tour,              // Named ranges tour/walkthrough
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

/// Inspector panel tab selection
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum InspectorTab {
    #[default]
    Inspector,  // Cell info, precedents, dependents
    Format,     // Formatting options (Ctrl+1)
    Names,      // Named ranges management
}

impl Mode {
    /// True if editing cell content (either regular Edit or Formula mode)
    pub fn is_editing(&self) -> bool {
        matches!(self, Mode::Edit | Mode::Formula)
    }

    /// True if in formula entry mode (grid nav inserts references)
    pub fn is_formula(&self) -> bool {
        matches!(self, Mode::Formula)
    }

    pub fn is_navigation(&self) -> bool {
        matches!(self, Mode::Navigation)
    }

    pub fn is_overlay(&self) -> bool {
        matches!(self, Mode::Command | Mode::GoTo | Mode::QuickOpen | Mode::Find | Mode::FontPicker | Mode::ThemePicker | Mode::About | Mode::RenameSymbol | Mode::CreateNamedRange | Mode::EditDescription | Mode::Tour)
    }
}
