/// Formula editing submode: determines how arrow keys behave in formula mode
///
/// Excel-like behavior: switches automatically based on caret position.
/// - Point mode: caret is at a ref insertion point (after `(`, `,`, operators)
/// - Caret mode: caret is inside a token (editing existing text)
///
/// F2 manually toggles between modes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FormulaNavMode {
    #[default]
    Point,   // Arrows move grid selection (ref-pick) - default when typing `=`
    Caret,   // Arrows move text cursor inside formula - default when F2 on existing formula
}

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
    ImpactPreview,     // Preview impact of rename/delete before applying
    RefactorLog,       // Show refactoring audit trail
    ExtractNamedRange, // Extract range literal to named range
    ImportReport,      // Excel import results report
    ExportReport,      // Excel export results report (shown when warnings exist)
    Preferences,       // User preferences dialog (Cmd+,)
    Hint,              // Keyboard hints visible (Vimium-style jump mode)
    License,           // Enter/view license dialog
    HubPasteToken,     // Hub: Paste device token (fallback auth)
    HubLink,           // Hub: Link workbook to dataset
    HubPublishConfirm, // Hub: Confirm publish when diverged
    ValidationDialog,  // Data validation dialog (Data > Validation)
    AISettings,        // AI configuration dialog (Help > AI Settings)
    AiDialog,          // AI dialog - Insert Formula or Analyze (Tools > AI)
    ExplainDiff,       // Explain Differences dialog (History right-click)
    PasteSpecial,      // Paste Special dialog (Ctrl+Alt+V)
    ColorPicker,       // Color picker modal (Fill Color)
    FormatPainter,     // Format Painter: next click applies captured format
    NumberFormatEditor, // Number format editor (Ctrl+1 when Format tab is open)
    TransformPreview,  // Transform diff preview dialog (Pro)
    ConvertPicker,     // Convert format picker dialog (palette → Convert)
    CloudOpen,         // Cloud sheet picker dialog (File > Open Cloud)
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

impl Menu {
    /// Next menu in order (wrapping)
    pub fn next(self) -> Menu {
        match self {
            Menu::File => Menu::Edit,
            Menu::Edit => Menu::View,
            Menu::View => Menu::Insert,
            Menu::Insert => Menu::Format,
            Menu::Format => Menu::Data,
            Menu::Data => Menu::Help,
            Menu::Help => Menu::File,
        }
    }

    /// Previous menu in order (wrapping)
    pub fn prev(self) -> Menu {
        match self {
            Menu::File => Menu::Help,
            Menu::Edit => Menu::File,
            Menu::View => Menu::Edit,
            Menu::Insert => Menu::View,
            Menu::Format => Menu::Insert,
            Menu::Data => Menu::Format,
            Menu::Help => Menu::Data,
        }
    }
}

/// Inspector panel tab selection
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum InspectorTab {
    #[default]
    Inspector,  // Cell info, precedents, dependents
    Format,     // Formatting options (Ctrl+1)
    Names,      // Named ranges management
    History,    // Undo history with Lua provenance (Ctrl+H)
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

    /// True if in keyboard hint mode (Vimium-style jump)
    pub fn is_hint(&self) -> bool {
        matches!(self, Mode::Hint)
    }

    pub fn is_overlay(&self) -> bool {
        matches!(self, Mode::Command | Mode::GoTo | Mode::QuickOpen | Mode::Find | Mode::FontPicker | Mode::ThemePicker | Mode::About | Mode::RenameSymbol | Mode::CreateNamedRange | Mode::EditDescription | Mode::Tour | Mode::ImpactPreview | Mode::RefactorLog | Mode::ExtractNamedRange | Mode::ImportReport | Mode::ExportReport | Mode::Preferences | Mode::License | Mode::HubPasteToken | Mode::HubLink | Mode::HubPublishConfirm | Mode::ValidationDialog | Mode::AISettings | Mode::ExplainDiff | Mode::PasteSpecial | Mode::ColorPicker | Mode::NumberFormatEditor | Mode::TransformPreview | Mode::ConvertPicker)
    }

    /// True if this mode has text input active (typing should work normally).
    /// Used to guard Option+letter accelerators on macOS to avoid
    /// conflicting with character composition (accents, special chars).
    pub fn has_text_input(&self) -> bool {
        matches!(
            self,
            Mode::Edit
                | Mode::Formula
                | Mode::GoTo           // GoTo cell input
                | Mode::Find           // Find/Replace input
                | Mode::Command        // Command palette input
                | Mode::RenameSymbol   // Rename dialog
                | Mode::CreateNamedRange // Named range dialog
                | Mode::EditDescription  // Description input
                | Mode::ExtractNamedRange // Extract dialog
                | Mode::License        // License key input
                | Mode::HubPasteToken  // Hub token input
                | Mode::HubLink        // Hub link input
                | Mode::AISettings     // API key input
                | Mode::AiDialog       // AI prompt input
                | Mode::ColorPicker    // Hex color input
                | Mode::NumberFormatEditor  // Currency symbol input
        )
    }
}
