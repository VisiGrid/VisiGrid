//! Help for the ribbon's existing commands. Shortcuts come from the live keymap.
use super::{disabled_reason, item_label, Item};
use crate::{actions as a, app::Spreadsheet, search::CommandId, theme::TokenKey};
use gpui::{prelude::FluentBuilder, *};

pub(super) fn shortcut(window: &Window, action: &dyn Action) -> Option<SharedString> {
    // Resolve in the sheet context even while the ribbon has keyboard focus.
    // GPUI excludes bindings shadowed by user overrides or explicit unbindings.
    window
        .highest_precedence_binding_for_action_in_context(
            action,
            KeyContext::parse("Spreadsheet").unwrap(),
        )
        .map(|binding| {
            binding
                .keystrokes()
                .iter()
                .map(|key| {
                    key.to_string()
                        .replace("ctrl-", "Ctrl+")
                        .replace("alt-", "Alt+")
                        .replace("shift-", "Shift+")
                        .replace("super-", "Super+")
                        .replace("win-", "Win+")
                })
                .collect::<Vec<_>>()
                .join(", ")
                .into()
        })
}

fn action(item: Item) -> Option<Box<dyn Action>> {
    use CommandId::*;
    let Item::Command(command, _) = item else {
        return None;
    };
    Some(match command {
        Paste => Box::new(a::Paste),
        Cut => Box::new(a::Cut),
        Copy => Box::new(a::Copy),
        PasteValues => Box::new(a::PasteValues),
        PasteSpecial => Box::new(a::PasteSpecial),
        FormatPainter => Box::new(a::FormatPainter),
        SelectFont => Box::new(a::ShowFontPicker),
        ToggleBold => Box::new(a::ToggleBold),
        ToggleItalic => Box::new(a::ToggleItalic),
        ToggleUnderline => Box::new(a::ToggleUnderline),
        FillColor => Box::new(a::ShowColorPicker),
        AlignLeft => Box::new(a::AlignLeft),
        AlignCenter => Box::new(a::AlignCenter),
        AlignRight => Box::new(a::AlignRight),
        FormatCurrency => Box::new(a::FormatCurrency),
        FormatPercent => Box::new(a::FormatPercent),
        AutoSum => Box::new(a::AutoSum),
        FindInCells => Box::new(a::FindInCells),
        FillDown => Box::new(a::FillDown),
        FillRight => Box::new(a::FillRight),
        AddSheet => Box::new(a::AddSheet),
        InsertPivotTable => Box::new(a::InsertPivotTable),
        Recalculate => Box::new(a::Recalculate),
        ToggleTrace => Box::new(a::ToggleTrace),
        CycleTracePrecedent => Box::new(a::CycleTracePrecedent),
        CycleTraceDependent => Box::new(a::CycleTraceDependent),
        ToggleInspector => Box::new(a::ToggleInspector),
        ToggleAutoFilter => Box::new(a::ToggleAutoFilter),
        TrimWhitespace => Box::new(a::TrimWhitespace),
        ValidationDialog => Box::new(a::ShowDataValidation),
        RefreshPivot => Box::new(a::RefreshPivot),
        RefreshAllPivots => Box::new(a::RefreshAllPivots),
        EditPivotFields => Box::new(a::EditPivotFields),
        FreezeTopRow => Box::new(a::FreezeTopRow),
        FreezeFirstColumn => Box::new(a::FreezeFirstColumn),
        FreezePanes => Box::new(a::FreezePanes),
        UnfreezePanes => Box::new(a::UnfreezePanes),
        ZoomIn => Box::new(a::ZoomIn),
        ZoomOut => Box::new(a::ZoomOut),
        ZoomReset => Box::new(a::ZoomReset),
        ToggleMinimap => Box::new(a::ToggleMinimap),
        ToggleCommentsSidebar => Box::new(a::ToggleCommentsSidebar),
        ToggleZenMode => Box::new(a::ToggleZenMode),
        // Commands without a matching GPUI action have no direct shortcut.
        // In particular, Ctrl+= is NOT the ribbon's insert-at-selection command.
        _ => return None,
    })
}

fn description(item: Item) -> &'static str {
    use CommandId::*;
    use Item::Command as C;
    match item {
        C(Paste, _) => "Paste clipboard content into the selection using your paste preference.",
        C(Cut, _) => "Move the selected cells when you paste them elsewhere.",
        C(Copy, _) => "Copy the selected cells to the clipboard.",
        C(PasteValues, _) => "Paste results without formulas or source formatting.",
        C(PasteSpecial, _) => "Choose which parts of the copied cells to paste.",
        C(FormatPainter, _) => "Copy the active cell's formatting, then apply it to another selection.",
        C(SelectFont, _) => "Choose a font family for the selected cells.",
        C(ToggleBold, _) => "Turn bold formatting on or off for the selected cells.",
        C(ToggleItalic, _) => "Turn italic formatting on or off for the selected cells.",
        C(ToggleUnderline, _) => "Turn underlining on or off for the selected cells.",
        C(FillColor, _) => "Choose a background color for the selected cells.",
        Item::FontSize => "Set the selected cells' font size in points (1–400).",
        Item::TextColor => "Choose a text color for the selected cells.",
        C(AlignLeft, _) => "Align the selected cells' content to the left.",
        C(AlignCenter, _) => "Center the selected cells' content horizontally.",
        C(AlignRight, _) => "Align the selected cells' content to the right.",
        Item::NumberFormat => "Choose how numbers are displayed without changing their values.",
        C(FormatCurrency, _) => "Display the selected numbers as currency.",
        C(FormatPercent, _) => "Display the selected numbers as percentages.",
        Item::Styles => "Apply a preset cell style to the selection.",
        C(AddConditionalFormat, _) => "Create a rule that formats cells when its condition is met.",
        C(ManageConditionalFormats, _) => "Review and edit conditional formatting rules.",
        C(AutoSum, _) => "Insert a SUM formula using nearby cells.",
        C(FindInCells, _) => "Find text or values in the sheet.",
        C(FillDown, _) => "Copy the top selected row down, adjusting relative formula references.",
        C(FillRight, _) => "Copy the leftmost selected column right, adjusting relative formula references.",
        C(AddSheet, _) => "Add a new worksheet to this workbook.",
        Item::InsertRows => "Insert entire sheet rows above the selection, one per selected row. Existing cells move down.",
        Item::InsertCols => "Insert entire sheet columns left of the selection, one per selected column. Existing cells move right.",
        C(InsertPivotTable, _) => "Summarize the selection or current data region in a pivot table.",
        C(ExtractNamedRange, _) => "Name a range referenced by the active formula and update matching formulas.",
        C(Recalculate, _) => "Recalculate the workbook's formulas.",
        C(ToggleTrace, _) => "Show or hide formula dependency tracing.",
        C(CycleTracePrecedent, _) => "Move to the next cell used by the traced formula.",
        C(CycleTraceDependent, _) => "Move to the next formula that uses the traced cell.",
        C(ToggleInspector, _) => "Show or hide details about the selected cell.",
        C(SortAscending, _) => "Sort the data in ascending order using the active column.",
        C(SortDescending, _) => "Sort the data in descending order using the active column.",
        C(ClearSort, _) => "Remove the current sort and restore the underlying row order.",
        C(ToggleAutoFilter, _) => "Show or hide filter controls for the data region.",
        C(TrimWhitespace, _) => "Remove leading and trailing whitespace from selected text cells.",
        C(ValidationDialog, _) => "Set rules for values allowed in the selected cells.",
        C(RefreshPivot, _) => "Refresh the pivot table at the active cell from its source data.",
        C(RefreshAllPivots, _) => "Refresh every pivot table in the workbook.",
        C(EditPivotFields, _) => "Choose the rows, columns, and values for the active pivot table.",
        C(FreezeTopRow, _) => "Keep the first row visible while scrolling.",
        C(FreezeFirstColumn, _) => "Keep the first column visible while scrolling.",
        C(FreezePanes, _) => "Keep rows above and columns left of the active cell visible while scrolling.",
        C(UnfreezePanes, _) => "Allow all rows and columns to scroll again.",
        C(ZoomIn, _) => "Increase the worksheet zoom.", C(ZoomOut, _) => "Decrease the worksheet zoom.",
        C(ZoomReset, _) => "Restore the worksheet zoom to 100%.",
        C(ToggleMinimap, _) => "Show or hide the sheet overview for quick navigation.",
        C(ToggleCommentsSidebar, _) => "Show or hide the workbook's comments sidebar.",
        C(SelectTheme, _) => "Choose the application's appearance theme.",
        C(ToggleZenMode, _) => "Hide interface controls to give the worksheet more space.",
        _ => "",
    }
}

#[derive(Clone)]
pub(in crate::views) struct RibbonTooltip {
    title: SharedString,
    pub(super) shortcut: Option<SharedString>,
    pub(super) description: SharedString,
    reason: Option<SharedString>,
    bg: Hsla,
    border: Hsla,
    text: Hsla,
    muted: Hsla,
}

impl RibbonTooltip {
    pub(super) fn new(
        app: &Spreadsheet,
        title: impl Into<SharedString>,
        description: impl Into<SharedString>,
        shortcut: Option<SharedString>,
        reason: Option<&'static str>,
    ) -> Self {
        Self {
            title: title.into(),
            description: description.into(),
            shortcut,
            reason: reason.map(Into::into),
            bg: app.token(TokenKey::PanelBg),
            border: app.token(TokenKey::PanelBorder),
            text: app.token(TokenKey::TextPrimary),
            muted: app.token(TokenKey::TextMuted),
        }
    }

    pub(super) fn item(app: &Spreadsheet, item: Item, window: &Window) -> Self {
        let title = match item {
            Item::Command(command, _) => command.name(),
            _ => item_label(item).trim_end_matches(" ▾"),
        };
        let key = action(item).and_then(|action| shortcut(window, action.as_ref()));
        Self::new(
            app,
            title,
            description(item),
            key,
            disabled_reason(app, item),
        )
    }
}

impl Render for RibbonTooltip {
    fn render(&mut self, _: &mut Window, _: &mut Context<Self>) -> impl IntoElement {
        div()
            .max_w(px(340.))
            .p_2()
            .rounded_md()
            .shadow_md()
            .border_1()
            .border_color(self.border)
            .bg(self.bg)
            .text_color(self.text)
            .text_size(px(12.))
            .flex()
            .flex_col()
            .gap_1()
            .child(
                div()
                    .flex()
                    .items_center()
                    .gap_3()
                    .child(
                        div()
                            .font_weight(FontWeight::SEMIBOLD)
                            .child(self.title.clone()),
                    )
                    .when_some(self.shortcut.clone(), |d, key| {
                        d.child(div().text_size(px(11.)).text_color(self.muted).child(key))
                    }),
            )
            .when(!self.description.is_empty(), |d| {
                d.child(div().child(self.description.clone()))
            })
            .when_some(self.reason.clone(), |d, reason| {
                d.child(div().text_color(self.muted).child(reason))
            })
    }
}
