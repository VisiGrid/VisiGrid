// Command palette - fuzzy search for all actions
// Activated with Ctrl+Shift+P

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Command {
    pub id: String,
    pub label: String,
    pub description: Option<String>,
    pub keybinding: Option<String>,
}

pub struct CommandPalette {
    pub visible: bool,
    pub query: String,
    pub commands: Vec<Command>,
    pub filtered: Vec<usize>,
    pub selected: usize,
}

impl CommandPalette {
    pub fn new() -> Self {
        Self {
            visible: false,
            query: String::new(),
            commands: default_commands(),
            filtered: Vec::new(),
            selected: 0,
        }
    }

    pub fn toggle(&mut self) {
        self.visible = !self.visible;
        if self.visible {
            self.query.clear();
            self.filter();
        }
    }

    pub fn filter(&mut self) {
        let query = self.query.to_lowercase();
        self.filtered = self
            .commands
            .iter()
            .enumerate()
            .filter(|(_, cmd)| {
                cmd.label.to_lowercase().contains(&query)
                    || cmd
                        .description
                        .as_ref()
                        .map(|d| d.to_lowercase().contains(&query))
                        .unwrap_or(false)
            })
            .map(|(i, _)| i)
            .collect();
        self.selected = 0;
    }
}

impl Default for CommandPalette {
    fn default() -> Self {
        Self::new()
    }
}

fn default_commands() -> Vec<Command> {
    vec![
        Command {
            id: "file.new".into(),
            label: "New Workbook".into(),
            description: Some("Create a new empty workbook".into()),
            keybinding: Some("Ctrl+N".into()),
        },
        Command {
            id: "file.open".into(),
            label: "Open File".into(),
            description: Some("Open an existing file".into()),
            keybinding: Some("Ctrl+O".into()),
        },
        Command {
            id: "file.save".into(),
            label: "Save".into(),
            description: Some("Save the current workbook".into()),
            keybinding: Some("Ctrl+S".into()),
        },
        Command {
            id: "edit.copy".into(),
            label: "Copy".into(),
            description: Some("Copy selected cells".into()),
            keybinding: Some("Ctrl+C".into()),
        },
        Command {
            id: "edit.paste".into(),
            label: "Paste".into(),
            description: Some("Paste from clipboard".into()),
            keybinding: Some("Ctrl+V".into()),
        },
        Command {
            id: "view.theme.toggle".into(),
            label: "Toggle Dark Mode".into(),
            description: Some("Switch between light and dark themes".into()),
            keybinding: None,
        },
    ]
}
