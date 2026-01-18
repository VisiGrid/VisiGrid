// Keybinding configuration

use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Keybinding {
    pub key: String,
    pub command: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub when: Option<String>,
}

#[derive(Debug, Clone)]
pub struct KeybindingManager {
    // Maps normalized key combo -> command id
    bindings: HashMap<String, String>,
    // Maps command id -> key combo (for display)
    shortcuts: HashMap<String, String>,
    config_path: PathBuf,
}

impl Default for KeybindingManager {
    fn default() -> Self {
        Self::new()
    }
}

impl KeybindingManager {
    pub fn new() -> Self {
        let config_path = Self::get_config_path();
        let mut manager = Self {
            bindings: HashMap::new(),
            shortcuts: HashMap::new(),
            config_path,
        };
        manager.load_defaults();
        manager.load_user_config();
        manager
    }

    fn get_config_path() -> PathBuf {
        let config_dir = dirs::config_dir()
            .unwrap_or_else(|| PathBuf::from("."))
            .join("visigrid");
        config_dir.join("keybindings.json")
    }

    fn load_defaults(&mut self) {
        let defaults = default_keybindings();
        for binding in defaults {
            let key = Self::normalize_key(&binding.key);
            self.shortcuts.insert(binding.command.clone(), binding.key.clone());
            self.bindings.insert(key, binding.command);
        }
    }

    fn load_user_config(&mut self) {
        if !self.config_path.exists() {
            // Create default config file
            self.create_default_config();
            return;
        }

        match fs::read_to_string(&self.config_path) {
            Ok(contents) => {
                match serde_json::from_str::<Vec<Keybinding>>(&contents) {
                    Ok(user_bindings) => {
                        for binding in user_bindings {
                            let key = Self::normalize_key(&binding.key);
                            self.shortcuts.insert(binding.command.clone(), binding.key.clone());
                            self.bindings.insert(key, binding.command);
                        }
                    }
                    Err(e) => {
                        eprintln!("Error parsing keybindings.json: {}", e);
                    }
                }
            }
            Err(e) => {
                eprintln!("Error reading keybindings.json: {}", e);
            }
        }
    }

    fn create_default_config(&self) {
        // Create config directory if it doesn't exist
        if let Some(parent) = self.config_path.parent() {
            if let Err(e) = fs::create_dir_all(parent) {
                eprintln!("Error creating config directory: {}", e);
                return;
            }
        }

        // Write default config with helpful comments
        let default_config = r#"[
    // Custom keybindings - these override defaults
    // Format: { "key": "Ctrl+Key", "command": "command.id" }
    //
    // Available commands:
    //   file.new, file.open, file.save
    //   edit.copy, edit.cut, edit.paste, edit.undo, edit.redo
    //   cell.clear, cell.edit, cell.moveUp/Down/Left/Right
    //   select.all, select.extend
    //   format.bold, format.italic, format.underline
    //   navigate.goto, navigate.find
    //   palette.toggle, theme.toggle
    //   formula.autosum
    //   data.fillDown, data.fillRight
    //
    // Example custom binding:
    // { "key": "Ctrl+G", "command": "navigate.goto" }
]
"#;

        if let Err(e) = fs::write(&self.config_path, default_config) {
            eprintln!("Error writing default keybindings.json: {}", e);
        }
    }

    /// Normalize key string to canonical form: "ctrl+shift+alt+key"
    fn normalize_key(key: &str) -> String {
        let key = key.to_lowercase();
        let parts: Vec<&str> = key.split('+').collect();

        let mut has_ctrl = false;
        let mut has_shift = false;
        let mut has_alt = false;
        let mut main_key = "";

        for part in parts {
            let part = part.trim();
            match part {
                "ctrl" | "control" => has_ctrl = true,
                "shift" => has_shift = true,
                "alt" => has_alt = true,
                _ => main_key = part,
            }
        }

        let mut result = String::new();
        if has_ctrl { result.push_str("ctrl+"); }
        if has_shift { result.push_str("shift+"); }
        if has_alt { result.push_str("alt+"); }
        result.push_str(main_key);
        result
    }

    /// Get command for a key combination
    pub fn get_command(&self, key: &str) -> Option<&String> {
        let normalized = Self::normalize_key(key);
        self.bindings.get(&normalized)
    }

    /// Get shortcut display string for a command
    pub fn get_shortcut(&self, command: &str) -> Option<&String> {
        self.shortcuts.get(command)
    }

    /// Get config file path for display
    pub fn config_path(&self) -> &PathBuf {
        &self.config_path
    }
}

pub fn default_keybindings() -> Vec<Keybinding> {
    vec![
        // File operations
        Keybinding { key: "ctrl+n".into(), command: "file.new".into(), when: None },
        Keybinding { key: "ctrl+o".into(), command: "file.open".into(), when: None },
        Keybinding { key: "ctrl+s".into(), command: "file.save".into(), when: None },

        // Navigation (Excel-style)
        Keybinding { key: "tab".into(), command: "cell.moveRight".into(), when: None },
        Keybinding { key: "shift+tab".into(), command: "cell.moveLeft".into(), when: None },
        Keybinding { key: "enter".into(), command: "cell.moveDown".into(), when: None },
        Keybinding { key: "shift+enter".into(), command: "cell.moveUp".into(), when: None },
        Keybinding { key: "ctrl+home".into(), command: "cell.goToStart".into(), when: None },
        Keybinding { key: "ctrl+end".into(), command: "cell.goToEnd".into(), when: None },

        // Editing
        Keybinding { key: "ctrl+c".into(), command: "edit.copy".into(), when: None },
        Keybinding { key: "ctrl+x".into(), command: "edit.cut".into(), when: None },
        Keybinding { key: "ctrl+v".into(), command: "edit.paste".into(), when: None },
        Keybinding { key: "ctrl+z".into(), command: "edit.undo".into(), when: None },
        Keybinding { key: "ctrl+y".into(), command: "edit.redo".into(), when: None },
        Keybinding { key: "delete".into(), command: "cell.clear".into(), when: None },
        Keybinding { key: "f2".into(), command: "cell.edit".into(), when: None },

        // Selection
        Keybinding { key: "ctrl+a".into(), command: "select.all".into(), when: None },
        Keybinding { key: "ctrl+shift+end".into(), command: "select.toEnd".into(), when: None },

        // Command palette
        Keybinding { key: "ctrl+shift+p".into(), command: "commandPalette.toggle".into(), when: None },

        // Problems panel
        Keybinding { key: "ctrl+shift+m".into(), command: "problems.toggle".into(), when: None },

        // Navigation dialogs
        Keybinding { key: "ctrl+g".into(), command: "navigate.goto".into(), when: None },
        Keybinding { key: "ctrl+p".into(), command: "file.quickOpen".into(), when: None },
        Keybinding { key: "ctrl+f".into(), command: "edit.find".into(), when: None },

        // Formatting
        Keybinding { key: "ctrl+b".into(), command: "format.bold".into(), when: None },
        Keybinding { key: "ctrl+i".into(), command: "format.italic".into(), when: None },
        Keybinding { key: "ctrl+u".into(), command: "format.underline".into(), when: None },

        // Column sizing
        Keybinding { key: "ctrl+0".into(), command: "column.autoSize".into(), when: None },

        // Data operations
        Keybinding { key: "ctrl+d".into(), command: "data.fillDown".into(), when: None },
        Keybinding { key: "ctrl+r".into(), command: "data.fillRight".into(), when: None },

        // Formula
        Keybinding { key: "alt+=".into(), command: "formula.autosum".into(), when: None },

        // Row/Column selection and manipulation
        Keybinding { key: "ctrl+space".into(), command: "select.column".into(), when: None },
        Keybinding { key: "shift+space".into(), command: "select.row".into(), when: None },
        Keybinding { key: "ctrl+-".into(), command: "edit.deleteRowCol".into(), when: None },
        Keybinding { key: "ctrl+shift+=".into(), command: "edit.insertRowCol".into(), when: None },

        // Date/Time insertion
        Keybinding { key: "ctrl+;".into(), command: "edit.insertDate".into(), when: None },
        Keybinding { key: "ctrl+shift+;".into(), command: "edit.insertTime".into(), when: None },

        // Formula view toggle
        Keybinding { key: "ctrl+`".into(), command: "view.toggleFormulas".into(), when: None },

        // Cell reference cycling (F4)
        Keybinding { key: "f4".into(), command: "edit.cycleReference".into(), when: None },

        // Split view
        Keybinding { key: "ctrl+\\".into(), command: "view.splitToggle".into(), when: None },
        Keybinding { key: "ctrl+w".into(), command: "view.splitSwitch".into(), when: None },

        // Zen mode
        Keybinding { key: "f11".into(), command: "view.zenMode".into(), when: None },

        // Cell inspector
        Keybinding { key: "ctrl+shift+i".into(), command: "view.inspector".into(), when: None },
    ]
}
