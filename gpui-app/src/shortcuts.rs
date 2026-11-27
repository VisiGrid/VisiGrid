//! The Keyboard Shortcuts window's contents (Help > Keyboard Shortcuts).
//!
//! Rows come from the commands themselves (`CommandId::shortcut`), so a
//! command that gains a shortcut appears here without a second list to keep.
//! Keys that have no command (moving, selecting, editing in a cell) are the
//! short `ESSENTIALS` list, and the person's own `keybindings.json` is shown
//! first because it wins over everything below it.

use crate::search::{CommandId, MenuCategory};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShortcutRow {
    pub name: String,
    pub keys: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShortcutSection {
    pub title: &'static str,
    pub rows: Vec<ShortcutRow>,
}

pub const CUSTOM_TITLE: &str = "Your keybindings";
const ESSENTIALS_TITLE: &str = "Moving and editing";
const OTHER_TITLE: &str = "Sheets and selection";

/// Keys bound in `keybindings.rs` that no command carries.
const ESSENTIALS: &[(&str, &str)] = &[
    ("Edit the cell", "F2"),
    ("Next cell", "Tab"),
    ("Previous cell", "Shift+Tab"),
    ("Jump to the edge of the data", "Ctrl+Arrow"),
    ("Extend the selection", "Shift+Arrow"),
    ("Extend the selection to the edge of the data", "Ctrl+Shift+Arrow"),
    ("Select the row", "Shift+Space"),
    ("Select the column", "Ctrl+Space"),
    ("First cell", "Ctrl+Home"),
    ("Last used cell", "Ctrl+End"),
    ("New line inside a cell", "Alt+Enter"),
    ("Cycle a reference between relative and absolute", "F4"),
    ("Open the cell's dropdown", "Alt+Down"),
    ("Open the context menu", "Shift+F10"),
];

const CATEGORIES: [MenuCategory; 7] = [
    MenuCategory::File,
    MenuCategory::Edit,
    MenuCategory::View,
    MenuCategory::Format,
    MenuCategory::Data,
    MenuCategory::Tools,
    MenuCategory::Help,
];

/// "ctrl-shift-d" or "ctrl+shift+d" → "Ctrl+Shift+D".
pub fn key_label(combo: &str) -> String {
    let combo = combo.trim();
    let mut parts: Vec<String> = Vec::new();
    let mut rest = combo;
    while !rest.is_empty() {
        // Skip the first character so a literal "-" or "+" key is kept
        let first = rest.chars().next().map_or(0, char::len_utf8);
        let (part, next) = match rest[first..].find(['-', '+']).map(|i| i + first) {
            Some(i) => (&rest[..i], &rest[i + 1..]),
            None => (rest, ""),
        };
        let mut chars = part.chars();
        parts.push(match chars.next() {
            Some(c) => c.to_uppercase().chain(chars).collect(),
            None => String::new(),
        });
        rest = next;
    }
    parts.join("+")
}

/// The person's own bindings from `keybindings.json`, as (keys, action).
pub fn custom_bindings() -> Vec<(String, String)> {
    let mut rows: Vec<(String, String)> = crate::user_keybindings::load_user_keybindings()
        .bindings
        .into_iter()
        .filter(|(key, _)| !key.trim_start().starts_with("//"))
        .map(|(key, action)| (key_label(&key), action))
        .collect();
    rows.sort();
    rows
}

fn matches(words: &[String], haystack: &str) -> bool {
    let haystack = haystack.to_lowercase();
    words.iter().all(|w| haystack.contains(w.as_str()))
}

/// Every section with at least one row matching `query` (all words, in the
/// name, the keys, the command's search keywords or the section title).
pub fn sections(custom: &[(String, String)], query: &str) -> Vec<ShortcutSection> {
    let words: Vec<String> = query.split_whitespace().map(str::to_lowercase).collect();
    let mut out: Vec<ShortcutSection> = Vec::new();
    let mut push = |title: &'static str, rows: Vec<ShortcutRow>| {
        if !rows.is_empty() {
            out.push(ShortcutSection { title, rows });
        }
    };

    push(
        CUSTOM_TITLE,
        custom
            .iter()
            .filter(|(keys, action)| matches(&words, &format!("{action} {keys} {CUSTOM_TITLE}")))
            .map(|(keys, action)| ShortcutRow { name: action.clone(), keys: keys.clone() })
            .collect(),
    );

    let commands = |category: Option<MenuCategory>, title: &'static str| -> Vec<ShortcutRow> {
        CommandId::all()
            .iter()
            .filter(|cmd| cmd.menu_category() == category)
            .filter_map(|cmd| {
                let keys = cmd.shortcut()?;
                matches(&words, &format!("{} {} {} {}", cmd.name(), keys, cmd.keywords(), title))
                    .then(|| ShortcutRow { name: cmd.name().to_string(), keys: keys.to_string() })
            })
            .collect()
    };

    let other = commands(None, OTHER_TITLE);
    let listed = |keys: &str| other.iter().any(|row| row.keys == keys);
    push(
        ESSENTIALS_TITLE,
        ESSENTIALS
            .iter()
            .filter(|(name, keys)| !listed(keys) && matches(&words, &format!("{name} {keys} {ESSENTIALS_TITLE}")))
            .map(|(name, keys)| ShortcutRow { name: name.to_string(), keys: keys.to_string() })
            .collect(),
    );
    push(OTHER_TITLE, other);
    for category in CATEGORIES {
        push(category.name(), commands(Some(category), category.name()));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn all_rows(sections: &[ShortcutSection]) -> Vec<&ShortcutRow> {
        sections.iter().flat_map(|s| s.rows.iter()).collect()
    }

    #[test]
    fn every_command_with_a_shortcut_is_listed_once() {
        let rows = sections(&[], "");
        let names: Vec<&str> = all_rows(&rows).iter().map(|r| r.name.as_str()).collect();
        for cmd in CommandId::all() {
            let listed = names.iter().filter(|n| **n == cmd.name()).count();
            assert_eq!(listed, usize::from(cmd.shortcut().is_some()), "{}", cmd.name());
        }
    }

    #[test]
    fn no_section_is_empty_and_essentials_do_not_repeat_a_command() {
        let rows = sections(&[], "");
        assert!(rows.iter().all(|s| !s.rows.is_empty()));
        let essentials = rows.iter().find(|s| s.title == ESSENTIALS_TITLE).expect("essentials");
        let other = rows.iter().find(|s| s.title == OTHER_TITLE).expect("other");
        for row in &essentials.rows {
            assert!(other.rows.iter().all(|o| o.keys != row.keys), "{}", row.keys);
        }
    }

    #[test]
    fn search_matches_names_keys_and_keywords() {
        let fill = sections(&[], "fill down");
        assert_eq!(all_rows(&fill).iter().map(|r| r.keys.as_str()).collect::<Vec<_>>(), ["Ctrl+D"]);
        assert!(all_rows(&sections(&[], "ctrl+d")).iter().any(|r| r.name == "Fill Down"));
        assert!(all_rows(&sections(&[], "F2")).iter().any(|r| r.name == "Edit the cell"));
        assert!(sections(&[], "no such shortcut anywhere").is_empty());
    }

    #[test]
    fn the_persons_own_bindings_come_first() {
        let custom = vec![("Ctrl+Shift+D".to_string(), "edit.fillDown".to_string())];
        let rows = sections(&custom, "");
        assert_eq!(rows[0].title, CUSTOM_TITLE);
        assert_eq!(rows[0].rows, [ShortcutRow { name: "edit.fillDown".into(), keys: "Ctrl+Shift+D".into() }]);
        assert_eq!(sections(&custom, "filldown")[0].title, CUSTOM_TITLE);
    }

    #[test]
    fn key_labels_read_like_the_built_in_ones() {
        assert_eq!(key_label("ctrl-shift-d"), "Ctrl+Shift+D");
        assert_eq!(key_label("ctrl+shift+d"), "Ctrl+Shift+D");
        assert_eq!(key_label("cmd-pageup"), "Cmd+Pageup");
        assert_eq!(key_label("ctrl--"), "Ctrl+-");
        assert_eq!(key_label("f9"), "F9");
    }
}
