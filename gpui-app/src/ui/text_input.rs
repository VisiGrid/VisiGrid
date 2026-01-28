//! Shared input-handling logic for manual text fields.
//!
//! Not a UI widget — just behaviour so every "mini text editor" in the app
//! (color picker hex input, find query, goto input, font picker query, etc.)
//! behaves consistently for typing, backspace, select-all, and paste.

/// What happened after processing a key or paste event.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InputAction {
    /// Buffer changed — caller should `cx.notify()`.
    Changed,
    /// User pressed Enter — caller should execute/submit.
    Submit,
    /// User pressed Escape — caller should cancel/close.
    Cancel,
    /// Key was not handled by the helper.
    Ignored,
}

/// Process a key-down event against a buffer + all-selected flag.
///
/// Handles: printable chars, backspace, enter, escape.
/// Does NOT handle paste (see [`handle_input_paste`]) because paste
/// comes via a separate action in gpui, not via on_key_down.
pub fn handle_input_key(
    buffer: &mut String,
    all_selected: &mut bool,
    key: &str,
    key_char: Option<&str>,
    has_modifier: bool,
) -> InputAction {
    match key {
        "enter" => return InputAction::Submit,
        "escape" => return InputAction::Cancel,
        "backspace" => {
            if *all_selected {
                buffer.clear();
                *all_selected = false;
            } else {
                buffer.pop();
            }
            return InputAction::Changed;
        }
        _ => {}
    }

    // Printable character (no ctrl/alt/platform modifier)
    if !has_modifier {
        if let Some(chars) = key_char {
            if *all_selected {
                buffer.clear();
                *all_selected = false;
            }
            for c in chars.chars() {
                buffer.push(c);
            }
            return InputAction::Changed;
        }
    }

    InputAction::Ignored
}

/// Handle a select-all gesture (Cmd+A / Ctrl+A).
///
/// Sets the `all_selected` flag. The next printable char, backspace,
/// or paste will replace/clear the entire buffer.
pub fn handle_input_select_all(all_selected: &mut bool) {
    *all_selected = true;
}

/// Handle a paste event against a buffer + all-selected flag.
///
/// Strips control characters and trims whitespace.
/// If `all_selected`, replaces the buffer; otherwise appends.
pub fn handle_input_paste(
    buffer: &mut String,
    all_selected: &mut bool,
    text: &str,
) {
    let clean: String = text.trim().chars()
        .filter(|c| !c.is_control())
        .collect();
    if clean.is_empty() {
        return;
    }
    if *all_selected {
        buffer.clear();
        *all_selected = false;
    }
    buffer.push_str(&clean);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_typing() {
        let mut buf = String::new();
        let mut sel = false;
        assert_eq!(handle_input_key(&mut buf, &mut sel, "a", Some("a"), false), InputAction::Changed);
        assert_eq!(buf, "a");
        assert_eq!(handle_input_key(&mut buf, &mut sel, "b", Some("b"), false), InputAction::Changed);
        assert_eq!(buf, "ab");
    }

    #[test]
    fn test_backspace() {
        let mut buf = "abc".to_string();
        let mut sel = false;
        assert_eq!(handle_input_key(&mut buf, &mut sel, "backspace", None, false), InputAction::Changed);
        assert_eq!(buf, "ab");
    }

    #[test]
    fn test_submit_cancel() {
        let mut buf = "x".to_string();
        let mut sel = false;
        assert_eq!(handle_input_key(&mut buf, &mut sel, "enter", None, false), InputAction::Submit);
        assert_eq!(handle_input_key(&mut buf, &mut sel, "escape", None, false), InputAction::Cancel);
    }

    #[test]
    fn test_select_all_then_type() {
        let mut buf = "hello".to_string();
        let mut sel = false;
        handle_input_select_all(&mut sel);
        assert!(sel);
        assert_eq!(handle_input_key(&mut buf, &mut sel, "x", Some("x"), false), InputAction::Changed);
        assert_eq!(buf, "x");
        assert!(!sel);
    }

    #[test]
    fn test_select_all_then_backspace() {
        let mut buf = "hello".to_string();
        let mut sel = false;
        handle_input_select_all(&mut sel);
        assert_eq!(handle_input_key(&mut buf, &mut sel, "backspace", None, false), InputAction::Changed);
        assert_eq!(buf, "");
    }

    #[test]
    fn test_select_all_then_paste() {
        let mut buf = "old".to_string();
        let mut sel = false;
        handle_input_select_all(&mut sel);
        handle_input_paste(&mut buf, &mut sel, "  #FF00AA  ");
        assert_eq!(buf, "#FF00AA");
        assert!(!sel);
    }

    #[test]
    fn test_paste_appends() {
        let mut buf = "#FF".to_string();
        let mut sel = false;
        handle_input_paste(&mut buf, &mut sel, "00AA");
        assert_eq!(buf, "#FF00AA");
    }

    #[test]
    fn test_modifier_ignored() {
        let mut buf = String::new();
        let mut sel = false;
        assert_eq!(handle_input_key(&mut buf, &mut sel, "a", Some("a"), true), InputAction::Ignored);
        assert_eq!(buf, "");
    }
}
