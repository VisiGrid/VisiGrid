use iced::widget::{text, Row};
use iced::{Color, Element};

#[derive(Debug, Clone, Copy, PartialEq)]
enum TokenType {
    Function,
    CellRef,
    Number,
    Operator,
    Paren,
    Colon,
    Comma,
    Text,
}

struct Token {
    text: String,
    token_type: TokenType,
}

/// Parse a formula into tokens for highlighting
fn tokenize(input: &str) -> Vec<Token> {
    let mut tokens = Vec::new();
    let chars: Vec<char> = input.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        let c = chars[i];

        // Function names (uppercase letters followed by open paren)
        if c.is_ascii_uppercase() {
            let start = i;
            while i < chars.len() && chars[i].is_ascii_uppercase() {
                i += 1;
            }
            let word: String = chars[start..i].iter().collect();

            // Check if followed by '(' to determine if it's a function
            if i < chars.len() && chars[i] == '(' {
                tokens.push(Token { text: word, token_type: TokenType::Function });
            } else {
                // Could be a cell reference like A1, AB12
                // Check if followed by digits
                let col_part = word;
                let num_start = i;
                while i < chars.len() && chars[i].is_ascii_digit() {
                    i += 1;
                }
                if i > num_start {
                    let num_part: String = chars[num_start..i].iter().collect();
                    tokens.push(Token {
                        text: format!("{}{}", col_part, num_part),
                        token_type: TokenType::CellRef
                    });
                } else {
                    // Just letters, treat as text
                    tokens.push(Token { text: col_part, token_type: TokenType::Text });
                }
            }
            continue;
        }

        // Numbers
        if c.is_ascii_digit() || (c == '.' && i + 1 < chars.len() && chars[i + 1].is_ascii_digit()) {
            let start = i;
            while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                i += 1;
            }
            let num: String = chars[start..i].iter().collect();
            tokens.push(Token { text: num, token_type: TokenType::Number });
            continue;
        }

        // Operators
        if "+-*/^%".contains(c) {
            tokens.push(Token { text: c.to_string(), token_type: TokenType::Operator });
            i += 1;
            continue;
        }

        // Comparison operators
        if c == '=' || c == '<' || c == '>' {
            let start = i;
            i += 1;
            if i < chars.len() && (chars[i] == '=' || chars[i] == '>') {
                i += 1;
            }
            let op: String = chars[start..i].iter().collect();
            tokens.push(Token { text: op, token_type: TokenType::Operator });
            continue;
        }

        // Parentheses
        if c == '(' || c == ')' {
            tokens.push(Token { text: c.to_string(), token_type: TokenType::Paren });
            i += 1;
            continue;
        }

        // Colon (range separator)
        if c == ':' {
            tokens.push(Token { text: c.to_string(), token_type: TokenType::Colon });
            i += 1;
            continue;
        }

        // Comma
        if c == ',' {
            tokens.push(Token { text: c.to_string(), token_type: TokenType::Comma });
            i += 1;
            continue;
        }

        // Whitespace and other characters
        tokens.push(Token { text: c.to_string(), token_type: TokenType::Text });
        i += 1;
    }

    tokens
}

/// Get color for token type (dark theme)
fn token_color_dark(token_type: TokenType) -> Color {
    match token_type {
        TokenType::Function => Color::from_rgb(0.4, 0.7, 1.0),  // Light blue
        TokenType::CellRef => Color::from_rgb(0.6, 0.9, 0.6),   // Light green
        TokenType::Number => Color::from_rgb(0.9, 0.7, 0.5),    // Orange/tan
        TokenType::Operator => Color::from_rgb(1.0, 0.8, 0.3),  // Yellow
        TokenType::Paren => Color::from_rgb(0.8, 0.6, 0.9),     // Purple
        TokenType::Colon => Color::from_rgb(0.7, 0.7, 0.7),     // Gray
        TokenType::Comma => Color::from_rgb(0.7, 0.7, 0.7),     // Gray
        TokenType::Text => Color::from_rgb(0.9, 0.9, 0.9),      // White
    }
}

/// Get color for token type (light theme)
fn token_color_light(token_type: TokenType) -> Color {
    match token_type {
        TokenType::Function => Color::from_rgb(0.0, 0.3, 0.7),  // Blue
        TokenType::CellRef => Color::from_rgb(0.0, 0.5, 0.0),   // Green
        TokenType::Number => Color::from_rgb(0.7, 0.4, 0.0),    // Brown/orange
        TokenType::Operator => Color::from_rgb(0.6, 0.4, 0.0),  // Dark yellow
        TokenType::Paren => Color::from_rgb(0.5, 0.2, 0.6),     // Purple
        TokenType::Colon => Color::from_rgb(0.4, 0.4, 0.4),     // Gray
        TokenType::Comma => Color::from_rgb(0.4, 0.4, 0.4),     // Gray
        TokenType::Text => Color::from_rgb(0.1, 0.1, 0.1),      // Black
    }
}

/// Create a highlighted formula display
pub fn highlight_formula<'a, Message: 'a>(formula: &str, dark_mode: bool) -> Element<'a, Message> {
    let tokens = tokenize(formula);
    let color_fn = if dark_mode { token_color_dark } else { token_color_light };

    let elements: Vec<Element<'a, Message>> = tokens
        .into_iter()
        .map(|token| {
            text(token.text)
                .size(13)
                .color(color_fn(token.token_type))
                .into()
        })
        .collect();

    Row::with_children(elements).into()
}

/// Check if a string looks like a formula
pub fn is_formula(s: &str) -> bool {
    s.starts_with('=')
}
