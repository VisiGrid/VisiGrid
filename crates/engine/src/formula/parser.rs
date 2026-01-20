// Formula parser - converts formula strings into AST
// Supports: numbers, cell refs (A1), ranges (A1:A5), functions (SUM), basic math (+, -, *, /)
// Also supports: comparison operators (<, >, =, <=, >=, <>), string literals, concatenation (&)

#[derive(Debug, Clone)]
pub enum Expr {
    Number(f64),
    Text(String),
    Boolean(bool),
    CellRef { col: usize, row: usize },
    Range {
        start_col: usize,
        start_row: usize,
        end_col: usize,
        end_row: usize,
    },
    Function {
        name: String,
        args: Vec<Expr>,
    },
    BinaryOp {
        op: Op,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    /// Named range reference (resolved at evaluation time)
    NamedRange(String),
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Op {
    // Arithmetic
    Add,
    Sub,
    Mul,
    Div,
    // Comparison
    Lt,      // <
    Gt,      // >
    Eq,      // =
    LtEq,    // <=
    GtEq,    // >=
    NotEq,   // <>
    // String
    Concat,  // &
}

pub fn parse(formula: &str) -> Result<Expr, String> {
    let formula = formula.trim();
    if !formula.starts_with('=') {
        return Err("Formula must start with =".to_string());
    }

    let input = &formula[1..]; // Skip the '='
    let tokens = tokenize(input)?;
    if tokens.is_empty() {
        return Err("Empty formula".to_string());
    }
    parse_expr(&tokens)
}

#[derive(Debug, Clone)]
enum Token {
    Number(f64),
    StringLit(String),
    CellRef { col: usize, row: usize },
    Ident(String),
    Plus,
    Minus,
    Star,
    Slash,
    LParen,
    RParen,
    Colon,
    Comma,
    // Comparison operators
    Lt,      // <
    Gt,      // >
    Eq,      // =
    LtEq,    // <=
    GtEq,    // >=
    NotEq,   // <>
    // String concatenation
    Ampersand, // &
}

fn tokenize(input: &str) -> Result<Vec<Token>, String> {
    let mut tokens = Vec::new();
    let mut chars = input.chars().peekable();

    while let Some(&c) = chars.peek() {
        match c {
            ' ' | '\t' => { chars.next(); }
            '+' => { tokens.push(Token::Plus); chars.next(); }
            '-' => { tokens.push(Token::Minus); chars.next(); }
            '*' => { tokens.push(Token::Star); chars.next(); }
            '/' => { tokens.push(Token::Slash); chars.next(); }
            '(' => { tokens.push(Token::LParen); chars.next(); }
            ')' => { tokens.push(Token::RParen); chars.next(); }
            ':' => { tokens.push(Token::Colon); chars.next(); }
            ',' => { tokens.push(Token::Comma); chars.next(); }
            '&' => { tokens.push(Token::Ampersand); chars.next(); }
            '<' => {
                chars.next();
                if let Some(&next) = chars.peek() {
                    match next {
                        '=' => { tokens.push(Token::LtEq); chars.next(); }
                        '>' => { tokens.push(Token::NotEq); chars.next(); }
                        _ => { tokens.push(Token::Lt); }
                    }
                } else {
                    tokens.push(Token::Lt);
                }
            }
            '>' => {
                chars.next();
                if let Some(&'=') = chars.peek() {
                    tokens.push(Token::GtEq);
                    chars.next();
                } else {
                    tokens.push(Token::Gt);
                }
            }
            '=' => { tokens.push(Token::Eq); chars.next(); }
            '"' => {
                // String literal
                chars.next(); // consume opening quote
                let mut s = String::new();
                loop {
                    match chars.next() {
                        Some('"') => break,
                        Some(ch) => s.push(ch),
                        None => return Err("Unterminated string literal".to_string()),
                    }
                }
                tokens.push(Token::StringLit(s));
            }
            'A'..='Z' | 'a'..='z' => {
                // Could be cell reference (A1) or function name (SUM)
                let mut ident = String::new();
                while let Some(&ch) = chars.peek() {
                    if ch.is_ascii_alphanumeric() || ch == '_' || ch == '$' {
                        ident.push(ch);
                        chars.next();
                    } else {
                        break;
                    }
                }

                // Check for TRUE/FALSE boolean literals
                let upper = ident.to_uppercase();
                if upper == "TRUE" {
                    tokens.push(Token::Ident("TRUE".to_string()));
                } else if upper == "FALSE" {
                    tokens.push(Token::Ident("FALSE".to_string()));
                } else if let Some(token) = try_parse_cell_ref(&ident) {
                    // Check if it's a cell reference (letters + digits)
                    tokens.push(token);
                } else {
                    // It's a function name or identifier
                    tokens.push(Token::Ident(upper));
                }
            }
            '$' => {
                // Absolute reference marker - collect with following letters/numbers
                let mut ident = String::new();
                while let Some(&ch) = chars.peek() {
                    if ch.is_ascii_alphanumeric() || ch == '$' {
                        ident.push(ch);
                        chars.next();
                    } else {
                        break;
                    }
                }
                if let Some(token) = try_parse_cell_ref(&ident) {
                    tokens.push(token);
                } else {
                    return Err(format!("Invalid cell reference: {}", ident));
                }
            }
            '0'..='9' | '.' => {
                let mut num_str = String::new();
                while let Some(&d) = chars.peek() {
                    if d.is_ascii_digit() || d == '.' {
                        num_str.push(d);
                        chars.next();
                    } else {
                        break;
                    }
                }
                let num: f64 = num_str.parse().map_err(|_| format!("Invalid number: {}", num_str))?;
                tokens.push(Token::Number(num));
            }
            _ => return Err(format!("Unexpected character: {}", c)),
        }
    }

    Ok(tokens)
}

fn try_parse_cell_ref(s: &str) -> Option<Token> {
    let s = s.to_uppercase();
    let mut chars = s.chars().peekable();

    // Skip leading $ for absolute column reference
    if chars.peek() == Some(&'$') {
        chars.next();
    }

    // Collect column letters (now supports multi-letter like AA, AB, etc.)
    let mut col_str = String::new();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_uppercase() {
            col_str.push(c);
            chars.next();
        } else {
            break;
        }
    }

    if col_str.is_empty() {
        return None;
    }

    // Skip $ for absolute row reference
    if chars.peek() == Some(&'$') {
        chars.next();
    }

    let row_str: String = chars.collect();
    if row_str.is_empty() || !row_str.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }

    let row: usize = row_str.parse().ok()?;
    if row == 0 {
        return None;
    }

    // Convert column letters to number (A=0, B=1, ..., Z=25, AA=26, AB=27, etc.)
    let col = col_str.chars().fold(0usize, |acc, c| {
        acc * 26 + (c as usize - 'A' as usize + 1)
    }) - 1;

    Some(Token::CellRef { col, row: row - 1 })
}

fn parse_expr(tokens: &[Token]) -> Result<Expr, String> {
    parse_comparison(tokens, 0).map(|(expr, _)| expr)
}

// Lowest precedence: comparison operators
fn parse_comparison(tokens: &[Token], pos: usize) -> Result<(Expr, usize), String> {
    let (mut left, mut pos) = parse_concat(tokens, pos)?;

    while pos < tokens.len() {
        let op = match &tokens[pos] {
            Token::Lt => Op::Lt,
            Token::Gt => Op::Gt,
            Token::Eq => Op::Eq,
            Token::LtEq => Op::LtEq,
            Token::GtEq => Op::GtEq,
            Token::NotEq => Op::NotEq,
            _ => break,
        };
        let (right, new_pos) = parse_concat(tokens, pos + 1)?;
        left = Expr::BinaryOp {
            op,
            left: Box::new(left),
            right: Box::new(right),
        };
        pos = new_pos;
    }

    Ok((left, pos))
}

// String concatenation (&)
fn parse_concat(tokens: &[Token], pos: usize) -> Result<(Expr, usize), String> {
    let (mut left, mut pos) = parse_add_sub(tokens, pos)?;

    while pos < tokens.len() {
        if let Token::Ampersand = &tokens[pos] {
            let (right, new_pos) = parse_add_sub(tokens, pos + 1)?;
            left = Expr::BinaryOp {
                op: Op::Concat,
                left: Box::new(left),
                right: Box::new(right),
            };
            pos = new_pos;
        } else {
            break;
        }
    }

    Ok((left, pos))
}

fn parse_add_sub(tokens: &[Token], pos: usize) -> Result<(Expr, usize), String> {
    let (mut left, mut pos) = parse_mul_div(tokens, pos)?;

    while pos < tokens.len() {
        match &tokens[pos] {
            Token::Plus => {
                let (right, new_pos) = parse_mul_div(tokens, pos + 1)?;
                left = Expr::BinaryOp {
                    op: Op::Add,
                    left: Box::new(left),
                    right: Box::new(right),
                };
                pos = new_pos;
            }
            Token::Minus => {
                let (right, new_pos) = parse_mul_div(tokens, pos + 1)?;
                left = Expr::BinaryOp {
                    op: Op::Sub,
                    left: Box::new(left),
                    right: Box::new(right),
                };
                pos = new_pos;
            }
            _ => break,
        }
    }

    Ok((left, pos))
}

fn parse_mul_div(tokens: &[Token], pos: usize) -> Result<(Expr, usize), String> {
    let (mut left, mut pos) = parse_primary(tokens, pos)?;

    while pos < tokens.len() {
        match &tokens[pos] {
            Token::Star => {
                let (right, new_pos) = parse_primary(tokens, pos + 1)?;
                left = Expr::BinaryOp {
                    op: Op::Mul,
                    left: Box::new(left),
                    right: Box::new(right),
                };
                pos = new_pos;
            }
            Token::Slash => {
                let (right, new_pos) = parse_primary(tokens, pos + 1)?;
                left = Expr::BinaryOp {
                    op: Op::Div,
                    left: Box::new(left),
                    right: Box::new(right),
                };
                pos = new_pos;
            }
            _ => break,
        }
    }

    Ok((left, pos))
}

fn parse_primary(tokens: &[Token], pos: usize) -> Result<(Expr, usize), String> {
    if pos >= tokens.len() {
        return Err("Unexpected end of expression".to_string());
    }

    match &tokens[pos] {
        Token::Number(n) => Ok((Expr::Number(*n), pos + 1)),
        Token::StringLit(s) => Ok((Expr::Text(s.clone()), pos + 1)),
        Token::CellRef { col, row } => {
            // Check if this is a range (A1:B5)
            if pos + 2 < tokens.len() {
                if let Token::Colon = &tokens[pos + 1] {
                    if let Token::CellRef { col: end_col, row: end_row } = &tokens[pos + 2] {
                        return Ok((
                            Expr::Range {
                                start_col: *col,
                                start_row: *row,
                                end_col: *end_col,
                                end_row: *end_row,
                            },
                            pos + 3,
                        ));
                    }
                }
            }
            Ok((Expr::CellRef { col: *col, row: *row }, pos + 1))
        }
        Token::Ident(name) => {
            // Check for boolean literals
            if name == "TRUE" {
                return Ok((Expr::Boolean(true), pos + 1));
            }
            if name == "FALSE" {
                return Ok((Expr::Boolean(false), pos + 1));
            }
            // Function call
            if pos + 1 < tokens.len() {
                if let Token::LParen = &tokens[pos + 1] {
                    let (args, new_pos) = parse_function_args(tokens, pos + 2)?;
                    return Ok((
                        Expr::Function {
                            name: name.clone(),
                            args,
                        },
                        new_pos,
                    ));
                }
            }
            // Not a function call - treat as a named range (resolved at evaluation time)
            Ok((Expr::NamedRange(name.clone()), pos + 1))
        }
        Token::LParen => {
            let (expr, pos) = parse_comparison(tokens, pos + 1)?;
            if pos >= tokens.len() {
                return Err("Missing closing parenthesis".to_string());
            }
            match &tokens[pos] {
                Token::RParen => Ok((expr, pos + 1)),
                _ => Err("Expected closing parenthesis".to_string()),
            }
        }
        Token::Minus => {
            // Unary minus
            let (expr, pos) = parse_primary(tokens, pos + 1)?;
            Ok((
                Expr::BinaryOp {
                    op: Op::Sub,
                    left: Box::new(Expr::Number(0.0)),
                    right: Box::new(expr),
                },
                pos,
            ))
        }
        _ => Err(format!("Unexpected token at position {}", pos)),
    }
}

fn parse_function_args(tokens: &[Token], pos: usize) -> Result<(Vec<Expr>, usize), String> {
    let mut args = Vec::new();
    let mut pos = pos;

    // Handle empty function call SUM()
    if pos < tokens.len() {
        if let Token::RParen = &tokens[pos] {
            return Ok((args, pos + 1));
        }
    }

    loop {
        let (arg, new_pos) = parse_comparison(tokens, pos)?;
        args.push(arg);
        pos = new_pos;

        if pos >= tokens.len() {
            return Err("Missing closing parenthesis in function call".to_string());
        }

        match &tokens[pos] {
            Token::RParen => return Ok((args, pos + 1)),
            Token::Comma => pos += 1,
            _ => return Err("Expected comma or closing parenthesis".to_string()),
        }
    }
}

/// Extract all cell references from an expression (for dependency tracking)
/// Returns a list of (row, col) tuples for single cells and expanded ranges
pub fn extract_cell_refs(expr: &Expr) -> Vec<(usize, usize)> {
    let mut refs = Vec::new();
    collect_cell_refs(expr, &mut refs);
    refs
}

fn collect_cell_refs(expr: &Expr, refs: &mut Vec<(usize, usize)>) {
    match expr {
        Expr::Number(_) | Expr::Text(_) | Expr::Boolean(_) | Expr::NamedRange(_) => {
            // NamedRange refs are resolved at evaluation time with access to NamedRangeStore
        }
        Expr::CellRef { col, row } => {
            refs.push((*row, *col));
        }
        Expr::Range { start_col, start_row, end_col, end_row } => {
            // Expand range to individual cells
            for r in *start_row..=*end_row {
                for c in *start_col..=*end_col {
                    refs.push((r, c));
                }
            }
        }
        Expr::Function { args, .. } => {
            for arg in args {
                collect_cell_refs(arg, refs);
            }
        }
        Expr::BinaryOp { left, right, .. } => {
            collect_cell_refs(left, refs);
            collect_cell_refs(right, refs);
        }
    }
}
