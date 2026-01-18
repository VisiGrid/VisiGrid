// Formula parser - converts formula strings into AST
// Supports: numbers, cell refs (A1), ranges (A1:A5), functions (SUM), basic math (+, -, *, /)

#[derive(Debug, Clone)]
pub enum Expr {
    Number(f64),
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
}

#[derive(Debug, Clone, Copy)]
pub enum Op {
    Add,
    Sub,
    Mul,
    Div,
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
            'A'..='Z' | 'a'..='z' => {
                // Could be cell reference (A1) or function name (SUM)
                let mut ident = String::new();
                while let Some(&ch) = chars.peek() {
                    if ch.is_ascii_alphanumeric() {
                        ident.push(ch);
                        chars.next();
                    } else {
                        break;
                    }
                }

                // Check if it's a cell reference (single letter + digits)
                if let Some(token) = try_parse_cell_ref(&ident) {
                    tokens.push(token);
                } else {
                    // It's a function name or identifier
                    tokens.push(Token::Ident(ident.to_uppercase()));
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
    let mut chars = s.chars();

    let col_char = chars.next()?;
    if !col_char.is_ascii_uppercase() {
        return None;
    }

    let row_str: String = chars.collect();
    if row_str.is_empty() || !row_str.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }

    let row: usize = row_str.parse().ok()?;
    if row == 0 {
        return None;
    }

    let col = (col_char as usize) - ('A' as usize);
    Some(Token::CellRef { col, row: row - 1 })
}

fn parse_expr(tokens: &[Token]) -> Result<Expr, String> {
    parse_add_sub(tokens, 0).map(|(expr, _)| expr)
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
            Err(format!("Unknown identifier: {}", name))
        }
        Token::LParen => {
            let (expr, pos) = parse_add_sub(tokens, pos + 1)?;
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
        let (arg, new_pos) = parse_add_sub(tokens, pos)?;
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
        Expr::Number(_) => {}
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
