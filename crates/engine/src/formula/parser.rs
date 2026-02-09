// Formula parser - converts formula strings into AST
// Supports: numbers, cell refs (A1), ranges (A1:A5), functions (SUM), basic math (+, -, *, /)
// Also supports: comparison operators (<, >, =, <=, >=, <>), string literals, concatenation (&)

use crate::sheet::{SheetId, SheetRef, UnboundSheetRef};

/// Generic expression AST, parameterized over sheet reference type.
/// - Parser outputs `ParsedExpr = Expr<UnboundSheetRef>` (sheet names unresolved)
/// - After binding, becomes `BoundExpr = Expr<SheetRef>` (sheet IDs resolved)
#[derive(Debug, Clone)]
pub enum Expr<S> {
    Number(f64),
    Text(String),
    Boolean(bool),
    /// Cell reference with sheet context
    /// - col_abs/row_abs: true if that component is absolute ($A vs A, $1 vs 1)
    CellRef {
        sheet: S,
        col: usize,
        row: usize,
        col_abs: bool,
        row_abs: bool,
    },
    /// Range reference with sheet context
    Range {
        sheet: S,
        start_col: usize,
        start_row: usize,
        end_col: usize,
        end_row: usize,
        start_col_abs: bool,
        start_row_abs: bool,
        end_col_abs: bool,
        end_row_abs: bool,
    },
    Function {
        name: String,
        args: Vec<Expr<S>>,
    },
    BinaryOp {
        op: Op,
        left: Box<Expr<S>>,
        right: Box<Expr<S>>,
    },
    /// Named range reference (resolved at evaluation time)
    NamedRange(String),
    /// Empty/omitted argument (e.g. the trailing slot in `=IF(a,b,)`)
    Empty,
}

/// Parser output: sheet references are unresolved names
pub type ParsedExpr = Expr<UnboundSheetRef>;

/// Bound expression: sheet references resolved to stable IDs
pub type BoundExpr = Expr<SheetRef>;

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
    // Exponentiation
    Pow,     // ^
}

/// Parse a formula string into an unbound AST (sheet names not yet resolved to IDs).
/// Call `bind_expr()` with workbook context to resolve sheet references before evaluation.
pub fn parse(formula: &str) -> Result<ParsedExpr, String> {
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
    /// Cell reference with absolute/relative flags
    CellRef {
        col: usize,
        row: usize,
        col_abs: bool,
        row_abs: bool,
    },
    /// Sheet name prefix (e.g., "Sheet1" from "Sheet1!A1")
    SheetPrefix(String),
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
    // Exponentiation and percent
    Caret,   // ^
    Percent, // %
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
            '^' => { tokens.push(Token::Caret); chars.next(); }
            '%' => { tokens.push(Token::Percent); chars.next(); }
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
            '\'' => {
                // Quoted sheet name (e.g., 'My Sheet'!A1 or 'Bob''s Sheet'!A1)
                // Doubled quotes ('') inside are escape for a single quote
                chars.next(); // consume opening quote
                let mut sheet_name = String::new();
                loop {
                    match chars.next() {
                        Some('\'') => {
                            // Check if it's an escaped quote (doubled)
                            if chars.peek() == Some(&'\'') {
                                chars.next(); // consume second quote
                                sheet_name.push('\''); // add single quote to name
                            } else {
                                break; // end of quoted name
                            }
                        }
                        Some(ch) => sheet_name.push(ch),
                        None => return Err("Unterminated sheet name".to_string()),
                    }
                }
                // Must be followed by !
                if chars.next() != Some('!') {
                    return Err("Quoted sheet name must be followed by !".to_string());
                }
                tokens.push(Token::SheetPrefix(sheet_name));
            }
            'A'..='Z' | 'a'..='z' => {
                // Could be cell reference (A1), function name (SUM), or sheet prefix (Sheet1!)
                let mut ident = String::new();
                while let Some(&ch) = chars.peek() {
                    if ch.is_ascii_alphanumeric() || ch == '_' || ch == '$' {
                        ident.push(ch);
                        chars.next();
                    } else {
                        break;
                    }
                }

                // Dotted function names (e.g. STDEV.P, NORM.S.DIST)
                // Consume .alpha segments while the ident is not a cell ref
                while chars.peek() == Some(&'.') {
                    // Peek past the dot to see if a letter follows
                    let mut lookahead = chars.clone();
                    lookahead.next(); // skip '.'
                    if let Some(&ch) = lookahead.peek() {
                        if ch.is_ascii_alphabetic() {
                            chars.next(); // consume '.'
                            ident.push('.');
                            while let Some(&ch) = chars.peek() {
                                if ch.is_ascii_alphabetic() {
                                    ident.push(ch);
                                    chars.next();
                                } else {
                                    break;
                                }
                            }
                        } else {
                            break;
                        }
                    } else {
                        break;
                    }
                }

                // Check if followed by ! (sheet reference prefix)
                if chars.peek() == Some(&'!') {
                    chars.next(); // consume the !
                    tokens.push(Token::SheetPrefix(ident));
                    continue;
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

    // Check for $ (absolute column reference)
    let col_abs = if chars.peek() == Some(&'$') {
        chars.next();
        true
    } else {
        false
    };

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

    // Check for $ (absolute row reference)
    let row_abs = if chars.peek() == Some(&'$') {
        chars.next();
        true
    } else {
        false
    };

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

    Some(Token::CellRef { col, row: row - 1, col_abs, row_abs })
}

fn parse_expr(tokens: &[Token]) -> Result<ParsedExpr, String> {
    parse_comparison(tokens, 0).map(|(expr, _)| expr)
}

// Lowest precedence: comparison operators
fn parse_comparison(tokens: &[Token], pos: usize) -> Result<(ParsedExpr, usize), String> {
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
fn parse_concat(tokens: &[Token], pos: usize) -> Result<(ParsedExpr, usize), String> {
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

fn parse_add_sub(tokens: &[Token], pos: usize) -> Result<(ParsedExpr, usize), String> {
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

fn parse_mul_div(tokens: &[Token], pos: usize) -> Result<(ParsedExpr, usize), String> {
    let (mut left, mut pos) = parse_power(tokens, pos)?;

    while pos < tokens.len() {
        match &tokens[pos] {
            Token::Star => {
                let (right, new_pos) = parse_power(tokens, pos + 1)?;
                left = Expr::BinaryOp {
                    op: Op::Mul,
                    left: Box::new(left),
                    right: Box::new(right),
                };
                pos = new_pos;
            }
            Token::Slash => {
                let (right, new_pos) = parse_power(tokens, pos + 1)?;
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

// Exponentiation (^) - right-associative, higher precedence than * /
fn parse_power(tokens: &[Token], pos: usize) -> Result<(ParsedExpr, usize), String> {
    let (base, pos) = parse_percent(tokens, pos)?;

    if pos < tokens.len() {
        if let Token::Caret = &tokens[pos] {
            // Right-associative: recurse into parse_power for the exponent
            let (exponent, new_pos) = parse_power(tokens, pos + 1)?;
            return Ok((
                Expr::BinaryOp {
                    op: Op::Pow,
                    left: Box::new(base),
                    right: Box::new(exponent),
                },
                new_pos,
            ));
        }
    }

    Ok((base, pos))
}

// Percent postfix (%) - highest precedence operator, desugars to * 0.01
fn parse_percent(tokens: &[Token], pos: usize) -> Result<(ParsedExpr, usize), String> {
    let (mut expr, mut pos) = parse_primary(tokens, pos)?;

    while pos < tokens.len() {
        if let Token::Percent = &tokens[pos] {
            expr = Expr::BinaryOp {
                op: Op::Mul,
                left: Box::new(expr),
                right: Box::new(Expr::Number(0.01)),
            };
            pos += 1;
        } else {
            break;
        }
    }

    Ok((expr, pos))
}

fn parse_primary(tokens: &[Token], pos: usize) -> Result<(ParsedExpr, usize), String> {
    if pos >= tokens.len() {
        return Err("Unexpected end of expression".to_string());
    }

    match &tokens[pos] {
        Token::Number(n) => Ok((Expr::Number(*n), pos + 1)),
        Token::StringLit(s) => Ok((Expr::Text(s.clone()), pos + 1)),
        Token::SheetPrefix(sheet_name) => {
            // Sheet prefix must be followed by a cell reference
            if pos + 1 >= tokens.len() {
                return Err("Sheet reference must be followed by cell reference".to_string());
            }
            let sheet = UnboundSheetRef::Named(sheet_name.clone());
            match &tokens[pos + 1] {
                Token::CellRef { col, row, col_abs, row_abs } => {
                    // Check if this is a range (Sheet1!A1:B5)
                    if pos + 3 < tokens.len() {
                        if let Token::Colon = &tokens[pos + 2] {
                            if let Token::CellRef { col: end_col, row: end_row, col_abs: end_col_abs, row_abs: end_row_abs } = &tokens[pos + 3] {
                                return Ok((
                                    Expr::Range {
                                        sheet,
                                        start_col: *col,
                                        start_row: *row,
                                        end_col: *end_col,
                                        end_row: *end_row,
                                        start_col_abs: *col_abs,
                                        start_row_abs: *row_abs,
                                        end_col_abs: *end_col_abs,
                                        end_row_abs: *end_row_abs,
                                    },
                                    pos + 4,
                                ));
                            }
                        }
                    }
                    Ok((Expr::CellRef { sheet, col: *col, row: *row, col_abs: *col_abs, row_abs: *row_abs }, pos + 2))
                }
                _ => Err("Sheet reference must be followed by cell reference".to_string()),
            }
        }
        Token::CellRef { col, row, col_abs, row_abs } => {
            // Check if this is a range (A1:B5)
            if pos + 2 < tokens.len() {
                if let Token::Colon = &tokens[pos + 1] {
                    if let Token::CellRef { col: end_col, row: end_row, col_abs: end_col_abs, row_abs: end_row_abs } = &tokens[pos + 2] {
                        return Ok((
                            Expr::Range {
                                sheet: UnboundSheetRef::Current,
                                start_col: *col,
                                start_row: *row,
                                end_col: *end_col,
                                end_row: *end_row,
                                start_col_abs: *col_abs,
                                start_row_abs: *row_abs,
                                end_col_abs: *end_col_abs,
                                end_row_abs: *end_row_abs,
                            },
                            pos + 3,
                        ));
                    }
                }
            }
            Ok((Expr::CellRef { sheet: UnboundSheetRef::Current, col: *col, row: *row, col_abs: *col_abs, row_abs: *row_abs }, pos + 1))
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
        Token::Plus => {
            // Unary plus (no-op, just parse the next expression)
            parse_primary(tokens, pos + 1)
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

fn parse_function_args(tokens: &[Token], pos: usize) -> Result<(Vec<ParsedExpr>, usize), String> {
    let mut args = Vec::new();
    let mut pos = pos;

    // Handle empty function call SUM()
    if pos < tokens.len() {
        if let Token::RParen = &tokens[pos] {
            return Ok((args, pos + 1));
        }
    }

    loop {
        // Empty argument: next token is , or ) immediately
        if pos < tokens.len() && matches!(&tokens[pos], Token::Comma | Token::RParen) {
            args.push(Expr::Empty);
            match &tokens[pos] {
                Token::RParen => return Ok((args, pos + 1)),
                Token::Comma => { pos += 1; continue; }
                _ => unreachable!(),
            }
        }

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

// =============================================================================
// Expression Binding - Convert ParsedExpr to BoundExpr
// =============================================================================

/// Bind a parsed expression by resolving sheet names to SheetIds.
///
/// The resolver function takes a sheet name and returns:
/// - Some(SheetId) if the sheet exists
/// - None if the sheet doesn't exist (will become #REF! error)
pub fn bind_expr<F>(expr: &ParsedExpr, resolver: F) -> BoundExpr
where
    F: Fn(&str) -> Option<SheetId> + Copy,
{
    match expr {
        Expr::Empty => Expr::Empty,
        Expr::Number(n) => Expr::Number(*n),
        Expr::Text(s) => Expr::Text(s.clone()),
        Expr::Boolean(b) => Expr::Boolean(*b),
        Expr::NamedRange(name) => Expr::NamedRange(name.clone()),
        Expr::CellRef { sheet, col, row, col_abs, row_abs } => {
            let bound_sheet = bind_sheet_ref(sheet, resolver);
            Expr::CellRef {
                sheet: bound_sheet,
                col: *col,
                row: *row,
                col_abs: *col_abs,
                row_abs: *row_abs,
            }
        }
        Expr::Range { sheet, start_col, start_row, end_col, end_row, start_col_abs, start_row_abs, end_col_abs, end_row_abs } => {
            let bound_sheet = bind_sheet_ref(sheet, resolver);
            Expr::Range {
                sheet: bound_sheet,
                start_col: *start_col,
                start_row: *start_row,
                end_col: *end_col,
                end_row: *end_row,
                start_col_abs: *start_col_abs,
                start_row_abs: *start_row_abs,
                end_col_abs: *end_col_abs,
                end_row_abs: *end_row_abs,
            }
        }
        Expr::Function { name, args } => {
            let bound_args = args.iter().map(|arg| bind_expr(arg, resolver)).collect();
            Expr::Function {
                name: name.clone(),
                args: bound_args,
            }
        }
        Expr::BinaryOp { op, left, right } => {
            Expr::BinaryOp {
                op: *op,
                left: Box::new(bind_expr(left, resolver)),
                right: Box::new(bind_expr(right, resolver)),
            }
        }
    }
}

/// Bind a sheet reference from unbound (name) to bound (ID).
fn bind_sheet_ref<F>(sheet: &UnboundSheetRef, resolver: F) -> SheetRef
where
    F: Fn(&str) -> Option<SheetId>,
{
    match sheet {
        UnboundSheetRef::Current => SheetRef::Current,
        UnboundSheetRef::Named(name) => {
            match resolver(name) {
                Some(id) => SheetRef::Id(id),
                None => SheetRef::RefError {
                    id: SheetId::from_raw(0), // Placeholder for unknown sheet
                    last_known_name: name.clone(),
                },
            }
        }
    }
}

/// Bind a parsed expression for same-sheet formulas (no cross-sheet references).
/// This is a convenience function that treats all Named refs as errors.
pub fn bind_expr_same_sheet(expr: &ParsedExpr) -> BoundExpr {
    bind_expr(expr, |_name| None)
}

// =============================================================================
// Formula Printing - Convert BoundExpr back to string
// =============================================================================

/// Format a bound expression as a formula string (with leading '=').
///
/// The `name_resolver` function takes a SheetId and returns the current sheet name.
/// This allows formulas to display updated names after sheet renames.
pub fn format_expr<F>(expr: &BoundExpr, name_resolver: F) -> String
where
    F: Fn(SheetId) -> Option<String> + Copy,
{
    format!("={}", format_expr_inner(expr, name_resolver))
}

/// Format a bound expression without the leading '='.
pub fn format_expr_inner<F>(expr: &BoundExpr, name_resolver: F) -> String
where
    F: Fn(SheetId) -> Option<String> + Copy,
{
    match expr {
        Expr::Empty => String::new(),
        Expr::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        Expr::Text(s) => format!("\"{}\"", s.replace('"', "\"\"")),
        Expr::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        Expr::NamedRange(name) => name.clone(),
        Expr::CellRef { sheet, col, row, col_abs, row_abs } => {
            let prefix = format_sheet_prefix(sheet, name_resolver);
            let addr = format_cell_addr(*col, *row, *col_abs, *row_abs);
            format!("{}{}", prefix, addr)
        }
        Expr::Range { sheet, start_col, start_row, end_col, end_row, start_col_abs, start_row_abs, end_col_abs, end_row_abs } => {
            let prefix = format_sheet_prefix(sheet, name_resolver);
            let start = format_cell_addr(*start_col, *start_row, *start_col_abs, *start_row_abs);
            let end = format_cell_addr(*end_col, *end_row, *end_col_abs, *end_row_abs);
            format!("{}{}:{}", prefix, start, end)
        }
        Expr::Function { name, args } => {
            let args_str: Vec<String> = args.iter()
                .map(|arg| format_expr_inner(arg, name_resolver))
                .collect();
            format!("{}({})", name, args_str.join(","))
        }
        Expr::BinaryOp { op, left, right } => {
            let left_str = format_expr_inner(left, name_resolver);
            let right_str = format_expr_inner(right, name_resolver);
            let op_str = match op {
                Op::Add => "+",
                Op::Sub => "-",
                Op::Mul => "*",
                Op::Div => "/",
                Op::Lt => "<",
                Op::Gt => ">",
                Op::Eq => "=",
                Op::LtEq => "<=",
                Op::GtEq => ">=",
                Op::NotEq => "<>",
                Op::Concat => "&",
                Op::Pow => "^",
            };
            format!("{}{}{}", left_str, op_str, right_str)
        }
    }
}

/// Format a sheet reference prefix (empty for Current, "SheetName!" for Id, "#REF!" for RefError)
fn format_sheet_prefix<F>(sheet: &SheetRef, name_resolver: F) -> String
where
    F: Fn(SheetId) -> Option<String>,
{
    match sheet {
        SheetRef::Current => String::new(),
        SheetRef::Id(id) => {
            match name_resolver(*id) {
                Some(name) => format!("{}!", format_sheet_name(&name)),
                None => "#REF!".to_string(), // Sheet was deleted
            }
        }
        SheetRef::RefError { last_known_name, .. } => {
            // Show #REF! - the original name is lost for display purposes
            // but we keep last_known_name for potential future "restore" features
            let _ = last_known_name; // Acknowledge we have it but don't use for now
            "#REF!".to_string()
        }
    }
}

/// Format a sheet name, adding quotes if necessary
pub fn format_sheet_name(name: &str) -> String {
    // Need quotes if name contains spaces, special characters, or starts with a digit
    let needs_quotes = name.contains(' ')
        || name.contains('!')
        || name.contains('\'')
        || name.contains(':')
        || name.contains('[')
        || name.contains(']')
        || name.chars().next().map(|c| c.is_ascii_digit()).unwrap_or(false);

    if needs_quotes {
        // Escape single quotes by doubling them
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

/// Format a cell address in A1 notation
fn format_cell_addr(col: usize, row: usize, col_abs: bool, row_abs: bool) -> String {
    let col_str = if col_abs {
        format!("${}", col_to_letters(col))
    } else {
        col_to_letters(col)
    };
    let row_str = if row_abs {
        format!("${}", row + 1)
    } else {
        format!("{}", row + 1)
    };
    format!("{}{}", col_str, row_str)
}

/// Convert column index to letter(s): 0 -> A, 25 -> Z, 26 -> AA, etc.
pub(crate) fn col_to_letters(col: usize) -> String {
    let mut result = String::new();
    let mut n = col + 1; // 1-indexed for calculation
    while n > 0 {
        n -= 1;
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    result
}

// =============================================================================
// Cell Reference Extraction
// =============================================================================

/// Extract all cell references from an expression (for dependency tracking)
/// Returns a list of (row, col) tuples for single cells and expanded ranges
pub fn extract_cell_refs<S>(expr: &Expr<S>) -> Vec<(usize, usize)> {
    let mut refs = Vec::new();
    collect_cell_refs(expr, &mut refs);
    refs
}

fn collect_cell_refs<S>(expr: &Expr<S>, refs: &mut Vec<(usize, usize)>) {
    match expr {
        Expr::Number(_) | Expr::Text(_) | Expr::Boolean(_) | Expr::NamedRange(_) | Expr::Empty => {
            // NamedRange refs are resolved at evaluation time with access to NamedRangeStore
        }
        Expr::CellRef { col, row, .. } => {
            refs.push((*row, *col));
        }
        Expr::Range { start_col, start_row, end_col, end_row, .. } => {
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

#[cfg(test)]
mod tests {
    use super::*;

    // =========================================================================
    // Absolute reference ($) parsing tests
    // =========================================================================

    #[test]
    fn test_parse_absolute_both() {
        // =$A$1
        let expr = parse("=$A$1").unwrap();
        match expr {
            Expr::CellRef { col, row, col_abs, row_abs, .. } => {
                assert_eq!(col, 0);
                assert_eq!(row, 0);
                assert!(col_abs, "col should be absolute");
                assert!(row_abs, "row should be absolute");
            }
            _ => panic!("Expected CellRef, got {:?}", expr),
        }
    }

    #[test]
    fn test_parse_absolute_col_only() {
        // =$A1
        let expr = parse("=$A1").unwrap();
        match expr {
            Expr::CellRef { col, row, col_abs, row_abs, .. } => {
                assert_eq!(col, 0);
                assert_eq!(row, 0);
                assert!(col_abs, "col should be absolute");
                assert!(!row_abs, "row should be relative");
            }
            _ => panic!("Expected CellRef, got {:?}", expr),
        }
    }

    #[test]
    fn test_parse_absolute_row_only() {
        // =A$1
        let expr = parse("=A$1").unwrap();
        match expr {
            Expr::CellRef { col, row, col_abs, row_abs, .. } => {
                assert_eq!(col, 0);
                assert_eq!(row, 0);
                assert!(!col_abs, "col should be relative");
                assert!(row_abs, "row should be absolute");
            }
            _ => panic!("Expected CellRef, got {:?}", expr),
        }
    }

    #[test]
    fn test_parse_relative() {
        // =A1
        let expr = parse("=A1").unwrap();
        match expr {
            Expr::CellRef { col_abs, row_abs, .. } => {
                assert!(!col_abs, "col should be relative");
                assert!(!row_abs, "row should be relative");
            }
            _ => panic!("Expected CellRef, got {:?}", expr),
        }
    }

    #[test]
    fn test_parse_absolute_multi_letter_col() {
        // =$O$95 (column O = index 14)
        let expr = parse("=$O$95").unwrap();
        match expr {
            Expr::CellRef { col, row, col_abs, row_abs, .. } => {
                assert_eq!(col, 14); // O = 14
                assert_eq!(row, 94); // 95 - 1 = 94
                assert!(col_abs);
                assert!(row_abs);
            }
            _ => panic!("Expected CellRef, got {:?}", expr),
        }
    }

    #[test]
    fn test_parse_absolute_range() {
        // =$O$95:$O$100
        let expr = parse("=$O$95:$O$100").unwrap();
        match expr {
            Expr::Range {
                start_col, start_row, end_col, end_row,
                start_col_abs, start_row_abs, end_col_abs, end_row_abs, ..
            } => {
                assert_eq!(start_col, 14);
                assert_eq!(start_row, 94);
                assert_eq!(end_col, 14);
                assert_eq!(end_row, 99);
                assert!(start_col_abs);
                assert!(start_row_abs);
                assert!(end_col_abs);
                assert!(end_row_abs);
            }
            _ => panic!("Expected Range, got {:?}", expr),
        }
    }

    #[test]
    fn test_parse_mixed_range() {
        // =$A1:A$5 (mixed absolute/relative)
        let expr = parse("=$A1:A$5").unwrap();
        match expr {
            Expr::Range {
                start_col_abs, start_row_abs, end_col_abs, end_row_abs, ..
            } => {
                assert!(start_col_abs, "$A");
                assert!(!start_row_abs, "1 relative");
                assert!(!end_col_abs, "A relative");
                assert!(end_row_abs, "$5");
            }
            _ => panic!("Expected Range, got {:?}", expr),
        }
    }

    // =========================================================================
    // Round-trip: parse → format_expr → parse again
    // =========================================================================

    #[test]
    fn test_roundtrip_absolute_both() {
        let parsed = parse("=$A$1").unwrap();
        let bound = bind_expr_same_sheet(&parsed);
        let formatted = format_expr(&bound, |_| None);
        assert_eq!(formatted, "=$A$1");
    }

    #[test]
    fn test_roundtrip_absolute_col() {
        let parsed = parse("=$A1").unwrap();
        let bound = bind_expr_same_sheet(&parsed);
        let formatted = format_expr(&bound, |_| None);
        assert_eq!(formatted, "=$A1");
    }

    #[test]
    fn test_roundtrip_absolute_row() {
        let parsed = parse("=A$1").unwrap();
        let bound = bind_expr_same_sheet(&parsed);
        let formatted = format_expr(&bound, |_| None);
        assert_eq!(formatted, "=A$1");
    }

    #[test]
    fn test_roundtrip_absolute_range() {
        let parsed = parse("=$O$95:$O$100").unwrap();
        let bound = bind_expr_same_sheet(&parsed);
        let formatted = format_expr(&bound, |_| None);
        assert_eq!(formatted, "=$O$95:$O$100");
    }

    #[test]
    fn test_roundtrip_mixed_range() {
        let parsed = parse("=$A1:A$5").unwrap();
        let bound = bind_expr_same_sheet(&parsed);
        let formatted = format_expr(&bound, |_| None);
        assert_eq!(formatted, "=$A1:A$5");
    }

    #[test]
    fn test_absolute_in_formula() {
        // =SUM($A$1:$A$10)+B2
        let parsed = parse("=SUM($A$1:$A$10)+B2").unwrap();
        let bound = bind_expr_same_sheet(&parsed);
        let formatted = format_expr(&bound, |_| None);
        assert_eq!(formatted, "=SUM($A$1:$A$10)+B2");
    }

    // =========================================================================
    // Power (^) and Percent (%) parsing tests
    // =========================================================================

    #[test]
    fn test_parse_power() {
        let expr = parse("=2^3").unwrap();
        match expr {
            Expr::BinaryOp { op: Op::Pow, .. } => {}
            _ => panic!("Expected Pow op, got {:?}", expr),
        }
    }

    #[test]
    fn test_parse_percent() {
        // 50% desugars to 50*0.01
        let expr = parse("=50%").unwrap();
        match expr {
            Expr::BinaryOp { op: Op::Mul, ref right, .. } => {
                match right.as_ref() {
                    Expr::Number(n) => assert_eq!(*n, 0.01),
                    _ => panic!("Expected Number(0.01), got {:?}", right),
                }
            }
            _ => panic!("Expected Mul op (desugared percent), got {:?}", expr),
        }
    }

    #[test]
    fn test_roundtrip_power() {
        let parsed = parse("=A1^2").unwrap();
        let bound = bind_expr_same_sheet(&parsed);
        let formatted = format_expr(&bound, |_| None);
        assert_eq!(formatted, "=A1^2");
    }

    // ── Unary plus tests ──────────────────────────────────────────

    #[test]
    fn test_unary_plus_cell_ref() {
        // =+A1 should parse to just A1 (unary plus is a no-op)
        let expr = parse("=+A1").unwrap();
        match &expr {
            Expr::CellRef { col, row, .. } => {
                assert_eq!(*col, 0); // A
                assert_eq!(*row, 0); // 1
            }
            _ => panic!("Expected CellRef, got {:?}", expr),
        }
    }

    #[test]
    fn test_unary_plus_number() {
        // =+1 should parse to just 1
        let expr = parse("=+1").unwrap();
        match &expr {
            Expr::Number(n) => assert_eq!(*n, 1.0),
            _ => panic!("Expected Number(1), got {:?}", expr),
        }
    }

    #[test]
    fn test_unary_plus_minus() {
        // =+-A1 should parse to unary minus of A1 (i.e. 0 - A1)
        let expr = parse("=+-A1").unwrap();
        match &expr {
            Expr::BinaryOp { op: Op::Sub, left, right } => {
                match left.as_ref() {
                    Expr::Number(n) => assert_eq!(*n, 0.0),
                    _ => panic!("Expected Number(0) on left"),
                }
                match right.as_ref() {
                    Expr::CellRef { col, row, .. } => {
                        assert_eq!(*col, 0);
                        assert_eq!(*row, 0);
                    }
                    _ => panic!("Expected CellRef on right"),
                }
            }
            _ => panic!("Expected Sub op (unary minus), got {:?}", expr),
        }
    }

    #[test]
    fn test_unary_plus_in_expression() {
        // =+A1/B2 should parse as A1/B2
        let expr = parse("=+A1/B2").unwrap();
        match &expr {
            Expr::BinaryOp { op: Op::Div, left, right } => {
                match left.as_ref() {
                    Expr::CellRef { col, row, .. } => {
                        assert_eq!(*col, 0); // A
                        assert_eq!(*row, 0); // 1
                    }
                    _ => panic!("Expected CellRef(A1) on left"),
                }
                match right.as_ref() {
                    Expr::CellRef { col, row, .. } => {
                        assert_eq!(*col, 1); // B
                        assert_eq!(*row, 1); // 2
                    }
                    _ => panic!("Expected CellRef(B2) on right"),
                }
            }
            _ => panic!("Expected Div op, got {:?}", expr),
        }
    }

    #[test]
    fn test_unary_plus_chained() {
        // =++1 should parse to just 1 (double unary plus)
        let expr = parse("=++1").unwrap();
        match &expr {
            Expr::Number(n) => assert_eq!(*n, 1.0),
            _ => panic!("Expected Number(1), got {:?}", expr),
        }
    }

    #[test]
    fn test_unary_plus_roundtrip_drops_plus() {
        // Round-trip formatting should drop the unary plus (it's a no-op)
        let parsed = parse("=+A1/B2").unwrap();
        let bound = bind_expr_same_sheet(&parsed);
        let formatted = format_expr(&bound, |_| None);
        assert_eq!(formatted, "=A1/B2");
    }

    #[test]
    fn test_unary_plus_with_function() {
        // =+SUM(A1:A3) should parse as SUM(A1:A3)
        let expr = parse("=+SUM(A1:A3)").unwrap();
        match &expr {
            Expr::Function { name, .. } => {
                assert_eq!(name, "SUM");
            }
            _ => panic!("Expected FunctionCall(SUM), got {:?}", expr),
        }
    }

    #[test]
    fn test_unary_plus_multiplication() {
        // =+H14*12 — real-world Excel formula from Haven import
        let expr = parse("=+H14*12").unwrap();
        match &expr {
            Expr::BinaryOp { op: Op::Mul, .. } => {}
            _ => panic!("Expected Mul op, got {:?}", expr),
        }
    }

    // ── Empty argument tests ─────────────────────────────────────

    fn extract_func_args(expr: &ParsedExpr) -> &[ParsedExpr] {
        match expr {
            Expr::Function { args, .. } => args,
            _ => panic!("Expected Function, got {:?}", expr),
        }
    }

    #[test]
    fn test_empty_arg_trailing() {
        // =IF(A1,B1,) → [CellRef, CellRef, Empty]
        let expr = parse("=IF(A1,B1,)").unwrap();
        let args = extract_func_args(&expr);
        assert_eq!(args.len(), 3);
        assert!(matches!(&args[0], Expr::CellRef { .. }));
        assert!(matches!(&args[1], Expr::CellRef { .. }));
        assert!(matches!(&args[2], Expr::Empty));
    }

    #[test]
    fn test_empty_arg_middle() {
        // =IF(A1,,C1) → [CellRef, Empty, CellRef]
        let expr = parse("=IF(A1,,C1)").unwrap();
        let args = extract_func_args(&expr);
        assert_eq!(args.len(), 3);
        assert!(matches!(&args[0], Expr::CellRef { .. }));
        assert!(matches!(&args[1], Expr::Empty));
        assert!(matches!(&args[2], Expr::CellRef { .. }));
    }

    #[test]
    fn test_empty_arg_leading() {
        // =IF(,A1,B1) → [Empty, CellRef, CellRef]
        let expr = parse("=IF(,A1,B1)").unwrap();
        let args = extract_func_args(&expr);
        assert_eq!(args.len(), 3);
        assert!(matches!(&args[0], Expr::Empty));
        assert!(matches!(&args[1], Expr::CellRef { .. }));
        assert!(matches!(&args[2], Expr::CellRef { .. }));
    }

    #[test]
    fn test_empty_arg_all_empty() {
        // =FUNC(,,) → [Empty, Empty, Empty]
        let expr = parse("=FUNC(,,)").unwrap();
        let args = extract_func_args(&expr);
        assert_eq!(args.len(), 3);
        assert!(args.iter().all(|a| matches!(a, Expr::Empty)));
    }

    #[test]
    fn test_empty_arg_no_regression_sum_empty() {
        // =SUM() → [] (no args, not [Empty])
        let expr = parse("=SUM()").unwrap();
        let args = extract_func_args(&expr);
        assert_eq!(args.len(), 0);
    }

    #[test]
    fn test_empty_arg_real_world_pattern() {
        // =FUNC(A1,B2,) — exact failure pattern from fcffsimpleginzu.xlsx import
        let expr = parse("=FUNC(A1,B2,)").unwrap();
        let args = extract_func_args(&expr);
        assert_eq!(args.len(), 3);
        assert!(matches!(&args[0], Expr::CellRef { .. }));
        assert!(matches!(&args[1], Expr::CellRef { .. }));
        assert!(matches!(&args[2], Expr::Empty));
    }

    #[test]
    fn test_empty_arg_roundtrip() {
        // =IF(A1,B1,) roundtrip through format_expr
        let parsed = parse("=IF(A1,B1,)").unwrap();
        let bound = bind_expr_same_sheet(&parsed);
        let formatted = format_expr(&bound, |_| None);
        assert_eq!(formatted, "=IF(A1,B1,)");
    }

    // ── Dotted function name tests ───────────────────────────────

    #[test]
    fn test_dotted_function_name() {
        // =NORM.S.DIST(0,TRUE) should parse as a single function
        let expr = parse("=NORM.S.DIST(0,TRUE)").unwrap();
        match &expr {
            Expr::Function { name, args } => {
                assert_eq!(name, "NORM.S.DIST");
                assert_eq!(args.len(), 2);
            }
            _ => panic!("Expected Function, got {:?}", expr),
        }
    }

    #[test]
    fn test_dotted_function_stdev_p() {
        let expr = parse("=STDEV.P(A1:A10)").unwrap();
        match &expr {
            Expr::Function { name, .. } => assert_eq!(name, "STDEV.P"),
            _ => panic!("Expected Function, got {:?}", expr),
        }
    }

    #[test]
    fn test_decimal_numbers_not_broken() {
        // Dot in numbers must still work
        let expr = parse("=1.23+4.56").unwrap();
        match &expr {
            Expr::BinaryOp { op: Op::Add, left, right } => {
                match left.as_ref() {
                    Expr::Number(n) => assert!((n - 1.23).abs() < 1e-10),
                    _ => panic!("Expected Number on left"),
                }
                match right.as_ref() {
                    Expr::Number(n) => assert!((n - 4.56).abs() < 1e-10),
                    _ => panic!("Expected Number on right"),
                }
            }
            _ => panic!("Expected Add op, got {:?}", expr),
        }
    }

    #[test]
    fn test_decimal_multiply_not_broken() {
        // =A1*0.5 must still work
        let expr = parse("=A1*0.5").unwrap();
        match &expr {
            Expr::BinaryOp { op: Op::Mul, right, .. } => {
                match right.as_ref() {
                    Expr::Number(n) => assert!((n - 0.5).abs() < 1e-10),
                    _ => panic!("Expected Number(0.5) on right"),
                }
            }
            _ => panic!("Expected Mul op, got {:?}", expr),
        }
    }
}
