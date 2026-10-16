// QUERY(data, query, [headers]) (Google Sheets): the Google Visualization API
// query language over a range or array.
//
// Clauses, in the order the language requires: select, where, group by, pivot,
// order by, limit, offset, label. `format` and `options` are refused with a
// message rather than ignored, so a result never looks formatted when it isn't.
//
// Columns are named by their sheet letters when data is a range (QUERY(D1:F9,
// "select E")), or Col1, Col2... (accepted for ranges too). Each column takes
// the type most of its values have; values of another type read as null, as in
// Sheets. Strings compare case-sensitively. Grouped results come out sorted by
// their group-by values, and aggregates head their column as "sum Amount".

use super::eval::{evaluate, CellLookup, EvalResult, Value};
use super::eval_helpers::{date_to_serial, serial_to_date};
use super::parser::{BoundExpr, Expr, RangeAxis};
use std::cmp::Ordering;

pub(super) fn evaluate_query<L: CellLookup>(args: &[BoundExpr], lookup: &L) -> EvalResult {
    if args.len() < 2 || args.len() > 3 {
        return EvalResult::Error("QUERY requires 2 or 3 arguments".to_string());
    }
    // Sheet letters name a range's columns; arrays use Col1, Col2...
    let first_col = match &args[0] {
        Expr::Range {
            start_col, end_col, ..
        } => Some(*start_col.min(end_col)),
        Expr::WholeRange {
            axis: RangeAxis::Column,
            start,
            end,
            ..
        } => Some(*start.min(end)),
        Expr::WholeRange {
            axis: RangeAxis::Row,
            ..
        } => Some(0),
        _ => None,
    };
    let (rows, cols, data) = match super::eval_array::grid_values(&args[0], lookup) {
        Ok(v) => v,
        Err(e) => return EvalResult::Error(e),
    };
    let text = match evaluate(&args[1], lookup) {
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Text(t) => t,
        EvalResult::Empty => String::new(),
        _ => return EvalResult::Error("#VALUE! QUERY's query must be text".to_string()),
    };
    let headers = match args
        .get(2)
        .filter(|a| !matches!(a, Expr::Empty))
        .map(|a| evaluate(a, lookup).to_number())
    {
        None => None,
        Some(Ok(n)) if n < 0.0 => None,
        Some(Ok(n)) => Some((n as usize).min(rows)),
        Some(Err(e)) => return EvalResult::Error(e),
    };
    match run(&data, cols, first_col, &text, headers) {
        Ok(out) => out,
        Err(message) => EvalResult::Error(message),
    }
}

// ---------------------------------------------------------------- tokens

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Word(String),
    Backtick(String),
    Str(String),
    Num(f64),
    Sym(&'static str),
}

fn tokenize(src: &str) -> Result<Vec<Tok>, String> {
    let chars: Vec<char> = src.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c.is_whitespace() {
            i += 1;
        } else if c == '\'' || c == '"' || c == '`' {
            // No escapes in this language: a quote of the other kind goes inside.
            let start = i + 1;
            let end = chars[start..]
                .iter()
                .position(|&x| x == c)
                .map(|p| start + p)
                .ok_or_else(|| err("a quote is not closed"))?;
            let s: String = chars[start..end].iter().collect();
            out.push(if c == '`' {
                Tok::Backtick(s)
            } else {
                Tok::Str(s)
            });
            i = end + 1;
        } else if c.is_ascii_digit()
            || (c == '.' && chars.get(i + 1).is_some_and(|d| d.is_ascii_digit()))
        {
            let start = i;
            while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                i += 1;
            }
            let s: String = chars[start..i].iter().collect();
            out.push(Tok::Num(
                s.parse()
                    .map_err(|_| err(&format!("{s} is not a number")))?,
            ));
        } else if c.is_alphanumeric() || c == '_' {
            let start = i;
            while i < chars.len() && (chars[i].is_alphanumeric() || chars[i] == '_') {
                i += 1;
            }
            out.push(Tok::Word(chars[start..i].iter().collect()));
        } else {
            let two: String = chars[i..(i + 2).min(chars.len())].iter().collect();
            let sym = ["!=", "<>", "<=", ">="].into_iter().find(|s| *s == two);
            let sym = match sym {
                Some(s) => s,
                None => match c {
                    '=' => "=",
                    '<' => "<",
                    '>' => ">",
                    '+' => "+",
                    '-' => "-",
                    '*' => "*",
                    '/' => "/",
                    '(' => "(",
                    ')' => ")",
                    ',' => ",",
                    _ => return Err(err(&format!("unexpected character {c}"))),
                },
            };
            i += sym.len();
            out.push(Tok::Sym(sym));
        }
    }
    Ok(out)
}

fn err(message: &str) -> String {
    format!("#VALUE! QUERY: {message}")
}

// ---------------------------------------------------------------- syntax

#[derive(Debug, Clone, Copy, PartialEq)]
enum Agg {
    Sum,
    Avg,
    Count,
    Min,
    Max,
}

impl Agg {
    fn name(self) -> &'static str {
        match self {
            Agg::Sum => "sum",
            Agg::Avg => "avg",
            Agg::Count => "count",
            Agg::Min => "min",
            Agg::Max => "max",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Cmp {
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    Contains,
    StartsWith,
    EndsWith,
    Matches,
    Like,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Scalar {
    Year,
    Month,
    Day,
    Upper,
    Lower,
}

#[derive(Debug, Clone, PartialEq)]
enum Q {
    /// A column, as an index into the data.
    Col(usize),
    Lit(Value),
    Agg(Agg, Box<Q>),
    Fn(Scalar, Box<Q>),
    Neg(Box<Q>),
    Arith(char, Box<Q>, Box<Q>),
    Cmp(Cmp, Box<Q>, Box<Q>),
    And(Box<Q>, Box<Q>),
    Or(Box<Q>, Box<Q>),
    Not(Box<Q>),
    IsNull(Box<Q>, bool),
}

impl Q {
    fn has_agg(&self) -> bool {
        match self {
            Q::Agg(..) => true,
            Q::Col(_) | Q::Lit(_) => false,
            Q::Fn(_, a) | Q::Neg(a) | Q::Not(a) | Q::IsNull(a, _) => a.has_agg(),
            Q::Arith(_, a, b) | Q::Cmp(_, a, b) | Q::And(a, b) | Q::Or(a, b) => {
                a.has_agg() || b.has_agg()
            }
        }
    }
}

#[derive(Default)]
struct Query {
    select: Option<Vec<Q>>,
    filter: Option<Q>,
    group_by: Vec<Q>,
    pivot: Vec<Q>,
    order_by: Vec<(Q, bool)>,
    limit: Option<usize>,
    offset: usize,
    labels: Vec<(Q, String)>,
}

struct Parser {
    toks: Vec<Tok>,
    pos: usize,
    cols: usize,
    first_col: Option<usize>,
}

impl Parser {
    fn peek_word(&self, word: &str) -> bool {
        matches!(self.toks.get(self.pos), Some(Tok::Word(w)) if w.eq_ignore_ascii_case(word))
    }
    fn peek_words(&self, words: &[&str]) -> bool {
        words.iter().enumerate().all(|(k, w)| matches!(self.toks.get(self.pos + k), Some(Tok::Word(t)) if t.eq_ignore_ascii_case(w)))
    }
    fn eat_word(&mut self, word: &str) -> bool {
        let yes = self.peek_word(word);
        if yes {
            self.pos += 1;
        }
        yes
    }
    fn eat_sym(&mut self, sym: &str) -> bool {
        let yes = matches!(self.toks.get(self.pos), Some(Tok::Sym(s)) if *s == sym);
        if yes {
            self.pos += 1;
        }
        yes
    }
    fn expect_sym(&mut self, sym: &str) -> Result<(), String> {
        if self.eat_sym(sym) {
            Ok(())
        } else {
            Err(err(&format!("expected {sym}")))
        }
    }
    fn query(&mut self) -> Result<Query, String> {
        let mut q = Query::default();
        if self.eat_word("select") {
            if self.eat_sym("*") {
                q.select = None;
            } else {
                q.select = Some(self.list()?);
            }
        }
        if self.eat_word("where") {
            q.filter = Some(self.expr()?);
        }
        if self.peek_words(&["group", "by"]) {
            self.pos += 2;
            q.group_by = self.list()?;
        }
        if self.eat_word("pivot") {
            q.pivot = self.list()?;
        }
        if self.peek_words(&["order", "by"]) {
            self.pos += 2;
            loop {
                let e = self.expr()?;
                let asc = if self.eat_word("desc") {
                    false
                } else {
                    self.eat_word("asc");
                    true
                };
                q.order_by.push((e, asc));
                if !self.eat_sym(",") {
                    break;
                }
            }
        }
        if self.eat_word("limit") {
            q.limit = Some(self.count("limit")?);
        }
        if self.eat_word("offset") {
            q.offset = self.count("offset")?;
        }
        if self.eat_word("label") {
            loop {
                let e = self.expr()?;
                match self.toks.get(self.pos) {
                    Some(Tok::Str(s)) => {
                        q.labels.push((e, s.clone()));
                        self.pos += 1;
                    }
                    _ => return Err(err("label needs a column and a quoted name")),
                }
                if !self.eat_sym(",") {
                    break;
                }
            }
        }
        for clause in ["format", "options"] {
            if self.peek_word(clause) {
                return Err(err(&format!("the {clause} clause isn't supported yet")));
            }
        }
        match self.toks.get(self.pos) {
            None => Ok(q),
            Some(t) => Err(err(&format!("unexpected {} (clauses go in the order select, where, group by, pivot, order by, limit, offset, label)", describe(t)))),
        }
    }

    fn count(&mut self, clause: &str) -> Result<usize, String> {
        match self.toks.get(self.pos) {
            Some(Tok::Num(n)) if *n >= 0.0 && n.fract() == 0.0 => {
                self.pos += 1;
                Ok(*n as usize)
            }
            _ => Err(err(&format!("{clause} needs a whole number"))),
        }
    }

    fn list(&mut self) -> Result<Vec<Q>, String> {
        let mut out = vec![self.expr()?];
        while self.eat_sym(",") {
            out.push(self.expr()?);
        }
        Ok(out)
    }

    fn expr(&mut self) -> Result<Q, String> {
        let mut left = self.and()?;
        while self.eat_word("or") {
            left = Q::Or(Box::new(left), Box::new(self.and()?));
        }
        Ok(left)
    }
    fn and(&mut self) -> Result<Q, String> {
        let mut left = self.not()?;
        while self.eat_word("and") {
            left = Q::And(Box::new(left), Box::new(self.not()?));
        }
        Ok(left)
    }
    fn not(&mut self) -> Result<Q, String> {
        if self.eat_word("not") {
            return Ok(Q::Not(Box::new(self.not()?)));
        }
        self.comparison()
    }
    fn comparison(&mut self) -> Result<Q, String> {
        let left = self.additive()?;
        if self.eat_word("is") {
            let negated = self.eat_word("not");
            if !self.eat_word("null") {
                return Err(err("expected null after is"));
            }
            return Ok(Q::IsNull(Box::new(left), negated));
        }
        let op = if self.eat_sym("=") {
            Cmp::Eq
        } else if self.eat_sym("!=") || self.eat_sym("<>") {
            Cmp::Ne
        } else if self.eat_sym("<=") {
            Cmp::Le
        } else if self.eat_sym(">=") {
            Cmp::Ge
        } else if self.eat_sym("<") {
            Cmp::Lt
        } else if self.eat_sym(">") {
            Cmp::Gt
        } else if self.eat_word("contains") {
            Cmp::Contains
        } else if self.peek_words(&["starts", "with"]) {
            self.pos += 2;
            Cmp::StartsWith
        } else if self.peek_words(&["ends", "with"]) {
            self.pos += 2;
            Cmp::EndsWith
        } else if self.eat_word("matches") {
            Cmp::Matches
        } else if self.eat_word("like") {
            Cmp::Like
        } else {
            return Ok(left);
        };
        Ok(Q::Cmp(op, Box::new(left), Box::new(self.additive()?)))
    }
    fn additive(&mut self) -> Result<Q, String> {
        let mut left = self.term()?;
        loop {
            let op = if self.eat_sym("+") {
                '+'
            } else if self.eat_sym("-") {
                '-'
            } else {
                return Ok(left);
            };
            left = Q::Arith(op, Box::new(left), Box::new(self.term()?));
        }
    }
    fn term(&mut self) -> Result<Q, String> {
        let mut left = self.unary()?;
        loop {
            let op = if self.eat_sym("*") {
                '*'
            } else if self.eat_sym("/") {
                '/'
            } else {
                return Ok(left);
            };
            left = Q::Arith(op, Box::new(left), Box::new(self.unary()?));
        }
    }
    fn unary(&mut self) -> Result<Q, String> {
        if self.eat_sym("-") {
            return Ok(Q::Neg(Box::new(self.unary()?)));
        }
        self.primary()
    }
    fn primary(&mut self) -> Result<Q, String> {
        let tok = self
            .toks
            .get(self.pos)
            .cloned()
            .ok_or_else(|| err("the query ends too soon"))?;
        self.pos += 1;
        match tok {
            Tok::Num(n) => Ok(Q::Lit(Value::Number(n))),
            Tok::Str(s) => Ok(Q::Lit(Value::Text(s))),
            Tok::Backtick(name) => self.column(&name),
            Tok::Sym("(") => {
                let e = self.expr()?;
                self.expect_sym(")")?;
                Ok(e)
            }
            Tok::Sym(s) => Err(err(&format!("unexpected {s}"))),
            Tok::Word(w) => {
                let lower = w.to_ascii_lowercase();
                match lower.as_str() {
                    "true" => return Ok(Q::Lit(Value::Boolean(true))),
                    "false" => return Ok(Q::Lit(Value::Boolean(false))),
                    "date" | "datetime" | "timeofday" => {
                        if let Some(Tok::Str(s)) = self.toks.get(self.pos).cloned() {
                            self.pos += 1;
                            return date_literal(&lower, &s).map(|n| Q::Lit(Value::Number(n)));
                        }
                    }
                    _ => {}
                }
                if matches!(self.toks.get(self.pos), Some(Tok::Sym("("))) {
                    self.pos += 1;
                    let agg = match lower.as_str() {
                        "sum" => Some(Agg::Sum),
                        "avg" => Some(Agg::Avg),
                        "count" => Some(Agg::Count),
                        "min" => Some(Agg::Min),
                        "max" => Some(Agg::Max),
                        _ => None,
                    };
                    let scalar = match lower.as_str() {
                        "year" => Some(Scalar::Year),
                        "month" => Some(Scalar::Month),
                        "day" => Some(Scalar::Day),
                        "upper" => Some(Scalar::Upper),
                        "lower" => Some(Scalar::Lower),
                        _ => None,
                    };
                    let arg = self.expr()?;
                    self.expect_sym(")")?;
                    return match (agg, scalar) {
                        (Some(a), _) => {
                            if arg.has_agg() {
                                return Err(err("aggregates can't be nested"));
                            }
                            Ok(Q::Agg(a, Box::new(arg)))
                        }
                        (_, Some(s)) => Ok(Q::Fn(s, Box::new(arg))),
                        _ => Err(err(&format!("the function {w}() isn't supported"))),
                    };
                }
                self.column(&w)
            }
        }
    }

    fn column(&self, name: &str) -> Result<Q, String> {
        let index = if let Some(n) = name
            .strip_prefix("Col")
            .and_then(|d| d.parse::<usize>().ok())
        {
            n.checked_sub(1)
        } else if !name.is_empty() && name.chars().all(|c| c.is_ascii_alphabetic()) {
            let letters = name
                .to_ascii_uppercase()
                .bytes()
                .fold(0usize, |acc, b| acc * 26 + (b - b'A' + 1) as usize)
                - 1;
            letters.checked_sub(self.first_col.unwrap_or(0))
        } else {
            None
        };
        match index {
            Some(i) if i < self.cols => Ok(Q::Col(i)),
            _ => Err(err(&format!("no column {name} in the data"))),
        }
    }
}

fn describe(t: &Tok) -> String {
    match t {
        Tok::Word(w) => w.clone(),
        Tok::Backtick(w) => format!("`{w}`"),
        Tok::Str(s) => format!("'{s}'"),
        Tok::Num(n) => n.to_string(),
        Tok::Sym(s) => s.to_string(),
    }
}

/// `date 'yyyy-mm-dd'`, `datetime 'yyyy-mm-dd HH:mm:ss'`, `timeofday 'HH:mm:ss'`
/// as serial numbers.
fn date_literal(kind: &str, s: &str) -> Result<f64, String> {
    let bad = || err(&format!("{kind} '{s}' is not valid"));
    let time = |t: &str| -> Result<f64, String> {
        let parts: Vec<f64> = t
            .split(':')
            .map(|p| p.parse::<f64>().map_err(|_| bad()))
            .collect::<Result<_, _>>()?;
        if parts.is_empty() || parts.len() > 3 {
            return Err(bad());
        }
        let (h, m, sec) = (
            parts[0],
            *parts.get(1).unwrap_or(&0.0),
            *parts.get(2).unwrap_or(&0.0),
        );
        Ok((h * 3600.0 + m * 60.0 + sec) / 86400.0)
    };
    let date = |d: &str| -> Result<f64, String> {
        let p: Vec<i32> = d
            .split('-')
            .map(|x| x.parse::<i32>().map_err(|_| bad()))
            .collect::<Result<_, _>>()?;
        match p.as_slice() {
            [y, m, d] if (1..=12).contains(m) && (1..=31).contains(d) => {
                Ok(date_to_serial(*y, *m, *d))
            }
            _ => Err(bad()),
        }
    };
    match kind {
        "date" => date(s.trim()),
        "timeofday" => time(s.trim()),
        _ => {
            let mut it = s.trim().splitn(2, ' ');
            let d = date(it.next().unwrap_or(""))?;
            Ok(d + it.next().map(time).transpose()?.unwrap_or(0.0))
        }
    }
}

// ---------------------------------------------------------------- evaluation

#[derive(Clone, Copy, PartialEq, PartialOrd)]
enum Kind {
    Number,
    Text,
    Boolean,
}

fn kind(v: &Value) -> Option<Kind> {
    match v {
        Value::Number(_) => Some(Kind::Number),
        Value::Text(t) if !t.is_empty() => Some(Kind::Text),
        Value::Boolean(_) => Some(Kind::Boolean),
        _ => None,
    }
}

/// How many leading rows are headers when the argument is omitted: one when
/// the first row is all text and some column below it holds mostly something
/// else; otherwise none (all-text data needs the argument, as in Sheets).
fn guess_headers(data: &[Vec<Value>], cols: usize) -> usize {
    let Some(first) = data.first() else { return 0 };
    if data.len() < 2 || !first.iter().all(|v| kind(v) == Some(Kind::Text)) {
        return 0;
    }
    let rest = &data[1..];
    (0..cols).any(|c| majority(rest, c).is_some_and(|k| k != Kind::Text)) as usize
}

/// A column's type: the one most of its values have (ties: number, text, boolean).
fn majority(rows: &[Vec<Value>], c: usize) -> Option<Kind> {
    let mut counts = [0usize; 3];
    for row in rows {
        if let Some(k) = kind(&row[c]) {
            counts[k as usize] += 1;
        }
    }
    let best = (0..3).max_by(|&a, &b| counts[a].cmp(&counts[b]).then(b.cmp(&a)))?;
    (counts[best] > 0).then(|| [Kind::Number, Kind::Text, Kind::Boolean][best])
}

fn run(
    data: &[Vec<Value>],
    cols: usize,
    first_col: Option<usize>,
    text: &str,
    headers: Option<usize>,
) -> Result<EvalResult, String> {
    let header_rows = headers.unwrap_or_else(|| guess_headers(data, cols));
    let labels: Vec<String> = (0..cols)
        .map(|c| {
            data[..header_rows]
                .iter()
                .map(|row| row[c].to_text().trim().to_string())
                .filter(|s| !s.is_empty())
                .collect::<Vec<_>>()
                .join(" ")
        })
        .collect();
    // Typed rows: a value of the column's minority type reads as null.
    let body = &data[header_rows..];
    let kinds: Vec<Option<Kind>> = (0..cols).map(|c| majority(body, c)).collect();
    let table: Vec<Vec<Value>> = body
        .iter()
        .map(|row| {
            row.iter()
                .zip(&kinds)
                .map(|(v, k)| {
                    if kind(v).is_some() && kind(v) == *k {
                        v.clone()
                    } else {
                        Value::Empty
                    }
                })
                .collect()
        })
        .collect();

    let names: Vec<String> = (0..cols)
        .map(|c| match first_col {
            Some(f) => column_letters(f + c),
            None => format!("Col{}", c + 1),
        })
        .collect();
    let mut parser = Parser {
        toks: tokenize(text)?,
        pos: 0,
        cols,
        first_col,
    };
    let q = parser.query()?;
    let select = q
        .select
        .clone()
        .unwrap_or_else(|| (0..cols).map(Q::Col).collect());

    if q.filter.as_ref().is_some_and(Q::has_agg) {
        return Err(err("where can't use aggregates"));
    }
    let rows: Vec<&Vec<Value>> = match &q.filter {
        Some(f) => {
            let mut kept = Vec::new();
            for row in &table {
                if truthy(&eval(f, &Ctx::Row(row))?) {
                    kept.push(row);
                }
            }
            kept
        }
        None => table.iter().collect(),
    };

    let grouped = !q.group_by.is_empty()
        || !q.pivot.is_empty()
        || select.iter().any(Q::has_agg)
        || q.order_by.iter().any(|(e, _)| e.has_agg());
    let header_for = |e: &Q| -> String {
        if let Some((_, l)) = q.labels.iter().find(|(le, _)| le == e) {
            return l.clone();
        }
        default_header(e, &labels, &names)
    };

    let (head, mut body_rows, order_keys): (Vec<String>, Vec<Vec<Value>>, Vec<Vec<Value>>) =
        if grouped {
            for e in select.iter().chain(q.order_by.iter().map(|(e, _)| e)) {
                check_grouped(e, &q.group_by, &q.pivot)?;
            }
            // Groups, sorted by their group-by values.
            let mut groups: Vec<(Vec<Value>, Vec<&Vec<Value>>)> = Vec::new();
            for row in &rows {
                let key = q
                    .group_by
                    .iter()
                    .map(|g| eval(g, &Ctx::Row(row)))
                    .collect::<Result<Vec<_>, _>>()?;
                match groups.iter_mut().find(|(k, _)| same_key(k, &key)) {
                    Some((_, members)) => members.push(row),
                    None => groups.push((key, vec![row])),
                }
            }
            groups.sort_by(|a, b| compare_keys(&a.0, &b.0));
            if groups.is_empty() && q.group_by.is_empty() {
                groups.push((Vec::new(), Vec::new()));
            }

            if q.pivot.is_empty() {
                let head = select.iter().map(header_for).collect();
                let mut out = Vec::new();
                let mut keys = Vec::new();
                for (_, members) in &groups {
                    let ctx = Ctx::Group(members);
                    out.push(
                        select
                            .iter()
                            .map(|e| eval(e, &ctx))
                            .collect::<Result<Vec<_>, _>>()?,
                    );
                    keys.push(
                        q.order_by
                            .iter()
                            .map(|(e, _)| eval(e, &ctx))
                            .collect::<Result<Vec<_>, _>>()?,
                    );
                }
                (head, out, keys)
            } else {
                // One column per pivot value (and aggregate), sorted by pivot value.
                let mut pivots: Vec<Vec<Value>> = Vec::new();
                for row in &rows {
                    let key = q
                        .pivot
                        .iter()
                        .map(|p| eval(p, &Ctx::Row(row)))
                        .collect::<Result<Vec<_>, _>>()?;
                    if !pivots.iter().any(|k| same_key(k, &key)) {
                        pivots.push(key);
                    }
                }
                pivots.sort_by(|a, b| compare_keys(a, b));
                let (plain, aggs): (Vec<&Q>, Vec<&Q>) = select.iter().partition(|e| !e.has_agg());
                let mut head: Vec<String> = plain.iter().map(|e| header_for(e)).collect();
                for p in &pivots {
                    let name = p.iter().map(display).collect::<Vec<_>>().join(", ");
                    for a in &aggs {
                        head.push(if aggs.len() == 1 {
                            name.clone()
                        } else {
                            format!("{name} {}", header_for(a))
                        });
                    }
                }
                let mut out = Vec::new();
                let mut keys = Vec::new();
                for (_, members) in &groups {
                    let mut row: Vec<Value> = plain
                        .iter()
                        .map(|e| eval(e, &Ctx::Group(members)))
                        .collect::<Result<_, _>>()?;
                    for p in &pivots {
                        let mut cell_rows = Vec::new();
                        for m in members {
                            let key = q
                                .pivot
                                .iter()
                                .map(|pe| eval(pe, &Ctx::Row(m)))
                                .collect::<Result<Vec<_>, _>>()?;
                            if same_key(&key, p) {
                                cell_rows.push(*m);
                            }
                        }
                        for a in &aggs {
                            // A pivot value with no rows in this group is empty, not 0.
                            row.push(if cell_rows.is_empty() {
                                Value::Empty
                            } else {
                                eval(a, &Ctx::Group(&cell_rows))?
                            });
                        }
                    }
                    out.push(row);
                    keys.push(
                        q.order_by
                            .iter()
                            .map(|(e, _)| eval(e, &Ctx::Group(members)))
                            .collect::<Result<Vec<_>, _>>()?,
                    );
                }
                (head, out, keys)
            }
        } else {
            let head = select.iter().map(header_for).collect();
            let mut out = Vec::new();
            let mut keys = Vec::new();
            for row in &rows {
                let ctx = Ctx::Row(row);
                out.push(
                    select
                        .iter()
                        .map(|e| eval(e, &ctx))
                        .collect::<Result<Vec<_>, _>>()?,
                );
                keys.push(
                    q.order_by
                        .iter()
                        .map(|(e, _)| eval(e, &ctx))
                        .collect::<Result<Vec<_>, _>>()?,
                );
            }
            (head, out, keys)
        };

    if !q.order_by.is_empty() {
        let mut order: Vec<usize> = (0..body_rows.len()).collect();
        order.sort_by(|&a, &b| {
            for (k, (_, asc)) in q.order_by.iter().enumerate() {
                let o = compare(&order_keys[a][k], &order_keys[b][k]);
                let o = if *asc { o } else { o.reverse() };
                if o != Ordering::Equal {
                    return o;
                }
            }
            Ordering::Equal
        });
        let mut sorted: Vec<Vec<Value>> = order
            .iter()
            .map(|&i| std::mem::take(&mut body_rows[i]))
            .collect();
        std::mem::swap(&mut body_rows, &mut sorted);
    }
    let body_rows: Vec<Vec<Value>> = body_rows
        .into_iter()
        .skip(q.offset)
        .take(q.limit.unwrap_or(usize::MAX))
        .collect();

    let with_header = header_rows > 0 || !q.labels.is_empty() || grouped;
    if body_rows.is_empty() && !with_header {
        return Ok(EvalResult::Error(
            "#N/A Query completed with an empty output".to_string(),
        ));
    }
    let mut out = Vec::with_capacity(body_rows.len() + 1);
    if with_header {
        out.push(
            head.into_iter()
                .map(|h| {
                    if h.is_empty() {
                        Value::Empty
                    } else {
                        Value::Text(h)
                    }
                })
                .collect(),
        );
    }
    out.extend(body_rows);
    Ok(EvalResult::Array(super::eval::Array2D::from_vec(out)))
}

fn column_letters(mut c: usize) -> String {
    let mut s = Vec::new();
    loop {
        s.push(b'A' + (c % 26) as u8);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    s.reverse();
    String::from_utf8(s).unwrap()
}

fn default_header(e: &Q, labels: &[String], names: &[String]) -> String {
    match e {
        Q::Col(i) => labels[*i].clone(),
        Q::Agg(a, inner) => format!("{} {}", a.name(), default_header(inner, labels, names))
            .trim()
            .to_string(),
        other => written(other, names),
    }
}

/// An expression as text, for a computed column's header.
fn written(e: &Q, names: &[String]) -> String {
    match e {
        Q::Col(i) => names[*i].clone(),
        Q::Lit(Value::Text(t)) => format!("'{t}'"),
        Q::Lit(v) => display(v),
        Q::Agg(a, x) => format!("{}({})", a.name(), written(x, names)),
        Q::Fn(f, x) => format!("{}({})", format!("{f:?}").to_lowercase(), written(x, names)),
        Q::Neg(x) => format!("-{}", written(x, names)),
        Q::Arith(op, a, b) => format!("{}{op}{}", written(a, names), written(b, names)),
        Q::Cmp(_, a, b) | Q::And(a, b) | Q::Or(a, b) => {
            format!("{} {}", written(a, names), written(b, names))
        }
        Q::Not(x) | Q::IsNull(x, _) => written(x, names),
    }
}

/// In a grouped query, a column outside an aggregate must be grouped by.
fn check_grouped(e: &Q, group_by: &[Q], pivot: &[Q]) -> Result<(), String> {
    if group_by.contains(e) || matches!(e, Q::Agg(..) | Q::Lit(_)) {
        return Ok(());
    }
    match e {
        Q::Col(_) if pivot.contains(e) => Err(err("a pivot column can't also be selected")),
        Q::Col(_) => Err(err(
            "every selected column must be in group by or inside an aggregate",
        )),
        Q::Fn(_, a) | Q::Neg(a) | Q::Not(a) | Q::IsNull(a, _) => check_grouped(a, group_by, pivot),
        Q::Arith(_, a, b) | Q::Cmp(_, a, b) | Q::And(a, b) | Q::Or(a, b) => {
            check_grouped(a, group_by, pivot)?;
            check_grouped(b, group_by, pivot)
        }
        Q::Agg(..) | Q::Lit(_) => Ok(()),
    }
}

enum Ctx<'a> {
    Row(&'a [Value]),
    Group(&'a [&'a Vec<Value>]),
}

fn eval(e: &Q, ctx: &Ctx) -> Result<Value, String> {
    Ok(match e {
        Q::Col(i) => match ctx {
            Ctx::Row(row) => row[*i].clone(),
            // Grouped by this column: every row in the group has the same value.
            Ctx::Group(rows) => rows.first().map(|r| r[*i].clone()).unwrap_or(Value::Empty),
        },
        Q::Lit(v) => v.clone(),
        Q::Agg(a, inner) => {
            let Ctx::Group(rows) = ctx else {
                return Err(err("aggregates need group by or a whole-column result"));
            };
            let mut values = Vec::new();
            for row in rows.iter() {
                let v = eval(inner, &Ctx::Row(row))?;
                if !matches!(v, Value::Empty) {
                    values.push(v);
                }
            }
            aggregate(*a, &values)
        }
        Q::Fn(f, inner) => {
            let v = eval(inner, ctx)?;
            match (f, v) {
                (Scalar::Year | Scalar::Month | Scalar::Day, Value::Number(n)) => {
                    let (y, m, d) = serial_to_date(n);
                    Value::Number(match f {
                        Scalar::Year => y,
                        Scalar::Month => m,
                        _ => d,
                    } as f64)
                }
                (Scalar::Upper, Value::Text(t)) => Value::Text(t.to_uppercase()),
                (Scalar::Lower, Value::Text(t)) => Value::Text(t.to_lowercase()),
                _ => Value::Empty,
            }
        }
        Q::Neg(inner) => match eval(inner, ctx)? {
            Value::Number(n) => Value::Number(-n),
            _ => Value::Empty,
        },
        Q::Arith(op, a, b) => match (eval(a, ctx)?, eval(b, ctx)?) {
            (Value::Number(x), Value::Number(y)) => match op {
                '+' => Value::Number(x + y),
                '-' => Value::Number(x - y),
                '*' => Value::Number(x * y),
                _ if y == 0.0 => Value::Empty,
                _ => Value::Number(x / y),
            },
            _ => Value::Empty,
        },
        Q::Cmp(op, a, b) => Value::Boolean(compare_op(*op, &eval(a, ctx)?, &eval(b, ctx)?)?),
        Q::And(a, b) => Value::Boolean(truthy(&eval(a, ctx)?) && truthy(&eval(b, ctx)?)),
        Q::Or(a, b) => Value::Boolean(truthy(&eval(a, ctx)?) || truthy(&eval(b, ctx)?)),
        Q::Not(a) => Value::Boolean(!truthy(&eval(a, ctx)?)),
        Q::IsNull(a, negated) => Value::Boolean(matches!(eval(a, ctx)?, Value::Empty) != *negated),
    })
}

fn aggregate(a: Agg, values: &[Value]) -> Value {
    if a == Agg::Count {
        return Value::Number(values.len() as f64);
    }
    if matches!(a, Agg::Min | Agg::Max) {
        // min and max work on text too.
        let pick = values.iter().reduce(|best, v| {
            let o = compare(v, best);
            if (a == Agg::Min && o == Ordering::Less) || (a == Agg::Max && o == Ordering::Greater) {
                v
            } else {
                best
            }
        });
        return pick.cloned().unwrap_or(Value::Empty);
    }
    let numbers: Vec<f64> = values
        .iter()
        .filter_map(|v| {
            if let Value::Number(n) = v {
                Some(*n)
            } else {
                None
            }
        })
        .collect();
    if numbers.is_empty() {
        return Value::Empty;
    }
    let sum: f64 = numbers.iter().sum();
    Value::Number(if a == Agg::Sum {
        sum
    } else {
        sum / numbers.len() as f64
    })
}

fn truthy(v: &Value) -> bool {
    matches!(v, Value::Boolean(true))
}

fn compare_op(op: Cmp, a: &Value, b: &Value) -> Result<bool, String> {
    if matches!(a, Value::Empty) || matches!(b, Value::Empty) {
        // A null compares unequal to everything; only "is null" finds it.
        return Ok(op == Cmp::Ne && !(matches!(a, Value::Empty) && matches!(b, Value::Empty)));
    }
    let text = |v: &Value| {
        if let Value::Text(t) = v {
            Some(t.clone())
        } else {
            None
        }
    };
    Ok(match op {
        Cmp::Eq | Cmp::Ne | Cmp::Lt | Cmp::Le | Cmp::Gt | Cmp::Ge => {
            let same_type = std::mem::discriminant(a) == std::mem::discriminant(b);
            if !same_type {
                return Ok(op == Cmp::Ne);
            }
            let o = compare_exact(a, b);
            match op {
                Cmp::Eq => o == Ordering::Equal,
                Cmp::Ne => o != Ordering::Equal,
                Cmp::Lt => o == Ordering::Less,
                Cmp::Le => o != Ordering::Greater,
                Cmp::Gt => o == Ordering::Greater,
                _ => o != Ordering::Less,
            }
        }
        _ => {
            let (Some(s), Some(p)) = (text(a), text(b)) else {
                return Ok(false);
            };
            match op {
                Cmp::Contains => s.contains(&p),
                Cmp::StartsWith => s.starts_with(&p),
                Cmp::EndsWith => s.ends_with(&p),
                Cmp::Matches => {
                    super::eval_regex::compiled(&format!("^(?:{p})$"), false)?.is_match(&s)
                }
                _ => super::eval_regex::compiled(&like_pattern(&p), false)?.is_match(&s),
            }
        }
    })
}

/// A like pattern (% any run, _ any one character) as an anchored regex.
fn like_pattern(p: &str) -> String {
    let mut out = String::from("^");
    for c in p.chars() {
        match c {
            '%' => out.push_str("(?s:.*)"),
            '_' => out.push_str("(?s:.)"),
            c => out.push_str(&regex::escape(&c.to_string())),
        }
    }
    out.push('$');
    out
}

/// Case-sensitive, as the language compares strings.
fn compare_exact(a: &Value, b: &Value) -> Ordering {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Value::Text(x), Value::Text(y)) => x.cmp(y),
        (Value::Boolean(x), Value::Boolean(y)) => x.cmp(y),
        _ => Ordering::Equal,
    }
}

/// Sort order: nulls first, then numbers, booleans and text; text by letter
/// ignoring case, then by case.
fn compare(a: &Value, b: &Value) -> Ordering {
    fn rank(v: &Value) -> u8 {
        match v {
            Value::Empty | Value::Error(_) => 0,
            Value::Number(_) => 1,
            Value::Boolean(_) => 2,
            Value::Text(_) => 3,
        }
    }
    rank(a).cmp(&rank(b)).then_with(|| match (a, b) {
        (Value::Text(x), Value::Text(y)) => x
            .to_lowercase()
            .cmp(&y.to_lowercase())
            .then_with(|| x.cmp(y)),
        _ => compare_exact(a, b),
    })
}

fn compare_keys(a: &[Value], b: &[Value]) -> Ordering {
    a.iter()
        .zip(b)
        .map(|(x, y)| compare(x, y))
        .find(|o| *o != Ordering::Equal)
        .unwrap_or(Ordering::Equal)
}

fn same_key(a: &[Value], b: &[Value]) -> bool {
    a.len() == b.len()
        && a.iter().zip(b).all(|(x, y)| {
            std::mem::discriminant(x) == std::mem::discriminant(y)
                && compare_exact(x, y) == Ordering::Equal
        })
}

fn display(v: &Value) -> String {
    match v {
        Value::Number(n) => EvalResult::Number(*n).to_display(),
        other => other.to_text(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tokens_and_column_names() {
        assert_eq!(
            tokenize("select A, `Col 2` where B >= 'x y' and C<>1.5").unwrap(),
            vec![
                Tok::Word("select".into()),
                Tok::Word("A".into()),
                Tok::Sym(","),
                Tok::Backtick("Col 2".into()),
                Tok::Word("where".into()),
                Tok::Word("B".into()),
                Tok::Sym(">="),
                Tok::Str("x y".into()),
                Tok::Word("and".into()),
                Tok::Word("C".into()),
                Tok::Sym("<>"),
                Tok::Num(1.5),
            ]
        );
        assert_eq!(column_letters(0), "A");
        assert_eq!(column_letters(25), "Z");
        assert_eq!(column_letters(26), "AA");
        assert_eq!(column_letters(701), "ZZ");
    }

    #[test]
    fn like_wildcards() {
        assert_eq!(like_pattern("a%b_c.d"), "^a(?s:.*)b(?s:.)c\\.d$");
    }

    #[test]
    fn date_literals() {
        assert_eq!(
            date_literal("date", "2026-01-31").unwrap(),
            date_to_serial(2026, 1, 31)
        );
        assert_eq!(date_literal("timeofday", "12:00:00").unwrap(), 0.5);
        assert_eq!(
            date_literal("datetime", "2026-01-31 06:00:00").unwrap(),
            date_to_serial(2026, 1, 31) + 0.25
        );
        assert!(date_literal("date", "2026-13-01").is_err());
    }
}
