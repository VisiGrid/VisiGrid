use std::fmt;

#[derive(Debug)]
pub enum ReconError {
    /// TOML parse / deserialization error.
    ConfigParse(String),
    /// Config validation error (missing role, bad pair reference, etc.).
    ConfigValidation(String),
    /// A referenced role in a pair does not exist.
    UnknownRole(String),
    /// Way value doesn't match pair count.
    WayMismatch { way: u8, pairs: usize },
    /// Missing required column in input data.
    MissingColumn { role: String, column: String },
    /// Date parse error.
    DateParse { role: String, record_id: String, value: String },
    /// Amount parse error.
    AmountParse { role: String, record_id: String, value: String },
    /// IO error (file read, etc.).
    Io(String),
}

impl fmt::Display for ReconError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ConfigParse(msg) => write!(f, "config parse error: {msg}"),
            Self::ConfigValidation(msg) => write!(f, "config validation error: {msg}"),
            Self::UnknownRole(role) => write!(f, "unknown role: {role}"),
            Self::WayMismatch { way, pairs } => {
                write!(f, "way={way} requires {} pair(s), found {pairs}", if *way == 2 { 1 } else { 2 })
            }
            Self::MissingColumn { role, column } => {
                write!(f, "role '{role}': missing column '{column}'")
            }
            Self::DateParse { role, record_id, value } => {
                write!(f, "role '{role}', record '{record_id}': cannot parse date '{value}'")
            }
            Self::AmountParse { role, record_id, value } => {
                write!(f, "role '{role}', record '{record_id}': cannot parse amount '{value}'")
            }
            Self::Io(msg) => write!(f, "IO error: {msg}"),
        }
    }
}

impl std::error::Error for ReconError {}
