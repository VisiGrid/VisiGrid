use std::collections::HashMap;

use chrono::NaiveDate;
use serde::Deserialize;

use crate::error::ReconError;

// ---------------------------------------------------------------------------
// Top-level config
// ---------------------------------------------------------------------------

#[derive(Debug, Deserialize)]
pub struct ReconConfig {
    pub name: String,
    pub way: u8,
    pub roles: HashMap<String, RoleConfig>,
    pub pairs: HashMap<String, PairConfig>,
    #[serde(default)]
    pub tolerance: ToleranceConfig,
    #[serde(default)]
    pub output: OutputConfig,
    #[serde(default)]
    pub settlement: Option<SettlementConfig>,
}

// ---------------------------------------------------------------------------
// Settlement
// ---------------------------------------------------------------------------

/// Settlement classification config.
///
/// `clock` specifies which role's date drives the settlement timeline.
/// For `*_only` groups there's only one aggregate so the clock is unambiguous.
/// For error groups (amount/timing mismatch), the clock role's date is used.
///
/// Each role's `columns.date` mapping already selects the right timestamp
/// (e.g. `effective_date` for Stripe payouts = payout date, `posted_date`
/// for bank = deposit date). The `clock` field selects which role's mapped
/// date to prefer when multiple are available.
#[derive(Debug, Clone, Deserialize)]
pub struct SettlementConfig {
    pub reference_date: NaiveDate,
    pub sla_days: u32,
    /// Role whose date is the settlement clock origin.
    /// Defaults to Processor if not specified.
    #[serde(default)]
    pub clock: SettlementClock,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, serde::Serialize)]
#[serde(rename_all = "snake_case")]
pub enum SettlementClock {
    Processor,
    Ledger,
    Bank,
}

impl Default for SettlementClock {
    fn default() -> Self {
        Self::Processor
    }
}

impl SettlementClock {
    /// The role name this clock corresponds to in aggregates.
    pub fn role_name(&self) -> &'static str {
        match self {
            Self::Processor => "processor",
            Self::Ledger => "ledger",
            Self::Bank => "bank",
        }
    }
}

impl std::fmt::Display for SettlementClock {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Processor => write!(f, "processor"),
            Self::Ledger => write!(f, "ledger"),
            Self::Bank => write!(f, "bank"),
        }
    }
}

// ---------------------------------------------------------------------------
// Role
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Deserialize)]
pub struct RoleConfig {
    pub kind: RoleKind,
    pub file: String,
    pub columns: ColumnMapping,
    #[serde(default)]
    pub filter: Option<RowFilter>,
    #[serde(default)]
    pub transform: Option<AmountTransform>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum RoleKind {
    Processor,
    Ledger,
    Bank,
}

impl std::fmt::Display for RoleKind {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Processor => write!(f, "processor"),
            Self::Ledger => write!(f, "ledger"),
            Self::Bank => write!(f, "bank"),
        }
    }
}

// ---------------------------------------------------------------------------
// Column mapping
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Deserialize)]
pub struct ColumnMapping {
    pub record_id: String,
    pub match_key: String,
    pub amount: String,
    pub date: String,
    pub currency: String,
    pub kind: String,
}

// ---------------------------------------------------------------------------
// Filter + Transform
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Deserialize)]
pub struct RowFilter {
    pub column: String,
    pub values: Vec<String>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct AmountTransform {
    pub multiply: Option<i64>,
    #[serde(default)]
    pub when_column: Option<String>,
    #[serde(default)]
    pub when_values: Option<Vec<String>>,
}

// ---------------------------------------------------------------------------
// Pair
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Deserialize)]
pub struct PairConfig {
    pub left: String,
    pub right: String,
    #[serde(default = "default_strategy")]
    pub strategy: MatchStrategy,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum MatchStrategy {
    ExactKey,
    FuzzyAmountDate,
}

fn default_strategy() -> MatchStrategy {
    MatchStrategy::ExactKey
}

// ---------------------------------------------------------------------------
// Tolerance + Output
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Deserialize)]
pub struct ToleranceConfig {
    #[serde(default)]
    pub amount_cents: i64,
    #[serde(default)]
    pub date_window_days: u32,
}

impl Default for ToleranceConfig {
    fn default() -> Self {
        Self {
            amount_cents: 0,
            date_window_days: 0,
        }
    }
}

#[derive(Debug, Clone, Default, Deserialize)]
pub struct OutputConfig {
    #[serde(default)]
    pub json: Option<String>,
}

// ---------------------------------------------------------------------------
// Parse + Validate
// ---------------------------------------------------------------------------

impl ReconConfig {
    pub fn from_toml(input: &str) -> Result<Self, ReconError> {
        let config: ReconConfig =
            toml::from_str(input).map_err(|e| ReconError::ConfigParse(e.to_string()))?;
        config.validate()?;
        Ok(config)
    }

    pub fn validate(&self) -> Result<(), ReconError> {
        // Way must be 2 or 3
        if self.way != 2 && self.way != 3 {
            return Err(ReconError::ConfigValidation(format!(
                "way must be 2 or 3, got {}",
                self.way
            )));
        }

        // Must have at least 2 roles
        if self.roles.len() < 2 {
            return Err(ReconError::ConfigValidation(
                "at least 2 roles are required".into(),
            ));
        }

        // Pair count must match way
        let expected_pairs = if self.way == 2 { 1 } else { 2 };
        if self.pairs.len() != expected_pairs {
            return Err(ReconError::WayMismatch {
                way: self.way,
                pairs: self.pairs.len(),
            });
        }

        // Each pair must reference existing roles
        for (pair_name, pair) in &self.pairs {
            if !self.roles.contains_key(&pair.left) {
                return Err(ReconError::UnknownRole(format!(
                    "pair '{pair_name}': left role '{}' not found",
                    pair.left
                )));
            }
            if !self.roles.contains_key(&pair.right) {
                return Err(ReconError::UnknownRole(format!(
                    "pair '{pair_name}': right role '{}' not found",
                    pair.right
                )));
            }
        }

        Ok(())
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    const VALID_2WAY: &str = r#"
name = "Test 2-Way"
way = 2

[roles.processor]
kind = "processor"
file = "stripe.csv"

[roles.processor.columns]
record_id  = "source_id"
match_key  = "group_id"
amount     = "amount_minor"
date       = "effective_date"
currency   = "currency"
kind       = "type"

[roles.ledger]
kind = "ledger"
file = "qbo.csv"

[roles.ledger.columns]
record_id  = "source_id"
match_key  = "source_id"
amount     = "amount_minor"
date       = "effective_date"
currency   = "currency"
kind       = "type"

[pairs.processor_ledger]
left = "processor"
right = "ledger"
strategy = "exact_key"

[tolerance]
amount_cents = 0
date_window_days = 2
"#;

    #[test]
    fn parse_valid_2way() {
        let config = ReconConfig::from_toml(VALID_2WAY).unwrap();
        assert_eq!(config.name, "Test 2-Way");
        assert_eq!(config.way, 2);
        assert_eq!(config.roles.len(), 2);
        assert_eq!(config.pairs.len(), 1);
        assert_eq!(config.tolerance.amount_cents, 0);
        assert_eq!(config.tolerance.date_window_days, 2);
        assert!(config.settlement.is_none());
    }

    #[test]
    fn parse_settlement_with_clock() {
        let input = format!(
            r#"{VALID_2WAY}

[settlement]
reference_date = "2026-01-31"
sla_days = 5
clock = "ledger"
"#
        );
        let config = ReconConfig::from_toml(&input).unwrap();
        let s = config.settlement.unwrap();
        assert_eq!(s.reference_date.to_string(), "2026-01-31");
        assert_eq!(s.sla_days, 5);
        assert_eq!(s.clock, SettlementClock::Ledger);
    }

    #[test]
    fn parse_settlement_clock_defaults_to_processor() {
        let input = format!(
            r#"{VALID_2WAY}

[settlement]
reference_date = "2026-01-31"
sla_days = 3
"#
        );
        let config = ReconConfig::from_toml(&input).unwrap();
        let s = config.settlement.unwrap();
        assert_eq!(s.clock, SettlementClock::Processor);
    }

    #[test]
    fn parse_settlement_rejects_invalid_clock() {
        let input = format!(
            r#"{VALID_2WAY}

[settlement]
reference_date = "2026-01-31"
sla_days = 5
clock = "processer"
"#
        );
        let err = ReconConfig::from_toml(&input);
        assert!(err.is_err(), "typo in clock should fail deserialization");
    }

    #[test]
    fn parse_with_filter_and_transform() {
        let input = r#"
name = "Filtered"
way = 2

[roles.processor]
kind = "processor"
file = "stripe.csv"

[roles.processor.columns]
record_id  = "source_id"
match_key  = "group_id"
amount     = "amount_minor"
date       = "effective_date"
currency   = "currency"
kind       = "type"

[roles.processor.transform]
multiply = -1
when_column = "type"
when_values = ["payout"]

[roles.processor.filter]
column = "type"
values = ["payout"]

[roles.ledger]
kind = "ledger"
file = "qbo.csv"

[roles.ledger.columns]
record_id  = "source_id"
match_key  = "source_id"
amount     = "amount_minor"
date       = "effective_date"
currency   = "currency"
kind       = "type"

[pairs.processor_ledger]
left = "processor"
right = "ledger"
"#;
        let config = ReconConfig::from_toml(input).unwrap();
        let proc = &config.roles["processor"];
        let xf = proc.transform.as_ref().unwrap();
        assert_eq!(xf.multiply, Some(-1));
        assert_eq!(xf.when_column.as_deref(), Some("type"));
        assert_eq!(xf.when_values.as_ref().unwrap(), &["payout"]);

        let filt = proc.filter.as_ref().unwrap();
        assert_eq!(filt.column, "type");
        assert_eq!(filt.values, vec!["payout"]);

        // Default strategy
        let pair = &config.pairs["processor_ledger"];
        assert_eq!(pair.strategy, MatchStrategy::ExactKey);
    }

    #[test]
    fn reject_way_mismatch() {
        let input = r#"
name = "Bad"
way = 3

[roles.a]
kind = "processor"
file = "a.csv"
[roles.a.columns]
record_id = "id"
match_key = "key"
amount = "amt"
date = "dt"
currency = "cur"
kind = "k"

[roles.b]
kind = "ledger"
file = "b.csv"
[roles.b.columns]
record_id = "id"
match_key = "key"
amount = "amt"
date = "dt"
currency = "cur"
kind = "k"

[pairs.ab]
left = "a"
right = "b"
"#;
        let err = ReconConfig::from_toml(input).unwrap_err();
        assert!(err.to_string().contains("way=3"));
    }

    #[test]
    fn reject_unknown_role_in_pair() {
        let input = r#"
name = "Bad"
way = 2

[roles.a]
kind = "processor"
file = "a.csv"
[roles.a.columns]
record_id = "id"
match_key = "key"
amount = "amt"
date = "dt"
currency = "cur"
kind = "k"

[roles.b]
kind = "ledger"
file = "b.csv"
[roles.b.columns]
record_id = "id"
match_key = "key"
amount = "amt"
date = "dt"
currency = "cur"
kind = "k"

[pairs.ac]
left = "a"
right = "c"
"#;
        let err = ReconConfig::from_toml(input).unwrap_err();
        assert!(err.to_string().contains("'c'"));
    }

    #[test]
    fn reject_invalid_way() {
        let input = r#"
name = "Bad"
way = 4

[roles.a]
kind = "processor"
file = "a.csv"
[roles.a.columns]
record_id = "id"
match_key = "key"
amount = "amt"
date = "dt"
currency = "cur"
kind = "k"

[roles.b]
kind = "ledger"
file = "b.csv"
[roles.b.columns]
record_id = "id"
match_key = "key"
amount = "amt"
date = "dt"
currency = "cur"
kind = "k"

[pairs.ab]
left = "a"
right = "b"
"#;
        let err = ReconConfig::from_toml(input).unwrap_err();
        assert!(err.to_string().contains("way must be 2 or 3"));
    }
}
