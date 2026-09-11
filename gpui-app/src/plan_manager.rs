//! GUI-owned lifecycle metadata for MCP-created Review Mode plans.
//!
//! The immutable prepared plan itself stays in the same pending-preview slot
//! used by local Lua Review Mode. This manager retains ownership, idempotency,
//! and terminal results so an MCP client may disconnect and later observe the
//! human's Apply or Dismiss decision without being in the commit path.

use std::collections::HashMap;

use serde_json::Value;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum McpPlanState {
    Ready,
    Invalid,
    Applied,
    Dismissed,
}

impl McpPlanState {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Ready => "ready",
            Self::Invalid => "invalid",
            Self::Applied => "applied",
            Self::Dismissed => "dismissed",
        }
    }

    pub fn is_terminal(self) -> bool {
        matches!(self, Self::Invalid | Self::Applied | Self::Dismissed)
    }
}

#[derive(Debug, Clone)]
pub struct McpPlanRecord {
    pub plan_id: String,
    pub owner: String,
    pub source_revision: u64,
    pub request_hash: String,
    pub state: McpPlanState,
    pub invalid_message: Option<String>,
    pub terminal_result: Option<Value>,
}

#[derive(Debug, Default)]
pub struct McpPlanManager {
    records: HashMap<String, McpPlanRecord>,
    create_keys: HashMap<(String, String), String>,
    active_plan_id: Option<String>,
}

impl McpPlanManager {
    pub fn active_plan_id(&self) -> Option<&str> {
        self.active_plan_id.as_deref()
    }

    pub fn record(&self, plan_id: &str) -> Option<&McpPlanRecord> {
        self.records.get(plan_id)
    }

    pub fn existing_for_key(
        &self,
        owner: &str,
        key: &str,
        request_hash: &str,
    ) -> Result<Option<&McpPlanRecord>, ()> {
        let Some(plan_id) = self.create_keys.get(&(owner.to_string(), key.to_string())) else {
            return Ok(None);
        };
        let record = self
            .records
            .get(plan_id)
            .expect("idempotency record must resolve");
        if record.request_hash == request_hash {
            Ok(Some(record))
        } else {
            Err(())
        }
    }

    pub fn active_record(&self) -> Option<&McpPlanRecord> {
        self.active_plan_id
            .as_deref()
            .and_then(|id| self.records.get(id))
    }

    pub fn insert(&mut self, record: McpPlanRecord, idempotency_key: String) {
        let plan_id = record.plan_id.clone();
        let owner = record.owner.clone();
        if !record.state.is_terminal() {
            self.active_plan_id = Some(plan_id.clone());
        }
        self.create_keys
            .insert((owner, idempotency_key), plan_id.clone());
        self.records.insert(plan_id, record);
    }

    pub fn mark_applied(&mut self, plan_id: &str, result: Value) {
        if let Some(record) = self.records.get_mut(plan_id) {
            record.state = McpPlanState::Applied;
            record.terminal_result = Some(result);
            if self.active_plan_id.as_deref() == Some(plan_id) {
                self.active_plan_id = None;
            }
        }
    }

    pub fn mark_dismissed(&mut self, plan_id: &str, result: Value) {
        if let Some(record) = self.records.get_mut(plan_id) {
            record.state = McpPlanState::Dismissed;
            record.terminal_result = Some(result);
            if self.active_plan_id.as_deref() == Some(plan_id) {
                self.active_plan_id = None;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(hash: &str) -> McpPlanRecord {
        McpPlanRecord {
            plan_id: "pv_test".into(),
            owner: "Codex".into(),
            source_revision: 7,
            request_hash: hash.into(),
            state: McpPlanState::Ready,
            invalid_message: None,
            terminal_result: None,
        }
    }

    #[test]
    fn create_idempotency_reuses_only_identical_payloads() {
        let mut manager = McpPlanManager::default();
        manager.insert(record("same"), "key-1".into());
        assert!(manager
            .existing_for_key("Codex", "key-1", "same")
            .unwrap()
            .is_some());
        assert!(manager
            .existing_for_key("Codex", "key-1", "different")
            .is_err());
    }

    #[test]
    fn terminal_plan_stops_blocking_the_next_plan() {
        let mut manager = McpPlanManager::default();
        manager.insert(record("same"), "key-1".into());
        assert_eq!(manager.active_plan_id(), Some("pv_test"));
        manager.mark_dismissed("pv_test", serde_json::json!({"state": "dismissed"}));
        assert_eq!(manager.active_plan_id(), None);
    }

    #[test]
    fn invalid_plan_does_not_block_a_corrected_retry() {
        let mut invalid = record("bad");
        invalid.state = McpPlanState::Invalid;
        let mut manager = McpPlanManager::default();
        manager.insert(invalid, "bad-key".into());
        assert_eq!(manager.active_plan_id(), None);
    }
}
