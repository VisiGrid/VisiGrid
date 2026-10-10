//! CPU-only MCP plan preparation. No GPUI entity or shared console Lua state
//! crosses into the worker; only a COW workbook snapshot and immutable inputs.
use crate::{review_mode::ReviewModeState, scripting, terminal::state::LuaPreviewData};
use visigrid_engine::{
    operation_plan::{PlanId, PlanProducer, VerificationDefinition},
    workbook::Workbook,
};

pub(crate) struct PlanJob {
    pub request: visigrid_protocol::CreatePlanMessage,
    pub record: crate::plan_manager::McpPlanRecord,
    pub workbook: Workbook,
    pub layout: crate::table_review::TableReviewLayout,
    pub source_layout: crate::table_structure::StructureLayout,
    pub source_frozen: (usize, usize),
    pub session_id: u64,
    pub context: scripting::ExecutionContextGenerationKey,
    pub blocks_delete: bool,
    pub verification: Vec<VerificationDefinition>,
}

pub(crate) struct BuiltPlan {
    pub preview: LuaPreviewData,
    pub review: ReviewModeState,
}

impl PlanJob {
    pub fn build(&self) -> Result<BuiltPlan, String> {
        if scripting::execution_context_generation_key(&self.workbook) != self.context {
            return Err("plan_stale: calculation context changed before preparation".into());
        }
        let visigrid_protocol::PlanProducerPayload::LuaScript { source } = &self.request.producer;
        // Keep the same sandbox and limits as the console, but never share
        // its globals with an agent or send a Lua VM between threads.
        let runtime = scripting::LuaRuntime::new().map_err(|error| error.to_string())?;
        let snapshot = scripting::SheetSnapshot::from_sheet(self.workbook.active_sheet());
        let result = runtime.eval_with_sheet(source, Box::new(snapshot));
        if let Some(error) = result.error {
            return Err(error);
        }
        if self.blocks_delete
            && result
                .ops
                .iter()
                .any(|op| matches!(op, scripting::LuaOp::DeleteRows { .. }))
        {
            return Err("unsupported_view_state: clear the active sort/filter before reviewing row deletion".into());
        }
        let script_hash = blake3::hash(source.as_bytes()).to_hex().to_string();
        let prepared = crate::ai_actions::prepare_lua_operation_plan_with_metadata(
            &self.workbook,
            self.session_id,
            PlanId(self.record.plan_id.clone()),
            PlanProducer {
                kind: "mcp_lua".into(),
                name: self.record.owner.clone(),
                source_path: None,
                source_hash: Some(script_hash.clone()),
            },
            self.request.title.trim().to_string(),
            self.request.description.clone(),
            &script_hash,
            &result.ops,
            self.verification.clone(),
        )
        .and_then(crate::ai_actions::require_visible_plan_changes)?;
        self.layout.validate(&prepared)?;
        let context =
            scripting::execution_context_fingerprint(&self.workbook, &prepared.plan().operations);
        if context != prepared.plan().execution_context
            || scripting::execution_context_generation_key(&self.workbook) != self.context
        {
            return Err("plan_stale: calculation context changed during preparation".into());
        }
        let review = ReviewModeState::from_prepared(&prepared, &self.workbook);
        let preview = LuaPreviewData {
            source_layout: self.source_layout.clone(),
            source_frozen: self.source_frozen,
            script_path: format!("mcp/{}.lua", self.record.plan_id).into(),
            script_hash,
            cells_written: prepared.plan().summary.total_changes(),
            cells_overwritten: crate::app::count_lua_overwrites(
                &result.ops,
                self.workbook.active_sheet(),
            ),
            source_sheet_index: self.workbook.active_sheet_index(),
            source_fingerprint: crate::app::sheet_fingerprint(self.workbook.active_sheet()),
            ops: result.ops,
            prepared_plan: Some(prepared),
            output: result.output,
            error: None,
        };
        Ok(BuiltPlan { preview, review })
    }
}
