//! Atomic named-range edits, including dependencies in hidden records and
//! dormant Table formulas. History stores authored deltas, never a workbook.
use super::{GuardedStructureCommit, Workbook};
use crate::{
    cell::{CellValue, ValueRef},
    named_range::{NamedRange, NamedRangeTarget},
    validation::{ConstraintValue, ListSource, ValidationType},
};

#[derive(Clone, Debug)]
pub enum NamedRangeEdit {
    Create(NamedRange),
    Delete(NamedRange),
    Rename {
        before: NamedRange,
        name: String,
    },
    Description {
        before: NamedRange,
        description: Option<String>,
    },
}

impl Workbook {
    /// Workbook-wide impact locations, including dormant rules and frozen
    /// formula sources. Labels explicitly distinguish metadata from cells.
    pub fn named_range_usages(&self, name: &str) -> Vec<(String, String)> {
        let mut usages = Vec::new();
        let mut add = |location: String, source: &str| {
            if crate::formula::names::references_name(source, name) {
                usages.push((location, source.to_string()));
            }
        };
        for sheet in self.sheets() {
            let prefix = format!("'{}'", sheet.name.replace('\'', "''"));
            for ((row, col), cell) in sheet.cells_iter() {
                let address = NamedRange::cell("", 0, row, col).reference_string();
                if let ValueRef::Formula { source, .. } = cell.value() {
                    add(format!("{prefix}!{address}"), source);
                }
                if let Some(source) = cell.frozen_formula() {
                    add(format!("{prefix}!{address} (frozen)"), source);
                }
            }
            for table in sheet.tables() {
                for col in &table.columns {
                    if let Some(source) = &col.formula {
                        add(format!("{}[{}] rule", table.name, col.name), source);
                    }
                }
                if let Some(totals) = &table.totals {
                    for (col, total) in table.columns.iter().zip(&totals.columns) {
                        if let Some(source) = &total.formula {
                            add(format!("{}[{}] totals", table.name, col.name), source);
                        }
                    }
                }
            }
            for rule in sheet.cond_formats.iter() {
                add(format!("{prefix} format rule {}", rule.id), &rule.predicate);
            }
            for (range, rule) in sheet.validations.iter() {
                let label = format!(
                    "{prefix}!{} validation",
                    NamedRange::range(
                        "",
                        0,
                        range.start_row,
                        range.start_col,
                        range.end_row,
                        range.end_col
                    )
                    .reference_string()
                );
                match &rule.rule_type {
                    ValidationType::Custom(source)
                    | ValidationType::List(ListSource::Range(source))
                    | ValidationType::List(ListSource::NamedRange(source)) => add(label, source),
                    ValidationType::List(_) => {}
                    ValidationType::WholeNumber(c)
                    | ValidationType::Decimal(c)
                    | ValidationType::Date(c)
                    | ValidationType::Time(c)
                    | ValidationType::TextLength(c) => {
                        for v in std::iter::once(&c.value1).chain(c.value2.iter()) {
                            if let ConstraintValue::CellRef(s) | ConstraintValue::Formula(s) = v {
                                add(label.clone(), s);
                            }
                        }
                    }
                }
            }
        }
        usages.sort();
        usages
    }

    /// Validate a private candidate and retain a guarded, sparse undo payload.
    /// Existing low-level store APIs remain available to loaders and batches.
    pub fn prepare_named_range_edit(
        &self,
        edit: &NamedRangeEdit,
    ) -> Result<(Workbook, GuardedStructureCommit), String> {
        self.ensure_writable()?;
        let mut candidate = self.clone();
        match edit {
            NamedRangeEdit::Create(range) => {
                if self.get_named_range(&range.name).is_some() {
                    return Err(format!("'{}' already exists.", range.name));
                }
                let (index, r0, c0, r1, c1) = match range.target {
                    NamedRangeTarget::Cell { sheet, row, col } => (sheet, row, col, row, col),
                    NamedRangeTarget::Range {
                        sheet,
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                    } => (sheet, start_row, start_col, end_row, end_col),
                };
                let sheet = self
                    .sheet(index)
                    .ok_or("The named range's sheet no longer exists.")?;
                if r0 > r1 || c0 > c1 || r1 >= sheet.rows || c1 >= sheet.cols {
                    return Err("The named range extends beyond the worksheet.".into());
                }
                candidate.named_ranges.set(range.clone())?;
            }
            NamedRangeEdit::Delete(before)
            | NamedRangeEdit::Rename { before, .. }
            | NamedRangeEdit::Description { before, .. } => {
                if self.get_named_range(&before.name) != Some(before) {
                    return Err("The named range changed while this dialog was open. Reopen it and try again.".into());
                }
                match edit {
                    NamedRangeEdit::Delete(_) => {
                        candidate.named_ranges.remove(&before.name);
                    }
                    NamedRangeEdit::Rename { name, .. } => {
                        candidate.named_ranges.rename(&before.name, name)?;
                        candidate.rewrite_named_range_sources(&before.name, name)?;
                    }
                    NamedRangeEdit::Description { description, .. } => {
                        candidate
                            .named_ranges
                            .set_description(&before.name, description.clone())?;
                    }
                    _ => unreachable!(),
                }
            }
        }
        if !matches!(edit, NamedRangeEdit::Description { .. }) {
            candidate.rebuild_dep_graph();
            let report = candidate.recompute_full_ordered();
            // Ordinary formula errors (notably #NAME? after deletion) are
            // valid results. Cycles and unsettled spills cannot be published.
            if report.had_cycles
                || report
                    .errors
                    .iter()
                    .any(|e| e.error.contains("spill not settled"))
            {
                return Err("The named range would create a cycle or an unsettled spill. Nothing was changed.".into());
            }
            // Formula-backed Table keys, spills and dependent pivots may change
            // on sheets that contain no directly rewritten formula sources.
            for sheet in &mut candidate.sheets {
                sheet.mark_table_changed();
                sheet.build_saved_table_view(sheet.rows.min(crate::sheet::NUM_ROWS))?;
            }
        }
        let commit = self.capture_guarded_batch(&candidate)?;
        Ok((candidate, commit))
    }

    fn rewrite_named_range_sources(&mut self, old: &str, new: &str) -> Result<(), String> {
        let rewrite = |source: &mut String| -> Result<bool, String> {
            let rewritten = crate::formula::names::rename_reference(source, old, new)?;
            let changed = *source != rewritten;
            *source = rewritten;
            Ok(changed)
        };
        for sheet in &mut self.sheets {
            let mut cells = Vec::new();
            for ((r, c), cell) in sheet.cells_iter() {
                let mut formula = match cell.value() {
                    ValueRef::Formula { source, .. } => Some(source.to_string()),
                    _ => None,
                };
                let mut frozen = cell.frozen_formula().map(str::to_string);
                let mut changed = false;
                if let Some(source) = &mut formula {
                    changed |= rewrite(source)?;
                }
                if let Some(source) = &mut frozen {
                    changed |= rewrite(source)?;
                }
                if changed {
                    let mut image = cell.to_cell();
                    if let Some(source) = formula {
                        image.value = CellValue::from_input(&source);
                    }
                    image.set_frozen_formula(frozen);
                    cells.push((r, c, image));
                }
            }
            for (r, c, image) in cells {
                sheet.restore_history_cell(r, c, Some(image));
            }
            for table in &mut sheet.data_tables {
                for column in &mut table.columns {
                    if let Some(source) = &mut column.formula {
                        rewrite(source)?;
                    }
                }
                if let Some(totals) = &mut table.totals {
                    for column in &mut totals.columns {
                        if let Some(source) = &mut column.formula {
                            rewrite(source)?;
                        }
                    }
                }
            }
            let ids: Vec<_> = sheet.cond_formats.iter().map(|r| r.id).collect();
            for id in ids {
                let rule = sheet.cond_formats.get_mut(id).unwrap();
                if rewrite(&mut rule.predicate)? {
                    rule.reparse();
                }
            }
            let validations: Vec<_> = sheet
                .validations
                .iter()
                .map(|(r, v)| (*r, v.clone()))
                .collect();
            for (range, mut rule) in validations {
                let mut changed = false;
                match &mut rule.rule_type {
                    ValidationType::Custom(source)
                    | ValidationType::List(ListSource::Range(source)) => {
                        changed |= rewrite(source)?;
                    }
                    ValidationType::List(ListSource::NamedRange(name))
                        if name.eq_ignore_ascii_case(old) =>
                    {
                        *name = new.into();
                        changed = true;
                    }
                    ValidationType::List(_) => {}
                    ValidationType::WholeNumber(c)
                    | ValidationType::Decimal(c)
                    | ValidationType::Date(c)
                    | ValidationType::Time(c)
                    | ValidationType::TextLength(c) => {
                        for v in std::iter::once(&mut c.value1).chain(c.value2.iter_mut()) {
                            if let ConstraintValue::CellRef(s) | ConstraintValue::Formula(s) = v {
                                changed |= rewrite(s)?;
                            }
                        }
                    }
                }
                if changed {
                    sheet.validations.set(range, rule);
                }
            }
        }
        Ok(())
    }
}
