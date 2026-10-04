//! Copy a reviewed sheet with fresh workbook-owned identities.
use super::Workbook;
use crate::{
    cell::ValueRef,
    formula::structured,
    sheet::SheetId,
    table::TableId,
    validation::{ConstraintValue, ListSource, ValidationType},
};
use std::collections::{BTreeMap, HashMap, HashSet};

fn rewrite(source: &str, names: &HashMap<String, String>) -> String {
    let mut result = source.to_string();
    for (start, end, mut reference) in structured::source_references(source).into_iter().rev() {
        let Some(name) = reference
            .table
            .as_ref()
            .and_then(|n| names.get(&n.to_ascii_lowercase()))
        else {
            continue;
        };
        reference.table = Some(name.clone());
        let replacement = if source[start..end].contains('[') {
            reference.format()
        } else {
            name.clone()
        };
        result.replace_range(start..end, &replacement);
    }
    result
}

impl Workbook {
    /// Stage a complete sheet copy. Formulas remain live; explicit references
    /// to other sheets keep their targets. Copied Table references bind locally.
    pub fn prepare_sheet_copy(
        &self,
        source: &Workbook,
        source_id: SheetId,
        name: &str,
    ) -> Result<(Workbook, usize), String> {
        self.ensure_writable()?;
        source.ensure_writable()?;
        if self.style_table != source.style_table {
            return Err(
                "Workbook styles changed since review. Prepare a new preview before copying."
                    .into(),
            );
        }
        if !super::is_valid_sheet_name(name) || self.sheet_name_exists(name) {
            return Err("The result sheet name is invalid or already used.".into());
        }
        let mut sheet = source
            .sheet_by_id(source_id)
            .ok_or("The reviewed source sheet is unavailable.")?
            .clone();
        if sheet.tables().iter().any(|t| t.totals.is_some()) {
            return Err("Copying sheets with totals-row Tables is not supported yet.".into());
        }
        let mut names = HashMap::new();
        let mut reserved: HashSet<String> = self
            .tables()
            .chain(source.tables())
            .map(|(_, t)| t.name.to_ascii_lowercase())
            .collect();
        // Do not bind previously unresolved formulas on existing sheets to a
        // generated name: those would also prevent removing the copy on undo.
        for book in [self, source] {
            for sheet in book.sheets() {
                let cell_formulas = sheet
                    .cells_iter()
                    .filter_map(|(_, cell)| match cell.value() {
                        ValueRef::Formula { source, .. } => Some(source),
                        _ => None,
                    });
                let rules = sheet
                    .tables()
                    .iter()
                    .flat_map(|t| t.columns.iter().filter_map(|c| c.formula.as_deref()));
                for formula in cell_formulas.chain(rules) {
                    for (_, _, reference) in structured::source_references(formula) {
                        if let Some(name) = reference.table {
                            reserved.insert(name.to_ascii_lowercase());
                        }
                    }
                }
            }
        }
        for table in sheet.tables() {
            let name = (1u64..)
                .map(|i| {
                    if i == 1 {
                        format!("{}_Copy", table.name)
                    } else {
                        format!("{}_Copy{i}", table.name)
                    }
                })
                .find(|name| {
                    !reserved.contains(&name.to_ascii_lowercase())
                        && self.get_named_range(name).is_none()
                        && source.get_named_range(name).is_none()
                })
                .ok_or("No available Table name.")?;
            reserved.insert(name.to_ascii_lowercase());
            names.insert(table.name.to_ascii_lowercase(), name);
        }
        let formulas: Vec<_> = sheet
            .cells_iter()
            .filter_map(|((r, c), cell)| match cell.value() {
                ValueRef::Formula { source, .. } => Some((r, c, rewrite(source, &names))),
                _ => None,
            })
            .collect();
        for (r, c, value) in formulas {
            sheet.set_value(r, c, &value);
        }
        let rules: Vec<_> = sheet
            .cond_formats
            .iter()
            .map(|r| (r.id, rewrite(&r.predicate, &names)))
            .collect();
        for (id, predicate) in rules {
            let rule = sheet.cond_formats.get_mut(id).unwrap();
            rule.predicate = predicate;
            rule.reparse();
        }
        let rules: Vec<_> = sheet
            .validations
            .iter()
            .map(|(range, rule)| (*range, rule.clone()))
            .collect();
        for (range, mut rule) in rules {
            match &mut rule.rule_type {
                ValidationType::Custom(s) => *s = rewrite(s, &names),
                ValidationType::List(ListSource::Range(s) | ListSource::NamedRange(s)) => {
                    *s = rewrite(s, &names)
                }
                ValidationType::WholeNumber(n)
                | ValidationType::Decimal(n)
                | ValidationType::Date(n)
                | ValidationType::Time(n)
                | ValidationType::TextLength(n) => {
                    for value in std::iter::once(&mut n.value1).chain(n.value2.iter_mut()) {
                        if let ConstraintValue::Formula(s) | ConstraintValue::CellRef(s) = value {
                            *s = rewrite(s, &names);
                        }
                    }
                }
                _ => {}
            }
            sheet.validations.set(range, rule);
        }
        let mut candidate = self.clone();
        let new_id = candidate.generate_sheet_id();
        sheet.id = new_id;
        sheet.set_name(name);
        let mut ids = HashMap::new();
        let mut allocators = BTreeMap::new();
        for table in &mut sheet.data_tables {
            let old = table.id;
            table.id = TableId(candidate.next_table_id);
            candidate.next_table_id = candidate
                .next_table_id
                .checked_add(1)
                .ok_or("Table identities exhausted.")?;
            ids.insert(old, table.id);
            table.name = names[&table.name.to_ascii_lowercase()].clone();
            for col in &mut table.columns {
                if let Some(formula) = &mut col.formula {
                    *formula = rewrite(formula, &names);
                }
            }
            allocators.insert(
                table.id.0,
                sheet
                    .table_column_allocators
                    .get(&old.0)
                    .copied()
                    .unwrap_or(table.next_column_id),
            );
            sheet.table_id_high_water = sheet.table_id_high_water.max(table.id.0);
        }
        sheet.table_column_allocators = allocators;
        if let Some(spec) = &mut sheet.table_view_spec {
            spec.table = *ids
                .get(&spec.table)
                .ok_or("The reviewed Table view is invalid.")?;
        }
        let mut pivot_id = candidate.next_pivot_id();
        let mut pivot_names: HashSet<_> = candidate
            .sheets()
            .iter()
            .flat_map(|s| s.pivots.iter().map(|p| p.name.to_lowercase()))
            .collect();
        for pivot in &mut sheet.pivots {
            pivot.id = pivot_id;
            pivot_id = pivot_id
                .checked_add(1)
                .ok_or("Pivot identities exhausted.")?;
            pivot.name = (1u64..)
                .map(|n| format!("PivotTable{n}"))
                .find(|n| !pivot_names.contains(&n.to_lowercase()))
                .unwrap();
            pivot_names.insert(pivot.name.to_lowercase());
            if pivot.source.sheet_id == source_id {
                pivot.source.sheet_id = new_id;
            }
            if let Some(id) = pivot.source.table_id {
                if let Some(copy) = ids.get(&id) {
                    pivot.source.table_id = Some(*copy);
                }
            }
            pivot.stale = true;
            pivot.source_generation = None;
        }
        let index = candidate.sheets.len();
        candidate.sheets.push(sheet);
        candidate.refresh_table_name_reservations();
        candidate.rebuild_dep_graph();
        // Formula errors are valid reviewed cell results. Recalculation still
        // must leave every saved view and pivot binding structurally safe.
        candidate.recompute_full_ordered();
        for sheet in candidate.sheets() {
            sheet.build_saved_table_view(sheet.rows)?;
        }
        for pivot in &candidate.sheets[index].pivots {
            if candidate.sheet_by_id(pivot.source.sheet_id).is_none() {
                return Err("A reviewed pivot source sheet no longer exists. Prepare a new preview before copying.".into());
            }
            candidate
                .resolve_pivot_source(&mut pivot.source.clone(), &mut pivot.definition.clone())
                .map_err(|e| e.to_string())?;
        }
        candidate.bump_revision_for_structure();
        Ok((candidate, index))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        cell::CellStyle,
        cond_format::CondStyle,
        sheet::Sheet,
        table::TableRange,
        validation::{CellRange, ValidationRule},
    };
    fn book() -> Workbook {
        let mut wb =
            Workbook::from_sheets(vec![Sheet::new_with_name(SheetId(7), 30, 8, "Data")], 0);
        for (r, c, v) in [
            (0, 0, "Amount"),
            (1, 0, "10"),
            (2, 0, "20"),
            (5, 0, "Amount"),
            (6, 0, "5"),
        ] {
            wb.set_cell_value_tracked(0, r, c, v);
        }
        wb.create_table(
            SheetId(7),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 2,
                end_col: 0,
            },
            "Alpha",
        )
        .unwrap();
        wb.create_table(
            SheetId(7),
            TableRange {
                start_row: 5,
                start_col: 0,
                end_row: 6,
                end_col: 0,
            },
            "Beta",
        )
        .unwrap();
        wb
    }
    #[test]
    fn rewrites_multiple_tables_and_formula_metadata_without_rewriting_literals() {
        let mut wb = book();
        wb.set_cell_value_tracked(0, 10, 0, "=SUM(Alpha[Amount])+SUM(Beta[Amount])");
        wb.set_cell_value_tracked(0, 11, 0, "=\"Alpha[Amount]\"");
        wb.active_sheet_mut().cond_formats.add(
            vec![CellRange::single(1, 0)],
            "=SUM(Alpha[Amount])>0",
            CondStyle::Named(CellStyle::Success),
        );
        wb.active_sheet_mut().validations.set(
            CellRange::single(2, 0),
            ValidationRule::custom("=SUM(Beta[Amount])>0"),
        );
        wb.active_sheet_mut()
            .validations
            .exclude(CellRange::single(3, 0));
        let (copy, index) = wb.prepare_sheet_copy(&wb, SheetId(7), "Result").unwrap();
        let s = copy.sheet(index).unwrap();
        assert_eq!(s.get_display(10, 0), "35");
        assert_eq!(
            s.get_raw(10, 0),
            "=SUM(Alpha_Copy[Amount])+SUM(Beta_Copy[Amount])"
        );
        assert_eq!(s.get_raw(11, 0), "=\"Alpha[Amount]\"");
        assert!(s
            .cond_formats
            .iter()
            .next()
            .unwrap()
            .predicate
            .contains("Alpha_Copy"));
        assert_eq!(
            s.validations.get(2, 0).unwrap().rule_type,
            ValidationType::Custom("=SUM(Beta_Copy[Amount])>0".into())
        );
        assert!(s.validations.is_excluded(3, 0));
        assert_eq!(wb.sheet_count(), 1);
    }
    #[test]
    fn pivots_get_fresh_ids_and_bind_to_copied_tables() {
        let mut wb = book();
        let table = wb.active_sheet().tables()[0].clone();
        let source = wb.table_pivot_source(table.id).unwrap();
        wb.active_sheet_mut().pivots.push(crate::pivot::PivotTable {
            id: 1,
            name: "PivotTable1".into(),
            source,
            definition: crate::pivot::PivotDefinition {
                rows: vec![],
                column: None,
                values: vec![crate::pivot::PivotValueField {
                    field: crate::pivot::PivotField {
                        column_id: Some(table.columns[0].id),
                        offset: 0,
                        header: "Amount".into(),
                    },
                    aggregation: crate::pivot::Aggregation::Sum,
                    number_format: None,
                }],
            },
            anchor_row: 20,
            anchor_col: 0,
            extent: None,
            last_refresh: None,
            stale: false,
            source_generation: None,
        });
        let (copy, index) = wb.prepare_sheet_copy(&wb, SheetId(7), "Result").unwrap();
        let s = copy.sheet(index).unwrap();
        assert_ne!(s.pivots[0].id, 1);
        assert_ne!(s.pivots[0].name, "PivotTable1");
        assert_eq!(s.pivots[0].source.sheet_id, s.id);
        assert_eq!(s.pivots[0].source.table_id, Some(s.tables()[0].id));
        assert!(s.pivots[0].stale);
        let mut stale = wb.clone();
        stale.active_sheet_mut().pivots[0].source.table_id = None;
        stale.active_sheet_mut().pivots[0].source.sheet_id = SheetId(999);
        assert!(wb
            .prepare_sheet_copy(&stale, SheetId(7), "Stale result")
            .unwrap_err()
            .contains("source sheet no longer exists"));
    }
    #[test]
    fn existing_external_references_stay_live_and_relative_formulas_use_copy() {
        let mut wb = book();
        wb.set_cell_value_tracked(0, 12, 0, "=A2");
        wb.set_cell_value_tracked(0, 13, 0, "=Data!A2");
        let mut preview = wb.clone();
        preview.set_cell_value_tracked(0, 1, 0, "77");
        let (copy, index) = wb
            .prepare_sheet_copy(&preview, SheetId(7), "Result")
            .unwrap();
        assert_eq!(copy.sheet(index).unwrap().get_display(12, 0), "77");
        assert_eq!(copy.sheet(index).unwrap().get_display(13, 0), "10");
    }
    #[test]
    fn generated_table_names_do_not_capture_existing_unresolved_references() {
        let mut wb = book();
        wb.set_cell_value_tracked(0, 15, 0, "=SUM(Alpha_Copy[Amount])");
        let (copy, index) = wb.prepare_sheet_copy(&wb, SheetId(7), "Result").unwrap();
        assert_eq!(copy.sheet(index).unwrap().tables()[0].name, "Alpha_Copy2");
        assert_eq!(
            copy.active_sheet().get_raw(15, 0),
            "=SUM(Alpha_Copy[Amount])"
        );
        assert!(!copy.has_external_table_references(copy.sheet(index).unwrap().id));
    }
}
