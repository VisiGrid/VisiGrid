use visigrid_engine::{
    filter::{ColumnFilter, NormalizedFilterKey},
    table::{TableRange, TableTotal, TableTotals},
    table_view::{TableFilter, TableViewSpec},
    workbook::Workbook,
};

fn book() -> Workbook {
    let mut wb = Workbook::new();
    for (row, data) in [
        ["Region", "Amount"],
        ["West", "10"],
        ["East", "20"],
        ["West", "30"],
    ]
    .iter()
    .enumerate()
    {
        for (col, value) in data.iter().enumerate() {
            wb.set_cell_value_tracked(0, row, col, value);
        }
    }
    wb.set_cell_value_tracked(0, 4, 0, "Total");
    wb.set_cell_value_tracked(0, 4, 1, "=SUBTOTAL(109,[Amount])");
    wb.create_table(
        wb.active_sheet_id(),
        TableRange {
            start_row: 0,
            start_col: 0,
            end_row: 3,
            end_col: 1,
        },
        "Sales",
    )
    .unwrap();
    let mut catalog = wb.saved_tables();
    catalog.version = 5;
    catalog.sheets[0].tables[0].totals = Some(TableTotals {
        visible: true,
        shown: Some(true),
        hidden_rows: Default::default(),
        columns: vec![
            TableTotal {
                label: Some("Total".into()),
                ..Default::default()
            },
            TableTotal {
                function: Some("sum".into()),
                ..Default::default()
            },
        ],
    });
    wb.restore_tables(catalog).unwrap();
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    wb
}

fn native_book() -> (Workbook, visigrid_engine::table::TableId) {
    let mut wb = Workbook::new();
    for (row, cells) in [
        ["Region", "Amount"],
        ["West", "10"],
        ["East", "20"],
        ["West", "30"],
    ]
    .iter()
    .enumerate()
    {
        for (col, cell) in cells.iter().enumerate() {
            wb.set_cell_value_tracked(0, row, col, cell);
        }
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 3,
                start_col: 0,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    (wb, id)
}

#[test]
fn table_subtotals_share_compensated_summation_with_sum_and_average() {
    let mut wb = book();
    for (row, value) in ["10000000000000000", "1", "-10000000000000000"].iter().enumerate() {
        wb.set_cell_value_tracked(0, row + 1, 1, value);
        wb.set_cell_value_tracked(0, row + 1, 0, "West");
    }
    let (sid, table) = wb.tables().next().unwrap();
    let mut spec = TableViewSpec::new(table.id);
    spec.filters.push(TableFilter {
        column: table.columns[0].id,
        criteria: ColumnFilter {
            selected: Some([NormalizedFilterKey::Text("west".into())].into()),
            text_filter: None,
        },
    });
    wb.set_table_view_spec(sid, Some(spec)).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_computed_value(4, 1).to_number().unwrap(), 1.0);
    for (formula, expected) in [
        ("=SUM(Sales[Amount])", 1.0),
        ("=SUBTOTAL(9,Sales[Amount])", 1.0),
        ("=SUBTOTAL(109,Sales[Amount])", 1.0),
        ("=AVERAGE(Sales[Amount])", 1.0 / 3.0),
        ("=SUBTOTAL(1,Sales[Amount])", 1.0 / 3.0),
        ("=SUBTOTAL(101,Sales[Amount])", 1.0 / 3.0),
    ] {
        wb.set_cell_value_tracked(0, 0, 4, formula);
        assert_eq!(wb.sheet(0).unwrap().get_computed_value(0, 4).to_number().unwrap(), expected, "{formula}");
    }
    wb.set_cell_value_tracked(0, 2, 0, "East");
    assert_eq!(wb.sheet(0).unwrap().get_computed_value(4, 1).to_number().unwrap(), 0.0);
}

#[test]
fn cycle_cache_can_mark_protected_totals_without_replacing_the_formula() {
    let mut wb = book();
    let sheet = wb.sheet_mut(0).unwrap();
    let formula = sheet.get_raw(4, 1);
    let generation = sheet.edit_generation();
    sheet.set_cycle_error(4, 1);
    assert!(sheet.is_cycle_error(4, 1));
    assert_eq!(sheet.get_display(4, 1), "#CYCLE!");
    assert_eq!(sheet.get_raw(4, 1), formula);
    assert_eq!(sheet.edit_generation(), generation);
    wb.recompute_full_ordered();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
}

#[test]
fn append_moves_footer_values_format_and_comments_in_one_replayable_commit() {
    use visigrid_engine::cell::{CellComment, NumberFormat};
    let (mut wb, id) = native_book();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    let mut format = wb.sheet(0).unwrap().get_format(4, 1);
    format.bold = true;
    format.number_format = NumberFormat::Number {
        decimals: 2,
        thousands: true,
        negative: Default::default(),
    };
    wb.sheet_mut(0).unwrap().set_format(4, 1, format.clone());
    wb.sheet_mut(0).unwrap().set_comment(
        4,
        1,
        Some(CellComment {
            text: "Visible records only".into(),
            author: "QA".into(),
        }),
    );
    let summary = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(summary, 0, 0, "=SUM(Sales[[#Totals],[Amount]])");
    wb.set_cell_value_tracked(0, 4, 4, "Neighbor");
    wb.set_cell_value_tracked(0, 8, 0, "Notes below");
    let generation = wb.sheet(summary).unwrap().edit_generation();
    let append = wb
        .append_table_rows(
            id,
            3,
            &[
                (4, 0, "North".into()),
                (4, 1, "40".into()),
                (6, 1, "50".into()),
            ],
        )
        .unwrap();
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(7));
    assert_eq!(wb.sheet(0).unwrap().get_display(7, 1), "150");
    assert_eq!(wb.sheet(summary).unwrap().get_display(0, 0), "150");
    assert!(wb.sheet(summary).unwrap().edit_generation() > generation);
    assert_eq!(wb.sheet(0).unwrap().get_format(7, 1), format);
    assert_eq!(
        wb.sheet(0).unwrap().get_cell(7, 1).comment().unwrap().text,
        "Visible records only"
    );
    assert!(wb.sheet(0).unwrap().get_cell(4, 1).comment().is_none());
    assert!(!wb.sheet(0).unwrap().get_format(4, 1).bold);
    assert_eq!(wb.sheet(0).unwrap().get_raw(5, 1), "");
    wb.apply_table_commit(&append, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(4));
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    assert_eq!(wb.sheet(0).unwrap().get_format(4, 1), format);
    assert!(wb.sheet(0).unwrap().get_cell(4, 1).comment().is_some());
    assert!(wb.sheet(0).unwrap().get_cell(7, 1).comment().is_none());
    assert_eq!(wb.sheet(0).unwrap().get_raw(6, 1), "");
    assert!(wb.sheet(0).unwrap().get_cell_opt(7, 1).is_none());
    wb.apply_table_commit(&append, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(7, 1), "150");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 4), "Neighbor");
    assert_eq!(wb.sheet(0).unwrap().get_raw(8, 0), "Notes below");
}

#[test]
fn footer_resize_expands_over_existing_records_and_shrinks_only_into_empty_cells() {
    let (mut wb, id) = native_book();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_cell_value_tracked(0, 5, 1, "25");
    let range = wb.table(id).unwrap().1.range;
    let grow = wb
        .resize_table(
            id,
            TableRange {
                end_row: 5,
                ..range
            },
        )
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "");
    assert_eq!(wb.sheet(0).unwrap().get_display(6, 1), "85");
    assert!(wb
        .resize_table(
            id,
            TableRange {
                end_row: 4,
                ..range
            }
        )
        .is_err());
    let shrink = wb.resize_table(id, range).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    assert_eq!(
        wb.sheet(0).unwrap().get_raw(5, 1),
        "25",
        "released records stay in place"
    );
    assert_eq!(wb.sheet(0).unwrap().get_raw(6, 1), "");
    wb.apply_table_commit(&shrink, true).unwrap();
    wb.apply_table_commit(&grow, true).unwrap();
    wb.apply_table_commit(&grow, false).unwrap();
    wb.apply_table_commit(&shrink, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
}

#[test]
fn footer_collision_and_stale_presentation_replay_fail_before_any_writes() {
    let (mut wb, id) = native_book();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_cell_value_tracked(0, 5, 1, "Note");
    let revision = wb.revision();
    assert!(wb
        .append_table_rows(id, 1, &[(1, 1, "999".into())])
        .is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.sheet(0).unwrap().get_raw(1, 1), "10");
    wb.clear_cell_tracked(0, 5, 1);
    let append = wb.append_table_rows(id, 1, &[]).unwrap();
    wb.sheet_mut(0).unwrap().toggle_bold(5, 1);
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&append, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(5));
    wb.sheet_mut(0).unwrap().toggle_bold(5, 1);
    wb.apply_table_commit(&append, true).unwrap();
    wb.set_cell_value_tracked(0, 5, 1, "Later note");
    assert!(wb.apply_table_commit(&append, false).is_err());
    assert_eq!(wb.sheet(0).unwrap().get_raw(5, 1), "Later note");
}

#[test]
fn fixed_and_structured_footer_references_follow_and_stale_replay_refuses() {
    let (mut wb, id) = native_book();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    let summary = wb.add_sheet_named("Summary").unwrap();
    let name = wb.sheet(0).unwrap().name.clone();
    wb.set_cell_value_tracked(summary, 0, 0, &format!("='{name}'!$B$5"));
    wb.set_cell_value_tracked(summary, 0, 1, "=SUM(Sales[[#Totals],[Amount]])");
    let append = wb.append_table_rows(id, 1, &[(4, 1, "10".into())]).unwrap();
    assert_eq!(wb.sheet(summary).unwrap().get_display(0, 0), "70");
    assert!(wb.sheet(summary).unwrap().get_raw(0, 0).ends_with("!$B$6"));
    assert_eq!(wb.sheet(summary).unwrap().get_display(0, 1), "70");
    let saved = wb.clone();
    // A later fixed reference to the old footer now means a body record.
    // Undo must not silently turn that reference into a link to totals.
    wb.set_cell_value_tracked(summary, 1, 0, &format!("='{name}'!$B$5"));
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&append, true).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.sheet(summary).unwrap().get_display(1, 0), "10");
    wb.restore_snapshot_monotonic(&saved);
    wb.apply_table_commit(&append, true).unwrap();
    assert_eq!(wb.sheet(summary).unwrap().get_display(0, 0), "60");
    for source in ["=SUM(((", "=INDIRECT(", "=OFFSET("] {
        wb.set_cell_value_tracked(0, 0, 4, source);
        let append = wb.append_table_rows(id, 1, &[]).unwrap();
        assert_eq!(wb.active_sheet().get_raw(0, 4), source);
        wb.apply_table_commit(&append, true).unwrap();
    }
    wb.clear_cell_tracked(0, 0, 4);
    assert!(
        wb.table_append_target(
            wb.active_sheet_id(),
            TableRange {
                start_row: 4,
                end_row: 4,
                start_col: 0,
                end_col: 0
            }
        )
        .is_err(),
        "typing in totals must not turn the footer into data"
    );
}

#[test]
fn footer_destination_ownership_and_bounds_are_preflighted() {
    use visigrid_engine::sheet::MergedRegion;
    for case in 0..3 {
        let (mut wb, id) = native_book();
        wb.set_table_totals_visible(
            id,
            true,
            Default::default(),
        )
        .unwrap();
        match case {
            0 => {
                wb.sheet_mut(0)
                    .unwrap()
                    .add_merge(MergedRegion::new(5, 0, 5, 1))
                    .unwrap();
            }
            1 => {
                wb.create_table(
                    wb.active_sheet_id(),
                    TableRange {
                        start_row: 5,
                        end_row: 6,
                        start_col: 0,
                        end_col: 1,
                    },
                    "Below",
                )
                .unwrap();
            }
            _ => wb.sheet_mut(0).unwrap().rows = 5,
        }
        let revision = wb.revision();
        assert!(wb.append_table_rows(id, 1, &[]).is_err(), "case {case}");
        assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(4));
        assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
        assert_eq!(wb.revision(), revision);
    }
}

#[test]
fn dormant_totals_allow_row_growth_without_creating_a_footer() {
    let (mut wb, id) = native_book();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_table_totals_visible(id, false, Default::default())
        .unwrap();
    let append = wb.append_table_rows(id, 1, &[(4, 1, "10".into())]).unwrap();
    assert!(wb.table(id).unwrap().1.totals_row().is_none());
    let show = wb
        .set_table_totals_visible(id, true, Default::default())
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(5, 1), "70");
    wb.apply_table_commit(&show, true).unwrap();
    wb.apply_table_commit(&append, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.range.end_row, 3);
}

#[test]
fn native_totals_show_edit_hide_and_replay_preserve_body_and_settings() {
    let (mut wb, id) = native_book();
    let show = wb
        .set_table_totals_visible(id, true, Default::default())
        .unwrap();
    assert!(show.is_totals_change());
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 0), "Total");
    wb.set_cell_value_tracked(0, 0, 4, "=SUM(Sales[[#Totals],[Amount]])");
    let summary = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(summary, 0, 0, "=SUM(Sales[[#Totals],[Amount]])");
    let generation = wb.sheet(summary).unwrap().edit_generation();
    let custom = wb
        .set_table_total(
            id,
            1,
            TableTotal {
                function: Some("custom".into()),
                formula: Some("= SUM([Amount]) * 2".into()),
                label: None,
            },
        )
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "120");
    assert_eq!(wb.sheet(summary).unwrap().get_display(0, 0), "120");
    assert!(
        wb.sheet(summary).unwrap().edit_generation() > generation,
        "cross-sheet pivot sources must become stale when totals change"
    );
    let label = wb
        .set_table_total(
            id,
            0,
            TableTotal {
                label: Some("=Literal label".into()),
                ..Default::default()
            },
        )
        .unwrap();
    assert!(matches!(
        wb.sheet(0).unwrap().get_cell(4, 0).value,
        visigrid_engine::cell::CellValue::Text(_)
    ));
    let hide = wb
        .set_table_totals_visible(id, false, Default::default())
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "");
    assert!(wb.sheet(0).unwrap().get_display(0, 4).starts_with("#REF!"));
    assert_eq!(wb.table(id).unwrap().1.range.data_rows(), 3);
    let reshow = wb
        .set_table_totals_visible(id, true, Default::default())
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "= SUM([Amount]) * 2");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 0), "=Literal label");
    for commit in [&reshow, &hide, &label, &custom, &show] {
        wb.apply_table_commit(commit, true).unwrap();
    }
    assert!(wb.table(id).unwrap().1.totals.is_none());
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "");
    for commit in [&show, &custom, &label, &hide, &reshow] {
        wb.apply_table_commit(commit, false).unwrap();
    }
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "120");
    assert_eq!(wb.sheet(0).unwrap().get_raw(2, 1), "20");
}

#[test]
fn native_totals_aggregate_settings_follow_filters_and_manual_hides() {
    let (mut wb, id) = native_book();
    wb.set_table_totals_visible(id, true, [2].into_iter().collect())
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "40");
    let table = wb.table(id).unwrap().1.clone();
    let mut spec = TableViewSpec::new(id);
    spec.filters.push(TableFilter {
        column: table.columns[0].id,
        criteria: ColumnFilter {
            selected: Some([NormalizedFilterKey::Text("east".into())].into()),
            text_filter: None,
        },
    });
    wb.set_table_view_spec(wb.active_sheet_id(), Some(spec))
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "0");
    wb.set_table_view_spec(wb.active_sheet_id(), None).unwrap();
    for (function, expected) in [
        ("average", "20"),
        ("min", "10"),
        ("max", "30"),
        ("count", "2"),
        ("countNums", "2"),
        ("var", "200"),
    ] {
        let commit = wb
            .set_table_total(
                id,
                1,
                TableTotal {
                    function: Some(function.into()),
                    ..Default::default()
                },
            )
            .unwrap();
        assert_eq!(
            wb.sheet(0).unwrap().get_display(4, 1),
            expected,
            "{function}"
        );
        wb.apply_table_commit(&commit, true).unwrap();
        wb.apply_table_commit(&commit, false).unwrap();
    }
}

#[test]
fn creating_a_manually_hidden_footer_preserves_visibility_and_totals_history() {
    for filtered in [false, true] {
        let (mut wb, id) = native_book();
        if filtered {
            let mut spec = TableViewSpec::new(id);
            spec.filters.push(TableFilter {
                column: wb.table(id).unwrap().1.columns[0].id,
                criteria: ColumnFilter {
                    selected: Some([NormalizedFilterKey::Text("west".into())].into()),
                    text_filter: None,
                },
            });
            wb.set_table_view_spec(wb.active_sheet_id(), Some(spec)).unwrap();
        }
        let flags = [1, 4].into();
        let (mut wb, _) = wb.prepare_table_row_visibility(wb.active_sheet_id(), flags).unwrap();
        wb.set_cell_value_tracked(0, 0, 3, "=SUBTOTAL(9,Sales[Amount])");
        wb.set_cell_value_tracked(0, 4, 3, "Keep neighboring cells");
        let spec = wb.active_sheet().table_view_spec().cloned();
        let show = wb.set_table_totals_visible(id, true, [1, 4].into()).unwrap();
        let expected = if filtered { "30" } else { "50" };
        assert_eq!(wb.active_sheet().get_display(4, 1), expected);
        assert_eq!(wb.active_sheet().get_display(0, 3), if filtered { "40" } else { "60" });
        assert_eq!(wb.active_sheet().manual_hidden_rows(), [1, 4].into());
        let hide = wb.set_table_totals_visible(id, false, [1, 4].into()).unwrap();
        assert_eq!(wb.active_sheet().get_raw(4, 1), "");
        let reshow = wb.set_table_totals_visible(id, true, [1, 4].into()).unwrap();
        assert_eq!(wb.active_sheet().get_display(4, 1), expected);
        for commit in [&reshow, &hide, &show] {
            wb.apply_table_commit(commit, true).unwrap();
            assert_eq!(wb.active_sheet().manual_hidden_rows(), [1, 4].into());
        }
        assert!(wb.table(id).unwrap().1.totals.is_none());
        for commit in [&show, &hide, &reshow] {
            wb.apply_table_commit(commit, false).unwrap();
            assert_eq!(wb.active_sheet().manual_hidden_rows(), [1, 4].into());
        }
        assert_eq!(wb.active_sheet().table_view_spec(), spec.as_ref());
        assert_eq!(wb.active_sheet().get_raw(4, 3), "Keep neighboring cells");
        assert_eq!(wb.active_sheet().get_display(4, 1), expected);
        let (mut unhidden, visibility) = wb.prepare_table_row_visibility(wb.active_sheet_id(), [1].into()).unwrap();
        assert!(unhidden.apply_table_commit(&reshow, true).is_err());
        visibility.replay(&mut unhidden, true).unwrap();
        unhidden.apply_table_commit(&reshow, true).unwrap();
    }
}

#[test]
fn native_totals_refuse_collisions_bad_formulas_and_stale_replay_atomically() {
    let (wb, id) = native_book();
    let (mut wb, _) = wb.prepare_table_row_visibility(wb.active_sheet_id(), [4].into()).unwrap();
    wb.set_cell_value_tracked(0, 4, 1, "Notes below");
    assert!(wb
        .set_table_totals_visible(id, true, Default::default())
        .is_err());
    assert!(wb.table(id).unwrap().1.totals.is_none());
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "Notes below");
    wb.clear_cell_tracked(0, 4, 1);
    let show = wb
        .set_table_totals_visible(id, true, [4].into_iter().collect())
        .unwrap();
    let before = wb.sheet(0).unwrap().get_raw(4, 1);
    assert!(wb
        .set_table_total(
            id,
            1,
            TableTotal {
                function: Some("custom".into()),
                formula: Some("=SUM(".into()),
                label: None
            }
        )
        .is_err());
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), before);
    wb.apply_table_commit(&show, true).unwrap();
    wb.set_cell_value_tracked(0, 4, 1, "Later data");
    assert!(wb.apply_table_commit(&show, false).is_err());
    assert!(wb.table(id).unwrap().1.totals.is_none());
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "Later data");
}

#[test]
fn native_footer_cannot_claim_merges_comments_other_tables_or_out_of_bounds_cells() {
    use visigrid_engine::{
        cell::CellComment,
        sheet::{MergedRegion, NUM_ROWS},
    };
    let (wb, id) = native_book();
    let (mut wb, _) = wb.prepare_table_row_visibility(wb.active_sheet_id(), [4].into()).unwrap();
    wb.sheet_mut(0)
        .unwrap()
        .add_merge(MergedRegion::new(4, 0, 4, 1))
        .unwrap();
    assert!(wb
        .set_table_totals_visible(id, true, Default::default())
        .is_err());
    wb.sheet_mut(0).unwrap().remove_merge((4, 0));
    wb.sheet_mut(0).unwrap().set_comment(
        4,
        1,
        Some(CellComment {
            text: "Keep me".into(),
            author: "QA".into(),
        }),
    );
    assert!(wb
        .set_table_totals_visible(id, true, Default::default())
        .is_err());
    wb.sheet_mut(0).unwrap().set_comment(4, 1, None);
    wb.create_table(
        wb.active_sheet_id(),
        TableRange {
            start_row: 4,
            end_row: 5,
            start_col: 0,
            end_col: 1,
        },
        "Below",
    )
    .unwrap();
    assert!(wb
        .set_table_totals_visible(id, true, Default::default())
        .is_err());
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: NUM_ROWS - 1,
                end_row: NUM_ROWS - 1,
                start_col: 3,
                end_col: 3,
            },
            "Last",
        )
        .unwrap()
        .table_id();
    assert!(wb
        .set_table_totals_visible(id, true, Default::default())
        .is_err());
}

#[test]
fn empty_native_table_and_escaped_column_name_have_valid_totals() {
    let mut wb = Workbook::new();
    wb.set_cell_value_tracked(0, 0, 0, "Amount [net]");
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                end_row: 0,
                start_col: 0,
                end_col: 0,
            },
            "Empty",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(1, 0), "0");
    assert_eq!(
        wb.sheet(0).unwrap().get_raw(1, 0),
        "=SUBTOTAL(109,[Amount '[net']])"
    );
}

#[test]
fn subtotal_function_numbers_and_hidden_rows_remain_distinct() {
    let mut wb = book();
    let mut catalog = wb.saved_tables();
    catalog.sheets[0].tables[0]
        .totals
        .as_mut()
        .unwrap()
        .hidden_rows
        .insert(2);
    wb.restore_tables(catalog).unwrap();
    for (function, expected) in [
        (1, "20"),
        (2, "3"),
        (3, "3"),
        (4, "30"),
        (5, "10"),
        (6, "6000"),
        (7, "10"),
        (9, "60"),
        (10, "100"),
        (101, "20"),
        (102, "2"),
        (103, "2"),
        (104, "30"),
        (105, "10"),
        (106, "300"),
        (109, "40"),
        (110, "200"),
    ] {
        wb.set_cell_value_tracked(0, 0, 4, &format!("=SUBTOTAL({function},Sales[Amount])"));
        assert_eq!(
            wb.sheet(0).unwrap().get_display(0, 4),
            expected,
            "code {function}"
        );
    }
    wb.set_cell_value_tracked(0, 0, 4, "=SUBTOTAL(12,Sales[Amount])");
    assert!(wb
        .sheet(0)
        .unwrap()
        .get_display(0, 4)
        .starts_with("#VALUE!"));
}

#[test]
fn filter_changes_and_cross_sheet_precedents_recalculate_totals_and_dependents() {
    let mut wb = book();
    let other = wb.add_sheet_named("Control").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "East");
    wb.set_cell_value_tracked(0, 2, 0, "=Control!A1");
    wb.set_cell_value_tracked(other, 0, 1, "=SUM(Sales[[#Totals],[Amount]])");
    let (sid, table) = wb.tables().next().unwrap();
    let mut spec = TableViewSpec::new(table.id);
    spec.filters.push(TableFilter {
        column: table.columns[0].id,
        criteria: ColumnFilter {
            selected: Some([NormalizedFilterKey::Text("west".into())].into()),
            text_filter: None,
        },
    });
    let generation = wb.sheet(other).unwrap().edit_generation();
    let commit = wb.set_table_view_spec(sid, Some(spec)).unwrap();
    assert!(
        wb.sheet(other).unwrap().edit_generation() > generation,
        "changed subtotal dependents must invalidate pivot sources"
    );
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "40");
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 1), "40");
    wb.set_cell_value_tracked(other, 0, 0, "West");
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 1), "60");
    wb.set_cell_value_tracked(other, 0, 0, "East");
    wb.apply_table_view_commit(&commit, true).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    wb.apply_table_view_commit(&commit, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "40");
}

#[test]
fn structured_totals_and_all_sections_include_footer_only_where_requested() {
    let mut wb = book();
    for (source, expected) in [
        ("=SUM(Sales[Amount])", "60"),
        ("=SUM(Sales[[#All],[Amount]])", "120"),
        ("=SUM(Sales[[#Totals],[Amount]])", "60"),
    ] {
        wb.set_cell_value_tracked(0, 0, 4, source);
        assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), expected);
    }
    let mut catalog = wb.saved_tables();
    catalog.sheets[0].tables[0].totals.as_mut().unwrap().visible = false;
    wb.restore_tables(catalog).unwrap();
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    assert!(wb.sheet(0).unwrap().get_display(0, 4).starts_with("#REF!"));
}

#[test]
fn nested_subtotals_are_ignored_and_numeric_looking_text_stays_text() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 2, 1, "=SUBTOTAL(9,G1:G1)");
    wb.set_cell_text_tracked(0, 3, 1, "30");
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "10");
    wb.set_cell_value_tracked(0, 0, 4, "=SUBTOTAL(103,Sales[Amount])");
    assert_eq!(wb.sheet(0).unwrap().get_display(0, 4), "2");
}

#[test]
fn totals_ownership_and_conversion_preserve_formulas_and_history() {
    let mut wb = book();
    let id = wb.tables().next().unwrap().1.id;
    let before = wb.saved_tables();
    let rename = wb.rename_table(id, "Renamed").unwrap();
    wb.apply_table_commit(&rename, true).unwrap();
    let append = wb.append_table_rows(id, 1, &[]).unwrap();
    wb.apply_table_commit(&append, true).unwrap();
    let (copied, index) = wb.prepare_sheet_copy(&wb, wb.active_sheet_id(), "Copy").unwrap();
    assert_ne!(copied.sheet(index).unwrap().tables()[0].id, id);
    assert_eq!(copied.sheet(index).unwrap().get_display(4, 1), "60");
    assert_eq!(wb.sheet_count(), 1);
    wb.set_cell_value_tracked(0, 4, 1, "999");
    wb.sheet_mut(0).unwrap().set_value(4, 1, "999");
    wb.sheet_mut(0).unwrap().clear_cell(4, 1);
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    let history = wb.prepare_table_row_history(0, 2, 1, false).unwrap().unwrap();
    wb.apply_table_row_history(&history, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(5, 1), "60");
    wb.apply_table_row_history(&history, true).unwrap();
    assert_eq!(
        serde_json::to_value(wb.saved_tables()).unwrap(),
        serde_json::to_value(before).unwrap()
    );
    let commit = wb.remove_table(id).unwrap();
    assert!(wb.sheet(0).unwrap().get_raw(4, 1).contains("$B$2:$B$4"));
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(
        wb.sheet(0).unwrap().get_raw(4, 1),
        "=SUBTOTAL(109,[Amount])"
    );
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
}

#[test]
fn totals_catalog_rejects_bad_versions_and_bounds_without_mutation() {
    let mut wb = book();
    let original = wb.saved_tables();
    let mut old = original.clone();
    old.version = 4;
    assert!(wb.restore_tables(old).is_err());
    let mut bad = original.clone();
    bad.sheets[0].tables[0]
        .totals
        .as_mut()
        .unwrap()
        .columns
        .pop();
    assert!(wb.restore_tables(bad).is_err());
    let mut bad = original.clone();
    bad.sheets[0].tables[0].range.end_row = visigrid_engine::sheet::NUM_ROWS - 1;
    assert!(wb.restore_tables(bad).is_err());
    assert_eq!(
        serde_json::to_value(wb.saved_tables()).unwrap(),
        serde_json::to_value(original).unwrap()
    );
}

#[test]
fn cross_sheet_structural_rewrite_of_a_footer_is_undoable() {
    let mut wb = book();
    let other = wb.add_sheet_named("Control").unwrap();
    wb.set_cell_value_tracked(other, 1, 0, "25");
    let mut catalog = wb.saved_tables();
    let mut cells_only = catalog.clone();
    cells_only.sheets[0].tables.clear();
    wb.restore_tables(cells_only).unwrap();
    wb.set_cell_value_tracked(0, 4, 1, "=Control!A2");
    let total = &mut catalog.sheets[0].tables[0].totals.as_mut().unwrap().columns[1];
    total.function = Some("custom".into());
    total.formula = Some("=Control!A2".into());
    wb.restore_tables(catalog).unwrap();
    wb.rebuild_dep_graph();
    wb.recompute_full_ordered();
    let history = wb.prepare_table_row_history(other, 0, 1, false).unwrap().unwrap();
    wb.apply_table_row_history(&history, false).unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_raw(2, 0), "25");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "=Control!A3");
    assert_eq!(wb.sheet(0).unwrap().tables()[0].totals.as_ref().unwrap().columns[1].formula.as_deref(), Some("=Control!A3"));
    wb.apply_table_row_history(&history, true).unwrap();
    assert_eq!(wb.sheet(other).unwrap().get_raw(1, 0), "25");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "=Control!A2");
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "25");
}

#[test]
fn totals_renames_rewrite_cells_rules_and_custom_settings_without_moving_records() {
    let (mut wb, id) = native_book();
    wb.set_calculated_column(id, 1, 1, "=IF([@Region]=\"West\",10,20)", true).unwrap();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    wb.set_table_total(id, 1, TableTotal {
        function: Some("custom".into()),
        formula: Some("=SUM([Amount])+SUM(Sales[Amount])+IF(\"Sales[Amount]\"=\"x\",1,0)".into()),
        label: None,
    }).unwrap();
    let other = wb.add_sheet_named("Summary").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=SUM(Sales[[#Totals],[Amount]])");
    let before = wb.saved_tables();
    let ids: Vec<_> = wb.table(id).unwrap().1.columns.iter().map(|c| c.id).collect();
    let rename = wb.rename_table(id, "Orders").unwrap();
    assert!(rename.is_name_change());
    let headers = wb.rename_table_columns(id, &["Area".into(), "Net [USD]".into()]).unwrap();
    assert!(headers.is_name_change());
    let expected = "=SUM([Net '[USD']])+SUM(Orders[Net '[USD']])+IF(\"Sales[Amount]\"=\"x\",1,0)";
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), expected);
    assert_eq!(wb.table(id).unwrap().1.totals.as_ref().unwrap().columns[1].formula.as_deref(), Some(expected));
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 0), "80");
    assert_eq!(wb.sheet(0).unwrap().get_raw(1, 1), "=IF([@[Area]]=\"West\",10,20)");
    assert_eq!(ids, wb.table(id).unwrap().1.columns.iter().map(|c| c.id).collect::<Vec<_>>());
    wb.apply_table_commit(&headers, true).unwrap();
    wb.apply_table_commit(&rename, true).unwrap();
    assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), serde_json::to_value(before).unwrap());
    wb.apply_table_commit(&rename, false).unwrap();
    wb.apply_table_commit(&headers, false).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), expected);
    assert_eq!(wb.sheet(other).unwrap().get_display(0, 0), "80");
}

#[test]
fn dormant_totals_rename_local_references_and_restore_them_when_shown() {
    let (mut wb, id) = native_book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    wb.set_table_total(id, 1, TableTotal { function: Some("custom".into()),
        formula: Some("=SUM([Amount])*2".into()), label: None }).unwrap();
    wb.set_table_totals_visible(id, false, Default::default()).unwrap();
    let rename = wb.rename_table_columns(id, &["Area".into(), "Revenue".into()]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "");
    assert_eq!(wb.table(id).unwrap().1.totals.as_ref().unwrap().columns[1].formula.as_deref(), Some("=SUM([Revenue])*2"));
    wb.apply_table_commit(&rename, true).unwrap();
    assert_eq!(wb.table(id).unwrap().1.totals.as_ref().unwrap().columns[1].formula.as_deref(), Some("=SUM([Amount])*2"));
    wb.apply_table_commit(&rename, false).unwrap();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "=SUM([Revenue])*2");
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "120");
}

#[test]
fn rename_updates_other_tables_visible_and_dormant_totals_and_rejects_stale_replay() {
    for visible in [true, false] {
        let (mut wb, id) = native_book();
        let other = wb.add_sheet_named("Other").unwrap();
        wb.set_cell_value_tracked(other, 0, 0, "Value");
        wb.set_cell_value_tracked(other, 1, 0, "1");
        let other_id = wb.create_table(wb.sheet(other).unwrap().id, TableRange {
            start_row: 0, end_row: 1, start_col: 0, end_col: 0,
        }, "Summary").unwrap().table_id();
        wb.set_table_totals_visible(other_id, true, Default::default()).unwrap();
        wb.set_table_total(other_id, 0, TableTotal { function: Some("custom".into()),
            formula: Some("=SUM(Sales[Amount])".into()), label: None }).unwrap();
        if !visible { wb.set_table_totals_visible(other_id, false, Default::default()).unwrap(); }
        let rename = wb.rename_table(id, "Orders").unwrap();
        assert_eq!(wb.table(other_id).unwrap().1.totals.as_ref().unwrap().columns[0].formula.as_deref(), Some("=SUM(Orders[Amount])"));
        if visible {
            assert_eq!(wb.sheet(other).unwrap().get_raw(2, 0), "=SUM(Orders[Amount])");
            assert_eq!(wb.sheet(other).unwrap().get_display(2, 0), "60");
        }
        wb.apply_table_commit(&rename, true).unwrap();
        wb.apply_table_commit(&rename, false).unwrap();
        if !visible { wb.set_table_totals_visible(other_id, true, Default::default()).unwrap(); }
        wb.set_table_total(other_id, 0, TableTotal { function: Some("custom".into()),
            formula: Some("=SUM(Orders[Amount])*2".into()), label: None }).unwrap();
        let revision = wb.revision();
        assert!(wb.apply_table_commit(&rename, true).is_err());
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.table(id).unwrap().1.name, "Orders");
        assert_eq!(wb.sheet(other).unwrap().get_display(2, 0), "120");
    }
}

#[test]
fn totals_header_swap_keeps_filters_and_labels_bound_to_ids_and_refuses_duplicate_names() {
    let mut wb = book();
    let id = wb.tables().next().unwrap().1.id;
    let table = wb.table(id).unwrap().1;
    let mut spec = TableViewSpec::new(id);
    spec.filters.push(TableFilter { column: table.columns[0].id, criteria: ColumnFilter {
        selected: Some([NormalizedFilterKey::Text("west".into())].into()), text_filter: None,
    }});
    wb.set_table_view_spec(wb.active_sheet_id(), Some(spec.clone())).unwrap();
    let rename = wb.rename_table_columns(id, &["Amount".into(), "Region".into()]).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 1), "=SUBTOTAL(109,[Region])");
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "40");
    assert_eq!(wb.sheet(0).unwrap().get_raw(4, 0), "Total");
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
    let revision = wb.revision();
    assert!(wb.rename_table_columns(id, &["Same".into(), "same".into()]).is_err());
    assert_eq!(wb.revision(), revision);
    wb.apply_table_commit(&rename, true).unwrap();
    assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "40");
}

#[test]
fn new_dormant_totals_reference_makes_rename_undo_stale_without_partial_writes() {
    let (mut wb, id) = native_book();
    let rename = wb.rename_table(id, "Orders").unwrap();
    let other = wb.add_sheet_named("Other").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "Value");
    let other_id = wb.create_table(wb.sheet(other).unwrap().id, TableRange {
        start_row: 0, end_row: 0, start_col: 0, end_col: 0,
    }, "Summary").unwrap().table_id();
    wb.set_table_totals_visible(other_id, true, Default::default()).unwrap();
    wb.set_table_total(other_id, 0, TableTotal { function: Some("custom".into()),
        formula: Some("=SUM(Orders[Amount])".into()), label: None }).unwrap();
    wb.set_table_totals_visible(other_id, false, Default::default()).unwrap();
    let revision = wb.revision();
    assert!(wb.apply_table_commit(&rename, true).unwrap_err().contains("New totals references"));
    assert_eq!(wb.revision(), revision);
    assert_eq!(wb.table(id).unwrap().1.name, "Orders");
}

#[test]
fn renaming_a_header_invalidates_cross_sheet_pivot_sources_when_custom_totals_change() {
    let (mut wb, id) = native_book();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    wb.set_table_total(id, 1, TableTotal { function: Some("custom".into()),
        formula: Some("=IF(COUNTIF(Sales[#Headers],\"Amount\"),1,2)".into()), label: None }).unwrap();
    let summary = wb.add_sheet_named("Summary").unwrap();
    // A fixed reference is not itself rewritten by the rename.
    wb.set_cell_value_tracked(summary, 0, 0, "=Sheet1!B5");
    assert_eq!(wb.sheet(summary).unwrap().get_display(0, 0), "1");
    let generation = wb.sheet(summary).unwrap().edit_generation();
    let rename = wb.rename_table_columns(id, &["Region".into(), "Revenue".into()]).unwrap();
    assert_eq!(wb.sheet(summary).unwrap().get_display(0, 0), "2");
    assert!(wb.sheet(summary).unwrap().edit_generation() > generation);
    let generation = wb.sheet(summary).unwrap().edit_generation();
    wb.apply_table_commit(&rename, true).unwrap();
    assert_eq!(wb.sheet(summary).unwrap().get_display(0, 0), "1");
    assert!(wb.sheet(summary).unwrap().edit_generation() > generation);
}

#[test]
fn manual_visibility_updates_totals_and_dependents_with_guarded_replay() {
    let mut wb = book();
    wb.set_cell_value_tracked(0, 6, 1, "=SUBTOTAL(9,Sales[Amount])");
    let other = wb.add_sheet_named("Dependents").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=SUM(Sales[#Totals])*2");
    let generation = wb.sheet(other).unwrap().edit_generation();
    let (mut after, commit) = wb.prepare_table_row_visibility(wb.sheet(0).unwrap().id, [2].into_iter().collect()).unwrap();
    assert_eq!(wb.sheet(0).unwrap().get_display(4, 1), "60");
    assert_eq!(after.sheet(0).unwrap().get_display(4, 1), "40");
    assert_eq!(after.sheet(0).unwrap().get_display(6, 1), "60");
    assert_eq!(after.sheet(other).unwrap().get_display(0, 0), "80");
    assert!(after.sheet(other).unwrap().edit_generation() > generation);
    assert_eq!(commit.changed_cell_count(), 0);
    commit.replay(&mut after, true).unwrap();
    assert_eq!(after.sheet(0).unwrap().get_display(4, 1), "60");
    assert_eq!(after.sheet(other).unwrap().get_display(0, 0), "120");
    commit.replay(&mut after, false).unwrap();
    assert_eq!(after.sheet(0).unwrap().get_display(4, 1), "40");
    after.set_cell_value_tracked(0, 1, 1, "123");
    assert!(commit.replay(&mut after, true).is_err());
}

#[test]
fn manual_visibility_keeps_dormant_settings_and_can_hide_or_unhide_footer() {
    let mut wb = book();
    let id = wb.tables().next().unwrap().1.id;
    wb.set_table_totals_visible(id, false, Default::default()).unwrap();
    let (mut hidden, _) = wb.prepare_table_row_visibility(wb.active_sheet_id(), [2].into_iter().collect()).unwrap();
    assert!(!hidden.table(id).unwrap().1.totals.as_ref().unwrap().visible);
    hidden.set_table_totals_visible(id, true, [2].into_iter().collect()).unwrap();
    assert_eq!(hidden.active_sheet().get_display(4, 1), "40");
    let (mut hidden, _) = hidden.prepare_table_row_visibility(hidden.active_sheet_id(), [2, 4].into_iter().collect()).unwrap();
    let append = hidden.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(hidden.active_sheet().manual_hidden_rows(), [2, 4].into());
    assert_eq!(hidden.active_sheet().get_display(5, 1), "40");
    hidden.apply_table_commit(&append, true).unwrap();
    let (mut shown, _) = hidden.prepare_table_row_visibility(hidden.active_sheet_id(), Default::default()).unwrap();
    assert_eq!(shown.active_sheet().get_display(4, 1), "60");
    shown.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(shown.active_sheet().get_display(5, 1), "60");
}

#[test]
fn manual_visibility_bounds_and_recovery_refuse_without_mutation() {
    let mut wb = book();
    let before = serde_json::to_value(wb.saved_tables()).unwrap();
    let revision = wb.revision();
    assert!(wb.prepare_table_row_visibility(wb.active_sheet_id(), [visigrid_engine::sheet::NUM_ROWS].into_iter().collect()).is_err());
    assert!(wb.prepare_table_row_visibility(visigrid_engine::sheet::SheetId(999), Default::default()).is_err());
    wb.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(wb.prepare_table_row_visibility(wb.active_sheet_id(), [2].into_iter().collect()).is_err());
    assert_eq!(wb.revision(), revision);
    assert_eq!(serde_json::to_value(wb.saved_tables()).unwrap(), before);
}

#[test]
fn manual_visibility_updates_all_totals_on_the_sheet_without_stale_flags() {
    let mut wb = book();
    for (row, label, amount) in [(8, "Region", "Amount"), (9, "West", "5"), (10, "East", "15")] {
        wb.set_cell_value_tracked(0, row, 0, label);
        wb.set_cell_value_tracked(0, row, 1, amount);
    }
    let id = wb.create_table(wb.active_sheet_id(), TableRange {
        start_row: 8, end_row: 10, start_col: 0, end_col: 1,
    }, "OtherData").unwrap().table_id();
    wb.set_table_totals_visible(id, true, Default::default()).unwrap();
    let flags = [2, 10].into_iter().collect();
    let (hidden, _) = wb.prepare_table_row_visibility(wb.active_sheet_id(), flags).unwrap();
    assert_eq!(hidden.active_sheet().get_display(4, 1), "40");
    assert_eq!(hidden.active_sheet().get_display(11, 1), "5");
    for (_, table) in hidden.tables() {
        assert_eq!(table.totals.as_ref().unwrap().hidden_rows, [2, 10].into_iter().collect());
    }
    let (shown, _) = hidden.prepare_table_row_visibility(hidden.active_sheet_id(), Default::default()).unwrap();
    assert_eq!(shown.active_sheet().get_display(4, 1), "60");
    assert_eq!(shown.active_sheet().get_display(11, 1), "20");
    assert!(shown.tables().all(|(_, t)| t.totals.as_ref().unwrap().hidden_rows.is_empty()));
}
