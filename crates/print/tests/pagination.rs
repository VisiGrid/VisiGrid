use visigrid_print::*;

fn sheet(rows: usize, columns: usize, height: f64, width: f64) -> LayoutInput {
    LayoutInput {
        rows: (0..rows)
            .map(|i| AxisItem {
                source_index: i,
                size_pt: height,
            })
            .collect(),
        columns: (0..columns)
            .map(|i| AxisItem {
                source_index: i,
                size_pt: width,
            })
            .collect(),
        ..Default::default()
    }
}

// Letter minus 36pt margins = 540 × 720pt body.
fn settings() -> PageSettings {
    PageSettings {
        paper: Paper::Letter,
        scale: Scale::Actual,
        ..Default::default()
    }
}

#[test]
fn exact_boundary_has_no_trailing_page() {
    let plan = paginate(&sheet(36, 3, 20.0, 180.0), &settings()).unwrap();
    assert_eq!(
        plan.pages(),
        &[Page {
            rows: 0..36,
            columns: 0..3
        }]
    );
    assert_eq!(
        plan.rect(0, 35..36, 2..3),
        Some(Rect {
            x: 396.0,
            y: 736.0,
            width: 180.0,
            height: 20.0
        })
    );
}

#[test]
fn pages_go_down_then_over_with_complete_coverage() {
    let input = sheet(80, 8, 20.0, 100.0);
    let plan = paginate(&input, &settings()).unwrap();
    assert_eq!(plan.pages().len(), 6);
    assert_eq!(
        plan.pages()[0],
        Page {
            rows: 0..36,
            columns: 0..5
        }
    );
    assert_eq!(
        plan.pages()[1],
        Page {
            rows: 36..72,
            columns: 0..5
        }
    );
    assert_eq!(
        plan.pages()[3],
        Page {
            rows: 0..36,
            columns: 5..8
        }
    );
    for row in 0..80 {
        for col in 0..8 {
            assert_eq!(
                (0..6)
                    .filter(|&p| plan.rect(p, row..row + 1, col..col + 1).is_some())
                    .count(),
                1
            );
        }
    }
}

#[test]
fn repeated_titles_reserve_space_and_occur_once_on_each_page() {
    let input = sheet(80, 8, 20.0, 100.0);
    let s = PageSettings {
        repeat_rows: 2,
        repeat_columns: 1,
        ..settings()
    };
    let plan = paginate(&input, &s).unwrap();
    assert_eq!(
        plan.pages()[0],
        Page {
            rows: 2..36,
            columns: 1..5
        }
    );
    assert_eq!(
        plan.pages()[1],
        Page {
            rows: 36..70,
            columns: 1..5
        }
    );
    for (p, page) in plan.pages().iter().enumerate() {
        assert_eq!(
            plan.rect(p, 0..2, 0..1),
            Some(Rect {
                x: 36.0,
                y: 36.0,
                width: 100.0,
                height: 40.0
            })
        );
        let body = plan
            .rect(
                p,
                page.rows.start..page.rows.start + 1,
                page.columns.start..page.columns.start + 1,
            )
            .unwrap();
        assert_eq!((body.x, body.y), (136.0, 76.0));
    }
}

#[test]
fn fit_width_keeps_height_unconstrained_and_reports_real_text_size() {
    let mut input = sheet(80, 9, 20.0, 100.0);
    input.text.push(TextCell {
        row: 0,
        column: 0,
        font_size_pt: 10.0,
    });
    let plan = paginate(
        &input,
        &PageSettings {
            scale: Scale::FitColumns,
            ..settings()
        },
    )
    .unwrap();
    assert_eq!(plan.scale(), 0.6);
    assert_eq!(plan.scale_constraint(), ScaleConstraint::Width);
    assert_eq!(plan.pages().len(), 2);
    assert_eq!(
        plan.readability(),
        &Readability {
            smallest_text_pt: Some(6.0),
            cells_below_threshold: 1
        }
    );
}

#[test]
fn fit_sheet_reports_height_constraint_and_never_enlarges() {
    let plan = paginate(
        &sheet(72, 2, 20.0, 100.0),
        &PageSettings {
            scale: Scale::FitSheet,
            ..settings()
        },
    )
    .unwrap();
    assert_eq!(plan.scale(), 0.5);
    assert_eq!(plan.scale_constraint(), ScaleConstraint::Height);
    assert_eq!(plan.pages().len(), 1);
    for scale in [Scale::FitSheet, Scale::FitColumns] {
        assert_eq!(
            paginate(
                &sheet(1, 1, 20.0, 100.0),
                &PageSettings {
                    scale,
                    ..settings()
                }
            )
            .unwrap()
            .scale(),
            1.0
        );
    }
}

#[test]
fn merge_moves_to_next_page_without_being_split() {
    let mut input = sheet(40, 3, 20.0, 100.0);
    input.merges.push(Merge {
        rows: 34..38,
        columns: 0..2,
    });
    let plan = paginate(&input, &settings()).unwrap();
    assert_eq!(plan.pages()[0].rows, 0..34);
    assert_eq!(plan.pages()[1].rows, 34..40);
    assert!(plan.rect(0, 34..38, 0..2).is_none());
    assert_eq!(plan.rect(1, 34..38, 0..2).unwrap().height, 80.0);
}

#[test]
fn oversized_and_transitively_overlapping_merge_spans_fail() {
    let mut input = sheet(50, 2, 20.0, 100.0);
    input.merges = vec![
        Merge {
            rows: 0..25,
            columns: 0..1,
        },
        Merge {
            rows: 24..49,
            columns: 1..2,
        },
    ];
    assert_eq!(
        paginate(&input, &settings()),
        Err(LayoutError::CannotFit {
            axis: Axis::Row,
            positions: 0..49
        })
    );
    assert!(matches!(
        paginate(&sheet(1, 1, 721.0, 100.0), &settings()),
        Err(LayoutError::CannotFit {
            axis: Axis::Row,
            ..
        })
    ));
    assert!(matches!(
        paginate(&sheet(1, 1, 20.0, 541.0), &settings()),
        Err(LayoutError::CannotFit {
            axis: Axis::Column,
            ..
        })
    ));
}

#[test]
fn titles_cannot_cut_merges_consume_whole_axis_or_consume_body() {
    let mut input = sheet(40, 3, 20.0, 100.0);
    input.merges.push(Merge {
        rows: 0..2,
        columns: 0..2,
    });
    assert_eq!(
        paginate(
            &input,
            &PageSettings {
                repeat_rows: 1,
                ..settings()
            }
        ),
        Err(LayoutError::TitlesCutMerge {
            axis: Axis::Row,
            merge: 0
        })
    );
    assert_eq!(
        paginate(
            &input,
            &PageSettings {
                repeat_columns: 1,
                ..settings()
            }
        ),
        Err(LayoutError::TitlesCutMerge {
            axis: Axis::Column,
            merge: 0
        })
    );
    assert_eq!(
        paginate(
            &input,
            &PageSettings {
                repeat_rows: 40,
                ..settings()
            }
        ),
        Err(LayoutError::InvalidTitles { axis: Axis::Row })
    );
    assert_eq!(
        paginate(
            &input,
            &PageSettings {
                repeat_rows: 36,
                ..settings()
            }
        ),
        Err(LayoutError::InvalidTitles { axis: Axis::Row })
    );
}

#[test]
fn validation_rejects_nonfinite_and_invalid_settings() {
    let input = sheet(1, 1, 20.0, 100.0);
    for scale in [f64::NAN, f64::INFINITY, 0.0, -1.0, 0.09, 4.01] {
        assert_eq!(
            paginate(
                &input,
                &PageSettings {
                    scale: Scale::Custom(scale),
                    ..settings()
                }
            ),
            Err(LayoutError::InvalidScale)
        );
    }
    for left in [f64::NAN, f64::INFINITY, -1.0, 600.0] {
        let s = PageSettings {
            margins: Margins {
                left,
                ..Default::default()
            },
            ..settings()
        };
        assert_eq!(paginate(&input, &s), Err(LayoutError::InvalidMargins));
    }
    for height in [f64::NAN, f64::INFINITY, 0.0, -1.0] {
        assert!(matches!(
            paginate(&sheet(1, 1, height, 100.0), &settings()),
            Err(LayoutError::InvalidAxis { .. })
        ));
    }
}

#[test]
fn no_content_and_resource_limits_fail_before_page_materialization() {
    assert_eq!(
        paginate(&LayoutInput::default(), &settings()),
        Err(LayoutError::NothingVisible)
    );
    assert_eq!(
        paginate(&sheet(1001, 1000, 1.0, 1.0), &settings()),
        Err(LayoutError::TooManyCells)
    );
    assert_eq!(
        paginate(&sheet(1001, 1, 720.0, 10.0), &settings()),
        Err(LayoutError::TooManyPages)
    );
    assert_eq!(
        paginate(&sheet(33, 33, 720.0, 540.0), &settings()),
        Err(LayoutError::TooManyPages)
    );
    assert!(matches!(
        paginate(
            &sheet(1, 100, 20.0, 100.0),
            &PageSettings {
                scale: Scale::FitColumns,
                ..settings()
            }
        ),
        Err(LayoutError::FitTooSmall { .. })
    ));
}

#[test]
fn source_mapping_preserves_gaps_and_sorted_display_order() {
    let mut input = sheet(3, 2, 20.0, 100.0);
    input.rows[0].source_index = 99;
    input.rows[1].source_index = 42;
    input.columns[1].source_index = 12;
    let plan = paginate(&input, &settings()).unwrap();
    assert_eq!(plan.source_cell(0, 1), Some((99, 12)));
    assert_eq!(plan.source_cell(1, 1), Some((42, 12)));
    assert_eq!(plan.rect(0, 1..2, 1..2).unwrap().y, 56.0);
    input.rows[2].source_index = 42;
    assert!(matches!(
        paginate(&input, &settings()),
        Err(LayoutError::DuplicateSourceIndex { .. })
    ));
}

#[test]
fn footer_and_orientation_affect_geometry_but_not_screen_state() {
    let input = sheet(36, 3, 20.0, 180.0);
    let plan = paginate(
        &input,
        &PageSettings {
            footer: true,
            ..settings()
        },
    )
    .unwrap();
    assert_eq!(plan.body().height, 702.0);
    assert_eq!(plan.pages().len(), 2);
    let plan = paginate(
        &input,
        &PageSettings {
            landscape: true,
            ..settings()
        },
    )
    .unwrap();
    assert_eq!(plan.page_size(), (792.0, 612.0));
    assert_eq!(plan.body().width, 720.0);
    assert_eq!(plan.pages()[0].rows, 0..27);
}

#[test]
fn invalid_merges_text_and_duplicate_cells_are_rejected() {
    let mut input = sheet(4, 4, 20.0, 100.0);
    input.merges.push(Merge {
        rows: 0..5,
        columns: 0..2,
    });
    assert_eq!(
        paginate(&input, &settings()),
        Err(LayoutError::InvalidMerge { index: 0 })
    );
    input.merges = vec![
        Merge {
            rows: 0..2,
            columns: 0..2,
        },
        Merge {
            rows: 1..3,
            columns: 1..3,
        },
    ];
    assert_eq!(
        paginate(&input, &settings()),
        Err(LayoutError::OverlappingMerges {
            first: 0,
            second: 1
        })
    );
    input.merges.clear();
    input.text = vec![TextCell {
        row: 0,
        column: 0,
        font_size_pt: f64::NAN,
    }];
    assert_eq!(
        paginate(&input, &settings()),
        Err(LayoutError::InvalidTextCell { index: 0 })
    );
    input.text = vec![
        TextCell {
            row: 0,
            column: 0,
            font_size_pt: 10.0
        };
        2
    ];
    assert_eq!(
        paginate(&input, &settings()),
        Err(LayoutError::InvalidTextCell { index: 1 })
    );
}

#[test]
fn fractional_sizes_do_not_create_a_spurious_page() {
    let plan = paginate(&sheet(100, 100, 7.2, 5.4), &settings()).unwrap();
    assert_eq!(plan.pages().len(), 1);
    assert!(plan.rect(0, 0..0, 0..1).is_none());
    assert!(plan.rect(0, 0..101, 0..1).is_none());
    assert!(plan.rect(1, 0..1, 0..1).is_none());
}

#[test]
fn many_shapes_have_complete_coverage_and_stay_inside_body() {
    for nrows in [1, 35, 36, 37, 95] {
        for ncols in [1, 5, 6, 11] {
            for scale in [
                Scale::Actual,
                Scale::FitColumns,
                Scale::FitSheet,
                Scale::Custom(1.25),
            ] {
                let input = sheet(nrows, ncols, 20.0, 100.0);
                let s = PageSettings {
                    scale,
                    repeat_rows: usize::from(nrows > 1),
                    repeat_columns: usize::from(ncols > 1),
                    ..settings()
                };
                let plan = paginate(&input, &s).unwrap();
                let body = plan.body();
                for row in s.repeat_rows..nrows {
                    for col in s.repeat_columns..ncols {
                        let rects: Vec<_> = (0..plan.pages().len())
                            .filter_map(|p| plan.rect(p, row..row + 1, col..col + 1))
                            .collect();
                        assert_eq!(rects.len(), 1);
                        let r = rects[0];
                        assert!(r.x >= body.x && r.y >= body.y);
                        assert!(r.x + r.width <= body.x + body.width + 1e-7);
                        assert!(r.y + r.height <= body.y + body.height + 1e-7);
                    }
                }
            }
        }
    }
}

#[test]
fn covered_merge_text_is_rejected_and_repeated_text_counts_once() {
    let mut input = sheet(80, 3, 20.0, 100.0);
    input.merges.push(Merge {
        rows: 0..1,
        columns: 0..2,
    });
    input.text.push(TextCell {
        row: 0,
        column: 1,
        font_size_pt: 7.0,
    });
    assert_eq!(
        paginate(&input, &settings()),
        Err(LayoutError::InvalidTextCell { index: 0 })
    );
    input.text[0].column = 0;
    let original = input.clone();
    let s = PageSettings {
        repeat_rows: 1,
        ..settings()
    };
    let plan = paginate(&input, &s).unwrap();
    assert_eq!(plan.pages().len(), 3);
    assert_eq!(plan.readability().cells_below_threshold, 1);
    assert_eq!(input, original);
    assert_eq!(plan, paginate(&input, &s).unwrap());
}

#[test]
fn axis_sizes_must_not_overflow_or_vanish_in_accumulation() {
    let mut input = sheet(2, 1, f64::MAX, 10.0);
    assert!(matches!(
        paginate(&input, &settings()),
        Err(LayoutError::InvalidAxis { .. })
    ));
    input.rows[1].size_pt = 1.0;
    assert!(matches!(
        paginate(&input, &settings()),
        Err(LayoutError::InvalidAxis { .. })
    ));
}
