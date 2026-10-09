//! Range reads visit only the rows a range covers (store::columns
//! `slots_between`). These check them against a brute-force read at the
//! places that bounding can get wrong: a 64-row bitmap word boundary, the
//! 1,024-row chunk boundary, and ranges that start or end mid-chunk, over
//! sparse chunks (a few cells) and dense ones (most rows filled).

use visigrid_engine::sheet::{Sheet, SheetId};

fn brute(sheet: &Sheet, r0: usize, c0: usize, r1: usize, c1: usize) -> Vec<f64> {
    let mut out = Vec::new();
    for r in r0..=r1 {
        for c in c0..=c1 {
            if let Ok(n) = sheet.get_display(r, c).parse::<f64>() {
                out.push(n);
            }
        }
    }
    out
}

fn check(sheet: &Sheet, r0: usize, c0: usize, r1: usize, c1: usize) {
    let mut got = Vec::new();
    sheet.numbers_in(r0, c0, r1, c1, &mut got);
    assert_eq!(
        got,
        brute(sheet, r0, c0, r1, c1),
        "range ({r0},{c0})..({r1},{c1})"
    );
}

const EDGES: [(usize, usize); 12] = [
    (0, 0),
    (0, 63),
    (63, 64),
    (64, 64),
    (60, 70),
    (1000, 1023),
    (1023, 1024),
    (1020, 1030),
    (1024, 2047),
    (500, 2100),
    (2047, 2048),
    (0, 3000),
];

#[test]
fn bounded_reads_match_brute_force_in_dense_chunks() {
    let mut sheet = Sheet::new(SheetId(1), 4000, 4);
    for r in 0..3000 {
        sheet.set_value(r, 0, &(r as f64 + 0.5).to_string());
        if r % 3 != 0 {
            sheet.set_value(r, 1, &r.to_string());
        }
    }
    for (r0, r1) in EDGES {
        check(&sheet, r0, 0, r1, 0);
        check(&sheet, r0, 0, r1, 1);
        check(&sheet, r0, 1, r1, 1);
    }
}

#[test]
fn bounded_reads_match_brute_force_in_sparse_chunks() {
    let mut sheet = Sheet::new(SheetId(1), 4000, 4);
    for r in [
        0, 1, 62, 63, 64, 65, 127, 1022, 1023, 1024, 1025, 2047, 2048, 2999,
    ] {
        sheet.set_value(r, 2, &(r * 10).to_string());
    }
    for (r0, r1) in EDGES {
        check(&sheet, r0, 2, r1, 2);
        check(&sheet, r0, 0, r1, 3);
    }
}
