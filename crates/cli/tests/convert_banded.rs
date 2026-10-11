// `vgrid convert -t json-full --banded -o DIR`: an import written the way
// VisiGrid stores a large sheet (a manifest naming row bands), so the server
// can import a file whose inline document would be tens of MB. Read back as
// the server and browser do (manifest, each band, finish), it must be the
// same workbook the inline json-full conversion gives.
//
// Run with: cargo test -p visigrid-cli --test convert_banded

use std::path::Path;
use std::process::Command;

use sha2::{Digest, Sha256};
use visigrid_io::json::bands;

fn vgrid(args: &[&str]) {
    let out = Command::new(env!("CARGO_BIN_EXE_vgrid")).args(args).output().unwrap();
    assert!(out.status.success(), "vgrid {args:?}: {}", String::from_utf8_lossy(&out.stderr));
}

/// A manifest and its bands, loaded as the collab host loads them.
fn load_banded(dir: &Path) -> visigrid_engine::workbook::Workbook {
    let manifest = std::fs::read_to_string(dir.join("manifest.json")).unwrap();
    let (mut wb, _, _) = visigrid_io::json::import_any(&manifest).unwrap();
    for (_, band) in bands::manifest_bands(&manifest).unwrap() {
        let data = std::fs::read(dir.join("bands").join(format!("{}.band", band.key))).unwrap();
        bands::apply(&mut wb, &data, Some(&band.key)).unwrap();
    }
    bands::finish(&mut wb).unwrap();
    wb
}

#[test]
fn a_large_csv_becomes_bands_that_read_back_as_the_same_workbook() {
    let dir = tempfile::tempdir().unwrap();
    let csv = dir.path().join("big.csv");
    let rows: usize = 70_000;
    let mut text = String::from("Name,Units,Double\n");
    for i in 1..rows {
        text.push_str(&format!("r{i},{i},=B{}*2\n", i + 1));
    }
    std::fs::write(&csv, text).unwrap();

    let out = dir.path().join("banded");
    vgrid(&["convert", csv.to_str().unwrap(), "-t", "json-full", "--banded", "-o", out.to_str().unwrap(), "-q"]);
    let inline = dir.path().join("inline.json");
    vgrid(&["convert", csv.to_str().unwrap(), "-t", "json-full", "-o", inline.to_str().unwrap()]);

    // 70,000 rows × 3 columns is over the band threshold: ceil(70,000 / 16,384) bands.
    let manifest = std::fs::read_to_string(out.join("manifest.json")).unwrap();
    let named = bands::manifest_bands(&manifest).unwrap();
    assert_eq!(named.len(), rows.div_ceil(bands::BAND_ROWS));
    assert!(manifest.len() < 64 * 1024, "the manifest carries no cells: {} bytes", manifest.len());
    for (_, band) in &named {
        let data = std::fs::read(out.join("bands").join(format!("{}.band", band.key))).unwrap();
        let digest: String = Sha256::digest(&data).iter().map(|b| format!("{b:02x}")).collect();
        assert_eq!(digest, band.key, "a band is named by the SHA-256 of its bytes");
        assert!(data.len() < 8 * 1024 * 1024, "each band fits the import sandbox's file limit");
    }

    let banded = load_banded(&out);
    let (expected, _, _) = visigrid_io::json::import_any(&std::fs::read_to_string(&inline).unwrap()).unwrap();
    let (a, b) = (&banded.sheets()[0], &expected.sheets()[0]);
    assert_eq!(a.cells_iter().count(), b.cells_iter().count());
    for row in [0, 1, 16_383, 16_384, 40_000, rows - 1] {
        for col in 0..3 {
            assert_eq!(a.get_raw(row, col), b.get_raw(row, col), "input at {row},{col}");
            assert_eq!(a.get_display(row, col), b.get_display(row, col), "value at {row},{col}");
        }
    }
    assert_eq!(banded.sheets()[0].get_display(rows - 1, 1), format!("{}", rows - 1));
    // CSV keeps a leading "=" as text (no formula injection), banded or not.
    assert_eq!(banded.sheets()[0].get_display(rows - 1, 2), format!("=B{rows}*2"));
}

#[test]
fn a_small_workbook_is_just_its_manifest_the_same_document_as_json_full() {
    let dir = tempfile::tempdir().unwrap();
    let csv = dir.path().join("small.csv");
    std::fs::write(&csv, "a,b\n1,=A2+1\n").unwrap();
    let out = dir.path().join("banded");
    vgrid(&["convert", csv.to_str().unwrap(), "-t", "json-full", "--banded", "-o", out.to_str().unwrap(), "-q"]);

    let manifest = std::fs::read_to_string(out.join("manifest.json")).unwrap();
    assert!(bands::manifest_bands(&manifest).unwrap().is_empty());
    assert_eq!(std::fs::read_dir(out.join("bands")).unwrap().count(), 0);
    let wb = load_banded(&out);
    assert_eq!(wb.sheets()[0].get_display(1, 0), "1");
    assert_eq!(wb.sheets()[0].get_display(1, 1), "=A2+1");
}

#[test]
fn banded_needs_json_full_and_an_output_directory() {
    let dir = tempfile::tempdir().unwrap();
    let csv = dir.path().join("x.csv");
    std::fs::write(&csv, "a\n1\n").unwrap();
    for args in [
        vec!["convert", csv.to_str().unwrap(), "-t", "csv", "--banded", "-o", "out"],
        vec!["convert", csv.to_str().unwrap(), "-t", "json-full", "--banded"],
    ] {
        let out = Command::new(env!("CARGO_BIN_EXE_vgrid")).args(&args).current_dir(dir.path()).output().unwrap();
        assert!(!out.status.success(), "{args:?} should be refused");
    }
}
