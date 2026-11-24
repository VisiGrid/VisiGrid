//! End-to-end check of the desktop's Grid cloud-sheet path against a running
//! Grid (Loco). Uses the same client and the same visigrid-json export the
//! desktop uses. Point XDG_CONFIG_HOME at a scratch config holding a
//! `grid-auth.json` from `vgrid grid login`; never at your real one.
//!
//!   grid_e2e create <file.csv>              create a sheet; prints its pid and id
//!   grid_e2e cycle <pid> [legacy ids...]    list, open, save, reload, conflict, legacy
use visigrid_engine::workbook::Workbook;
use visigrid_hub_client::grid::{GridClient, GridError};
use visigrid_io::json;

fn document(wb: &Workbook) -> String {
    json::export_workbook(wb, &[], wb.active_sheet_index()).expect("export visigrid-json")
}

fn open(bytes: &[u8]) -> Workbook {
    json::import_any(std::str::from_utf8(bytes).expect("utf-8")).expect("visigrid-json").0
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let client = GridClient::from_saved().expect("signed in to Grid (grid-auth.json)");
    match args.first().map(String::as_str) {
        Some("create") => {
            let sheet = visigrid_io::csv::import(std::path::Path::new(&args[1])).expect("csv");
            let wb = Workbook::from_sheets(vec![sheet], 0);
            let doc: serde_json::Value = serde_json::from_str(&document(&wb)).unwrap();
            let created = client.create("Desktop E2E", &doc).expect("create");
            println!("{} {} {}", created.pid, created.id, created.revision);
        }
        Some("cycle") => {
            let pid = &args[1];
            // List shows it.
            let listed = client.list().expect("list");
            let entry = listed.iter().find(|s| &s.pid == pid).expect("listed");
            println!("list: {} sheet(s); this one at revision {}", listed.len(), entry.revision);
            // Open: revision, then the document, then the workbook the desktop would write.
            let revision = client.get(pid).expect("get").revision;
            let mut wb = open(&client.data(pid).expect("data"));
            assert_eq!(wb.active_sheet().get_raw(1, 0), "Apple");
            println!("open: revision {revision}, A2 = {}", wb.active_sheet().get_raw(1, 0));
            // Edit and save over the revision we opened.
            wb.active_sheet_mut().set_value(1, 1, "7");
            let saved = client.save(pid, revision, document(&wb).as_bytes()).expect("save");
            assert_eq!(saved.revision, revision + 1);
            println!("save: revision {} -> {}", revision, saved.revision);
            // Reload shows the edit.
            let reloaded = open(&client.data(pid).expect("data"));
            assert_eq!(reloaded.active_sheet().get_raw(1, 1), "7");
            println!("reload: B2 = {}", reloaded.active_sheet().get_raw(1, 1));
            // A save from the stale revision is a conflict, not an overwrite.
            match client.save(pid, revision, document(&wb).as_bytes()) {
                Err(GridError::Conflict) => println!("stale save: conflict (as expected)"),
                other => panic!("stale save should conflict: {other:?}"),
            }
            // Old Rails ids resolve to this sheet.
            for id in &args[2..] {
                let resolved = client.resolve_legacy(id).expect("legacy");
                assert_eq!(&resolved, pid);
                println!("legacy {id} -> {resolved}");
            }
        }
        _ => eprintln!("usage: grid_e2e create <file.csv> | grid_e2e cycle <pid> [legacy ids...]"),
    }
}
