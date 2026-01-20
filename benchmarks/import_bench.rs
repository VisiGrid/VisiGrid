//! Simple import benchmark tool
//!
//! Build and run:
//!   rustc -O benchmarks/import_bench.rs -o benchmarks/import_bench
//!   ./benchmarks/import_bench benchmarks/fixtures/small.xlsx
//!
//! Or use the integrated test (preferred):
//!   cargo test -p visigrid-io --release -- import_benchmark --nocapture

use std::env;
use std::path::Path;
use std::time::Instant;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <xlsx_file>", args[0]);
        eprintln!("Example: {} benchmarks/fixtures/small.xlsx", args[0]);
        std::process::exit(1);
    }

    let path = Path::new(&args[1]);
    if !path.exists() {
        eprintln!("File not found: {}", path.display());
        std::process::exit(1);
    }

    println!("Benchmarking: {}", path.display());
    println!("File size: {} bytes", std::fs::metadata(path).map(|m| m.len()).unwrap_or(0));
    println!();

    // Would need to link against visigrid_io here, which isn't practical for a standalone binary.
    // Use the integrated test instead.
    println!("Note: This standalone benchmark is incomplete.");
    println!("Use the integrated test instead:");
    println!("  cargo test -p visigrid-io --release -- import_benchmark --nocapture");
}
