//! Manual Linux portal QA; confirm only a test file destination, never a real printer.
#[cfg(target_os = "linux")]
fn main() -> Result<(), String> {
    use visigrid_print::{native, PageSettings, Paper};
    let args: Vec<_> = std::env::args().collect();
    let input = args
        .get(1)
        .ok_or("Usage: system_print INPUT.pdf [a4|letter|legal] [landscape]")?;
    let paper = match args.get(2).map(String::as_str).unwrap_or("a4") {
        "a4" => Paper::A4,
        "letter" => Paper::Letter,
        "legal" => Paper::Legal,
        _ => return Err("Unknown paper size".into()),
    };
    let page = PageSettings {
        paper,
        landscape: args.get(3).is_some_and(|s| s == "landscape"),
        ..Default::default()
    };
    let bytes = std::fs::read(input).map_err(|e| e.to_string())?;
    let outcome = smol::block_on(native::print_pdf(
        &bytes,
        &page,
        "VisiGrid print integration test",
    ))?;
    println!("{outcome:?}");
    Ok(())
}
#[cfg(not(target_os = "linux"))]
fn main() {
    eprintln!("This manual print-portal test requires Linux.");
}
