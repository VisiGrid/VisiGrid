//! Rasterize a VisiGrid-generated PDF for preview QA (not a shipped CLI command).
use std::{io::Write, sync::Arc};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<_> = std::env::args().collect();
    if args.len() < 3 || args.len() > 4 {
        return Err("Usage: preview_pdf INPUT.pdf OUTPUT.ppm [PAGE_NUMBER]".into());
    }
    let page = args
        .get(3)
        .map(|s| s.parse::<usize>())
        .transpose()?
        .unwrap_or(1);
    let data = Arc::new(std::fs::read(&args[1])?);
    let result = visigrid_print::preview::rasterize(
        data,
        page.checked_sub(1).ok_or("Page numbers start at 1")?,
        2400,
    )?;
    let mut output = std::fs::File::create(&args[2])?;
    write!(output, "P6\n{} {}\n255\n", result.width, result.height)?;
    for pixel in result.rgba.as_chunks::<4>().0 {
        output.write_all(&pixel[..3])?;
    }
    Ok(())
}
