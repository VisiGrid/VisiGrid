//! Two-page rendering spike, deliberately not a production worksheet exporter.
//! Run with an explicit output path; fonts are the app's existing OFL assets.
use krilla::color::rgb;
use krilla::geom::{PathBuilder, Point};
use krilla::page::PageSettings as PdfPageSettings;
use krilla::paint::{Fill, FillRule, Stroke};
use krilla::text::{Font, TextDirection};
use krilla::Document;
use visigrid_print::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args_os()
        .nth(1)
        .ok_or("Usage: pdf_spike OUTPUT.pdf")?;
    // create_new protects an existing output during repeated prototype runs.
    let regular = Font::new(
        include_bytes!("../../../gpui-app/assets/fonts/ibm-plex-sans/IBMPlexSans-Regular.ttf")
            .to_vec()
            .into(),
        0,
    )
    .ok_or("Invalid regular font")?;
    let bold = Font::new(
        include_bytes!("../../../gpui-app/assets/fonts/ibm-plex-sans/IBMPlexSans-SemiBold.ttf")
            .to_vec()
            .into(),
        0,
    )
    .ok_or("Invalid bold font")?;
    let input = LayoutInput {
        rows: (0..60)
            .map(|i| AxisItem {
                source_index: i,
                size_pt: if i < 2 { 26.0 } else { 20.0 },
            })
            .collect(),
        columns: [220.0, 140.0, 140.0]
            .into_iter()
            .enumerate()
            .map(|(i, size_pt)| AxisItem {
                source_index: i,
                size_pt,
            })
            .collect(),
        merges: vec![Merge {
            rows: 0..1,
            columns: 0..3,
        }],
        text: (0..60)
            .flat_map(|row| {
                (0..3)
                    .filter(move |&column| row != 0 || column == 0)
                    .map(move |column| TextCell {
                        row,
                        column,
                        font_size_pt: if row == 0 { 14.0 } else { 10.0 },
                    })
            })
            .collect(),
    };
    let settings = PageSettings {
        paper: Paper::Letter,
        repeat_rows: 2,
        footer: true,
        ..Default::default()
    };
    let plan = paginate(&input, &settings)?;
    assert_eq!(plan.pages().len(), 2);
    let mut document = Document::new();
    for page_index in 0..plan.pages().len() {
        let (width, height) = plan.page_size();
        let mut page = document.start_page_with(
            PdfPageSettings::from_wh(width as f32, height as f32).ok_or("Invalid page size")?,
        );
        let mut surface = page.surface();
        for row in 0..input.rows.len() {
            for column in 0..3 {
                if row == 0 && column != 0 {
                    continue;
                }
                let cols = if row == 0 { 0..3 } else { column..column + 1 };
                let Some(rect) = plan.rect(page_index, row..row + 1, cols) else {
                    continue;
                };
                let mut path = PathBuilder::new();
                path.move_to(rect.x as f32, rect.y as f32);
                path.line_to((rect.x + rect.width) as f32, rect.y as f32);
                path.line_to((rect.x + rect.width) as f32, (rect.y + rect.height) as f32);
                path.line_to(rect.x as f32, (rect.y + rect.height) as f32);
                path.close();
                let path = path.finish().ok_or("Invalid cell rectangle")?;
                let color = if row < 2 {
                    rgb::Color::new(225, 234, 246)
                } else if row % 2 == 0 {
                    rgb::Color::new(246, 248, 251)
                } else {
                    rgb::Color::new(255, 255, 255)
                };
                surface.set_fill(Some(Fill {
                    paint: color.into(),
                    ..Default::default()
                }));
                surface.set_stroke(Some(Stroke {
                    paint: rgb::Color::new(180, 192, 208).into(),
                    width: 0.5,
                    ..Default::default()
                }));
                surface.draw_path(&path);
                surface.set_stroke(None);
                surface.set_fill(Some(Fill {
                    paint: rgb::Color::new(25, 38, 56).into(),
                    ..Default::default()
                }));
                let text = if row == 0 {
                    "VisiGrid print prototype — Café / €".to_string()
                } else if row == 1 {
                    ["Description", "Reference", "Amount"][column].to_string()
                } else {
                    match column {
                        0 => format!("Report item {:02}", row - 1),
                        1 => format!("INV-{:04}", row + 1000),
                        _ => format!("€ {:.2}", row as f64 * 12.75),
                    }
                };
                surface.push_clip_path(&path, &FillRule::NonZero);
                let size = if row == 0 { 14.0 } else { 10.0 } * plan.scale() as f32;
                surface.draw_text(
                    Point::from_xy(
                        (rect.x + 6.0) as f32,
                        (rect.y + rect.height / 2.0 + f64::from(size) * 0.35) as f32,
                    ),
                    if row < 2 {
                        bold.clone()
                    } else {
                        regular.clone()
                    },
                    size,
                    &text,
                    false,
                    TextDirection::Auto,
                );
                surface.pop();
            }
        }
        surface.draw_text(
            Point::from_xy(
                plan.body().x as f32,
                (plan.body().y + plan.body().height + 12.0) as f32,
            ),
            regular.clone(),
            9.0,
            &format!(
                "Prototype fixture • Page {} of {}",
                page_index + 1,
                plan.pages().len()
            ),
            false,
            TextDirection::Auto,
        );
        surface.finish();
        page.finish();
    }
    let bytes = document
        .finish()
        .map_err(|e| format!("PDF rendering failed: {e:?}"))?;
    let mut file = std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(output)?;
    use std::io::Write;
    file.write_all(&bytes)?;
    println!(
        "Rendered {} pages at {:.0}%; minimum text {} pt",
        plan.pages().len(),
        plan.scale() * 100.0,
        plan.readability().smallest_text_pt.unwrap()
    );
    Ok(())
}
