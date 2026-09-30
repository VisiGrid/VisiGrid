//! Searchable vector PDF output with externally shaped Unicode text.
use crate::snapshot::{PrintCell, Snapshot};
use crate::{paginate, PageSettings, Rect};
use cosmic_text::{Attrs, Buffer, Family, FontSystem, Metrics, Shaping, Style, Weight, Wrap};
use krilla::color::rgb;
use krilla::geom::{PathBuilder, Point};
use krilla::paint::{Fill, FillRule, Stroke};
use krilla::surface::Surface;
use krilla::text::{Font, GlyphId, KrillaGlyph};
use krilla::Document;
use std::collections::{HashMap, HashSet};
use std::io::Write;
use std::path::Path;
use visigrid_engine::cell::{Alignment, CellBorder, CellFormat, TextOverflow, VerticalAlignment};

#[derive(Debug)]
pub struct PdfOutput {
    pub bytes: Vec<u8>,
    pub pages: usize,
    pub scale: f64,
    pub clipped_cells: usize,
    /// First ten source addresses, in sheet order, for actionable diagnostics.
    pub clipped_addresses: Vec<String>,
    pub small_text_cells: usize,
    pub substituted_fonts: Vec<String>,
}

impl AsRef<[u8]> for PdfOutput {
    fn as_ref(&self) -> &[u8] {
        &self.bytes
    }
}

/// Render entirely before opening the destination. An error leaves it intact.
pub fn render(snapshot: &Snapshot, settings: &PageSettings) -> Result<PdfOutput, String> {
    render_cancellable(snapshot, settings, || false)
}

pub fn render_cancellable(
    snapshot: &Snapshot,
    settings: &PageSettings,
    cancelled: impl Fn() -> bool,
) -> Result<PdfOutput, String> {
    if cancelled() {
        return Err("Export cancelled".into());
    }
    let plan = paginate(&snapshot.layout, settings).map_err(|e| e.to_string())?;
    let mut fonts = FontSystem::new();
    for data in [
        include_bytes!("../../../gpui-app/assets/fonts/ibm-plex-sans/IBMPlexSans-Regular.ttf")
            .as_slice(),
        include_bytes!("../../../gpui-app/assets/fonts/ibm-plex-sans/IBMPlexSans-SemiBold.ttf")
            .as_slice(),
        include_bytes!("../../../gpui-app/assets/fonts/ibm-plex-sans/IBMPlexSans-Italic.ttf")
            .as_slice(),
        include_bytes!(
            "../../../gpui-app/assets/fonts/ibm-plex-sans/IBMPlexSans-SemiBoldItalic.ttf"
        )
        .as_slice(),
    ] {
        fonts.db_mut().load_font_data(data.to_vec());
    }
    fonts.db_mut().set_sans_serif_family("IBM Plex Sans");
    let families: HashSet<String> = fonts
        .db()
        .faces()
        .flat_map(|f| f.families.iter().map(|(name, _)| name.to_lowercase()))
        .collect();
    let mut font_cache = HashMap::new();
    let mut substituted = HashSet::new();
    let mut clipped = HashSet::new();
    let cells: HashMap<_, _> = snapshot
        .cells
        .iter()
        .map(|cell| ((cell.rows.start, cell.columns.start), cell))
        .collect();
    let mut document = Document::new();
    for (page_index, band) in plan.pages().iter().enumerate() {
        if cancelled() {
            return Err("Export cancelled".into());
        }
        let (width, height) = plan.page_size();
        let mut page = document.start_page_with(
            krilla::page::PageSettings::from_wh(width as f32, height as f32)
                .ok_or("Invalid paper size")?,
        );
        let mut surface = page.surface();
        let mut on_page = Vec::new();
        for r in (0..settings.repeat_rows).chain(band.rows.clone()) {
            for c in (0..settings.repeat_columns).chain(band.columns.clone()) {
                if let Some(cell) = cells.get(&(r, c)) {
                    if let Some(rect) =
                        plan.rect(page_index, cell.rows.clone(), cell.columns.clone())
                    {
                        on_page.push((*cell, rect));
                    }
                }
            }
        }
        // Gridlines are page decoration, never worksheet borders. Draw them
        // under fills and explicit borders; merged interiors have no gridlines.
        if settings.gridlines {
            surface.set_fill(None);
            surface.set_stroke(Some(Stroke {
                paint: rgb::Color::new(180, 180, 180).into(),
                width: (0.4 * plan.scale() as f32).max(0.25),
                ..Default::default()
            }));
            for (cell, rect) in &on_page {
                if cancelled() {
                    return Err("Export cancelled".into());
                }
                if cell.format.background_color.is_none() {
                    surface.draw_path(&rectangle(*rect)?);
                }
            }
            surface.set_stroke(None);
        }
        // Paint all fills before text. Text overflow can then cross blank cells.
        for (cell, rect) in &on_page {
            if cancelled() {
                return Err("Export cancelled".into());
            }
            if let Some(color) = cell.format.background_color {
                fill(&mut surface, color);
                surface.set_stroke(None);
                surface.draw_path(&rectangle(*rect)?);
            }
        }
        for (cell, rect) in &on_page {
            if cancelled() {
                return Err("Export cancelled".into());
            }
            let f = &cell.format;
            for (border, start, end) in [
                (
                    f.border_top,
                    (rect.x, rect.y),
                    (rect.x + rect.width, rect.y),
                ),
                (
                    f.border_bottom,
                    (rect.x, rect.y + rect.height),
                    (rect.x + rect.width, rect.y + rect.height),
                ),
                (
                    f.border_left,
                    (rect.x, rect.y),
                    (rect.x, rect.y + rect.height),
                ),
                (
                    f.border_right,
                    (rect.x + rect.width, rect.y),
                    (rect.x + rect.width, rect.y + rect.height),
                ),
            ] {
                border_line(&mut surface, border, start, end, plan.scale());
            }
            if cell.text.is_empty() {
                continue;
            }
            let mut text_rect = *rect;
            if f.text_overflow == TextOverflow::Overflow
                && f.alignment == Alignment::Left
                && cell.columns.len() == 1
                && cell.rows.len() == 1
            {
                for col in cell.columns.end..band.columns.end {
                    let Some(next) = cells.get(&(cell.rows.start, col)) else {
                        break;
                    };
                    if !next.text.is_empty()
                        || next.columns.len() != 1
                        || next.rows.len() != 1
                        || next.format.background_color.is_some()
                        || next.format.has_any_border()
                    {
                        break;
                    }
                    let Some(next_rect) =
                        plan.rect(page_index, next.rows.clone(), next.columns.clone())
                    else {
                        break;
                    };
                    text_rect.width = next_rect.x + next_rect.width - text_rect.x;
                }
            }
            let requested = f.font_family.as_deref().unwrap_or(&snapshot.default_font);
            let family = if families.contains(&requested.to_lowercase()) {
                requested
            } else {
                substituted.insert(requested.to_string());
                "IBM Plex Sans"
            };
            if draw_cell(
                &mut surface,
                cell,
                text_rect,
                plan.scale() as f32,
                snapshot.default_size,
                family,
                &mut fonts,
                &mut font_cache,
            )? {
                clipped.insert(cell.source);
            }
        }
        if settings.footer {
            let footer = PrintCell {
                rows: 0..1,
                columns: 0..1,
                source: (0, 0),
                text: format!(
                    "{}  |  Page {} of {}",
                    snapshot.name,
                    page_index + 1,
                    plan.pages().len()
                ),
                format: CellFormat {
                    font_size: Some(9.0),
                    ..Default::default()
                },
            };
            draw_cell(
                &mut surface,
                &footer,
                Rect {
                    x: plan.body().x,
                    y: plan.body().y + plan.body().height,
                    width: plan.body().width,
                    height: 18.0,
                },
                1.0,
                9.0,
                "IBM Plex Sans",
                &mut fonts,
                &mut font_cache,
            )?;
        }
        surface.finish();
        page.finish();
    }
    let bytes = document
        .finish()
        .map_err(|e| format!("PDF rendering failed: {e:?}"))?;
    let mut substituted_fonts: Vec<_> = substituted.into_iter().collect();
    substituted_fonts.sort();
    let mut clipped_positions: Vec<_> = clipped.iter().copied().collect();
    clipped_positions.sort_unstable();
    let clipped_addresses = clipped_positions
        .into_iter()
        .take(10)
        .map(|(r, c)| address(r, c))
        .collect();
    Ok(PdfOutput {
        bytes,
        pages: plan.pages().len(),
        scale: plan.scale(),
        clipped_cells: clipped.len(),
        clipped_addresses,
        small_text_cells: plan.readability().cells_below_threshold,
        substituted_fonts,
    })
}

fn address(row: usize, col: usize) -> String {
    let mut col = col + 1;
    let mut label = String::new();
    while col > 0 {
        col -= 1;
        label.insert(0, (b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    format!("{label}{}", row + 1)
}

pub fn save_atomic(path: &Path, bytes: &[u8]) -> Result<(), String> {
    let parent = path
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let mut temp = tempfile::NamedTempFile::new_in(parent).map_err(|e| e.to_string())?;
    temp.write_all(bytes)
        .and_then(|_| temp.as_file().sync_all())
        .map_err(|e| e.to_string())?;
    temp.persist(path).map_err(|e| e.error.to_string())?;
    Ok(())
}

fn rectangle(r: Rect) -> Result<krilla::geom::Path, String> {
    let mut p = PathBuilder::new();
    p.move_to(r.x as f32, r.y as f32);
    p.line_to((r.x + r.width) as f32, r.y as f32);
    p.line_to((r.x + r.width) as f32, (r.y + r.height) as f32);
    p.line_to(r.x as f32, (r.y + r.height) as f32);
    p.close();
    p.finish().ok_or("Invalid cell rectangle".into())
}

fn fill(surface: &mut Surface<'_>, color: [u8; 4]) {
    // Composite transparent cell paint onto white paper, never the UI theme.
    let a = f32::from(color[3]) / 255.0;
    let rgb = [color[0], color[1], color[2]]
        .map(|c| (f32::from(c) * a + 255.0 * (1.0 - a)).round() as u8);
    surface.set_fill(Some(Fill {
        paint: rgb::Color::new(rgb[0], rgb[1], rgb[2]).into(),
        ..Default::default()
    }));
    surface.set_stroke(None);
}

fn border_line(
    surface: &mut Surface<'_>,
    border: CellBorder,
    start: (f64, f64),
    end: (f64, f64),
    scale: f64,
) {
    if !border.is_set() {
        return;
    }
    let color = border.color.unwrap_or([0, 0, 0, 255]);
    surface.set_fill(None);
    surface.set_stroke(Some(Stroke {
        paint: rgb::Color::new(color[0], color[1], color[2]).into(),
        width: (f64::from(border.style.weight()) * 0.75 * scale) as f32,
        ..Default::default()
    }));
    let mut path = PathBuilder::new();
    path.move_to(start.0 as f32, start.1 as f32);
    path.line_to(end.0 as f32, end.1 as f32);
    if let Some(path) = path.finish() {
        surface.draw_path(&path);
    }
    surface.set_stroke(None);
}

#[allow(clippy::too_many_arguments)]
fn draw_cell(
    surface: &mut Surface<'_>,
    cell: &PrintCell,
    rect: Rect,
    scale: f32,
    default_size: f32,
    family: &str,
    fonts: &mut FontSystem,
    cache: &mut HashMap<cosmic_text::fontdb::ID, Font>,
) -> Result<bool, String> {
    let f = &cell.format;
    let size = f.font_size.unwrap_or(default_size) * scale;
    if !size.is_finite() || size <= 0.0 {
        return Err("Invalid cell font size".into());
    }
    let padding = 3.0 * scale;
    let width = (rect.width as f32 - padding * 2.0).max(0.01);
    let height = (rect.height as f32 - 2.0 * scale).max(0.01);
    let mut buffer = Buffer::new(fonts, Metrics::new(size, size * 1.2));
    buffer.set_size(Some(width), None);
    buffer.set_wrap(if f.text_overflow == TextOverflow::Wrap {
        Wrap::WordOrGlyph
    } else {
        Wrap::None
    });
    let attrs = Attrs::new()
        .family(Family::Name(family))
        .weight(if !f.bold {
            Weight::NORMAL
        } else if family.eq_ignore_ascii_case("IBM Plex Sans") {
            // Match the app's bundled bold face (600); requesting 700 can
            // select an unrelated system fallback despite this family existing.
            Weight::SEMIBOLD
        } else {
            Weight::BOLD
        })
        .style(if f.italic {
            Style::Italic
        } else {
            Style::Normal
        });
    buffer.set_text(
        &cell.text,
        &attrs,
        Shaping::Advanced,
        Some(cosmic_text::Align::Left),
    );
    buffer.shape_until_scroll(fonts, false);
    let text_height = buffer
        .layout_runs()
        .map(|l| l.line_top + l.line_height)
        .fold(0.0, f32::max);
    let y_offset = match f.vertical_alignment {
        VerticalAlignment::Top => 0.0,
        VerticalAlignment::Middle => ((height - text_height) / 2.0).max(0.0),
        VerticalAlignment::Bottom => (height - text_height).max(0.0),
    };
    let mut clipped = text_height > height + 0.1;
    fill(
        surface,
        f.font_color.unwrap_or_else(|| {
            let bg = f.background_color.unwrap_or([255, 255, 255, 255]);
            if bg[3] > 128
                && (u32::from(bg[0]) * 299 + u32::from(bg[1]) * 587 + u32::from(bg[2]) * 114)
                    < 128000
            {
                [255, 255, 255, 255]
            } else {
                [0, 0, 0, 255]
            }
        }),
    );
    surface.push_clip_path(&rectangle(rect)?, &FillRule::NonZero);
    for run in buffer.layout_runs() {
        let x_offset = match f.alignment {
            Alignment::Right => width - run.line_w,
            Alignment::Center | Alignment::CenterAcrossSelection => (width - run.line_w) / 2.0,
            _ => 0.0,
        };
        // Padding guides alignment/wrapping, but the actual clip is the outer
        // cell rectangle. Text extending into padding is still visible.
        let left = padding + x_offset;
        clipped |= left < -0.1 || left + run.line_w > rect.width as f32 + 0.1;
        let baseline = rect.y as f32 + scale + y_offset + run.line_y;
        let mut start = 0;
        while start < run.glyphs.len() {
            let font_id = run.glyphs[start].font_id;
            let mut end = start + 1;
            while end < run.glyphs.len() && run.glyphs[end].font_id == font_id {
                end += 1;
            }
            if let std::collections::hash_map::Entry::Vacant(entry) = cache.entry(font_id) {
                let font = fonts
                    .db()
                    .with_face_data(font_id, |data, index| {
                        Font::new(data.to_vec().into(), index)
                    })
                    .flatten()
                    .ok_or("Unable to embed a required font")?;
                entry.insert(font);
            }
            let mut advance = 0.0;
            let mut glyphs = Vec::new();
            for g in &run.glyphs[start..end] {
                if g.glyph_id == 0 {
                    return Err(format!("No installed font can render text at row {}, column {}. Install a font covering that script and try again.", cell.source.0+1, cell.source.1+1));
                }
                glyphs.push(KrillaGlyph::new(
                    GlyphId::new(u32::from(g.glyph_id)),
                    g.w / size,
                    (g.x - advance) / size + g.x_offset,
                    g.y_offset - g.y / size,
                    0.0,
                    g.start..g.end,
                    None,
                ));
                advance += g.w;
            }
            surface.draw_glyphs(
                Point::from_xy(rect.x as f32 + padding + x_offset, baseline),
                &glyphs,
                cache[&font_id].clone(),
                run.text,
                size,
                false,
            );
            start = end;
        }
        if f.underline || f.strikethrough {
            // Decoration width follows the same shaped line metrics.
            let x = rect.x as f32 + padding + x_offset;
            for y in [
                f.underline.then_some(baseline + size * 0.12),
                f.strikethrough.then_some(baseline - size * 0.3),
            ]
            .into_iter()
            .flatten()
            {
                let r = Rect {
                    x: f64::from(x),
                    y: f64::from(y),
                    width: f64::from(run.line_w),
                    height: f64::from((size / 16.0).max(0.3)),
                };
                if r.width > 0.0 {
                    surface.draw_path(&rectangle(r)?);
                }
            }
        }
    }
    surface.pop();
    Ok(clipped)
}
