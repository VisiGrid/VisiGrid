//! XLSX style parser: extracts formatting from styles.xml and per-cell style IDs
//! from worksheet XML within XLSX (ZIP) archives.
//!
//! Phase 1B of the XLSX Formatting Import plan.

use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::path::Path;
use visigrid_engine::cell::{
    Alignment, BorderStyle, CellBorder, CellFormat, NumberFormat, TextOverflow,
    VerticalAlignment,
};
use zip::ZipArchive;

// =============================================================================
// Public types
// =============================================================================

/// Parsed style table from styles.xml — maps cellXfs index → CellFormat.
pub struct StyleTable {
    pub styles: Vec<CellFormat>,
}

impl StyleTable {
    pub fn get(&self, id: usize) -> Option<&CellFormat> {
        self.styles.get(id)
    }

    pub fn len(&self) -> usize {
        self.styles.len()
    }
}

/// Per-cell style references extracted from a worksheet XML.
pub struct SheetFormatting {
    /// (row, col, style_id) triples
    pub cell_styles: Vec<(usize, usize, usize)>,
    /// Column widths in raw Excel character-width units
    pub col_widths: HashMap<usize, f64>,
    /// Row heights in raw Excel point units
    pub row_heights: HashMap<usize, f64>,
    /// Merged cell regions: (start_row, start_col, end_row, end_col)
    pub merged_regions: Vec<(usize, usize, usize, usize)>,
}

/// Stats about style parsing for the import report.
#[derive(Debug, Default)]
pub struct StyleImportStats {
    pub styles_imported: usize,
    pub unique_styles: usize,
    pub unsupported_features: Vec<String>,
}

// =============================================================================
// XML entity unescaping
// =============================================================================

/// Unescape the 5 predefined XML entities: &amp; &lt; &gt; &quot; &apos;
fn unescape_xml(s: &str) -> String {
    if !s.contains('&') {
        return s.to_string();
    }
    s.replace("&quot;", "\"")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&apos;", "'")
}

// Built-in number format mapping
// =============================================================================

fn builtin_number_format(id: u16) -> NumberFormat {
    match id {
        0 => NumberFormat::General,
        1 => NumberFormat::Number { decimals: 0 },
        2 => NumberFormat::Number { decimals: 2 },
        3 => NumberFormat::Number { decimals: 0 }, // #,##0
        4 => NumberFormat::Number { decimals: 2 }, // #,##0.00
        9 => NumberFormat::Percent { decimals: 0 },
        10 => NumberFormat::Percent { decimals: 2 },
        11 => NumberFormat::Number { decimals: 2 }, // 0.00E+00 (scientific)
        14 => NumberFormat::Date {
            style: visigrid_engine::cell::DateStyle::Short,
        },
        15 => NumberFormat::Date {
            style: visigrid_engine::cell::DateStyle::Long,
        },
        16 => NumberFormat::Date {
            style: visigrid_engine::cell::DateStyle::Long,
        },
        17 => NumberFormat::Date {
            style: visigrid_engine::cell::DateStyle::Short,
        },
        18 => NumberFormat::Time,
        19 => NumberFormat::Time,
        20 => NumberFormat::Time,
        21 => NumberFormat::Time,
        22 => NumberFormat::DateTime,
        37 => NumberFormat::Number { decimals: 0 }, // #,##0;(#,##0)
        38 => NumberFormat::Number { decimals: 0 }, // #,##0;[Red](#,##0)
        39 => NumberFormat::Number { decimals: 2 }, // #,##0.00;(#,##0.00)
        40 => NumberFormat::Number { decimals: 2 }, // #,##0.00;[Red](#,##0.00)
        44 => NumberFormat::Currency { decimals: 2 },
        45 => NumberFormat::Time,
        46 => NumberFormat::Time,
        47 => NumberFormat::Time,
        48 => NumberFormat::Number { decimals: 2 }, // ##0.0E+0
        49 => NumberFormat::General,                // @ (text)
        _ => NumberFormat::General,
    }
}

// =============================================================================
// Indexed color palette (standard 64 Excel colors)
// =============================================================================

/// Standard Excel indexed color palette (RGBA).
/// Index 0-7 are the primary colors, 8-63 are extended.
fn indexed_color(idx: u8) -> Option<[u8; 4]> {
    let rgb: [u8; 3] = match idx {
        0 => [0, 0, 0],       // Black
        1 => [255, 255, 255], // White
        2 => [255, 0, 0],     // Red
        3 => [0, 255, 0],     // Green
        4 => [0, 0, 255],     // Blue
        5 => [255, 255, 0],   // Yellow
        6 => [255, 0, 255],   // Magenta
        7 => [0, 255, 255],   // Cyan
        8 => [0, 0, 0],       // Black (duplicate)
        9 => [255, 255, 255], // White (duplicate)
        10 => [255, 0, 0],
        11 => [0, 255, 0],
        12 => [0, 0, 255],
        13 => [255, 255, 0],
        14 => [255, 0, 255],
        15 => [0, 255, 255],
        16 => [128, 0, 0],    // Dark Red
        17 => [0, 128, 0],    // Dark Green
        18 => [0, 0, 128],    // Dark Blue
        19 => [128, 128, 0],  // Olive
        20 => [128, 0, 128],  // Purple
        21 => [0, 128, 128],  // Teal
        22 => [192, 192, 192], // Silver
        23 => [128, 128, 128], // Gray
        24 => [153, 153, 255],
        25 => [153, 51, 102],
        26 => [255, 255, 204],
        27 => [204, 255, 255],
        28 => [102, 0, 102],
        29 => [255, 128, 128],
        30 => [0, 102, 204],
        31 => [204, 204, 255],
        32 => [0, 0, 128],
        33 => [255, 0, 255],
        34 => [255, 255, 0],
        35 => [0, 255, 255],
        36 => [128, 0, 128],
        37 => [128, 0, 0],
        38 => [0, 128, 128],
        39 => [0, 0, 255],
        40 => [0, 204, 255],
        41 => [204, 255, 255],
        42 => [204, 255, 204],
        43 => [255, 255, 153],
        44 => [153, 204, 255],
        45 => [255, 153, 204],
        46 => [204, 153, 255],
        47 => [255, 204, 153],
        48 => [51, 102, 255],
        49 => [51, 204, 204],
        50 => [153, 204, 0],
        51 => [255, 204, 0],
        52 => [255, 153, 0],
        53 => [255, 102, 0],
        54 => [102, 102, 153],
        55 => [150, 150, 150],
        56 => [0, 51, 102],
        57 => [51, 153, 102],
        58 => [0, 51, 0],
        59 => [51, 51, 0],
        60 => [153, 51, 0],
        61 => [153, 51, 51],
        62 => [51, 51, 153],
        63 => [51, 51, 51],
        64 => return Some([0, 0, 0, 255]),       // System foreground (black)
        65 => return Some([255, 255, 255, 255]),  // System background (white)
        _ => return None,
    };
    Some([rgb[0], rgb[1], rgb[2], 255])
}

/// Flat theme color defaults (approximate, no tint math).
/// theme="0" through theme="9" map to Excel's default theme.
fn theme_color_default(idx: u8) -> Option<[u8; 4]> {
    let rgb: [u8; 3] = match idx {
        0 => [255, 255, 255], // Background 1 (lt1)
        1 => [0, 0, 0],       // Text 1 (dk1)
        2 => [238, 236, 225],  // Background 2 (lt2)
        3 => [31, 73, 125],    // Text 2 (dk2)
        4 => [79, 129, 189],   // Accent 1
        5 => [192, 80, 77],    // Accent 2
        6 => [155, 187, 89],   // Accent 3
        7 => [128, 100, 162],  // Accent 4
        8 => [75, 172, 198],   // Accent 5
        9 => [247, 150, 70],   // Accent 6
        _ => return None,
    };
    Some([rgb[0], rgb[1], rgb[2], 255])
}

// =============================================================================
// Color parsing
// =============================================================================

/// Parse a color from XML attributes (rgb, indexed, or theme).
/// Returns RGBA as [u8; 4], or None if no color found.
fn parse_color_attrs(
    attrs: &[(Vec<u8>, Vec<u8>)],
    unsupported: &mut Vec<String>,
) -> Option<[u8; 4]> {
    let mut rgb_val: Option<Vec<u8>> = None;
    let mut indexed_val: Option<u8> = None;
    let mut theme_val: Option<u8> = None;

    for (key, value) in attrs {
        match key.as_slice() {
            b"rgb" => rgb_val = Some(value.clone()),
            b"indexed" => {
                indexed_val = std::str::from_utf8(value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            b"theme" => {
                theme_val = std::str::from_utf8(value)
                    .ok()
                    .and_then(|s| s.parse().ok());
            }
            _ => {}
        }
    }

    // Prefer rgb > indexed > theme
    if let Some(hex) = rgb_val {
        return parse_argb_hex(&hex);
    }
    if let Some(idx) = indexed_val {
        return indexed_color(idx);
    }
    if let Some(idx) = theme_val {
        let color = theme_color_default(idx);
        if color.is_some() {
            // Log as approximate
            if !unsupported.iter().any(|s| s.starts_with("theme tints")) {
                unsupported.push("theme tints approximated".to_string());
            }
        }
        return color;
    }
    None
}

/// Parse AARRGGBB hex string to RGBA [u8; 4].
fn parse_argb_hex(hex: &[u8]) -> Option<[u8; 4]> {
    let s = std::str::from_utf8(hex).ok()?;
    let s = s.trim_start_matches('#');

    if s.len() == 8 {
        // AARRGGBB → RGBA
        let a = u8::from_str_radix(&s[0..2], 16).ok()?;
        let r = u8::from_str_radix(&s[2..4], 16).ok()?;
        let g = u8::from_str_radix(&s[4..6], 16).ok()?;
        let b = u8::from_str_radix(&s[6..8], 16).ok()?;
        Some([r, g, b, a])
    } else if s.len() == 6 {
        // RRGGBB → RGBA (alpha=255)
        let r = u8::from_str_radix(&s[0..2], 16).ok()?;
        let g = u8::from_str_radix(&s[2..4], 16).ok()?;
        let b = u8::from_str_radix(&s[4..6], 16).ok()?;
        Some([r, g, b, 255])
    } else {
        None
    }
}

// =============================================================================
// Internal parsed components
// =============================================================================

/// Parsed font entry from <fonts>.
#[derive(Debug, Clone, Default)]
struct ParsedFont {
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    size: Option<f32>,
    color: Option<[u8; 4]>,
    family: Option<String>,
}

/// Parsed fill entry from <fills>.
#[derive(Debug, Clone, Default)]
struct ParsedFill {
    bg_color: Option<[u8; 4]>,
}

/// Parsed border entry from <borders>.
#[derive(Debug, Clone, Default)]
struct ParsedBorder {
    top: CellBorder,
    right: CellBorder,
    bottom: CellBorder,
    left: CellBorder,
}

// =============================================================================
// styles.xml parser
// =============================================================================

/// Parse styles.xml content into a StyleTable.
pub fn parse_styles_xml(xml: &str) -> (StyleTable, Vec<String>) {
    let mut unsupported: Vec<String> = Vec::new();

    // Step 1: Parse sub-sections
    let custom_num_fmts = parse_num_fmts(xml);
    let fonts = parse_fonts(xml, &mut unsupported);
    let fills = parse_fills(xml, &mut unsupported);
    let borders = parse_borders(xml, &mut unsupported);

    // Step 2: Parse cellXfs and resolve each <xf> into a CellFormat
    let styles = parse_cell_xfs(xml, &custom_num_fmts, &fonts, &fills, &borders, &mut unsupported);

    (StyleTable { styles }, unsupported)
}

/// Parse <numFmts> section → HashMap<formatId, formatCode>
fn parse_num_fmts(xml: &str) -> HashMap<u16, String> {
    let mut map = HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_num_fmts = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.name().as_ref() == b"numFmts" => {
                in_num_fmts = true;
            }
            Ok(Event::End(ref e)) if e.name().as_ref() == b"numFmts" => {
                break;
            }
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if in_num_fmts && e.name().as_ref() == b"numFmt" =>
            {
                let mut id: Option<u16> = None;
                let mut code: Option<String> = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"numFmtId" => {
                            id = std::str::from_utf8(&attr.value)
                                .ok()
                                .and_then(|s| s.parse().ok());
                        }
                        b"formatCode" => {
                            // Must unescape XML entities: &quot; → " (e.g. "$"#,##0)
                            let raw = String::from_utf8_lossy(&attr.value).to_string();
                            code = Some(unescape_xml(&raw));
                        }
                        _ => {}
                    }
                }
                if let (Some(id), Some(code)) = (id, code) {
                    map.insert(id, code);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    map
}

/// Parse <fonts> section into Vec<ParsedFont>.
fn parse_fonts(xml: &str, unsupported: &mut Vec<String>) -> Vec<ParsedFont> {
    let mut fonts = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut depth = 0; // 0 = outside, 1 = inside <fonts>, 2 = inside <font>

    let mut current_font = ParsedFont::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                match name.as_ref() {
                    b"fonts" if depth == 0 => depth = 1,
                    b"font" if depth == 1 => {
                        depth = 2;
                        current_font = ParsedFont::default();
                    }
                    b"color" if depth == 2 => {
                        let attrs = collect_attrs(e);
                        current_font.color = parse_color_attrs(&attrs, unsupported);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if depth == 2 => {
                let name = e.name();
                match name.as_ref() {
                    b"b" => current_font.bold = true,
                    b"i" => current_font.italic = true,
                    b"u" => current_font.underline = true,
                    b"strike" => current_font.strikethrough = true,
                    b"sz" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                current_font.size = std::str::from_utf8(&attr.value)
                                    .ok()
                                    .and_then(|s| s.parse().ok());
                            }
                        }
                    }
                    b"color" => {
                        let attrs = collect_attrs(e);
                        current_font.color = parse_color_attrs(&attrs, unsupported);
                    }
                    b"name" | b"rFont" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"val" {
                                current_font.family =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                match e.name().as_ref() {
                    b"font" if depth == 2 => {
                        fonts.push(current_font.clone());
                        depth = 1;
                    }
                    b"fonts" if depth == 1 => break,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    fonts
}

/// Parse <fills> section into Vec<ParsedFill>.
fn parse_fills(xml: &str, unsupported: &mut Vec<String>) -> Vec<ParsedFill> {
    let mut fills = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut depth = 0; // 0 = outside, 1 = inside <fills>, 2 = inside <fill>
    let mut in_pattern_fill = false;
    let mut current_fill = ParsedFill::default();
    let mut _is_gradient = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                match name.as_ref() {
                    b"fills" if depth == 0 => depth = 1,
                    b"fill" if depth == 1 => {
                        depth = 2;
                        current_fill = ParsedFill::default();
                        _is_gradient = false;
                    }
                    b"patternFill" if depth == 2 => {
                        in_pattern_fill = true;
                        // Check for solid pattern
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"patternType" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if val != "solid" && val != "none" {
                                    // Non-solid patterns: just treat as none for now
                                }
                            }
                        }
                    }
                    b"gradientFill" if depth == 2 => {
                        _is_gradient = true;
                        if !unsupported.iter().any(|s| s.starts_with("gradient fills")) {
                            unsupported.push("gradient fills".to_string());
                        }
                    }
                    b"fgColor" if in_pattern_fill => {
                        let attrs = collect_attrs(e);
                        current_fill.bg_color = parse_color_attrs(&attrs, unsupported);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                match name.as_ref() {
                    b"patternFill" if depth == 2 => {
                        // Self-closing <patternFill patternType="none"/>
                    }
                    b"fgColor" if in_pattern_fill => {
                        let attrs = collect_attrs(e);
                        current_fill.bg_color = parse_color_attrs(&attrs, unsupported);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                match e.name().as_ref() {
                    b"patternFill" => in_pattern_fill = false,
                    b"fill" if depth == 2 => {
                        fills.push(current_fill.clone());
                        depth = 1;
                        in_pattern_fill = false;
                    }
                    b"fills" if depth == 1 => break,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    fills
}

/// Parse <borders> section into Vec<ParsedBorder>.
fn parse_borders(xml: &str, unsupported: &mut Vec<String>) -> Vec<ParsedBorder> {
    let mut borders = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut depth = 0; // 0 = outside, 1 = inside <borders>, 2 = inside <border>
    let mut current_side: Option<&'static str> = None;
    let mut current_border = ParsedBorder::default();
    let mut side_style = BorderStyle::None;
    let mut side_color: Option<[u8; 4]> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                match name.as_ref() {
                    b"borders" if depth == 0 => depth = 1,
                    b"border" if depth == 1 => {
                        depth = 2;
                        current_border = ParsedBorder::default();
                    }
                    b"left" | b"right" | b"top" | b"bottom" if depth == 2 => {
                        let side_name = match name.as_ref() {
                            b"left" => "left",
                            b"right" => "right",
                            b"top" => "top",
                            b"bottom" => "bottom",
                            _ => unreachable!(),
                        };
                        current_side = Some(side_name);
                        side_style = BorderStyle::None;
                        side_color = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style" {
                                side_style =
                                    parse_border_style(&String::from_utf8_lossy(&attr.value));
                            }
                        }
                    }
                    b"color" if current_side.is_some() => {
                        let attrs = collect_attrs(e);
                        side_color = parse_color_attrs(&attrs, unsupported);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                match name.as_ref() {
                    b"left" | b"right" | b"top" | b"bottom" if depth == 2 => {
                        // Self-closing border side with no style
                        let mut style = BorderStyle::None;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"style" {
                                style =
                                    parse_border_style(&String::from_utf8_lossy(&attr.value));
                            }
                        }
                        if style != BorderStyle::None {
                            let border = CellBorder {
                                style,
                                color: None,
                            };
                            match name.as_ref() {
                                b"left" => current_border.left = border,
                                b"right" => current_border.right = border,
                                b"top" => current_border.top = border,
                                b"bottom" => current_border.bottom = border,
                                _ => {}
                            }
                        }
                    }
                    b"color" if current_side.is_some() => {
                        let attrs = collect_attrs(e);
                        side_color = parse_color_attrs(&attrs, unsupported);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                match e.name().as_ref() {
                    b"left" | b"right" | b"top" | b"bottom" if depth == 2 => {
                        if let Some(side) = current_side.take() {
                            let border = CellBorder {
                                style: side_style,
                                color: side_color,
                            };
                            match side {
                                "left" => current_border.left = border,
                                "right" => current_border.right = border,
                                "top" => current_border.top = border,
                                "bottom" => current_border.bottom = border,
                                _ => {}
                            }
                        }
                        side_style = BorderStyle::None;
                        side_color = None;
                    }
                    b"border" if depth == 2 => {
                        borders.push(current_border.clone());
                        depth = 1;
                    }
                    b"borders" if depth == 1 => break,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    borders
}

fn parse_border_style(s: &str) -> BorderStyle {
    match s {
        "thin" | "hair" => BorderStyle::Thin,
        "medium" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" => BorderStyle::Medium,
        "thick" | "double" => BorderStyle::Thick,
        _ => BorderStyle::None,
    }
}

/// Parse <cellXfs> section and resolve each <xf> into a CellFormat.
fn parse_cell_xfs(
    xml: &str,
    custom_num_fmts: &HashMap<u16, String>,
    fonts: &[ParsedFont],
    fills: &[ParsedFill],
    borders: &[ParsedBorder],
    _unsupported: &mut Vec<String>,
) -> Vec<CellFormat> {
    let mut styles = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_cell_xfs = false;
    let mut in_xf = false;
    let mut current_xf = XfEntry::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"cellXfs" => {
                        in_cell_xfs = true;
                    }
                    b"xf" if in_cell_xfs => {
                        in_xf = true;
                        current_xf = XfEntry::default();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    current_xf.num_fmt_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"fontId" => {
                                    current_xf.font_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"fillId" => {
                                    current_xf.fill_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"borderId" => {
                                    current_xf.border_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"applyFont" => {
                                    current_xf.apply_font =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                b"applyFill" => {
                                    current_xf.apply_fill =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                b"applyBorder" => {
                                    current_xf.apply_border =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                b"applyNumberFormat" => {
                                    current_xf.apply_number_format =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                b"applyAlignment" => {
                                    current_xf.apply_alignment =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                _ => {}
                            }
                        }
                    }
                    b"alignment" if in_xf => {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    current_xf.h_align = Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                }
                                b"vertical" => {
                                    current_xf.v_align = Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                }
                                b"wrapText" => {
                                    current_xf.wrap_text =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                match e.name().as_ref() {
                    b"xf" if in_cell_xfs => {
                        // Self-closing <xf .../> — parse and push immediately
                        current_xf = XfEntry::default();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    current_xf.num_fmt_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"fontId" => {
                                    current_xf.font_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"fillId" => {
                                    current_xf.fill_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"borderId" => {
                                    current_xf.border_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"applyFont" => {
                                    current_xf.apply_font =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                b"applyFill" => {
                                    current_xf.apply_fill =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                b"applyBorder" => {
                                    current_xf.apply_border =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                b"applyNumberFormat" => {
                                    current_xf.apply_number_format =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                b"applyAlignment" => {
                                    current_xf.apply_alignment =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                _ => {}
                            }
                        }
                        styles.push(resolve_xf(&current_xf, custom_num_fmts, fonts, fills, borders));
                    }
                    b"alignment" if in_xf => {
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    current_xf.h_align = Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                }
                                b"vertical" => {
                                    current_xf.v_align = Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                }
                                b"wrapText" => {
                                    current_xf.wrap_text =
                                        attr.value.as_ref() == b"1" || attr.value.as_ref() == b"true";
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                match e.name().as_ref() {
                    b"xf" if in_xf => {
                        styles.push(resolve_xf(&current_xf, custom_num_fmts, fonts, fills, borders));
                        in_xf = false;
                    }
                    b"cellXfs" => break,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    styles
}

#[derive(Debug, Default)]
struct XfEntry {
    num_fmt_id: Option<u16>,
    font_id: Option<usize>,
    fill_id: Option<usize>,
    border_id: Option<usize>,
    apply_font: bool,
    apply_fill: bool,
    apply_border: bool,
    apply_number_format: bool,
    apply_alignment: bool,
    h_align: Option<String>,
    v_align: Option<String>,
    wrap_text: bool,
}

/// Resolve an XfEntry into a CellFormat using the parsed component tables.
fn resolve_xf(
    xf: &XfEntry,
    custom_num_fmts: &HashMap<u16, String>,
    fonts: &[ParsedFont],
    fills: &[ParsedFill],
    borders: &[ParsedBorder],
) -> CellFormat {
    let mut format = CellFormat::default();

    // Font
    if let Some(font_id) = xf.font_id {
        if let Some(font) = fonts.get(font_id) {
            format.bold = font.bold;
            format.italic = font.italic;
            format.underline = font.underline;
            format.strikethrough = font.strikethrough;
            // Only store non-default font sizes (Excel default is ~11pt)
            if let Some(size) = font.size {
                format.font_size = Some(size);
            }
            format.font_color = font.color;
            format.font_family = font.family.clone();
        }
    }

    // Fill
    if let Some(fill_id) = xf.fill_id {
        if let Some(fill) = fills.get(fill_id) {
            format.background_color = fill.bg_color;
        }
    }

    // Border
    if let Some(border_id) = xf.border_id {
        if let Some(border) = borders.get(border_id) {
            format.border_top = border.top;
            format.border_right = border.right;
            format.border_bottom = border.bottom;
            format.border_left = border.left;
        }
    }

    // Number format
    if let Some(num_fmt_id) = xf.num_fmt_id {
        if let Some(code) = custom_num_fmts.get(&num_fmt_id) {
            format.number_format = NumberFormat::Custom(code.clone());
        } else {
            format.number_format = builtin_number_format(num_fmt_id);
        }
    }

    // Alignment
    if let Some(ref h) = xf.h_align {
        format.alignment = match h.as_str() {
            "left" => Alignment::Left,
            "center" => Alignment::Center,
            "right" => Alignment::Right,
            "general" => Alignment::General,
            "centerContinuous" => Alignment::CenterAcrossSelection,
            _ => Alignment::General,
        };
    }

    if let Some(ref v) = xf.v_align {
        format.vertical_alignment = match v.as_str() {
            "top" => VerticalAlignment::Top,
            "center" => VerticalAlignment::Middle,
            "bottom" => VerticalAlignment::Bottom,
            _ => VerticalAlignment::Middle,
        };
    }

    if xf.wrap_text {
        format.text_overflow = TextOverflow::Wrap;
    }

    format
}

// =============================================================================
// Worksheet XML parser — per-cell style IDs + layout
// =============================================================================

/// Parse a worksheet XML to extract per-cell style IDs and layout dimensions.
pub fn parse_sheet_formatting(xml: &str) -> SheetFormatting {
    let mut cell_styles = Vec::new();
    let mut col_widths = HashMap::new();
    let mut row_heights = HashMap::new();
    let mut merged_regions = Vec::new();

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut _current_row: Option<usize> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match e.name().as_ref() {
                    b"row" => {
                        let mut row_idx: Option<usize> = None;
                        let mut custom_height = false;
                        let mut ht: Option<f64> = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    row_idx = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse::<usize>().ok())
                                        .map(|r| r.saturating_sub(1)); // 1-based → 0-based
                                }
                                b"ht" => {
                                    ht = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"customHeight" => {
                                    custom_height = attr.value.as_ref() == b"1"
                                        || attr.value.as_ref() == b"true";
                                }
                                _ => {}
                            }
                        }

                        _current_row = row_idx;

                        if custom_height {
                            if let (Some(row), Some(height)) = (row_idx, ht) {
                                row_heights.insert(row, height);
                            }
                        }
                    }
                    b"c" => {
                        // Cell element: extract style ID from s="N" attribute
                        let mut style_id: Option<usize> = None;
                        let mut cell_ref: Option<String> = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"s" => {
                                    style_id = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"r" => {
                                    cell_ref = Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                }
                                _ => {}
                            }
                        }

                        if let (Some(style_id), Some(ref cell_ref)) = (style_id, &cell_ref) {
                            if style_id > 0 {
                                // style_id 0 = default, skip
                                if let Some((row, col)) = parse_cell_ref(cell_ref) {
                                    cell_styles.push((row, col, style_id));
                                }
                            }
                        }
                    }
                    b"col" => {
                        // Column width
                        let mut min_col: Option<usize> = None;
                        let mut max_col: Option<usize> = None;
                        let mut width: Option<f64> = None;
                        let mut custom_width = false;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"min" => {
                                    min_col = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse::<usize>().ok())
                                        .map(|c| c.saturating_sub(1));
                                }
                                b"max" => {
                                    max_col = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse::<usize>().ok())
                                        .map(|c| c.saturating_sub(1));
                                }
                                b"width" => {
                                    width = std::str::from_utf8(&attr.value)
                                        .ok()
                                        .and_then(|s| s.parse().ok());
                                }
                                b"customWidth" => {
                                    custom_width = attr.value.as_ref() == b"1"
                                        || attr.value.as_ref() == b"true";
                                }
                                _ => {}
                            }
                        }

                        if custom_width {
                            if let (Some(min), Some(max), Some(w)) = (min_col, max_col, width) {
                                for col in min..=max {
                                    col_widths.insert(col, w);
                                }
                            }
                        }
                    }
                    b"mergeCell" => {
                        // Parse ref="A1:C3" attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                let ref_str = String::from_utf8_lossy(&attr.value);
                                if let Some(region) = parse_merge_ref(&ref_str) {
                                    merged_regions.push(region);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if e.name().as_ref() == b"row" {
                    _current_row = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    SheetFormatting {
        cell_styles,
        col_widths,
        row_heights,
        merged_regions,
    }
}

/// Parse a merge range reference like "A1:C3" into (start_row, start_col, end_row, end_col).
pub fn parse_merge_ref(r: &str) -> Option<(usize, usize, usize, usize)> {
    let parts: Vec<&str> = r.split(':').collect();
    if parts.len() != 2 {
        return None;
    }
    let (sr, sc) = parse_cell_ref(parts[0])?;
    let (er, ec) = parse_cell_ref(parts[1])?;
    Some((sr, sc, er, ec))
}

/// Parse a cell reference like "B5" into (row, col) = (4, 1).
fn parse_cell_ref(r: &str) -> Option<(usize, usize)> {
    let mut col_part = String::new();
    let mut row_part = String::new();

    for ch in r.chars() {
        if ch.is_ascii_alphabetic() {
            col_part.push(ch);
        } else if ch.is_ascii_digit() {
            row_part.push(ch);
        }
    }

    if col_part.is_empty() || row_part.is_empty() {
        return None;
    }

    let mut col: usize = 0;
    for ch in col_part.chars() {
        col = col * 26 + (ch.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    col = col.saturating_sub(1); // 1-based → 0-based

    let row: usize = row_part.parse().ok()?;
    let row = row.saturating_sub(1); // 1-based → 0-based

    Some((row, col))
}

// =============================================================================
// Top-level import entry point
// =============================================================================

/// Parse all formatting data from an XLSX file.
/// Returns (style_table, per_sheet_formatting, stats).
/// `sheet_names` must match the order of sheets in the workbook.
pub fn parse_xlsx_formatting(
    path: &Path,
    sheet_names: &[String],
) -> Result<(StyleTable, Vec<SheetFormatting>, StyleImportStats), String> {
    let file = std::fs::File::open(path)
        .map_err(|e| format!("Failed to open XLSX file for styles: {}", e))?;
    let mut archive = ZipArchive::new(file)
        .map_err(|e| format!("Failed to read XLSX as ZIP for styles: {}", e))?;

    let mut stats = StyleImportStats::default();

    // Step 1: Parse styles.xml
    let (style_table, unsupported) = match read_zip_file(&mut archive, "xl/styles.xml") {
        Ok(xml) => parse_styles_xml(&xml),
        Err(_) => {
            // No styles.xml — return empty
            return Ok((
                StyleTable {
                    styles: Vec::new(),
                },
                sheet_names.iter().map(|_| SheetFormatting {
                    cell_styles: Vec::new(),
                    col_widths: HashMap::new(),
                    row_heights: HashMap::new(),
                    merged_regions: Vec::new(),
                }).collect(),
                stats,
            ));
        }
    };
    stats.unsupported_features = unsupported;
    stats.unique_styles = style_table.len();

    // Step 2: Resolve worksheet paths
    let workbook_xml = read_zip_file(&mut archive, "xl/workbook.xml")
        .unwrap_or_default();
    let rels_xml = read_zip_file(&mut archive, "xl/_rels/workbook.xml.rels")
        .unwrap_or_default();
    let worksheet_paths = resolve_worksheet_paths_for_sheets(&workbook_xml, &rels_xml, sheet_names);

    // Step 3: Parse each worksheet for per-cell style IDs and layout
    let mut sheet_formats = Vec::new();
    for ws_path in &worksheet_paths {
        let formatting = match read_zip_file(&mut archive, ws_path) {
            Ok(xml) => {
                let sf = parse_sheet_formatting(&xml);
                stats.styles_imported += sf.cell_styles.len();
                sf
            }
            Err(_) => SheetFormatting {
                cell_styles: Vec::new(),
                col_widths: HashMap::new(),
                row_heights: HashMap::new(),
                merged_regions: Vec::new(),
            },
        };
        sheet_formats.push(formatting);
    }

    // Pad with empty formatting if we have fewer worksheet paths than sheets
    while sheet_formats.len() < sheet_names.len() {
        sheet_formats.push(SheetFormatting {
            cell_styles: Vec::new(),
            col_widths: HashMap::new(),
            row_heights: HashMap::new(),
            merged_regions: Vec::new(),
        });
    }

    Ok((style_table, sheet_formats, stats))
}

/// Check if a style is visually relevant for empty cells.
/// Only create cell entries for styles with background, borders, or center-across.
pub fn is_style_visually_relevant(format: &CellFormat) -> bool {
    format.background_color.is_some()
        || format.border_top.style != BorderStyle::None
        || format.border_right.style != BorderStyle::None
        || format.border_bottom.style != BorderStyle::None
        || format.border_left.style != BorderStyle::None
        || format.alignment == Alignment::CenterAcrossSelection
}

// =============================================================================
// Helpers
// =============================================================================

/// Collect XML attributes into a Vec of (key, value) pairs.
fn collect_attrs(e: &quick_xml::events::BytesStart) -> Vec<(Vec<u8>, Vec<u8>)> {
    e.attributes()
        .flatten()
        .map(|a| (a.key.as_ref().to_vec(), a.value.to_vec()))
        .collect()
}

/// Read a file from a ZIP archive.
fn read_zip_file<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
) -> Result<String, String> {
    let mut file = archive
        .by_name(path)
        .map_err(|e| format!("File '{}' not found in XLSX: {}", path, e))?;
    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| format!("Failed to read '{}': {}", path, e))?;
    Ok(content)
}

/// Resolve worksheet XML paths for specific sheet names (in order).
fn resolve_worksheet_paths_for_sheets(
    workbook_xml: &str,
    rels_xml: &str,
    sheet_names: &[String],
) -> Vec<String> {
    // Parse workbook.xml to get (name, rId) pairs
    let mut name_to_rid: Vec<(String, String)> = Vec::new();
    let mut reader = Reader::from_str(workbook_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"sheet" =>
            {
                let mut name = None;
                let mut rid = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"name" => {
                            name = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        b"r:id" => {
                            rid = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        _ => {}
                    }
                }
                if let (Some(name), Some(rid)) = (name, rid) {
                    name_to_rid.push((name, rid));
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Parse rels to get rid → target
    let mut rid_to_target: HashMap<String, String> = HashMap::new();
    let mut reader = Reader::from_str(rels_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = Some(String::from_utf8_lossy(&attr.value).to_string()),
                        b"Target" => {
                            target = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    rid_to_target.insert(id, target);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Map sheet_names → paths in order
    let name_rid_map: HashMap<&str, &str> = name_to_rid
        .iter()
        .map(|(n, r)| (n.as_str(), r.as_str()))
        .collect();

    sheet_names
        .iter()
        .map(|name| {
            name_rid_map
                .get(name.as_str())
                .and_then(|rid| rid_to_target.get(*rid))
                .map(|target| format!("xl/{}", target))
                .unwrap_or_default()
        })
        .collect()
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_argb_hex() {
        // AARRGGBB
        assert_eq!(parse_argb_hex(b"FF0000FF"), Some([0, 0, 255, 255]));
        // RRGGBB
        assert_eq!(parse_argb_hex(b"FF0000"), Some([255, 0, 0, 255]));
        // With alpha
        assert_eq!(parse_argb_hex(b"80FF0000"), Some([255, 0, 0, 128]));
    }

    #[test]
    fn test_indexed_color() {
        assert_eq!(indexed_color(0), Some([0, 0, 0, 255]));       // Black
        assert_eq!(indexed_color(1), Some([255, 255, 255, 255])); // White
        assert_eq!(indexed_color(2), Some([255, 0, 0, 255]));     // Red
        assert_eq!(indexed_color(99), None);                       // Out of range
    }

    #[test]
    fn test_builtin_number_format() {
        assert_eq!(builtin_number_format(0), NumberFormat::General);
        assert_eq!(
            builtin_number_format(2),
            NumberFormat::Number { decimals: 2 }
        );
        assert_eq!(
            builtin_number_format(9),
            NumberFormat::Percent { decimals: 0 }
        );
        assert_eq!(
            builtin_number_format(14),
            NumberFormat::Date {
                style: visigrid_engine::cell::DateStyle::Short,
            }
        );
        assert_eq!(
            builtin_number_format(44),
            NumberFormat::Currency { decimals: 2 }
        );
    }

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("B5"), Some((4, 1)));
        assert_eq!(parse_cell_ref("Z1"), Some((0, 25)));
        assert_eq!(parse_cell_ref("AA1"), Some((0, 26)));
        assert_eq!(parse_cell_ref("AZ10"), Some((9, 51)));
    }

    #[test]
    fn test_parse_border_style() {
        assert_eq!(parse_border_style("thin"), BorderStyle::Thin);
        assert_eq!(parse_border_style("medium"), BorderStyle::Medium);
        assert_eq!(parse_border_style("thick"), BorderStyle::Thick);
        assert_eq!(parse_border_style("none"), BorderStyle::None);
        assert_eq!(parse_border_style("hair"), BorderStyle::Thin);
    }

    #[test]
    fn test_parse_minimal_styles_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>
  </borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="1" borderId="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>"#;

        let (table, _unsupported) = parse_styles_xml(xml);
        assert_eq!(table.len(), 3);

        // Style 0: default
        let s0 = &table.styles[0];
        assert!(!s0.bold);
        assert_eq!(s0.font_size, Some(11.0));

        // Style 1: bold, red, 14pt Arial
        let s1 = &table.styles[1];
        assert!(s1.bold);
        assert_eq!(s1.font_size, Some(14.0));
        assert_eq!(s1.font_color, Some([255, 0, 0, 255]));
        assert_eq!(s1.font_family, Some("Arial".to_string()));

        // Style 2: yellow fill, thin borders
        let s2 = &table.styles[2];
        assert_eq!(s2.background_color, Some([255, 255, 0, 255]));
        assert_eq!(s2.border_top.style, BorderStyle::Thin);
        assert_eq!(s2.border_bottom.style, BorderStyle::Thin);
        assert_eq!(s2.border_left.style, BorderStyle::Thin);
        assert_eq!(s2.border_right.style, BorderStyle::Thin);
    }

    #[test]
    fn test_parse_custom_number_format() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="1"><font><sz val="11"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"##;

        let (table, _) = parse_styles_xml(xml);
        assert_eq!(table.len(), 2);
        assert!(matches!(&table.styles[1].number_format, NumberFormat::Custom(code) if code == "#,##0.00"));
    }

    #[test]
    fn test_parse_empty_styles_xml() {
        // Minimal valid but empty
        let xml = r#"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</styleSheet>"#;

        let (table, unsupported) = parse_styles_xml(xml);
        assert_eq!(table.len(), 0);
        assert!(unsupported.is_empty());
    }

    #[test]
    fn test_parse_sheet_formatting_cell_styles() {
        let xml = r#"<?xml version="1.0"?>
<worksheet>
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><v>100</v></c>
      <c r="B1" s="0"><v>200</v></c>
      <c r="C1" s="2"><v>300</v></c>
    </row>
    <row r="2">
      <c r="A2" s="3"><v>400</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let sf = parse_sheet_formatting(xml);
        // Style 0 is default, should be skipped
        assert_eq!(sf.cell_styles.len(), 3);
        assert!(sf.cell_styles.contains(&(0, 0, 1))); // A1 → style 1
        assert!(sf.cell_styles.contains(&(0, 2, 2))); // C1 → style 2
        assert!(sf.cell_styles.contains(&(1, 0, 3))); // A2 → style 3
    }

    #[test]
    fn test_parse_sheet_formatting_col_widths() {
        let xml = r#"<?xml version="1.0"?>
<worksheet>
  <cols>
    <col min="1" max="1" width="15.5" customWidth="1"/>
    <col min="3" max="5" width="20.0" customWidth="1"/>
    <col min="6" max="6" width="8.0"/>
  </cols>
  <sheetData></sheetData>
</worksheet>"#;

        let sf = parse_sheet_formatting(xml);
        assert_eq!(sf.col_widths.len(), 4); // col 0, 2, 3, 4 (custom only)
        assert_eq!(sf.col_widths[&0], 15.5);
        assert_eq!(sf.col_widths[&2], 20.0);
        assert_eq!(sf.col_widths[&3], 20.0);
        assert_eq!(sf.col_widths[&4], 20.0);
        assert!(!sf.col_widths.contains_key(&5)); // Not custom
    }

    #[test]
    fn test_parse_sheet_formatting_row_heights() {
        let xml = r#"<?xml version="1.0"?>
<worksheet>
  <sheetData>
    <row r="1" ht="30.0" customHeight="1">
      <c r="A1"><v>1</v></c>
    </row>
    <row r="2" ht="15.0">
      <c r="A2"><v>2</v></c>
    </row>
    <row r="3" ht="45.5" customHeight="1">
      <c r="A3"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>"#;

        let sf = parse_sheet_formatting(xml);
        assert_eq!(sf.row_heights.len(), 2); // Only custom heights
        assert_eq!(sf.row_heights[&0], 30.0);
        assert_eq!(sf.row_heights[&2], 45.5);
        assert!(!sf.row_heights.contains_key(&1)); // Not custom
    }

    #[test]
    fn test_is_style_visually_relevant() {
        let default = CellFormat::default();
        assert!(!is_style_visually_relevant(&default));

        let with_bg = CellFormat {
            background_color: Some([255, 255, 0, 255]),
            ..Default::default()
        };
        assert!(is_style_visually_relevant(&with_bg));

        let with_border = CellFormat {
            border_top: CellBorder::thin(),
            ..Default::default()
        };
        assert!(is_style_visually_relevant(&with_border));

        let center_across = CellFormat {
            alignment: Alignment::CenterAcrossSelection,
            ..Default::default()
        };
        assert!(is_style_visually_relevant(&center_across));

        let bold_only = CellFormat {
            bold: true,
            ..Default::default()
        };
        assert!(!is_style_visually_relevant(&bold_only));
    }

    #[test]
    fn test_parse_alignment() {
        let xml = r#"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment horizontal="center" vertical="top" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">
      <alignment horizontal="centerContinuous"/>
    </xf>
  </cellXfs>
</styleSheet>"#;

        let (table, _) = parse_styles_xml(xml);
        assert_eq!(table.len(), 3);

        let s1 = &table.styles[1];
        assert_eq!(s1.alignment, Alignment::Center);
        assert_eq!(s1.vertical_alignment, VerticalAlignment::Top);
        assert_eq!(s1.text_overflow, TextOverflow::Wrap);

        let s2 = &table.styles[2];
        assert_eq!(s2.alignment, Alignment::CenterAcrossSelection);
    }

    #[test]
    fn test_parse_gradient_fill_reported() {
        let xml = r#"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><gradientFill><stop position="0"><color rgb="FF000000"/></stop></gradientFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>"#;

        let (_, unsupported) = parse_styles_xml(xml);
        assert!(unsupported.iter().any(|s| s.contains("gradient")));
    }

    #[test]
    fn test_theme_color_approximate_warning() {
        let xml = r#"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/></font>
    <font><sz val="11"/><color theme="4"/></font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>"#;

        let (table, unsupported) = parse_styles_xml(xml);
        // Theme 4 = Accent 1 = [79, 129, 189]
        assert_eq!(table.styles[1].font_color, Some([79, 129, 189, 255]));
        assert!(unsupported.iter().any(|s| s.contains("theme tints")));
    }

    #[test]
    fn test_unescape_xml() {
        assert_eq!(unescape_xml("hello"), "hello");
        assert_eq!(unescape_xml("&quot;$&quot;#,##0"), "\"$\"#,##0");
        assert_eq!(unescape_xml("&amp;"), "&");
        assert_eq!(unescape_xml("&lt;&gt;"), "<>");
        assert_eq!(unescape_xml("no entities"), "no entities");
    }

    #[test]
    fn test_numfmt_xml_entity_unescaping() {
        // Format codes in styles.xml use XML entities for quotes:
        // formatCode="&quot;$&quot;#,##0.00" should become "$"#,##0.00
        let xml = r##"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>
  </numFmts>
  <fonts count="1"><font><sz val="11"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"##;

        let (table, _unsupported) = parse_styles_xml(xml);
        assert_eq!(table.styles.len(), 1);
        match &table.styles[0].number_format {
            NumberFormat::Custom(code) => {
                assert_eq!(code, r##""$"#,##0.00"##, "Format code should have unescaped quotes");
            }
            other => panic!("Expected Custom format, got {:?}", other),
        }
    }

    #[test]
    fn test_parse_merge_ref() {
        assert_eq!(parse_merge_ref("A1:C3"), Some((0, 0, 2, 2)));
        assert_eq!(parse_merge_ref("B2:D5"), Some((1, 1, 4, 3)));
        assert_eq!(parse_merge_ref("AA1:AB10"), Some((0, 26, 9, 27)));
        assert_eq!(parse_merge_ref("A1"), None); // no colon
        assert_eq!(parse_merge_ref(""), None);
    }

    #[test]
    fn test_parse_sheet_formatting_merges() {
        let xml = r#"<?xml version="1.0"?>
<worksheet>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  <mergeCells count="2">
    <mergeCell ref="A1:C3"/>
    <mergeCell ref="E5:F10"/>
  </mergeCells>
</worksheet>"#;

        let sf = parse_sheet_formatting(xml);
        assert_eq!(sf.merged_regions.len(), 2);
        assert_eq!(sf.merged_regions[0], (0, 0, 2, 2));   // A1:C3
        assert_eq!(sf.merged_regions[1], (4, 4, 9, 5));   // E5:F10
    }
}
