//! Resolve display fonts without changing the workbook's requested family.
use crate::app::Spreadsheet;
use crate::settings::DEFAULT_FONT_FAMILY;
use gpui::{App, Font, SharedString};
use std::collections::HashMap;

pub struct FontCatalog {
    families: HashMap<String, SharedString>,
}

impl FontCatalog {
    pub fn new(cx: &App) -> Self {
        let text = cx.text_system();
        let names = text.all_font_names().into_iter().filter(|name| {
            if name.starts_with('.') || name == crate::SYMBOL_FONT_FAMILY {
                return false;
            }
            // GPUI appends these fallback candidates even when they aren't installed.
            if matches!(
                name.as_str(),
                "Helvetica"
                    | "Segoe UI"
                    | "Ubuntu"
                    | "Adwaita Sans"
                    | "Cantarell"
                    | "Noto Sans"
                    | "DejaVu Sans"
                    | "Arial"
            ) {
                let font = gpui::font(name.clone());
                return text
                    .get_font_for_id(text.resolve_font(&font))
                    .is_some_and(|resolved| resolved.family.as_ref() == name);
            }
            true
        });
        Self::from_names(names)
    }

    fn from_names(names: impl IntoIterator<Item = String>) -> Self {
        Self {
            families: names
                .into_iter()
                .map(|name| (name.to_lowercase(), name.into()))
                .collect(),
        }
    }

    pub fn names(&self) -> Vec<String> {
        let mut names: Vec<_> = self.families.values().map(ToString::to_string).collect();
        names.sort_unstable();
        names
    }

    pub fn resolve(&self, requested: &str) -> SharedString {
        self.families
            .get(&requested.to_lowercase())
            .cloned()
            .unwrap_or_else(|| DEFAULT_FONT_FAMILY.into())
    }

    pub fn is_missing(&self, requested: &str) -> bool {
        !self.families.contains_key(&requested.to_lowercase())
    }
}

impl Spreadsheet {
    pub fn cell_font_family(&self, explicit: Option<&str>) -> SharedString {
        self.font_catalog
            .resolve(explicit.unwrap_or(&self.cell_font.family))
    }

    pub fn cell_font_size(&self, explicit_points: Option<f32>) -> f32 {
        self.cell_font.pixels(explicit_points, self.metrics.zoom)
    }

    pub fn cell_font(&self, format: &visigrid_engine::cell::CellFormat) -> Font {
        Font {
            family: self.cell_font_family(format.font_family.as_deref()),
            fallbacks: Some(gpui::FontFallbacks::from_fonts(vec![
                crate::SYMBOL_FONT_FAMILY.into(),
            ])),
            weight: if format.bold {
                gpui::FontWeight::BOLD
            } else {
                gpui::FontWeight::NORMAL
            },
            style: if format.italic {
                gpui::FontStyle::Italic
            } else {
                gpui::FontStyle::Normal
            },
            ..Font::default()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn missing_fonts_use_bundled_fallback_without_rewriting_the_request() {
        let catalog = FontCatalog::from_names([DEFAULT_FONT_FAMILY.into(), "Example Sans".into()]);
        let requested = "Calibri".to_string();
        assert!(catalog.is_missing(&requested));
        assert_eq!(catalog.resolve(&requested).as_ref(), DEFAULT_FONT_FAMILY);
        assert_eq!(requested, "Calibri");
        assert!(!catalog.is_missing("example sans"));
        assert_eq!(catalog.resolve("example sans").as_ref(), "Example Sans");
    }
}
