//! Cell font sizes are stored in points, independently of zoom and display DPI.
use super::UserSettings;

pub const DEFAULT_FONT_FAMILY: &str = "IBM Plex Sans";
pub const DEFAULT_FONT_SIZE: f32 = 11.0;
pub const FONT_SIZE_LIMITS: (f32, f32) = (1.0, 400.0);

pub fn points_to_pixels(points: f32) -> f32 {
    points * (96.0 / 72.0)
}

#[derive(Debug, Clone, PartialEq)]
pub struct CellFontDefaults {
    pub family: String,
    pub size: f32,
}

impl CellFontDefaults {
    pub fn from_user(user: &UserSettings) -> Self {
        let family = user
            .appearance
            .default_font_family
            .resolve(DEFAULT_FONT_FAMILY.to_string());
        let size = user.appearance.default_font_size.resolve(DEFAULT_FONT_SIZE);
        Self {
            family: if family.trim().is_empty() {
                DEFAULT_FONT_FAMILY.into()
            } else {
                family.trim().into()
            },
            size: if size.is_finite() {
                size.clamp(FONT_SIZE_LIMITS.0, FONT_SIZE_LIMITS.1)
            } else {
                DEFAULT_FONT_SIZE
            },
        }
    }

    pub fn pixels(&self, explicit_points: Option<f32>, zoom: f32) -> f32 {
        points_to_pixels(
            explicit_points
                .filter(|s| s.is_finite() && *s > 0.0)
                .unwrap_or(self.size),
        ) * zoom
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::settings::Setting;

    #[test]
    fn sizes_are_points_and_explicit_sizes_win_at_every_zoom() {
        let user: UserSettings = serde_json::from_str(
            r#"{"appearance":{"default_font_family":"IBM Plex Mono","default_font_size":12.5}}"#,
        )
        .unwrap();
        let defaults = CellFontDefaults::from_user(&user);
        assert_eq!(defaults.family, "IBM Plex Mono");
        for zoom in [0.5, 1.0, 1.5, 2.0] {
            assert!((defaults.pixels(Some(11.0), zoom) - 14.666667 * zoom).abs() < 0.0001);
            assert!((defaults.pixels(None, zoom) - 16.666667 * zoom).abs() < 0.0001);
        }
        let restored = serde_json::from_str(&serde_json::to_string(&user).unwrap()).unwrap();
        assert_eq!(defaults, CellFontDefaults::from_user(&restored));
    }

    #[test]
    fn missing_and_invalid_defaults_are_safe() {
        let mut user = UserSettings::default();
        let defaults = CellFontDefaults::from_user(&user);
        assert_eq!(defaults.family, DEFAULT_FONT_FAMILY);
        assert_eq!(defaults.size, 11.0);
        user.appearance.default_font_family = Setting::Value("  ".into());
        user.appearance.default_font_size = Setting::Value(f32::NAN);
        assert_eq!(CellFontDefaults::from_user(&user), defaults);
        user.appearance.default_font_size = Setting::Value(500.0);
        assert_eq!(CellFontDefaults::from_user(&user).size, 400.0);
    }
}
