//! Validated global defaults for otherwise unsized rows and columns.
use super::UserSettings;

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CellSizeDefaults {
    pub column_width: f32,
    pub row_height: f32,
}

impl CellSizeDefaults {
    pub const COLUMN_WIDTH: f32 = 96.0;
    pub const ROW_HEIGHT: f32 = 28.0;
    pub const COLUMN_LIMITS: (f32, f32) = (20.0, 500.0);
    pub const ROW_LIMITS: (f32, f32) = (12.0, 200.0);

    pub fn from_user(user: &UserSettings) -> Self {
        fn dimension(value: f32, fallback: f32, limits: (f32, f32)) -> f32 {
            if value.is_finite() {
                value.clamp(limits.0, limits.1)
            } else {
                fallback
            }
        }
        Self {
            column_width: dimension(
                user.appearance
                    .default_column_width
                    .resolve(Self::COLUMN_WIDTH),
                Self::COLUMN_WIDTH,
                Self::COLUMN_LIMITS,
            ),
            row_height: dimension(
                user.appearance.default_row_height.resolve(Self::ROW_HEIGHT),
                Self::ROW_HEIGHT,
                Self::ROW_LIMITS,
            ),
        }
    }
}

impl Default for CellSizeDefaults {
    fn default() -> Self {
        Self {
            column_width: Self::COLUMN_WIDTH,
            row_height: Self::ROW_HEIGHT,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::settings::Setting;

    #[test]
    fn json_sizes_roundtrip_and_missing_values_inherit() {
        let user: UserSettings = serde_json::from_str(
            r#"{"appearance":{"default_column_width":80,"default_row_height":24}}"#,
        )
        .unwrap();
        let sizes = CellSizeDefaults::from_user(&user);
        assert_eq!(
            sizes,
            CellSizeDefaults {
                column_width: 80.0,
                row_height: 24.0
            }
        );
        let saved = serde_json::to_string(&user).unwrap();
        let restored = serde_json::from_str(&saved).unwrap();
        assert_eq!(CellSizeDefaults::from_user(&restored), sizes);
        for json in [
            "{}",
            r#"{"appearance":{}}"#,
            r#"{"appearance":{"default_column_width":null}}"#,
        ] {
            assert_eq!(
                CellSizeDefaults::from_user(&serde_json::from_str(json).unwrap()),
                CellSizeDefaults::default()
            );
        }
    }

    #[test]
    fn dimensions_are_finite_and_within_resize_limits() {
        let mut user = UserSettings::default();
        user.appearance.default_column_width = Setting::Value(-80.0);
        user.appearance.default_row_height = Setting::Value(10000.0);
        assert_eq!(
            CellSizeDefaults::from_user(&user),
            CellSizeDefaults {
                column_width: 20.0,
                row_height: 200.0
            }
        );
        user.appearance.default_column_width = Setting::Value(f32::NAN);
        user.appearance.default_row_height = Setting::Value(f32::INFINITY);
        assert_eq!(
            CellSizeDefaults::from_user(&user),
            CellSizeDefaults::default()
        );
    }
}
