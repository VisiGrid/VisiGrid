//! Toolbar preferences tolerate newer fields without resetting unrelated settings.
use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ToolbarLayout {
    #[default]
    Compact,
    Ribbon,
}

/// Retain unknown fields/command IDs on round trips, including future schemas.
/// Resolve known fields individually so one malformed field is not fatal.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(transparent)]
pub struct ToolbarSettings(Value);

impl Default for ToolbarSettings {
    fn default() -> Self {
        Self(Value::Object(Map::new()))
    }
}

impl ToolbarSettings {
    pub fn layout(&self) -> ToolbarLayout {
        if self.is_supported() && self.0.get("layout").and_then(Value::as_str) == Some("ribbon") {
            ToolbarLayout::Ribbon
        } else {
            ToolbarLayout::Compact
        }
    }

    pub fn collapsed(&self) -> bool {
        self.is_supported()
            && self
                .0
                .get("ribbon_collapsed")
                .and_then(Value::as_bool)
                .unwrap_or(false)
    }

    pub fn is_supported(&self) -> bool {
        self.0
            .get("schema_version")
            .and_then(Value::as_u64)
            .unwrap_or(1)
            <= 1
    }

    fn fields(&mut self) -> Option<&mut Map<String, Value>> {
        if !self.is_supported() {
            return None;
        }
        if !self.0.is_object() {
            self.0 = Value::Object(Map::new());
        }
        let fields = self.0.as_object_mut().unwrap();
        fields.insert("schema_version".into(), Value::from(1));
        Some(fields)
    }

    pub fn set_layout(&mut self, layout: ToolbarLayout) -> bool {
        let Some(fields) = self.fields() else {
            return false;
        };
        fields.insert(
            "layout".into(),
            Value::from(match layout {
                ToolbarLayout::Compact => "compact",
                ToolbarLayout::Ribbon => "ribbon",
            }),
        );
        true
    }

    pub fn set_collapsed(&mut self, collapsed: bool) -> bool {
        let Some(fields) = self.fields() else {
            return false;
        };
        fields.insert("ribbon_collapsed".into(), Value::from(collapsed));
        true
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::settings::UserSettings;

    #[test]
    fn legacy_hidden_toolbar_and_other_preferences_survive() {
        let settings: UserSettings =
            serde_json::from_str(r#"{"appearance":{"show_format_bar":false,"theme_id":"test"}}"#)
                .unwrap();
        assert_eq!(settings.appearance.toolbar.layout(), ToolbarLayout::Compact);
        assert!(!settings.appearance.show_format_bar.resolve(true));
        assert_eq!(settings.appearance.theme_id.as_value().unwrap(), "test");
    }

    #[test]
    fn malformed_fields_fall_back_independently() {
        for raw in [
            "null",
            "[]",
            "42",
            r#""bad""#,
            r#"{"layout":false,"ribbon_collapsed":[]}"#,
        ] {
            let settings: UserSettings = serde_json::from_str(&format!(
                r#"{{"appearance":{{"theme_id":"test","toolbar":{raw}}}}}"#
            ))
            .unwrap();
            assert_eq!(settings.appearance.toolbar.layout(), ToolbarLayout::Compact);
            assert!(!settings.appearance.toolbar.collapsed());
            assert_eq!(settings.appearance.theme_id.as_value().unwrap(), "test");
        }
        let settings: ToolbarSettings =
            serde_json::from_str(r#"{"layout":"ribbon","ribbon_collapsed":"bad"}"#).unwrap();
        assert_eq!(settings.layout(), ToolbarLayout::Ribbon);
        assert!(!settings.collapsed());
    }

    #[test]
    fn updates_preserve_unknown_customization_and_future_schema() {
        let mut settings: ToolbarSettings =
            serde_json::from_str(r#"{"quick_access":["future.command"],"custom":{"x":1}}"#)
                .unwrap();
        assert!(settings.set_layout(ToolbarLayout::Ribbon));
        assert!(settings.set_collapsed(true));
        let value = serde_json::to_value(&settings).unwrap();
        assert_eq!(value["quick_access"][0], "future.command");
        assert_eq!(value["custom"]["x"], 1);
        let roundtrip: ToolbarSettings = serde_json::from_value(value).unwrap();
        assert_eq!(roundtrip.layout(), ToolbarLayout::Ribbon);
        assert!(roundtrip.collapsed());
        let raw = serde_json::json!({"schema_version": 2, "layout":"future", "custom": [1,2]});
        let mut future: ToolbarSettings = serde_json::from_value(raw.clone()).unwrap();
        assert_eq!(future.layout(), ToolbarLayout::Compact);
        assert!(!future.set_layout(ToolbarLayout::Ribbon));
        assert_eq!(serde_json::to_value(future).unwrap(), raw);
    }
}
