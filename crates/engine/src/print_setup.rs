//! Sheet-owned presentation settings. These do not affect formula results or semantic hashes.
use crate::sheet::{NUM_COLS, NUM_ROWS};
use serde::{Deserialize, Serialize};

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum PrintPaper {
    #[default]
    A4,
    Letter,
    Legal,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum PrintScale {
    #[default]
    FitColumns,
    Actual,
    FitSheet,
}

/// Inclusive, zero-based source coordinates, independent of filtering and sorting.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct PrintArea {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct PrintRows {
    pub start: usize,
    pub end: usize,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct PrintSetup {
    pub paper: PrintPaper,
    pub landscape: bool,
    pub scale: PrintScale,
    pub gridlines: bool,
    pub page_numbers: bool,
    pub area: Option<PrintArea>,
    /// Must resolve to a leading band of the captured print area. Hidden rows are omitted.
    pub repeat_rows: Option<PrintRows>,
}

impl PrintSetup {
    pub fn validate(&self) -> Result<(), String> {
        if self.area.is_some_and(|a| {
            a.start_row > a.end_row
                || a.start_col > a.end_col
                || a.end_row >= NUM_ROWS
                || a.end_col >= NUM_COLS
        }) {
            return Err("Invalid saved print area.".into());
        }
        if self
            .repeat_rows
            .is_some_and(|r| r.start > r.end || r.end >= NUM_ROWS)
        {
            return Err("Invalid repeated header rows.".into());
        }
        Ok(())
    }

    pub fn is_default(&self) -> bool {
        self == &Self::default()
    }

    pub(crate) fn adjust(&mut self, rows: bool, at: usize, count: usize, delete: bool) {
        if count == 0 {
            return;
        }
        let limit = if rows { NUM_ROWS } else { NUM_COLS };
        if let Some(mut area) = self.area {
            let (start, end) = if rows {
                (area.start_row, area.end_row)
            } else {
                (area.start_col, area.end_col)
            };
            self.area = adjusted(start, end, at, count, delete, limit).map(|(start, end)| {
                if rows {
                    area.start_row = start;
                    area.end_row = end;
                } else {
                    area.start_col = start;
                    area.end_col = end;
                }
                area
            });
        }
        if rows {
            self.repeat_rows = self.repeat_rows.and_then(|r| {
                adjusted(r.start, r.end, at, count, delete, limit)
                    .map(|(start, end)| PrintRows { start, end })
            });
        }
    }
}

fn adjusted(
    start: usize,
    end: usize,
    at: usize,
    count: usize,
    delete: bool,
    limit: usize,
) -> Option<(usize, usize)> {
    let (start, end) = if delete {
        let lo = start - start.saturating_sub(at).min(count);
        let hi = end.saturating_add(1) - end.saturating_add(1).saturating_sub(at).min(count);
        if lo >= hi {
            return None;
        }
        (lo, hi - 1)
    } else {
        (
            if at <= start {
                start.saturating_add(count)
            } else {
                start
            },
            if at <= end {
                end.saturating_add(count)
            } else {
                end
            },
        )
    };
    (start < limit).then_some((start, end.min(limit - 1)))
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn ranges_follow_inserts_deletes_and_disappear_when_fully_deleted() {
        let mut setup = PrintSetup {
            area: Some(PrintArea {
                start_row: 3,
                end_row: 20,
                start_col: 2,
                end_col: 6,
            }),
            repeat_rows: Some(PrintRows { start: 3, end: 5 }),
            ..Default::default()
        };
        setup.adjust(true, 0, 2, false);
        assert_eq!(setup.repeat_rows, Some(PrintRows { start: 5, end: 7 }));
        setup.adjust(true, 6, 2, false);
        assert_eq!(setup.repeat_rows, Some(PrintRows { start: 5, end: 9 }));
        setup.adjust(true, 4, 3, true);
        assert_eq!(setup.repeat_rows, Some(PrintRows { start: 4, end: 6 }));
        setup.adjust(false, 3, 2, false);
        assert_eq!(setup.area.unwrap().end_col, 8);
        setup.adjust(false, 0, 4, true);
        assert_eq!(
            (setup.area.unwrap().start_col, setup.area.unwrap().end_col),
            (0, 4)
        );
        setup.adjust(true, 4, 3, true);
        assert_eq!(setup.repeat_rows, None);
        setup.adjust(false, 0, 5, true);
        assert_eq!(setup.area, None);
    }
    #[test]
    fn invalid_ranges_are_rejected_and_old_serialization_defaults() {
        assert_eq!(
            serde_json::from_str::<PrintSetup>("{}").unwrap(),
            PrintSetup::default()
        );
        let setup = PrintSetup {
            repeat_rows: Some(PrintRows {
                start: 0,
                end: usize::MAX,
            }),
            ..Default::default()
        };
        assert!(setup.validate().is_err());
    }
}
