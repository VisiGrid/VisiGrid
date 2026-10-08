//! Authored-cell comparison and compact guarded-history patches. Derived ASTs
//! and spill state are rebuilt; the workbook fingerprint protects presentation.
use crate::{
    cell::{Cell, CellFormat, CellRef, CellValue, ValueRef},
    formula::parser,
    sheet::{Sheet, SheetId},
};
use std::{borrow::Cow, collections::HashMap};

fn same_metadata(a: CellRef<'_>, b: CellRef<'_>) -> bool {
    (std::ptr::eq(a.format(), b.format())
        || serde_json::to_value(a.format()).unwrap() == serde_json::to_value(b.format()).unwrap())
        && a.comment() == b.comment()
        && a.style_id() == b.style_id()
        && a.frozen_formula() == b.frozen_formula()
}
pub(super) fn same_cells(a: Option<CellRef<'_>>, b: Option<CellRef<'_>>) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(a), Some(b)) => {
            let value = match (a.value(), b.value()) {
                (ValueRef::Empty, ValueRef::Empty) => true,
                (ValueRef::Number(a), ValueRef::Number(b)) => a.to_bits() == b.to_bits(),
                (ValueRef::Text(a), ValueRef::Text(b)) => a == b,
                (ValueRef::Formula { source: a, .. }, ValueRef::Formula { source: b, .. }) => {
                    a == b
                }
                _ => false,
            };
            value && same_metadata(a, b)
        }
        _ => false,
    }
}
fn owned(cell: Option<CellRef<'_>>) -> Option<Cell> {
    cell.map(|c| {
        let mut cell = c.to_cell();
        cell.clear_spill_state();
        cell
    })
}
#[derive(Clone, Debug, serde::Serialize)]
struct Images {
    #[serde(serialize_with = "crate::history_size::serialize_optional_cell")]
    before: Option<Cell>,
    #[serde(serialize_with = "crate::history_size::serialize_optional_cell")]
    after: Option<Cell>,
}
#[derive(Clone, Debug)]
enum Change {
    Formula { before: Box<str>, after: Box<str> },
    Images(Box<Images>),
}
#[derive(Clone, Debug)]
pub(super) struct CellPatch {
    pub sheet: SheetId,
    row: usize,
    col: usize,
    change: Change,
}
impl CellPatch {
    pub(super) fn estimated_history_bytes(&self) -> usize {
        std::mem::size_of::<Self>() + match &self.change {
            Change::Formula { before, after } => before.len().next_multiple_of(16) + after.len().next_multiple_of(16) + 32,
            Change::Images(images) => crate::history_size::serialized_bytes(images),
        }
    }

    pub(super) fn capture(
        sheet: SheetId,
        row: usize,
        col: usize,
        before: Option<CellRef<'_>>,
        after: Option<CellRef<'_>>,
    ) -> Self {
        let change = match (before, after) {
            (Some(b), Some(a)) if same_metadata(b, a) => match (b.value(), a.value()) {
                (ValueRef::Formula { source: b, .. }, ValueRef::Formula { source: a, .. }) => {
                    Change::Formula {
                        before: b.into(),
                        after: a.into(),
                    }
                }
                _ => Change::Images(Box::new(Images {
                    before: owned(before),
                    after: owned(after),
                })),
            },
            _ => Change::Images(Box::new(Images {
                before: owned(before),
                after: owned(after),
            })),
        };
        Self {
            sheet,
            row,
            col,
            change,
        }
    }
    /// Called only after the full authored-workbook fingerprint matches. That
    /// check includes all presentation retained in place by a formula patch.
    pub(super) fn matches(&self, sheet: &Sheet, undo: bool) -> bool {
        let actual = sheet.get_cell_opt(self.row, self.col);
        match &self.change {
            Change::Formula { before, after } => {
                matches!(actual.map(|c|c.value()), Some(ValueRef::Formula { source,.. }) if source == if undo { after.as_ref() } else { before.as_ref() })
            }
            Change::Images(images) => same_cells(
                actual,
                (if undo { &images.after } else { &images.before })
                    .as_ref()
                    .map(CellRef::new),
            ),
        }
    }
    pub(super) fn install(&self, sheet: &mut Sheet, undo: bool) {
        match &self.change {
            Change::Formula { before, after } => {
                let source = if undo { before } else { after };
                // Preserve the Formula variant even for a malformed imported
                // source. Re-parsing is derived state, not an authored change.
                sheet.write_table_header(
                    self.row,
                    self.col,
                    CellValue::Formula {
                        source: source.to_string(),
                        ast: parser::parse(source).ok().map(Box::new),
                    },
                );
                sheet.mark_table_changed();
            }
            Change::Images(images) => sheet.restore_history_cell(
                self.row,
                self.col,
                if undo {
                    images.before.clone()
                } else {
                    images.after.clone()
                },
            ),
        }
    }
}

/// Emit the exact original JSON signature without building a tree or cloning
/// a cell/AST. Format objects keep the canonical key ordering used by Value.
/// The cache borrows stable format addresses only for one immutable sheet scan.
#[derive(Default)]
pub(super) struct SignatureEncoder {
    buffer: Vec<u8>,
    formats: HashMap<*const CellFormat, Vec<u8>>,
    format_bytes: usize,
}
impl SignatureEncoder {
    pub(super) fn encode(&mut self, cell: CellRef<'_>) -> &[u8] {
        self.buffer.clear();
        self.buffer.push(b'[');
        let value = match cell.value() {
            ValueRef::Empty => ("empty", Cow::Borrowed("")),
            ValueRef::Number(n) => ("number", Cow::Owned(n.to_bits().to_string())),
            ValueRef::Text(s) => ("text", Cow::Borrowed(s)),
            ValueRef::Formula { source, .. } => ("formula", Cow::Borrowed(source)),
        };
        serde_json::to_writer(&mut self.buffer, &value).unwrap();
        self.buffer.push(b',');
        let key = cell.format() as *const CellFormat;
        if let Some(bytes) = self.formats.get(&key) {
            self.buffer.extend_from_slice(bytes);
        } else {
            let bytes = serde_json::to_vec(&serde_json::to_value(cell.format()).unwrap()).unwrap();
            self.buffer.extend_from_slice(&bytes);
            if self.formats.len() < 256 && self.format_bytes + bytes.len() <= 1024 * 1024 {
                self.format_bytes += bytes.len();
                self.formats.insert(key, bytes);
            }
        }
        self.buffer.push(b',');
        serde_json::to_writer(
            &mut self.buffer,
            &serde_json::to_value(cell.comment()).unwrap(),
        )
        .unwrap();
        self.buffer.push(b',');
        serde_json::to_writer(&mut self.buffer, &cell.style_id()).unwrap();
        self.buffer.push(b',');
        serde_json::to_writer(&mut self.buffer, &cell.frozen_formula()).unwrap();
        self.buffer.push(b']');
        &self.buffer
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellComment;
    use std::sync::Arc;

    #[test]
    fn borrowed_comparison_matches_original_signature_for_authored_and_derived_state() {
        let mut cells = vec![None, Some(Cell::new())];
        for value in [
            CellValue::Number(0.0),
            CellValue::Number(-0.0),
            CellValue::Number(f64::NAN),
            CellValue::Number(f64::from_bits(f64::NAN.to_bits() + 1)),
            CellValue::Text("escaped \"\n é".into()),
            CellValue::from_input("=(A1+2)*3"),
        ] {
            let mut cell = Cell::new();
            cell.value = value;
            cells.push(Some(cell));
        }
        let formula = cells.last().unwrap().as_ref().unwrap().clone();
        for variant in 0..9 {
            let mut cell = formula.clone();
            match variant {
                0 => cell.set_comment(Some(CellComment {
                    text: "note".into(),
                    author: "writer".into(),
                })),
                1 => cell.set_style_id(Some(12)),
                2 => cell.set_frozen_formula(Some("=A2".into())),
                3 => Arc::make_mut(&mut cell.format).bold = true,
                4 => Arc::make_mut(&mut cell.format).font_size = Some(f32::NAN),
                5 => Arc::make_mut(&mut cell.format).font_size = Some(f32::INFINITY),
                6 => cell.set_spill_parent(Some((3, 4))),
                7 => {
                    if let CellValue::Formula { ast, .. } = &mut cell.value {
                        *ast = None;
                    }
                }
                _ => cell.format = Arc::new((*cell.format).clone()),
            }
            cells.push(Some(cell));
        }
        for (i, a) in cells.iter().enumerate() {
            for (j, b) in cells.iter().enumerate() {
                assert_eq!(
                    same_cells(a.as_ref().map(CellRef::new), b.as_ref().map(CellRef::new)),
                    super::super::signature(a) == super::super::signature(b),
                    "{i}, {j}"
                );
            }
        }
    }

    #[test]
    fn encoding_matches_original_after_format_cache_limit() {
        // Keep every address alive throughout the scan, as an immutable Sheet does.
        let cells: Vec<_> = (0..300)
            .map(|i| {
                let mut cell = Cell::new();
                cell.value = CellValue::Text(format!("row {i}: \"é\"\n\\"));
                Arc::make_mut(&mut cell.format).font_size = Some(i as f32);
                cell.set_comment(Some(CellComment {
                    text: "line\n\"quote\"".into(),
                    author: "é".into(),
                }));
                cell.set_style_id(Some(i));
                cell.set_frozen_formula(Some("='Sheet 1'!$A$2".into()));
                cell
            })
            .collect();
        let mut encoder = SignatureEncoder::default();
        for cell in cells.iter().chain(cells.iter().rev()) {
            assert_eq!(
                encoder.encode(CellRef::new(cell)),
                serde_json::to_vec(&super::super::signature(&Some(cell.clone()))).unwrap()
            );
        }
        assert_eq!(encoder.formats.len(), 256);
        assert!(encoder.format_bytes <= 1024 * 1024);
    }

    #[test]
    fn compact_formula_history_preserves_presentation_and_malformed_formula_variant() {
        let mut before = Cell::new();
        before.value = CellValue::Formula {
            source: "=invalid(".into(),
            ast: None,
        };
        Arc::make_mut(&mut before.format).bold = true;
        before.set_comment(Some(CellComment {
            text: "retain".into(),
            author: "tester".into(),
        }));
        before.set_style_id(Some(42));
        before.set_frozen_formula(Some("=A1".into()));
        let mut after = before.clone();
        after.value = CellValue::from_input("=1+2");
        let patch = CellPatch::capture(
            SheetId(1),
            0,
            0,
            Some(before.as_ref()),
            Some(after.as_ref()),
        );
        assert!(matches!(patch.change, Change::Formula { .. }));
        let mut sheet = Sheet::new(SheetId(1), 10, 10);
        sheet.restore_history_cell(0, 0, Some(after.clone()));
        assert!(patch.matches(&sheet, true));
        patch.install(&mut sheet, true);
        assert!(same_cells(sheet.get_cell_opt(0, 0), Some(before.as_ref())));
        assert!(matches!(
            sheet.get_cell_opt(0, 0).unwrap().value(),
            ValueRef::Formula { ast: None, .. }
        ));
        assert!(patch.matches(&sheet, false));
        patch.install(&mut sheet, false);
        assert!(same_cells(sheet.get_cell_opt(0, 0), Some(after.as_ref())));
        // Presentation changes must use full images instead of source-only patches.
        Arc::make_mut(&mut after.format).italic = true;
        let full = CellPatch::capture(
            SheetId(1),
            0,
            0,
            Some(before.as_ref()),
            Some(after.as_ref()),
        );
        assert!(matches!(full.change, Change::Images(_)));
        full.install(&mut sheet, false);
        assert!(same_cells(sheet.get_cell_opt(0, 0), Some(after.as_ref())));
        full.install(&mut sheet, true);
        assert!(same_cells(sheet.get_cell_opt(0, 0), Some(before.as_ref())));
    }
}
