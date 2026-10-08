//! Retained history payload accounting. Count a binary serialization without
//! allocating an output buffer (including maps with non-string keys). This is
//! a payload budget, not an allocator/RSS measurement: shared data is charged
//! per owner. Snapshot estimates use cell counts rather than formatting ASTs.
use serde::{Serialize, ser::*};

pub fn serialized_bytes(value: &impl Serialize) -> usize {
    let mut counter = Counter(0, false);
    // A failing Serialize must not panic while history is recorded. Keep the
    // partial count and a small pad so the budget still sees the value.
    if value.serialize(&mut counter).is_err() {
        return counter.0.saturating_add(1024);
    }
    counter.0
}

pub fn sheet_bytes(sheet: &crate::sheet::Sheet) -> usize {
    // Snapshot history is rare. Walk populated cells so a multi-megabyte
    // value counts toward the cap; a flat 256 bytes per cell does not.
    std::mem::size_of_val(sheet).saturating_add(sheet.cells_iter().map(|(_, cell)| {
        std::mem::size_of::<crate::cell::Cell>()
            + value_ref_bytes(&cell.value())
            + serialized_bytes(&(cell.format(), cell.comment(), cell.frozen_formula()))
    }).sum::<usize>())
}
pub fn workbook_bytes(wb: &crate::workbook::Workbook) -> usize {
    std::mem::size_of_val(wb) + wb.sheets().iter().map(sheet_bytes).sum::<usize>()
}
struct Counter(usize, bool);
impl Counter { fn add(&mut self, bytes: usize) { self.0 = self.0.saturating_add(bytes); } }
impl<'a> serde::Serializer for &'a mut Counter {
    type Ok = ();
    type Error = serde::de::value::Error;
    type SerializeSeq = Self;
    type SerializeTuple = Self;
    type SerializeTupleStruct = Self;
    type SerializeTupleVariant = Self;
    type SerializeMap = Self;
    type SerializeStruct = Self;
    type SerializeStructVariant = Self;
    fn serialize_bool(self, _: bool) -> Result<(), Self::Error> { self.add(std::mem::size_of::<bool>()); Ok(()) }
    fn serialize_i8(self, _: i8) -> Result<(), Self::Error> { self.add(std::mem::size_of::<i8>()); Ok(()) }
    fn serialize_i16(self, _: i16) -> Result<(), Self::Error> { self.add(std::mem::size_of::<i16>()); Ok(()) }
    fn serialize_i32(self, _: i32) -> Result<(), Self::Error> { self.add(std::mem::size_of::<i32>()); Ok(()) }
    fn serialize_i64(self, _: i64) -> Result<(), Self::Error> { self.add(std::mem::size_of::<i64>()); Ok(()) }
    fn serialize_i128(self, _: i128) -> Result<(), Self::Error> { self.add(std::mem::size_of::<i128>()); Ok(()) }
    fn serialize_u8(self, _: u8) -> Result<(), Self::Error> { self.add(std::mem::size_of::<u8>()); Ok(()) }
    fn serialize_u16(self, _: u16) -> Result<(), Self::Error> { self.add(std::mem::size_of::<u16>()); Ok(()) }
    fn serialize_u32(self, _: u32) -> Result<(), Self::Error> { self.add(std::mem::size_of::<u32>()); Ok(()) }
    fn serialize_u64(self, v: u64) -> Result<(), Self::Error> { let bytes = if self.1 { self.1 = false; v as usize } else { 8 }; self.add(bytes); Ok(()) }
    fn serialize_u128(self, _: u128) -> Result<(), Self::Error> { self.add(std::mem::size_of::<u128>()); Ok(()) }
    fn serialize_f32(self, _: f32) -> Result<(), Self::Error> { self.add(std::mem::size_of::<f32>()); Ok(()) }
    fn serialize_f64(self, _: f64) -> Result<(), Self::Error> { self.add(std::mem::size_of::<f64>()); Ok(()) }
    fn serialize_char(self, _: char) -> Result<(), Self::Error> { self.add(std::mem::size_of::<char>()); Ok(()) }
    fn serialize_str(self, v: &str) -> Result<(), Self::Error> { self.add(24 + v.len()); Ok(()) }
    fn serialize_bytes(self, v: &[u8]) -> Result<(), Self::Error> { self.add(24 + v.len()); Ok(()) }
    fn serialize_none(self) -> Result<(), Self::Error> { self.add(1); Ok(()) }
    fn serialize_some<T: ?Sized + Serialize>(self, v: &T) -> Result<(), Self::Error> { self.add(1); v.serialize(self) }
    fn serialize_unit(self) -> Result<(), Self::Error> { Ok(()) }
    fn serialize_unit_struct(self, _: &'static str) -> Result<(), Self::Error> { Ok(()) }
    fn serialize_unit_variant(self, _: &'static str, _: u32, _: &'static str) -> Result<(), Self::Error> { self.add(4); Ok(()) }
    fn serialize_newtype_struct<T: ?Sized + Serialize>(self, name: &'static str, v: &T) -> Result<(), Self::Error> { self.1 = name == "HistoryPayloadBytes"; v.serialize(self) }
    fn serialize_newtype_variant<T: ?Sized + Serialize>(self, _: &'static str, _: u32, _: &'static str, v: &T) -> Result<(), Self::Error> { self.add(4); v.serialize(self) }
    fn serialize_seq(self, _: Option<usize>) -> Result<Self, Self::Error> { self.add(24); Ok(self) }
    fn serialize_tuple(self, _: usize) -> Result<Self, Self::Error> { Ok(self) }
    fn serialize_tuple_struct(self, _: &'static str, _: usize) -> Result<Self, Self::Error> { Ok(self) }
    fn serialize_tuple_variant(self, _: &'static str, _: u32, _: &'static str, _: usize) -> Result<Self, Self::Error> { self.add(4); Ok(self) }
    fn serialize_map(self, len: Option<usize>) -> Result<Self, Self::Error> { self.add(24 + len.unwrap_or(0).saturating_mul(16)); Ok(self) }
    fn serialize_struct(self, _: &'static str, _: usize) -> Result<Self, Self::Error> { Ok(self) }
    fn serialize_struct_variant(self, _: &'static str, _: u32, _: &'static str, _: usize) -> Result<Self, Self::Error> { self.add(4); Ok(self) }
}
impl SerializeSeq for &mut Counter {
    type Ok = ();
    type Error = serde::de::value::Error;
    fn serialize_element<T: ?Sized + Serialize>(&mut self, v: &T) -> Result<(), Self::Error> { v.serialize(&mut **self) }
    fn end(self) -> Result<(), Self::Error> { Ok(()) }
}
impl SerializeTuple for &mut Counter {
    type Ok = ();
    type Error = serde::de::value::Error;
    fn serialize_element<T: ?Sized + Serialize>(&mut self, v: &T) -> Result<(), Self::Error> { v.serialize(&mut **self) }
    fn end(self) -> Result<(), Self::Error> { Ok(()) }
}
impl SerializeTupleStruct for &mut Counter {
    type Ok = ();
    type Error = serde::de::value::Error;
    fn serialize_field<T: ?Sized + Serialize>(&mut self, v: &T) -> Result<(), Self::Error> { v.serialize(&mut **self) }
    fn end(self) -> Result<(), Self::Error> { Ok(()) }
}
impl SerializeTupleVariant for &mut Counter {
    type Ok = ();
    type Error = serde::de::value::Error;
    fn serialize_field<T: ?Sized + Serialize>(&mut self, v: &T) -> Result<(), Self::Error> { v.serialize(&mut **self) }
    fn end(self) -> Result<(), Self::Error> { Ok(()) }
}
impl SerializeStruct for &mut Counter {
    type Ok = ();
    type Error = serde::de::value::Error;
    fn serialize_field<T: ?Sized + Serialize>(&mut self, _: &'static str, v: &T) -> Result<(), Self::Error> { v.serialize(&mut **self) }
    fn end(self) -> Result<(), Self::Error> { Ok(()) }
}
impl SerializeStructVariant for &mut Counter {
    type Ok = ();
    type Error = serde::de::value::Error;
    fn serialize_field<T: ?Sized + Serialize>(&mut self, _: &'static str, v: &T) -> Result<(), Self::Error> { v.serialize(&mut **self) }
    fn end(self) -> Result<(), Self::Error> { Ok(()) }
}
impl SerializeMap for &mut Counter {
    type Ok = ();
    type Error = serde::de::value::Error;
    fn serialize_key<T: ?Sized + Serialize>(&mut self, v: &T) -> Result<(), Self::Error> { v.serialize(&mut **self) }
    fn serialize_value<T: ?Sized + Serialize>(&mut self, v: &T) -> Result<(), Self::Error> { v.serialize(&mut **self) }
    fn end(self) -> Result<(), Self::Error> { Ok(()) }
}

/// A counted opaque payload for nested history commits.
pub fn serialize_count<S: serde::Serializer>(bytes: usize, serializer: S) -> Result<S::Ok, S::Error> {
    // Size serializer recognizes this marker. Not an on-disk representation.
    serializer.serialize_newtype_struct("HistoryPayloadBytes", &(bytes as u64))
}

fn ast_bytes(expr: &crate::formula::parser::ParsedExpr) -> usize {
    use crate::{formula::parser::Expr, sheet::UnboundSheetRef};
    std::mem::size_of_val(expr) + match expr {
        Expr::Text(s) | Expr::NamedRange(s) | Expr::ReferenceError(s) => s.capacity(),
        Expr::CellRef { sheet, .. } | Expr::Range { sheet, .. } | Expr::WholeRange { sheet, .. } => match sheet { UnboundSheetRef::Named(s) => s.capacity(), _ => 0 },
        Expr::Function { name, args } => name.capacity() + args.iter().map(ast_bytes).sum::<usize>()
            + (args.capacity() - args.len()) * std::mem::size_of_val(expr),
        Expr::BinaryOp { left, right, .. } => ast_bytes(left) + ast_bytes(right),
        Expr::StructuredRef(r) => r.table.as_ref().map_or(0, String::capacity) + r.columns.as_ref().map_or(0, |(a,b)| a.capacity()+b.capacity()),
        _ => 0,
    }
}
fn value_ref_bytes(value: &crate::cell::ValueRef<'_>) -> usize {
    use crate::cell::ValueRef;
    std::mem::size_of::<crate::cell::CellValue>() + match value {
        ValueRef::Text(s) => s.len(),
        ValueRef::Formula { source, ast } => source.len() + ast.map_or(0, ast_bytes),
        _ => 0,
    }
}
fn value_bytes(value: &crate::cell::CellValue) -> usize {
    use crate::cell::CellValue;
    std::mem::size_of_val(value) + match value {
        CellValue::Text(s) => s.capacity(),
        CellValue::Formula { source, ast } => source.capacity() + ast.as_ref().map_or(0, |a| ast_bytes(a)),
        _ => 0,
    }
}
pub fn cell_bytes(cell: &crate::cell::Cell) -> usize {
    std::mem::size_of_val(cell) + value_bytes(&cell.value) + serialized_bytes(&(&cell.format, cell.comment(), cell.frozen_formula()))
}
pub fn serialize_value<S: serde::Serializer>(value: &crate::cell::CellValue, serializer: S) -> Result<S::Ok, S::Error> { serialize_count(value_bytes(value), serializer) }
pub fn serialize_cell<S: serde::Serializer>(cell: &crate::cell::Cell, serializer: S) -> Result<S::Ok, S::Error> { serialize_count(cell_bytes(cell), serializer) }
pub fn serialize_optional_cell<S: serde::Serializer>(cell: &Option<crate::cell::Cell>, serializer: S) -> Result<S::Ok, S::Error> { serialize_count(cell.as_ref().map_or(0, cell_bytes), serializer) }

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sheet_bytes_counts_large_cell_text() {
        let mut sheet = crate::sheet::Sheet::new(crate::sheet::SheetId(1), 10, 4);
        sheet.set_text(0, 0, &"x".repeat(100_000));
        assert!(sheet_bytes(&sheet) > 100_000);
    }
}
