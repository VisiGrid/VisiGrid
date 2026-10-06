//! Conservative retained-payload estimates for bounded undo history. This is
//! accounting, not an allocator measurement; shared payloads may be charged twice.
use std::fmt::{Debug, Write};

pub fn estimated_debug_bytes(value: &impl Debug) -> usize {
    struct Counter(usize);
    impl Write for Counter {
        fn write_str(&mut self, text: &str) -> std::fmt::Result {
            self.0 = self.0.saturating_add(text.len().saturating_mul(2));
            if self.0 > 1024 * 1024 * 1024 {
                Err(std::fmt::Error)
            } else {
                Ok(())
            }
        }
    }
    let mut counter = Counter(4096);
    let _ = write!(counter, "{value:?}");
    counter.0
}
