//! Exact calculation context carried beside a sequenced operation. Seeds are
//! decimal strings on the wire so JavaScript never rounds a u64 before Rust
//! sees it. Parsing completes before a caller changes its replica.
use serde_json::Value;
use visigrid_engine::RecalcClock;

pub fn install_frame_clock(
    client: &mut crate::client::Client,
    frame: &Value,
) -> Result<(), String> {
    let Some(value) = frame.get("clock") else {
        return Ok(());
    };
    let clock = parse_clock(value)?;
    client.wb.ensure_writable()?;
    client.confirmed.ensure_writable()?;
    client.wb.set_recalc_clock(Some(clock));
    client.confirmed.set_recalc_clock(Some(clock));
    // A value edit can advance NOW/RAND in cells elsewhere in the workbook.
    // Rebase starts from confirmed, so it must include that recalculation.
    if client.confirmed.volatile_cell_count() > 0 || client.wb.volatile_cell_count() > 0 {
        client.confirmed.recompute_full_ordered();
        client.wb.recompute_full_ordered();
        client.record_calculation_refresh();
    }
    Ok(())
}

pub fn parse_clock(value: &Value) -> Result<RecalcClock, String> {
    let now_ms = value
        .get("now_ms")
        .and_then(Value::as_i64)
        .ok_or("Calculation clock needs integer now_ms")?;
    let utc_offset_seconds = value
        .get("utc_offset_seconds")
        .and_then(Value::as_i64)
        .filter(|offset| (-86400..=86400).contains(offset))
        .ok_or("Calculation clock needs a valid timezone offset")?;
    let seed = value
        .get("seed")
        .and_then(Value::as_str)
        .filter(|seed| !seed.is_empty() && seed.bytes().all(|c| c.is_ascii_digit()))
        .and_then(|seed| seed.parse::<u64>().ok())
        .ok_or("Calculation clock needs an exact decimal seed")?;
    Ok(RecalcClock {
        now_ms: Some(now_ms),
        utc_offset_seconds: Some(utc_offset_seconds),
        seed: Some(seed),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn clock_recalculation_requests_a_repaint_for_volatile_cells() {
        let mut client = crate::client::Client::new(1);
        client.wb.set_cell_value_tracked_at(0, 0, 0, "=RAND()");
        client.confirmed = client.wb.clone();
        client.record_changes();
        let frame = json!({"clock":{"now_ms":1,"utc_offset_seconds":0,"seed":"42"}});
        install_frame_clock(&mut client, &frame).unwrap();
        assert!(client.take_changes().full);
        let mut ordinary = crate::client::Client::new(1);
        ordinary.record_changes();
        install_frame_clock(&mut ordinary, &frame).unwrap();
        assert!(!ordinary.take_changes().full);
    }

    #[test]
    fn retains_seed_beyond_javascript_integer_precision() {
        let clock = parse_clock(&json!({"now_ms":1791028800123_i64,
            "utc_offset_seconds":-18000,"seed":"18446744073709551615"}))
        .unwrap();
        assert_eq!(clock.seed, Some(u64::MAX));
        assert_eq!(clock.now_ms, Some(1791028800123));
        assert_eq!(clock.utc_offset_seconds, Some(-18000));
    }

    #[test]
    fn refuses_rounded_seeds_and_incomplete_contexts() {
        for value in [
            json!({"now_ms":1,"utc_offset_seconds":0,"seed":9007199254740993_u64}),
            json!({"now_ms":1,"utc_offset_seconds":0,"seed":"18446744073709551616"}),
            json!({"now_ms":1,"utc_offset_seconds":90000,"seed":"1"}),
            json!({"now_ms":1,"seed":"1"}),
            json!({"now_ms":1.5,"utc_offset_seconds":0,"seed":"1"}),
        ] {
            assert!(parse_clock(&value).is_err(), "accepted {value}");
        }
    }
}
