//! Exact calculation context carried beside a sequenced operation. Seeds are
//! decimal strings on the wire so JavaScript never rounds a u64 before Rust
//! sees it. Parsing completes before a caller changes its replica.
use serde_json::Value;
use visigrid_engine::RecalcClock;

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
