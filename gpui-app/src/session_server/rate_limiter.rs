//! Token bucket rate limiter for session server.
//!
//! Limits are ops-based, not request-based. A single message with 50k ops
//! counts as 50k against the rate limit.
//!
//! Design decisions:
//! - Per-connection bucket (keyed by connection, not client string)
//! - If message exceeds available tokens: reject immediately (no partial consume)
//! - Includes retry_after_ms for well-behaved agents
//! - Clock trait for deterministic testing

use std::time::{Duration, Instant};

/// Clock abstraction for testability.
/// In production, uses std::time::Instant.
/// In tests, can be mocked for deterministic behavior.
pub trait Clock: Send + Sync {
    fn now(&self) -> Instant;
}

/// Real clock using std::time::Instant.
#[derive(Clone, Copy, Default)]
pub struct RealClock;

impl Clock for RealClock {
    fn now(&self) -> Instant {
        Instant::now()
    }
}

/// Rate limiter configuration.
#[derive(Debug, Clone, Copy)]
pub struct RateLimiterConfig {
    /// Maximum burst capacity (tokens).
    pub burst_ops: u32,
    /// Refill rate (tokens per second).
    pub ops_per_sec: u32,
    /// Cost for apply_ops is ops.len() (variable)
    /// Cost for inspect message.
    pub inspect_cost: u32,
    /// Cost for subscribe message.
    pub subscribe_cost: u32,
    /// Cost for unsubscribe message.
    pub unsubscribe_cost: u32,
    /// Cost for ping message.
    pub ping_cost: u32,
}

impl Default for RateLimiterConfig {
    fn default() -> Self {
        Self {
            burst_ops: 50_000,
            ops_per_sec: 20_000,
            inspect_cost: 10,
            subscribe_cost: 10,
            unsubscribe_cost: 1,
            ping_cost: 1,
        }
    }
}

/// Error returned when rate limit is exceeded.
#[derive(Debug, Clone, Copy)]
pub struct RateLimitedError {
    /// Milliseconds until enough tokens are available.
    pub retry_after_ms: u64,
    /// Tokens requested.
    pub requested: u32,
    /// Tokens available.
    pub available: u32,
}

/// Token bucket rate limiter.
///
/// Thread-safe for use across connection handler threads.
pub struct RateLimiter<C: Clock = RealClock> {
    /// Current token count (as f64 for fractional refill).
    tokens: f64,
    /// Maximum tokens (burst capacity).
    max_tokens: u32,
    /// Tokens added per second.
    refill_rate: u32,
    /// Last refill timestamp.
    last_refill: Instant,
    /// Clock for time measurement.
    clock: C,
    /// Configuration for message costs.
    config: RateLimiterConfig,
}

impl RateLimiter<RealClock> {
    /// Create a new rate limiter with default clock.
    pub fn new(config: RateLimiterConfig) -> Self {
        Self::with_clock(config, RealClock)
    }
}

impl<C: Clock> RateLimiter<C> {
    /// Create a new rate limiter with custom clock.
    pub fn with_clock(config: RateLimiterConfig, clock: C) -> Self {
        let now = clock.now();
        Self {
            tokens: config.burst_ops as f64,
            max_tokens: config.burst_ops,
            refill_rate: config.ops_per_sec,
            last_refill: now,
            clock,
            config,
        }
    }

    /// Refill tokens based on elapsed time.
    fn refill(&mut self) {
        let now = self.clock.now();
        let elapsed = now.duration_since(self.last_refill);

        // Guard against clock going backwards (rare but possible)
        if elapsed.is_zero() {
            return;
        }

        let elapsed_secs = elapsed.as_secs_f64();
        let refill_amount = elapsed_secs * self.refill_rate as f64;

        self.tokens = (self.tokens + refill_amount).min(self.max_tokens as f64);
        self.last_refill = now;
    }

    /// Try to consume tokens for an apply_ops request.
    /// Cost is the number of ops in the batch.
    pub fn try_apply_ops(&mut self, ops_count: usize) -> Result<(), RateLimitedError> {
        self.try_consume(ops_count as u32)
    }

    /// Try to consume tokens for an inspect request.
    pub fn try_inspect(&mut self) -> Result<(), RateLimitedError> {
        self.try_consume(self.config.inspect_cost)
    }

    /// Try to consume tokens for a subscribe request.
    pub fn try_subscribe(&mut self) -> Result<(), RateLimitedError> {
        self.try_consume(self.config.subscribe_cost)
    }

    /// Try to consume tokens for an unsubscribe request.
    pub fn try_unsubscribe(&mut self) -> Result<(), RateLimitedError> {
        self.try_consume(self.config.unsubscribe_cost)
    }

    /// Try to consume tokens for a ping request.
    pub fn try_ping(&mut self) -> Result<(), RateLimitedError> {
        self.try_consume(self.config.ping_cost)
    }

    /// Try to consume the specified number of tokens.
    ///
    /// If tokens are available: consume and return Ok.
    /// If not enough tokens: return Err with retry_after_ms.
    ///
    /// IMPORTANT: This is all-or-nothing. If request exceeds available,
    /// nothing is consumed.
    fn try_consume(&mut self, cost: u32) -> Result<(), RateLimitedError> {
        self.refill();

        let available = self.tokens.floor() as u32;

        if cost == 0 {
            return Ok(());
        }

        if cost > self.max_tokens {
            // Request exceeds burst capacity - will never succeed
            return Err(RateLimitedError {
                retry_after_ms: u64::MAX, // Never
                requested: cost,
                available,
            });
        }

        if (cost as f64) <= self.tokens {
            self.tokens -= cost as f64;
            Ok(())
        } else {
            // Calculate time until enough tokens are available
            let tokens_needed = cost as f64 - self.tokens;
            let seconds_to_wait = tokens_needed / self.refill_rate as f64;
            let retry_after_ms = (seconds_to_wait * 1000.0).ceil() as u64;

            Err(RateLimitedError {
                retry_after_ms,
                requested: cost,
                available,
            })
        }
    }

    /// Get current available tokens (for debugging/monitoring).
    pub fn available_tokens(&mut self) -> u32 {
        self.refill();
        self.tokens.floor() as u32
    }

    /// Get the configuration.
    pub fn config(&self) -> &RateLimiterConfig {
        &self.config
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::{AtomicU64, Ordering};
    use std::sync::Arc;

    /// Mock clock for deterministic testing.
    /// Uses an atomic offset from a base instant for thread safety.
    struct MockClock {
        base: Instant,
        offset_nanos: Arc<AtomicU64>,
    }

    impl MockClock {
        fn new() -> (Self, Arc<AtomicU64>) {
            let offset_nanos = Arc::new(AtomicU64::new(0));
            (
                Self {
                    base: Instant::now(),
                    offset_nanos: offset_nanos.clone(),
                },
                offset_nanos,
            )
        }

        fn advance(offset_nanos: &Arc<AtomicU64>, duration: Duration) {
            offset_nanos.fetch_add(duration.as_nanos() as u64, Ordering::SeqCst);
        }
    }

    impl Clock for MockClock {
        fn now(&self) -> Instant {
            let offset = Duration::from_nanos(self.offset_nanos.load(Ordering::SeqCst));
            self.base + offset
        }
    }

    #[test]
    fn test_burst_boundary_exact() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, _offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // Exact burst should pass
        assert!(limiter.try_apply_ops(100).is_ok());
        assert_eq!(limiter.available_tokens(), 0);
    }

    #[test]
    fn test_burst_boundary_plus_one_fails() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, _offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // One more than burst should fail
        let err = limiter.try_apply_ops(101).unwrap_err();
        assert_eq!(err.requested, 101);
        assert_eq!(err.available, 100);
        // Should never succeed (exceeds max capacity)
        assert_eq!(err.retry_after_ms, u64::MAX);
    }

    #[test]
    fn test_refill_after_time() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // Consume all tokens
        assert!(limiter.try_apply_ops(100).is_ok());
        assert_eq!(limiter.available_tokens(), 0);

        // Request should fail
        assert!(limiter.try_apply_ops(10).is_err());

        // Advance time by 1 second (10 tokens refill)
        MockClock::advance(&offset, Duration::from_secs(1));

        // Now should have 10 tokens
        assert_eq!(limiter.available_tokens(), 10);
        assert!(limiter.try_apply_ops(10).is_ok());
    }

    #[test]
    fn test_retry_after_calculation() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, _offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // Consume all tokens
        assert!(limiter.try_apply_ops(100).is_ok());

        // Request 20 tokens (need 2 seconds to refill)
        let err = limiter.try_apply_ops(20).unwrap_err();
        assert_eq!(err.retry_after_ms, 2000);
    }

    #[test]
    fn test_costing_apply_ops() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, _offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // apply_ops costs ops.len()
        assert!(limiter.try_apply_ops(30).is_ok());
        assert_eq!(limiter.available_tokens(), 70);

        assert!(limiter.try_apply_ops(50).is_ok());
        assert_eq!(limiter.available_tokens(), 20);
    }

    #[test]
    fn test_costing_inspect() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            inspect_cost: 10,
            ..Default::default()
        };
        let (clock, _offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // inspect costs 10
        assert!(limiter.try_inspect().is_ok());
        assert_eq!(limiter.available_tokens(), 90);
    }

    #[test]
    fn test_costing_ping() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ping_cost: 1,
            ..Default::default()
        };
        let (clock, _offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // ping costs 1
        assert!(limiter.try_ping().is_ok());
        assert_eq!(limiter.available_tokens(), 99);
    }

    #[test]
    fn test_no_partial_consume() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, _offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // Consume 80 tokens
        assert!(limiter.try_apply_ops(80).is_ok());
        assert_eq!(limiter.available_tokens(), 20);

        // Try to consume 30 (more than available)
        // Should fail AND not consume anything
        assert!(limiter.try_apply_ops(30).is_err());
        assert_eq!(limiter.available_tokens(), 20); // Still 20, not partially consumed
    }

    #[test]
    fn test_zero_cost_always_succeeds() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, _offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // Drain tokens
        assert!(limiter.try_apply_ops(100).is_ok());

        // Zero-cost should still succeed
        assert!(limiter.try_apply_ops(0).is_ok());
    }

    #[test]
    fn test_refill_caps_at_burst() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // Start at max
        assert_eq!(limiter.available_tokens(), 100);

        // Advance a lot of time
        MockClock::advance(&offset, Duration::from_secs(1000));

        // Should still cap at burst
        assert_eq!(limiter.available_tokens(), 100);
    }

    #[test]
    fn test_fractional_refill_accumulates() {
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // Consume all
        assert!(limiter.try_apply_ops(100).is_ok());

        // Advance 100ms (1 token)
        MockClock::advance(&offset, Duration::from_millis(100));
        assert_eq!(limiter.available_tokens(), 1);

        // Advance another 100ms
        MockClock::advance(&offset, Duration::from_millis(100));
        assert_eq!(limiter.available_tokens(), 2);
    }

    #[test]
    fn test_clock_backwards_clamped() {
        // This test verifies behavior if time goes backwards
        // We can't easily simulate Instant going backwards, but we verify
        // that zero-duration elapsed is handled correctly
        let config = RateLimiterConfig {
            burst_ops: 100,
            ops_per_sec: 10,
            ..Default::default()
        };
        let (clock, _offset) = MockClock::new();
        let mut limiter = RateLimiter::with_clock(config, clock);

        // Consume some
        assert!(limiter.try_apply_ops(50).is_ok());
        assert_eq!(limiter.available_tokens(), 50);

        // Call refill again without advancing time - should not crash or overflow
        assert_eq!(limiter.available_tokens(), 50);
    }
}
