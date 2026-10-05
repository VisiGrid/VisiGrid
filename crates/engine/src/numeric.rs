//! Summing many numbers without float drift.

/// A running sum with Neumaier compensation: the rounding error of each
/// addition is kept and added back at the end, so 300,000 two-decimal
/// amounts total to .09, not .08999842. Plain summation loses a little on
/// every addition and the losses grow with the count; this keeps the result
/// within one rounding of the exact sum of the values as stored.
///
/// SUM, AVERAGE, SUMIF(S), AVERAGEIF(S), pivot sums and recipe Group by all
/// use it, so a total agrees wherever it is computed.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct Sum {
    sum: f64,
    compensation: f64,
}

impl Sum {
    pub fn add(&mut self, x: f64) {
        let t = self.sum + x;
        // Infinities and NaN take over the sum; compensating them would
        // only turn an infinity into NaN
        if t.is_finite() {
            self.compensation += if self.sum.abs() >= x.abs() { (self.sum - t) + x } else { (x - t) + self.sum };
        }
        self.sum = t;
    }

    /// Fold another partial sum in (pivot groups merging).
    pub fn merge(&mut self, other: Sum) {
        self.add(other.sum);
        self.compensation += other.compensation;
    }

    pub fn value(self) -> f64 {
        if self.sum.is_finite() { self.sum + self.compensation } else { self.sum }
    }
}

impl std::ops::AddAssign<f64> for Sum {
    fn add_assign(&mut self, x: f64) {
        self.add(x);
    }
}

impl std::ops::AddAssign<&f64> for Sum {
    fn add_assign(&mut self, x: &f64) {
        self.add(*x);
    }
}

/// The compensated sum of `values`.
pub fn sum<'a>(values: impl IntoIterator<Item = &'a f64>) -> f64 {
    let mut s = Sum::default();
    for v in values {
        s.add(*v);
    }
    s.value()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn many_cents_total_exactly() {
        // Varied two-decimal amounts, as in an export; the exact total in cents
        let mut seed = 1u64;
        let cents: Vec<i64> = (0..300_000)
            .map(|_| {
                seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1442695040888963407);
                (seed >> 33) as i64 % 500_000
            })
            .collect();
        let values: Vec<f64> = cents.iter().map(|c| *c as f64 / 100.0).collect();
        let exact = cents.iter().sum::<i64>() as f64 / 100.0;
        let naive: f64 = values.iter().sum();
        assert_eq!(sum(&values), exact);
        assert_ne!(naive, exact, "the plain sum drifts (else this test proves nothing)");
        assert_eq!(sum(&[0.1, 0.2, 0.3]), 0.6);
    }

    #[test]
    fn infinities_and_merges() {
        assert_eq!(sum(&[1.0, f64::INFINITY, 2.0]), f64::INFINITY);
        assert!(sum(&[f64::INFINITY, f64::NEG_INFINITY]).is_nan());
        let (mut a, mut b) = (Sum::default(), Sum::default());
        for i in 0..1000 {
            a += 0.1;
            b += i as f64 * 0.01;
        }
        a.merge(b);
        let all: Vec<f64> = (0..1000).map(|_| 0.1).chain((0..1000).map(|i| i as f64 * 0.01)).collect();
        assert_eq!(a.value(), sum(&all));
        assert_eq!(a.value(), 100.0 + 4995.0);
    }
}
