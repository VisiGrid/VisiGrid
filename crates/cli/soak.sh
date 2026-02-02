#!/bin/bash
CASES="${1:-10000}"
echo "Soaking with PROPTEST_CASES=$CASES"
PROPTEST_CASES="$CASES" cargo test -p visigrid-cli --test property_diff --release -- --nocapture
