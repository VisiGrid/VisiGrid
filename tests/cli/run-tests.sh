#!/bin/bash
# Golden test runner for VisiGrid CLI
# Usage: ./run-tests.sh [path/to/visigrid]

set +e  # Don't exit on error so we can see all test results

VISIGRID="${1:-../../target/release/visigrid}"
TESTS_DIR="$(dirname "$0")"
PASSED=0
FAILED=0

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
NC='\033[0m' # No Color

run_test() {
    local test_dir="$1"
    local test_name="${test_dir#$TESTS_DIR/}"

    # Read args (one per line)
    if [[ ! -f "$test_dir/args.txt" ]]; then
        echo -e "${RED}SKIP${NC} $test_name (no args.txt)"
        return
    fi

    # Read args into array
    local -a args=()
    while IFS= read -r line || [[ -n "$line" ]]; do
        [[ -n "$line" ]] && args+=("$line")
    done < "$test_dir/args.txt"

    # Determine input file
    local input_file=""
    for f in "$test_dir"/input.*; do
        if [[ -f "$f" ]]; then
            input_file="$f"
            break
        fi
    done

    # Run command
    local actual_stdout=$(mktemp)
    local actual_stderr=$(mktemp)
    local actual_exit=0

    if [[ -n "$input_file" ]]; then
        "$VISIGRID" "${args[@]}" < "$input_file" > "$actual_stdout" 2> "$actual_stderr" || actual_exit=$?
    else
        "$VISIGRID" "${args[@]}" > "$actual_stdout" 2> "$actual_stderr" || actual_exit=$?
    fi

    # Compare results
    local failed=0

    # Check stdout
    if [[ -f "$test_dir/expected.stdout" ]]; then
        if ! diff -q "$test_dir/expected.stdout" "$actual_stdout" > /dev/null 2>&1; then
            echo -e "${RED}FAIL${NC} $test_name: stdout mismatch"
            echo "  Expected:"
            head -5 "$test_dir/expected.stdout" | sed 's/^/    /'
            echo "  Actual:"
            head -5 "$actual_stdout" | sed 's/^/    /'
            failed=1
        fi
    fi

    # Check stderr
    if [[ -f "$test_dir/expected.stderr" ]]; then
        if ! diff -q "$test_dir/expected.stderr" "$actual_stderr" > /dev/null 2>&1; then
            echo -e "${RED}FAIL${NC} $test_name: stderr mismatch"
            echo "  Expected:"
            cat "$test_dir/expected.stderr" | sed 's/^/    /'
            echo "  Actual:"
            cat "$actual_stderr" | sed 's/^/    /'
            failed=1
        fi
    fi

    # Check exit code
    if [[ -f "$test_dir/expected.exit" ]]; then
        local expected_exit=$(cat "$test_dir/expected.exit" | tr -d '\n')
        if [[ "$actual_exit" != "$expected_exit" ]]; then
            echo -e "${RED}FAIL${NC} $test_name: exit code mismatch (expected $expected_exit, got $actual_exit)"
            failed=1
        fi
    fi

    # Cleanup
    rm -f "$actual_stdout" "$actual_stderr"

    if [[ $failed -eq 0 ]]; then
        echo -e "${GREEN}PASS${NC} $test_name"
        ((PASSED++))
    else
        ((FAILED++))
    fi
}

# Find and run all tests
echo "Running CLI golden tests..."
echo ""

for test_dir in $(find "$TESTS_DIR" -type f -name "args.txt" | xargs -I {} dirname {} | sort -u); do
    run_test "$test_dir"
done

echo ""
echo "Results: $PASSED passed, $FAILED failed"

if [[ $FAILED -gt 0 ]]; then
    exit 1
fi
