#!/usr/bin/env python3
"""Generate the local 30,000-row Review Mode performance fixture."""

import argparse
import csv
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("output", type=Path)
    parser.add_argument("--rows", type=int, default=30_000)
    args = parser.parse_args()
    if args.rows < 30_000:
        parser.error("--rows must be at least 30000 for large_sparse_review.lua")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Transaction", "Vendor", "Amount"])
        writer.writerow(["tx-001", "Amazon.com", "100"])
        writer.writerow(["tx-002", "AMZN", "200"])
        writer.writerow(["tx-001", "Amazon.com", "100"])
        writer.writerow(["", "", ""])
        writer.writerow(["tx-003", "Acme", "50"])
        writer.writerow(["", "Total", "=SUM(C2:C4)"])
        for spreadsheet_row in range(8, args.rows + 1):
            writer.writerow([
                f"tx-{spreadsheet_row - 1:06d}",
                f"Vendor {spreadsheet_row}",
                f"{(spreadsheet_row % 997) + 1}.00",
            ])

    print(f"Wrote {args.rows} rows to {args.output}")


if __name__ == "__main__":
    main()
