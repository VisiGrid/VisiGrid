# Canonical Transaction Schema (v1)

The interchange format between adapters, CI, and VisiGrid/VisiHub.
Every adapter emits this CSV format. `vgrid fill` consumes it.

## Columns

| Column           | Type    | Required | Description                                       |
|------------------|---------|----------|---------------------------------------------------|
| `effective_date` | string  | **yes**  | `YYYY-MM-DD` UTC — when the transaction took effect |
| `posted_date`    | string  | **yes**  | `YYYY-MM-DD` UTC — when it was posted/available    |
| `amount_minor`   | integer | **yes**  | Integer minor units (cents). `125050` for $1,250.50 |
| `currency`       | string  | **yes**  | ISO 4217 uppercase (`USD`, `EUR`, etc.)            |
| `type`           | string  | **yes**  | `charge`, `payout`, `fee`, `refund`, `adjustment`  |
| `source`         | string  | **yes**  | Adapter name (`stripe`, `mercury`, `qbo`, etc.)    |
| `source_id`      | string  | **yes**  | External transaction ID (unique within source)     |
| `group_id`       | string  | **yes**  | Grouping key (e.g., payout ID). Empty string if N/A |
| `description`    | string  | **yes**  | Human-readable description                         |
| `amount`         | string  | optional | Decimal string, 2 decimal places (e.g., `"1250.50"`) |

## Rules

- **`amount_minor` is the source of truth.** Stripe gives integer cents.
  The compute engine derives decimal display. This eliminates float
  formatting disputes.
- `amount` is optional in v1. If present, it must satisfy
  `amount_minor == round(amount * 100)` for USD. The CLI validates
  this if both columns exist but does not require `amount`.
- All amounts are signed (negative for fees/refunds).
- **No currency symbols, commas, or whitespace in `amount`** — these
  are adapter bugs.
- `effective_date` and `posted_date` are both UTC.
- **Date attribution rule:** Use `posted_date` for payouts (when funds
  land), `effective_date` for activity (charges, fees, refunds).
- `source_id` must be unique within a source.
- Schema is append-only: new columns may be added, existing columns
  never renamed or removed.

## Decimal Handling Guard

The canonical schema forbids currency symbols. Adapters must normalize
before output. The CSV parser in `vgrid fill` enforces strict rules:

- Parse `amount_minor` as integer only — reject non-integer values
- If `amount` column present: parse as decimal string only (`-?\d+\.\d{2}`)
- Reject `$1,250.50`, `1250.5`, `1,250.50` — these are adapter bugs
- Fail with a clear error on format violation, never silently coerce

## CSV Format Rules

- UTF-8 encoding (no BOM)
- Unix newlines (`\n`)
- Trailing newline after last row
- Standard RFC 4180 quoting (double-quote fields containing commas,
  newlines, or quotes)
- No formula injection: values must not start with `=`
