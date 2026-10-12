# Midas Terminal v5.8.1

## Stock Pitch automated financial readout

- Adds deterministic, rule-based commentary directly below the annual income statement, annual balance sheet, and quarterly valuation tables in exported Stock Pitch DOCX reports.
- Uses priority-ranked composite signals first (for example High-Quality Growth, Growth Quality Warning, Deleveraging, Receivable Build-up, Multiple Compression, and Growing Into Valuation), then supporting single-metric signals.
- Supports both English and Korean report exports without an additional model/API call.
- Adds safeguards for non-meaningful P/E and EV/EBITDA values when earnings or EBITDA are non-positive.
- Suppresses standard industrial leverage/EV-EBITDA heuristics for financial companies and inserts a sector-specific interpretation notice instead.
- `FINANCIAL_COMMENT_MAX` controls the maximum number of commentary lines inserted below each table (default: 8).
