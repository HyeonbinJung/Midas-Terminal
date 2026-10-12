# Midas Terminal v5.8.2 — Advanced Financial Commentary

## Automated Stock Pitch commentary upgrades
- Core diagnostics are now guaranteed below each relevant table instead of being dropped when the comment limit is reached.
  - Income statement: revenue growth, profitability/EBITDA margin, and cash-flow quality.
  - Balance sheet: liquidity, leverage, ROE, and ROIC.
  - Valuation: EV/EBITDA, P/E, and valuation/multiple trend.
- Default commentary limit increased from 8 to 10 lines (`FINANCIAL_COMMENT_MAX`, max 14).
- Added severity labels: `POSITIVE`, `NEUTRAL`, `WEAK`, `WARNING`, `SEVERE`, and `CRITICAL`.
- Added a compact section diagnosis summarizing signal severity counts.

## New multi-year and cross-metric diagnostics
- Three-period sustained high growth / structural contraction.
- Multi-year growth deceleration and multi-year EBITDA-margin erosion.
- Profitability collapse, persistent cash burn, earnings-vs-cash-flow divergence, and growth-with-low-capital-returns warnings.
- Critical balance-sheet stress, multi-year cash depletion/debt build-up, persistent receivable/inventory build-up, working-capital stress, liquidity deterioration, capital-structure deterioration, and high goodwill/intangible exposure.
- Explicit ROE and ROIC profile commentary across good, neutral, weak, and poor ranges.

## Forward and valuation intelligence
- Uses already-filled consensus FY+1/FY+2 revenue and EPS estimates to flag forward contraction, growth normalization/acceleration, and earnings leverage/lag.
- Uses auto-selected peers to calculate relative EV/EBITDA and P/E discounts/premiums when comparable peer data are available.
- Adds growth-adjusted relative valuation signals, premium-at-risk warnings, value-trap risk, and expensive-with-deteriorating-fundamentals warnings.
- Retains N/M safeguards for non-positive EBITDA or net income.
- Financial-sector suppression remains in place for industrial-company leverage/EV-EBITDA rules that are not appropriate for banks/insurers.
