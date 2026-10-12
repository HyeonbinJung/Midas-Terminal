# Midas Terminal v5.9 — Peer Analysis

## F6 Peer Analysis
- Added a dedicated F6 Peer Analysis workstation; System moved to F7.
- Supports one subject company plus up to seven analyst-entered peers.
- Added valuation/fundamental comparison matrix: Market Cap, EV, Net Debt, trailing/forward P/E, EV/Revenue, EV/EBITDA, P/S, P/B, revenue growth, EBITDA/operating/net margin, ROE, ROA, Net Debt/EBITDA, FCF yield and beta.
- Added latest three fiscal-year revenue comparison table and chart.
- Added up to eight quarterly revenue periods in a separate table and chart.
- Added grouped multiples and profitability/return comparison charts.
- Added 3-year weekly relative-price comparison normalized to 100 at each security's first available observation.
- Added `/api/peer-analysis/{ticker}` endpoint with up to seven peers, TTL caching, existing Yahoo concurrency protection, partial-peer warnings and missing-field preservation.

## Compatibility
- Existing F1-F5 behavior preserved.
- Existing System screen preserved as F7.
- Uses the existing yfinance/Yahoo Finance data pipeline and Plotly bundle; no new runtime dependency was added.
