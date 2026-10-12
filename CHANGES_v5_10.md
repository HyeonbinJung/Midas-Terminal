# v5.10 — Shared Virtual Portfolio (F7)

- Added **F7 PORTFOLIO** and moved SYSTEM to **F8**.
- One global MIDAS Capital virtual portfolio is shared by every visitor. There are no per-user books or browser-local balances.
- Initial cash is **$1,000,000**.
- BUY/SELL requires the server-side trade password. Default: `3824` (override with `PORTFOLIO_TRADE_PASSWORD`).
- Orders are accepted only while the **NYSE regular session is open**. `exchange-calendars` handles weekends, exchange holidays, and early closes; the server enforces the rule even if the UI is bypassed.
- Every fill obtains a fresh **Yahoo Finance / yfinance** quote on the server. Yahoo/yfinance quotes are not exchange-certified real-time data and can be delayed or throttled.
- Short selling is disabled: SELL quantity cannot exceed the shared position.
- Added a persistent SQLite ledger at `data/midas_portfolio.sqlite3` by default (override with `PORTFOLIO_DB_PATH`).
- SQLite stores the singleton account, open positions/average cost, BUY/SELL transaction history, realized P/L, and portfolio-equity snapshots.
- Added open-position valuation, weights, unrealized/realized P/L, quote timestamps, trade notes, and a Plotly equity/cash history chart.
- Portfolio quote refresh runs every 30 seconds while F7 is open. Equity snapshots are rate-limited to once per minute during regular market hours; every trade forces a snapshot.
- Added SQLite WAL + `BEGIN IMMEDIATE` transaction handling so simultaneous users cannot spend the same cash or oversell the same shares.
