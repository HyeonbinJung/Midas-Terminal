# v5.11 — Shared Virtual Trading Desk

- Upgraded F7 Portfolio into a shared virtual trading desk backed by SQLite.
- Added MARKET and LIMIT orders, DAY/GTC time-in-force, order cancellation, open-order and order-history ledgers.
- Added Yahoo/yfinance bid/ask when available; otherwise an explicitly labeled indicative spread model.
- Added volume-aware simulated execution: bid/ask crossing, square-root slippage, recent 1-minute-volume / ADV liquidity cap, partial fills, and IOC market-order remainder cancellation.
- Added reserved cash / reserved sell-share checks so concurrent open orders do not over-allocate the shared book.
- Added a 15-second server-side open-limit-order monitor; fills are capped to once per order per minute bucket.
- Added 1m / 5m / 15m / daily / weekly / monthly candlestick charts with volume, VWAP, SMA20, EMA20, fill markers and a live quote line.
- Quote display refreshes every 5 seconds; chart data refreshes every 10 seconds while F7 is open.
- Expanded F7 KPIs with buying power, cash %, gross exposure, largest position and open-order count.
- Extended fill records with order ID/type, reference mid, bid, ask, slippage bps, participation %, and liquidity source.

Important: Yahoo/yfinance is not a Level-II, exchange-certified, or guaranteed real-time execution feed. Bid/ask fallback, slippage, liquidity, partial fills, and market impact are simulations for research/portfolio training.
