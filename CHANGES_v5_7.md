# v5.7 — Expanded Alpha Factor Lab

- Fixed Alpha Model charts not rendering reliably by bundling Plotly locally (`static/plotly.min.js`) instead of relying on the CDN, purging stale plot containers, forcing terminal-black plot backgrounds, and resizing plots after render/view changes.
- Replaced the top normalized overlay with a true interactive Plotly `scatter3d` factor lattice: X = quarter, Y = selected factor, Z = within-series z-score. Mouse drag rotates, Shift+drag pans, wheel zooms, double-click resets.
- Expanded the Alpha Factor catalog to 66 metrics across Momentum, Value / Multiple, Growth, Profitability / Quality, Cash Flow Quality, Leverage / Balance Sheet, Efficiency, Volatility / Risk, Size / Liquidity, Estimate Revision, Earnings Surprise, Sentiment, Shareholder Return and Capital Allocation.
- Added dynamic factor-selector generation from the backend catalog, with grouped scrolling controls and default/all/clear selection modes.
- Added quarterly calculations for RSI, MACD, multiple momentum horizons, P/E, EV multiples, growth, margins, ROE/ROA/ROIC, cash-flow quality, leverage, efficiency, volatility, liquidity, earnings drift, analyst actions, insider activity, payout yields and capital-allocation metrics where Yahoo Finance exposes the required data.
- Current-only Yahoo metrics (for example forward P/E, short interest, bid/ask and some estimate data) are plotted only in the latest quarter instead of fabricating history. Unsupported historical revision series remain intentionally blank.
