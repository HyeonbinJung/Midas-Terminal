# v5.6 — Screener card hardening + Alpha Factor Modeling

- Fixed MATCHED SECURITIES card overflow: 13 score cells now render in a bounded 3-column mini matrix instead of concatenating into one unbroken line.
- Added F5 ALPHA MODEL; SYSTEM moved to F6.
- Added `/api/alpha-factors/{ticker}` powered by yfinance / Yahoo Finance.
- Added selectable quarterly factor families: Momentum, Valuation/Multiples, Growth, Quality/Profitability, Leverage/Balance Sheet, Risk.
- Added 23 metrics including 3M/6M/12M momentum, P/E, EV/EBITDA, P/S, P/B, FCF Yield, growth factors, ROE/margins, Debt/Equity, Net Debt/EBITDA, Current Ratio, rolling beta, realized volatility and max drawdown.
- Added normalized quarterly Z-score overlay, raw factor mini-charts, 3Y price history, latest factor/Z-score/forward-quarter Spearman diagnostics, and data-coverage status.
- All new chart surfaces retain the black/graphite terminal palette and interactive Plotly zoom/pan behavior.
