# Midas Terminal Web v3 changes

## New
- Automatic Catalyst table population in Research API, web UI and exported DOCX.
- Recent SEC 8-K / 6-K is fetched and used as a catalyst source when available.
- Scheduled earnings catalyst from yfinance calendar when available.
- Revenue-growth, EPS/profitability, street-target and margin/FCF catalyst rules.
- Catalyst table is displayed directly on the F2 Research screen; no Accept/Edit/Delete workflow.

## DOCX visual rules
- Any auto-filled numeric cell beginning with a negative value is colored red.
- Revenue Growth (%) values >= `REVENUE_GROWTH_HIGHLIGHT_PCT` are highlighted yellow.
- Default highlight threshold: 20%.
- Negative catalyst PRICE IMPACT values are also red.

## Web visual rules
- Negative peer metrics are red.
- Peer revenue growth >= 20% is highlighted.
- Negative catalyst price-impact cells are red.

## Preserved from v2
- UUID export names, POST export endpoint, export click lock.
- Per-IP limits, TTL caches, single-flight, Yahoo concurrency cap.
- SEC global pacing, peer lightweight fetch, same-ticker export lock.
- Auto-cleanup of generated DOCX files.
