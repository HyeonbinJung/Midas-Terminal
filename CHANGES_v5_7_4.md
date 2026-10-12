# v5.7.4 — Alpha price overlay + 2D/3D switch

- Added the stock price directly to the Alpha normalized factor space.
- Price is sampled at each factor quarter end from the 3-year adjusted price history, then z-normalized across available quarters so it is comparable with factor z-scores.
- 3D mode now plots X=quarter, Y=factor/price, Z=normalized z-score. Price is rendered as a thicker white/lime trace.
- Added 2D / 3D view buttons above the normalized Alpha chart.
- 2D mode overlays every selected factor and normalized price on one quarter-by-quarter z-score chart.
- Kept zoom/pan/reset interactions in both modes and orbit rotation in 3D mode.
- Bumped static asset cache version to 5.7.4.
