# v5.7.5 — Continuous price path in Alpha Factor Space

- Alpha price overlay now uses the full daily 3-year adjusted close path instead of quarter-end samples.
- Quarterly factor observations stay at their actual quarter-end dates.
- 3D view: X = elapsed date, Y = factor/price plane, Z = normalized score; price is rendered as a dense continuous daily line.
- 2D view: daily normalized price path overlays quarterly factor points on a true date axis.
- Hover keeps raw price values while the shared visual axis remains normalized Z-score.
- Static cache/version bumped to v5.7.5.
