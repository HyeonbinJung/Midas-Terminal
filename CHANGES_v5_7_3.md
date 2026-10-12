# v5.7.3 Alpha WebGL visibility fix

- Fixed Alpha 3D canvas being covered by an opaque `.main-svg` background.
- Alpha SVG overlay is now transparent; Plotly paper/scene backgrounds remain graphite.
- Alpha `scatter3d` receives finite-only XYZ arrays, matching the F4 Screener rendering pattern.
- Alpha 3D now uses the direct F4-style `purge -> newPlot -> resize -> relayout` lifecycle.
