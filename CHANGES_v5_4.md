# v5.4 — Interactive 3D Quant Cockpit

- Rebuilt F4 Screener as three independently scrollable terminal windows.
- Added draggable splitters between the left filter, center visualization, and right results windows.
- Replaced faux-3D canvas with Plotly WebGL `scatter3d`: drag-to-rotate, wheel zoom, pan, reset, and modebar controls.
- Added five stacked 3-year weekly adjusted-close charts under the factor graph.
- Added compact security cards plus the raw factor matrix in the right window.
- Screener API now attaches compact 3-year price history only to the five displayed matches.
