# v5.5

- Fixed white/blank 3D factor plot on the second screener run by purging stale Plotly/WebGL contexts before re-rendering.
- Switched the five individual 3-year price charts from WebGL `scattergl` to SVG `scatter` to prevent WebGL-context exhaustion.
- Added a combined 3-year normalized price overlay for the five matched stocks (start = 100) with wheel zoom, drag pan, and reset.
- Forced Plotly plot containers/canvases to retain the terminal black background across rerenders.
