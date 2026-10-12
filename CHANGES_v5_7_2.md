# v5.7.2 Alpha graph render fix

- Rebuilt Alpha charts on the same purge -> Plotly.newPlot -> resize lifecycle used by the working F4 Screener.
- Removed the Alpha-only manual width/height + null relayout path that could leave a blank graphite canvas.
- Added defensive array alignment for quarters, factor series and price history.
- Preserved interactive 3D orbit/pan/wheel zoom and 2D factor/price charts.
