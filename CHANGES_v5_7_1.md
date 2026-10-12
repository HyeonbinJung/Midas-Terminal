# v5.7.1 — Alpha chart renderer hotfix

- Fixed the locally bundled Plotly UMD export (`moduleName` -> `Plotly`) that prevented Alpha Model plots from rendering.
- Added app-side compatibility recovery if a cached/alternate Plotly bundle still exposes `window.moduleName`.
- Alpha plots now render only after the visible pane has measurable dimensions and receive explicit first-frame width/height.
- Re-entering F5 Alpha Model re-renders the current payload instead of resize-only behavior.
- Bumped static asset cache keys to force browsers/ngrok sessions to load the repaired renderer.
