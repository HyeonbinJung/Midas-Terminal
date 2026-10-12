# MIDAS Terminal Web v5.1 hotfix

## Fixed: F2 Key Financials not appearing
- Static HTML/CSS/JS now use cache-busting version strings.
- `/`, `/static/*`, and API responses use no-cache/no-store headers so replacing the local project while keeping the same ngrok URL cannot leave an old v4/v5 frontend in the browser.
- F2 Key Financials now reports the number of populated cells in the panel title and explicitly renders an empty-state table if Yahoo has no usable values.
- Research API fetches use `cache: no-store` in the browser.

## Fixed: Korean export failing with `untranslated English research field`
- Removed the fatal all-or-nothing translation check.
- `googletrans.Translator` remains the primary translation path.
- Long SEC paragraphs are split into smaller chunks and retried at finer sentence-sized granularity.
- If googletrans returns the English source unchanged, a secondary Google Translate endpoint is attempted.
- If Google translation is temporarily unavailable, report generation is best-effort instead of returning HTTP 400/504. Export is not blocked by one failed field.
- Korean translation cache key bumped so a previous failed/partial translation is not reused.

- Default Korean translation timeout increased to 300 seconds for long SEC-heavy reports.
