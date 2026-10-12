# v5 changes

## F2 Key Financials
- `RUN` on the Research tab now returns and renders Key Financials directly in the browser.
- Added 2.1 Income Statement table: FY2023A-FY2027E.
- Added 2.2 Balance Sheet table: FY2023A-FY2027E; forecast balance-sheet cells remain blank when no reliable public consensus data exists.
- Monetary values are shown in millions, consistent with the DOCX template.
- Negative financial values are red in the web table.
- Revenue Growth at or above `REVENUE_GROWTH_HIGHLIGHT_PCT` is highlighted in the web table.
- Tables are horizontally scrollable inside their own panel on mobile, rather than widening the whole F2 page.

## Korean DOCX translation
- Korean export still uses `from googletrans import Translator`.
- Translation now covers the full automatically populated narrative fields, including:
  - Company Background filing excerpts
  - Highlight arguments and evidence/source text
  - Catalyst names, expected-date prose, and verification text
  - Risk names and the full Mitigation / Monitored Indicator paragraphs, including SEC-derived excerpts
- Long SEC excerpts are split into translation-safe chunks.
- Up to four translation requests run concurrently to reduce export time.
- Failed/unchanged long English fields are retried instead of silently producing a mostly-English Korean report.
- Fixed DOCX labels for Company Background, Key Financials, Catalysts, Risk Factors, source rows, and core analyst-table headers are localized to Korean after the English-keyed template fill completes.
- `High / med / low` placeholders are localized as `높음 / 중간 / 낮음` in Korean exports.
- Translation timeout default increased to 180 seconds for long annual-filing excerpts.

## Existing behavior retained
- Automatic Highlights, Catalysts, Risk Factors, Company Background and Peer selection.
- Negative DOCX numbers render red.
- Revenue Growth >= threshold renders with yellow highlight.
- UUID report names, cache/single-flight logic, report queue, Yahoo concurrency control, SEC pacing, rate limits and automatic generated-file cleanup.
- Recommendation, analyst-authored Thesis and all DCF sections remain outside automatic idea generation.
