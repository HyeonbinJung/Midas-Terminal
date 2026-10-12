# v4 changes

## Research automation
- Added fact-based Highlights builder and DOCX population (`# / ARGUMENT / EVIDENCE · SOURCE`).
- Added rule-based Risk Factors builder and DOCX population (`RISK / LIKELIHOOD / SEVERITY / TP IMPACT / MITIGATION · MONITORED INDICATOR`).
- Risk generation combines financial trends with SEC annual-filing risk disclosures.
- TP-impact ranges are negative and inherit v3 red-number DOCX styling.
- v3 Revenue Growth yellow highlighting remains active.

## Bilingual DOCX
- Added `googletrans==4.0.2`.
- Added `from googletrans import Translator` based translation service.
- Research page now includes `ENGLISH / 한국어` export selector.
- `POST /api/report/{ticker}?language=en|ko`.
- Korean export translates auto-generated Company Background, Highlights, Catalysts and Risk Factors.
- Translation output is cached for the research TTL.

## Mobile
- Prevented F1 controls from overflowing the viewport.
- Mobile RESEARCH button becomes full width.
- F2 action controls reflow on narrow screens.
- Wide research tables use local horizontal scrolling.
- View switching resets retained horizontal scroll.
- Added aggressive `min-width:0`, wrapping and viewport containment rules.

## QA
- Python source compile checks passed.
- JavaScript syntax check passed.
- Template generation tested with populated Highlights, Catalysts and 6 Risk rows.
- DOCX rendered to PNG and visually inspected across all pages.
- Negative risk TP impacts render red.
- Revenue Growth >= 20% renders with yellow highlight.
